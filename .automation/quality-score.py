#!/usr/bin/env python3
"""
quality-score.py — Real code-quality grading for the GAS operational tools.

This is the grader that actually tries to answer "how good is this code", as
opposed to code-audit.py (cosmetic conventions) and risk-audit.py (regression
metal-detector). It scores SIX axes from function-level + project-level signals,
normalized by size so cosmetic edits can't swing the number, and rolls them into
a weighted letter grade.

  Axes: correctness, reliability, performance, maintainability, security, handoff.

Deterministic by default (pure stdlib, safe for the nightly cron). With --llm it
also asks Claude to read the code and grade it semantically — the one thing regex
can't do — printed as a SEPARATE verdict, never folded into the deterministic
score (so the deterministic number stays reproducible).

  python3 quality-score.py                      # all projects, deterministic
  python3 quality-score.py --project training-tracker
  python3 quality-score.py --llm                # + Claude semantic judge (opt-in)
  python3 quality-score.py --project X --llm --json

The deterministic analyzer is honest about being a heuristic floor; the --llm pass
is where semantic correctness gets judged. Neither replaces a human reading the diff.
"""

import os
import re
import sys
import json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import analysis_common as ac  # noqa: E402

with open(os.path.join(SCRIPT_DIR, 'config.json')) as f:
    CONFIG = json.load(f)
REPO_ROOT = CONFIG['repo_root']
QUALITY_CFG = CONFIG.get('quality', {})

# Axis weights → overall. Correctness and maintainability carry the most because
# they're what actually bite an operator who inherits the tool.
AXIS_WEIGHTS = {
    'correctness':     0.25,
    'reliability':     0.15,
    'performance':     0.15,
    'maintainability': 0.20,
    'security':        0.15,
    'handoff':         0.10,
}

DUP_MIN_LINES = 12
DUP_SIMILARITY = 0.75
DUP_SHINGLE = 5

# --- Security patterns (shared shape with risk-audit, broadened) -------------
SECRET_KEY_RE = re.compile(
    r'(passcode|password|passwd|secret|api[_-]?key|apikey|access[_-]?code|client[_-]?secret|auth[_-]?token)\s*[:=]',
    re.IGNORECASE)
SECRET_VAL_RE = re.compile(r'[:=]\s*(?:[\'"][^\'"]{4,}[\'"]|\d{4,})')
URL_SECRET_RE = re.compile(r'[?&](?:token|key|secret|passcode|auth[_-]?token)=[\w\-]{8,}', re.IGNORECASE)
HTML_TAG_STRING_RE = re.compile(r"""['"`]\s*<\s*\w""")
ESCAPE_HINT_RE = re.compile(r'escapeHtml|encodeURI|DOMPurify|sanitize|attrSafe|\bsafe[A-Z_]\w*')
# Broadened from risk-audit's hardcoded field list: ANY identifier concatenated into
# an HTML-tag string literal in client code, minus obviously-safe numeric/date/count
# variables. Catches the real stored-XSS class (user-typed names) the old rule missed.
HTML_VAR_CONCAT_RE = re.compile(r"""\+\s*[A-Za-z_$][\w.$\[\]']*""")
SAFE_VAR_RE = re.compile(r'(count|index|idx|total|len|length|num|rowindex|^i$|^n$|date|time|width|height|size)', re.IGNORECASE)
# Template-literal interpolation into HTML. Must be checked on the RAW line — the
# sanitizer blanks template contents, which is exactly how `${user.name}` XSS evaded
# the concat rule. Field-name targeted (user-content fields), because flagging every
# ${var} would bury the signal under system-generated ids/amounts/dates.
TPL_TAG_RE = re.compile(r'<\s*\w')
TPL_EXPR_RE = re.compile(r'\$\{([^}]*)\}')
RISKY_FIELD_RE = re.compile(r'(name|notes|comment|message|reason|desc|email|title)', re.IGNORECASE)


def grade_letter(score):
    if score >= 90: return 'A'
    if score >= 80: return 'B'
    if score >= 70: return 'C'
    if score >= 60: return 'D'
    return 'F'


def clamp(x):
    return max(0, min(100, int(round(x))))


def _duplication_pairs(funcs):
    cand = [f for f in funcs if f['length'] >= DUP_MIN_LINES and len(f['tokens']) >= DUP_SHINGLE]
    shingles = []
    for f in cand:
        t = f['tokens']
        shingles.append(set(tuple(t[k:k + DUP_SHINGLE]) for k in range(len(t) - DUP_SHINGLE + 1)))
    pairs = 0
    for a in range(len(cand)):
        for b in range(a + 1, len(cand)):
            if min(cand[a]['length'], cand[b]['length']) / max(cand[a]['length'], cand[b]['length']) < 0.5:
                continue
            sa, sb = shingles[a], shingles[b]
            if not sa or not sb:
                continue
            jac = len(sa & sb) / len(sa | sb)
            if jac >= DUP_SIMILARITY:
                pairs += 1
    return pairs


def analyze_project(project):
    path = os.path.join(REPO_ROOT, project['path'])
    if not os.path.isdir(path):
        return None
    bundles = ac.read_code_files(path, os)
    if not bundles:
        return None

    all_funcs = []
    gs_names = {}
    all_code = ''
    gs_code = ''   # server-side only — trigger/reliability checks must not see client JS
    findings = {k: [] for k in AXIS_WEIGHTS}

    # File-level passes
    cell_loops = 0          # getRange().get/setValue inside a loop  (perf, real N+1)
    full_reads_in_loop = 0  # getDataRange/getValues inside a loop
    flush_in_loop = 0
    unguarded_sheet = 0     # getSheetByName(...). chained with no obvious guard
    empty_catch = 0
    sec_hits = []

    for fname, orig, san in bundles:
        all_code += '\n'.join(orig) + '\n'
        if fname.endswith('.gs'):
            gs_code += '\n'.join(orig) + '\n'
        is_client = fname.endswith(('.html', '.js'))
        funcs = ac.extract_functions(san)
        for fn in funcs:
            fn['file'] = fname
            all_funcs.append(fn)
            if fn['file'].endswith('.gs'):
                gs_names.setdefault(fn['name'], []).append(fn)
        mask = ac.loop_mask(san)

        for idx, sline in enumerate(san):
            raw = orig[idx] if idx < len(orig) else ''
            in_loop = mask[idx] if idx < len(mask) else False

            # \b after "Value" so the batched setValues()/getValues() (correct) aren't
            # mis-counted as cell-by-cell — "setValue" is a prefix of "setValues".
            if in_loop and re.search(r'getRange\s*\([^)]*\)\s*\.\s*(?:get|set)Value\b', sline):
                cell_loops += 1
            if in_loop and re.search(r'getDataRange\s*\(\s*\)|getRange\s*\([^)]*\)\s*\.\s*getValues', sline):
                # A handle created fresh inside the loop (iterating over DIFFERENT
                # sheets, one read each) is the batch pattern, not a redundant
                # re-read. Walk back through the loop body for the declaration.
                m2 = re.search(r'(\w+)\s*\.\s*(?:getDataRange|getRange)', sline)
                recv = m2.group(1) if m2 else ''
                fresh = False
                if recv:
                    decl = re.compile(
                        r'(?:(?:const|let|var)\s+' + recv + r'\s*=' +
                        r'|for\s*\(\s*(?:const|let|var)\s+' + recv + r'\s+of' +
                        r'|forEach\s*\(\s*(?:function\s*\(\s*)?\(?\s*' + recv + r'\s*[,)=]' +
                        r'|\.map\s*\(\s*\(?\s*' + recv + r'\s*[,)=])')
                    for back in range(idx - 1, max(-1, idx - 40), -1):
                        if decl.search(san[back]):
                            fresh = True
                            break
                        if back < len(mask) and not mask[back]:
                            break  # left the loop body without finding it
                if not fresh:
                    full_reads_in_loop += 1
            if in_loop and 'flush' in sline:
                flush_in_loop += 1
            if re.search(r'getSheetByName\s*\([^)]*\)\s*\.\s*(?:getRange|getLastRow|getDataRange|getValues)', sline):
                unguarded_sheet += 1

            # Security (key in real code, value from original line)
            if SECRET_KEY_RE.search(sline) and SECRET_VAL_RE.search(raw) and not \
               re.search(r'type\s*=\s*[\'"]password|getElementById|placeholder', raw, re.IGNORECASE):
                sec_hits.append(('high', fname, idx + 1, 'hardcoded secret/passcode in source'))
            if URL_SECRET_RE.search(raw):
                sec_hits.append(('high', fname, idx + 1, 'secret/token embedded in a URL'))
            # Concat rule checks the RAW line — the sanitizer blanks string
            # contents (including the '<tag' part), which had silently killed
            # this rule. Same field-name targeting as the template rule.
            if is_client and HTML_TAG_STRING_RE.search(raw) and '${' not in raw \
               and HTML_VAR_CONCAT_RE.search(raw) and not ESCAPE_HINT_RE.search(raw):
                for concat in HTML_VAR_CONCAT_RE.findall(raw):
                    if RISKY_FIELD_RE.search(concat) and not SAFE_VAR_RE.search(concat):
                        sec_hits.append(('medium', fname, idx + 1, 'variable concatenated into HTML without escaping (XSS)'))
                        break
            if is_client and '${' in raw and TPL_TAG_RE.search(raw) and not ESCAPE_HINT_RE.search(raw):
                for expr in TPL_EXPR_RE.findall(raw):
                    if RISKY_FIELD_RE.search(expr) and not SAFE_VAR_RE.search(expr):
                        sec_hits.append(('medium', fname, idx + 1,
                                         'template literal interpolates user-content field into HTML without escaping (XSS)'))
                        break

        # empty catch blocks (swallowed errors) — counted on ORIGINAL text, so a
        # catch documented with a comment (intentional ignore, e.g. releaseLock
        # guards) is not treated as silent swallowing.
        empty_catch += len(re.findall(r'catch\s*\([^)]*\)\s*\{\s*\}', '\n'.join(orig)))

    # Duplicate gs function names (last-wins footgun)
    dup_fn = [n for n, defs in gs_names.items() if len(defs) > 1]

    # Reliability: trigger handlers with/without try/catch + alerting.
    # .gs code only — client HTML defines onChange-style handlers (onCategoryChange
    # etc.) that the on[A-Z] heuristic would otherwise miscount as GAS triggers.
    trigger_handlers = set(re.findall(r"ScriptApp\.newTrigger\(\s*['\"](\w+)['\"]", gs_code))
    trigger_handlers |= set(re.findall(r'function\s+(on[A-Z]\w*)\s*\(', gs_code))
    unguarded_triggers = 0
    for h in trigger_handlers:
        m = re.search(r'function\s+' + re.escape(h) + r'\s*\([^)]*\)\s*\{', gs_code)
        if not m:
            continue
        body = gs_code[m.end():m.end() + 1500]
        if 'try' not in body:
            unguarded_triggers += 1
    has_alerting = bool(re.search(r'catch[\s\S]{0,200}?(?:MailApp|GmailApp)\.sendEmail|catch[\s\S]{0,200}?sendAlert', gs_code))

    # Maintainability metrics
    func_count = len(all_funcs) or 1
    avg_complexity = sum(f['complexity'] for f in all_funcs) / func_count
    long_funcs = sum(1 for f in all_funcs if f['length'] > 80)
    max_nesting = max((f['max_depth'] for f in all_funcs), default=0)
    dup_pairs = _duplication_pairs(all_funcs)
    code_lines = sum(1 for ln in all_code.split('\n') if ln.strip())
    comment_lines = len(re.findall(r'^\s*(?://|/\*|\*)', all_code, re.MULTILINE))
    comment_density = comment_lines / max(code_lines, 1)

    # Handoff signals
    files = set(os.listdir(path))
    has_claude = 'CLAUDE.md' in files
    has_readme = any(x.lower() in ('readme.md', 'setup.md', 'setup_guide.md', 'complete_setup_guide.md') for x in files)
    has_init = bool(re.search(r'function\s+(?:initializeSheet|runInitialSetup|initSheets|setupAccountability|manualInit)', all_code))
    has_config_sheet = bool(re.search(r"getSheetByName\(['\"][\w ]*(?:Settings|Config|config)['\"]", all_code))
    has_trigger_cleanup = bool(re.search(r'function\s+(?:delete|remove|cleanup)\w*[Tt]rigger', all_code))
    has_dynamic_detect = bool(re.search(r'getValues\(\)[\s\S]{0,400}?indexOf|headers?\.indexOf|detect\w*Column', all_code))

    # ---- Axis scoring -------------------------------------------------------
    scores = {}

    c = 100
    c -= min(len(dup_fn) * 15, 45)
    c -= min(unguarded_sheet * 2, 20)
    c -= min(empty_catch * 6, 18)
    scores['correctness'] = clamp(c)
    if dup_fn:
        findings['correctness'].append('duplicate .gs function name(s): ' + ', '.join(dup_fn))
    if unguarded_sheet:
        findings['correctness'].append('%d getSheetByName() chained with no missing-sheet guard' % unguarded_sheet)
    if empty_catch:
        findings['correctness'].append('%d empty catch block(s) swallow errors' % empty_catch)

    r = 100
    r -= min(unguarded_triggers * 12, 48)
    if not has_alerting and trigger_handlers:
        r -= 12
    scores['reliability'] = clamp(r)
    if unguarded_triggers:
        findings['reliability'].append('%d trigger handler(s) without try/catch' % unguarded_triggers)
    if not has_alerting and trigger_handlers:
        findings['reliability'].append('no failure-alert (catch → MailApp/sendAlert) on any trigger')

    p = 100
    p -= min(cell_loops * 8, 50)
    p -= min(full_reads_in_loop * 6, 24)
    p -= min(flush_in_loop * 5, 15)
    scores['performance'] = clamp(p)
    if cell_loops:
        findings['performance'].append('%d cell-by-cell getRange().get/setValue inside loops (batch instead)' % cell_loops)
    if full_reads_in_loop:
        findings['performance'].append('%d full-range read(s) inside loops' % full_reads_in_loop)
    if flush_in_loop:
        findings['performance'].append('%d SpreadsheetApp.flush() inside loops' % flush_in_loop)

    m = 100
    if avg_complexity > 8:
        m -= min((avg_complexity - 8) * 3, 24)
    m -= min(long_funcs * 4, 28)
    if max_nesting > 5:
        m -= min((max_nesting - 5) * 4, 16)
    m -= min(dup_pairs * 5, 20)
    if comment_density < 0.03:
        m -= 6
    scores['maintainability'] = clamp(m)
    findings['maintainability'].append(
        'avg complexity %.1f, %d fn>80 lines, max nesting %d, %d near-dup pair(s), comment density %.0f%%'
        % (avg_complexity, long_funcs, max_nesting, dup_pairs, comment_density * 100))

    s = 100
    for sev, _f, _l, _msg in sec_hits:
        s -= {'high': 25, 'medium': 10, 'low': 4}.get(sev, 4)
    scores['security'] = clamp(s)
    for sev, f_, l_, msg in sec_hits[:6]:
        findings['security'].append('%s: %s:%d — %s' % (sev, f_, l_, msg))

    h = 100
    for ok, pts, label in [
        (has_claude, 18, 'CLAUDE.md'), (has_readme, 16, 'README/setup guide'),
        (has_init, 18, 'idempotent init fn'), (has_config_sheet, 18, 'config-in-sheet'),
        (has_trigger_cleanup, 15, 'trigger-cleanup util'), (has_dynamic_detect, 15, 'dynamic column detection')]:
        if not ok:
            h -= pts
            findings['handoff'].append('missing: ' + label)
    scores['handoff'] = clamp(h)

    overall = sum(scores[a] * w for a, w in AXIS_WEIGHTS.items())
    return {
        'name': project['name'],
        'scores': scores,
        'overall': clamp(overall),
        'grade': grade_letter(overall),
        'findings': findings,
        'metrics': {
            'functions': func_count, 'lines': code_lines,
            'avg_complexity': round(avg_complexity, 1), 'max_nesting': max_nesting,
            'cell_loops': cell_loops, 'dup_fn': dup_fn,
        },
        '_code': all_code, '_path': path,
    }


# ---------------------------------------------------------------------------
# LLM semantic judge (opt-in). Uses the official anthropic SDK per the
# claude-api guidance; degrades gracefully if the SDK / key is absent.
# ---------------------------------------------------------------------------

JUDGE_SCHEMA = {
    'type': 'object',
    'properties': {
        'scores': {
            'type': 'object',
            'properties': {a: {'type': 'integer'} for a in AXIS_WEIGHTS},
            'required': list(AXIS_WEIGHTS), 'additionalProperties': False,
        },
        'overall': {'type': 'integer'},
        'grade': {'type': 'string'},
        'top_issues': {
            'type': 'array',
            'items': {
                'type': 'object',
                'properties': {
                    'severity': {'type': 'string', 'enum': ['critical', 'high', 'medium', 'low']},
                    'file': {'type': 'string'},
                    'summary': {'type': 'string'},
                },
                'required': ['severity', 'file', 'summary'], 'additionalProperties': False,
            },
        },
        'verdict': {'type': 'string'},
    },
    'required': ['scores', 'overall', 'grade', 'top_issues', 'verdict'],
    'additionalProperties': False,
}

JUDGE_SYSTEM = (
    "You are a senior reviewer grading Google Apps Script tools that run a pair of "
    "Chick-fil-A restaurants. These tools are handed off to NON-TECHNICAL operators, so "
    "judge them on: correctness of the actual logic (sheet I/O, data flow, edge cases a "
    "regex linter cannot see), reliability (trigger error handling, idempotent setup, "
    "alert-on-failure), performance (batch sheet reads/writes vs cell-by-cell), "
    "maintainability (clarity for a self-taught maintainer, not abstraction for its own "
    "sake), security (secrets in source, XSS in HTML Service UIs), and handoff readiness "
    "(operator-editable config in sheets, docs, safe-to-re-run setup). Score each axis "
    "0-100 and give an overall 0-100 with a letter grade (A>=90,B>=80,C>=70,D>=60,F). "
    "Be specific and blunt. Reward simple, self-contained GAS code; do not penalize the "
    "absence of tests, TypeScript, or CI — those are deliberately out of scope. Return "
    "ONLY the structured object."
)


def _import_anthropic():
    """Import the anthropic SDK, falling back to the bundled .venv if the running
    interpreter doesn't have it (this machine has several Pythons; the SDK lives in
    .automation/.venv). Returns the module or None."""
    try:
        import anthropic
        return anthropic
    except ImportError:
        pass
    import glob
    for sp in glob.glob(os.path.join(SCRIPT_DIR, '.venv', 'lib', 'python*', 'site-packages')):
        if sp not in sys.path:
            sys.path.insert(0, sp)
    try:
        import anthropic
        return anthropic
    except ImportError:
        return None


def _load_api_key():
    """Resolve the API key: env var first, then a gitignored .automation/.anthropic_key
    file (so the key persists for both interactive and nightly runs without exporting)."""
    if os.environ.get('ANTHROPIC_API_KEY') or os.environ.get('ANTHROPIC_AUTH_TOKEN'):
        return True
    keyfile = os.path.join(SCRIPT_DIR, '.anthropic_key')
    if os.path.exists(keyfile):
        try:
            with open(keyfile) as fh:
                key = fh.read().strip()
            if key:
                os.environ['ANTHROPIC_API_KEY'] = key
                return True
        except Exception:
            pass
    return False


def run_llm_judge(result, model):
    anthropic = _import_anthropic()
    if anthropic is None:
        return {'_error': 'anthropic SDK not found — run: .automation/.venv/bin/pip install anthropic'}
    if not _load_api_key():
        return {'_error': 'no API key — set ANTHROPIC_API_KEY or put it in .automation/.anthropic_key'}

    code = result['_code']
    budget = 120_000  # ~30K tokens of code; bound cost on big projects
    truncated = len(code) > budget
    if truncated:
        code = code[:budget]

    det = {a: result['scores'][a] for a in AXIS_WEIGHTS}
    user = (
        "Project: %s\n"
        "Deterministic analyzer's axis scores (for reference, not authoritative): %s\n"
        "%s\n\n"
        "Source code below%s:\n\n%s"
    ) % (result['name'], json.dumps(det),
         'Metrics: ' + json.dumps(result['metrics'].get('dup_fn') and result['metrics'] or result['metrics']),
         ' (TRUNCATED — judge what you can see)' if truncated else '', code)

    try:
        client = anthropic.Anthropic()
        with client.messages.stream(
            model=model,
            max_tokens=8000,
            thinking={'type': 'adaptive'},
            output_config={'effort': 'high', 'format': {'type': 'json_schema', 'schema': JUDGE_SCHEMA}},
            system=JUDGE_SYSTEM,
            messages=[{'role': 'user', 'content': user}],
        ) as stream:
            msg = stream.get_final_message()
    except Exception as e:
        return {'_error': '%s: %s' % (type(e).__name__, e)}

    if msg.stop_reason == 'refusal':
        return {'_error': 'model declined to grade (refusal)'}
    text = next((b.text for b in msg.content if b.type == 'text'), '')
    try:
        return json.loads(text)
    except Exception:
        return {'_error': 'could not parse judge output', '_raw': text[:400]}


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def _bar(score, width=20):
    filled = int(round(score / 100 * width))
    return '█' * filled + '·' * (width - filled)


def print_report(result, llm=None):
    print("  %s — %s/100  (grade %s)" % (result['name'], result['overall'], result['grade']))
    for axis in AXIS_WEIGHTS:
        sc = result['scores'][axis]
        print("      %-15s %s %3d  %s" % (axis, _bar(sc), sc,
              result['findings'][axis][0] if result['findings'][axis] else ''))
        for extra in result['findings'][axis][1:]:
            print("      %-15s %s      %s" % ('', ' ' * 20, extra))
    if llm is not None:
        print()
        if llm.get('_error'):
            print("      LLM judge: skipped — %s" % llm['_error'])
        else:
            ls = llm.get('scores', {})
            print("      LLM verdict: %s/100 (grade %s) vs deterministic %s/100"
                  % (llm.get('overall', '?'), llm.get('grade', '?'), result['overall']))
            axis_str = '  '.join('%s %s' % (a[:4], ls.get(a, '?')) for a in AXIS_WEIGHTS)
            print("      LLM axes:    %s" % axis_str)
            if llm.get('verdict'):
                print("      \"%s\"" % llm['verdict'].strip())
            for iss in llm.get('top_issues', [])[:5]:
                print("        (%s) %s — %s" % (iss.get('severity'), iss.get('file'), iss.get('summary')))
    print()


if __name__ == '__main__':
    args = sys.argv[1:]
    want_llm = '--llm' in args
    want_json = '--json' in args
    project_filter = None
    if '--project' in args:
        pi = args.index('--project')
        project_filter = args[pi + 1] if pi + 1 < len(args) else None

    model = QUALITY_CFG.get('llm_model', 'claude-opus-4-8')
    out = []
    if not want_json:
        print("[quality-score] Multi-axis code grade (deterministic%s)\n"
              % (' + Claude semantic judge' if want_llm else ''))

    for project in CONFIG['projects']:
        if project_filter and project['name'] != project_filter:
            continue
        result = analyze_project(project)
        if not result:
            continue
        llm = run_llm_judge(result, model) if want_llm else None
        if want_json:
            rec = {k: v for k, v in result.items() if not k.startswith('_')}
            if llm is not None:
                rec['llm'] = llm
            out.append(rec)
        else:
            print_report(result, llm)

    if want_json:
        print(json.dumps(out, indent=2))