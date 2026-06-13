#!/usr/bin/env python3
"""
risk-audit.py — Risk-focused static analysis (security, correctness-smell, duplication, complexity).

This is the deliberate counterpart to code-audit.py:
  - code-audit.py  = HYGIENE (GAS conventions, handoff readiness) — cosmetic, gameable, fine.
  - risk-audit.py  = RISK (things that can actually lose money / leak data / break) — a triaged
                     findings list, NOT a single gameable score.

Design principles (learned the hard way):
  * Deterministic Python is a tireless *metal detector*: it sweeps every file nightly and catches
    KNOWN patterns + regressions. It does NOT understand intent — it cannot tell you whether two
    functions are *equivalent*, only that they *look* duplicated. Semantic correctness still needs
    a human/LLM pass. This tool's job is to AIM that pass, not replace it.
  * Findings are fixed or explicitly ACKNOWLEDGED (with a reason). An acknowledgment is fingerprinted
    to the flagged code, so it stays quiet until that code changes — then it re-surfaces. This avoids
    nag-fatigue without letting a real regression hide.
  * Any unacknowledged critical/high finding flags the project "AT RISK" regardless of hygiene score.

CLI:
  python3 risk-audit.py                      # report (suppresses acknowledged findings)
  python3 risk-audit.py --list               # list ALL findings WITH fingerprints (incl. acknowledged)
  python3 risk-audit.py --project payroll-system
  python3 risk-audit.py --ack <fingerprint> --reason "why this is acceptable"
"""

import os
import re
import sys
import json
import hashlib
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SCRIPT_DIR, 'config.json')) as f:
    CONFIG = json.load(f)

REPO_ROOT = CONFIG['repo_root']
ACK_FILE = os.path.join(SCRIPT_DIR, 'risk-acks.json')

CODE_EXTENSIONS = ('.gs', '.html', '.js', '.css')

# Severity ordering / which severities flag a project "AT RISK"
SEVERITY_ORDER = {'critical': 0, 'high': 1, 'medium': 2, 'low': 3}
AT_RISK_SEVERITIES = ('critical', 'high')

# Tunables
DUP_MIN_LINES = 12          # only compare functions at least this long
DUP_SIMILARITY = 0.75       # Jaccard threshold to call two functions near-duplicates
DUP_SHINGLE = 5             # token n-gram size for similarity
LONG_FUNCTION_LINES = 80
DEEP_NESTING = 5
LARGE_FILE_LINES = 3000
MAX_FUNCTION_SCAN = 800     # safety cap so a brace-balance miss can't run to EOF


# ---------------------------------------------------------------------------
# Source sanitizing + function extraction
# ---------------------------------------------------------------------------

def sanitize(text):
    """Blank out string/template/comment content (preserving newlines and structural braces)
    so brace-matching and token analysis aren't fooled by punctuation inside strings/comments.
    Returns sanitized text aligned line-for-line with the original."""
    out = []
    i, n = 0, len(text)
    state = 'code'   # code | line_comment | block_comment | string
    delim = ''
    while i < n:
        c = text[i]
        nxt = text[i + 1] if i + 1 < n else ''
        if state == 'code':
            if c == '/' and nxt == '/':
                state = 'line_comment'; out.append('  '); i += 2; continue
            if c == '/' and nxt == '*':
                state = 'block_comment'; out.append('  '); i += 2; continue
            if c in ('"', "'", '`'):
                state = 'string'; delim = c; out.append(' '); i += 1; continue
            out.append(c); i += 1; continue
        if state == 'line_comment':
            if c == '\n':
                state = 'code'; out.append('\n')
            else:
                out.append(' ')
            i += 1; continue
        if state == 'block_comment':
            if c == '*' and nxt == '/':
                state = 'code'; out.append('  '); i += 2; continue
            out.append('\n' if c == '\n' else ' '); i += 1; continue
        if state == 'string':
            if c == '\\':
                out.append('  '); i += 2; continue
            if c == delim:
                state = 'code'; out.append(' '); i += 1; continue
            out.append('\n' if c == '\n' else ' '); i += 1; continue
    return ''.join(out)


FUNC_DECL_RE = re.compile(r'\bfunction\s+(\w+)\s*\(')
FUNC_EXPR_RE = re.compile(r'(?:var|let|const)\s+(\w+)\s*=\s*(?:async\s+)?function\s*\(')


def extract_functions(orig_lines, san_lines):
    """Return [{name, start, end, depth, length, body_norm}] using sanitized lines for brace
    matching. start/end are 1-based inclusive line numbers into the original file."""
    functions = []
    n = len(san_lines)
    i = 0
    while i < n:
        sline = san_lines[i]
        m = FUNC_DECL_RE.search(sline) or FUNC_EXPR_RE.search(sline)
        if not m:
            i += 1
            continue
        name = m.group(1)
        # Find the opening brace at/after this line
        depth = 0
        started = False
        max_depth = 0
        end = None
        for j in range(i, min(n, i + MAX_FUNCTION_SCAN)):
            for ch in san_lines[j]:
                if ch == '{':
                    depth += 1; started = True
                    max_depth = max(max_depth, depth)
                elif ch == '}':
                    depth -= 1
                    if started and depth == 0:
                        end = j
                        break
            if end is not None:
                break
        if end is None:
            i += 1
            continue
        start = i
        body_san = ' '.join(san_lines[start:end + 1])
        tokens = [t for t in re.split(r'[^A-Za-z0-9_]+', body_san.lower()) if t]
        functions.append({
            'name': name,
            'start': start + 1,
            'end': end + 1,
            'length': end - start + 1,
            'max_depth': max_depth,
            'tokens': tokens,
        })
        i = end + 1
    return functions


# ---------------------------------------------------------------------------
# Acknowledgments
# ---------------------------------------------------------------------------

def load_acks():
    if not os.path.exists(ACK_FILE):
        return {}
    try:
        with open(ACK_FILE) as fh:
            data = json.load(fh)
        return {a['fingerprint']: a for a in data.get('acks', [])}
    except Exception:
        return {}


def save_ack(fingerprint, reason):
    data = {'acks': []}
    if os.path.exists(ACK_FILE):
        try:
            with open(ACK_FILE) as fh:
                data = json.load(fh)
        except Exception:
            data = {'acks': []}
    data.setdefault('acks', [])
    data['acks'] = [a for a in data['acks'] if a.get('fingerprint') != fingerprint]
    data['acks'].append({
        'fingerprint': fingerprint,
        'reason': reason,
        'added': datetime.now().isoformat(timespec='seconds'),
    })
    with open(ACK_FILE, 'w') as fh:
        json.dump(data, fh, indent=2)


def fingerprint(rule, project, where, snippet):
    norm = re.sub(r'\s+', ' ', snippet).strip().lower()
    raw = f'{rule}|{project}|{where}|{norm}'
    return hashlib.sha1(raw.encode('utf-8')).hexdigest()[:12]


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------

# Secret detection is split so we can require the KEYWORD to be real code (checked against the
# sanitized line, where string/comment contents are blanked) while reading the literal VALUE from
# the original line. This avoids false positives like  showToast('Error updating passcode: ' + e)
# where "passcode" merely appears inside a display string.
# No word boundaries: real constants embed these as substrings (DEFAULT_ADMIN_ACCESS_PASSCODE,
# MANAGER_PASSCODE, apiKey). Matching against the SANITIZED line already prevents matching a
# keyword that only appears inside a display string.
SECRET_KEY_RE = re.compile(
    r'(passcode|password|passwd|secret|api[_-]?key|apikey|access[_-]?code|client[_-]?secret|auth[_-]?token)\s*[:=]',
    re.IGNORECASE)
SECRET_VAL_RE = re.compile(r'[:=]\s*(?:[\'"][^\'"]{4,}[\'"]|\d{4,})')
INNERHTML_RE = re.compile(r'(?:innerHTML|outerHTML|insertAdjacentHTML)\b')
ESCAPE_HINT_RE = re.compile(r'escapeHtml|escapeHtmlHealth|encodeURI|DOMPurify|sanitize', re.IGNORECASE)
TZ_RE = re.compile(r'toISOString\(\)\s*\.\s*(?:split\(\s*[\'"]T[\'"]\s*\)|slice\(\s*0\s*,\s*10\s*\)|substr(?:ing)?\(\s*0\s*,\s*10\s*\))')

# --- Hardened checks (added 2026-06-13 to cover blind spots the original rules missed) ---
# Concatenation XSS: a variable concatenated into an HTML-tag string literal, unescaped. The
# ${}-template rule above missed the employee-name stored XSS, which used + concatenation.
HTML_TAG_STRING_RE = re.compile(r"""['"`]\s*<\s*\w""")
# Concatenation of a USER-DATA-ish field (name/notes/comment/...) into HTML — targets the real
# XSS class without flagging every safe HTML concatenation (dates, counts, class names).
USER_DATA_CONCAT_RE = re.compile(r"""\+\s*[\w.$\[\]']*(?:employeeName|fullName|displayName|firstName|lastName|\.notes?\b|\.reason\b|\.comments?\b|commentText|\.description\b|\.message\b|guestName|\.email\b)[\w.$\[\]']*""", re.IGNORECASE)
# Secret/token embedded in a URL query string (also used when scanning .json/.sh configs).
URL_SECRET_RE = re.compile(r'[?&](?:token|key|secret|passcode|auth[_-]?token)=[\w\-]{8,}', re.IGNORECASE)
# UTC date extraction beyond toISOString (getUTC* / toDateString) — same day-shift hazard.
TZ_UTC_EXTRACT_RE = re.compile(r'\.getUTC(?:FullYear|Month|Date)\(\)|new Date\([^)]*\)\s*\.\s*toDateString\(\)')


def scan_lines(project, relpath, orig_lines, san_lines):
    """Line-based security + correctness-smell checks. Returns list of findings.
    orig_lines = source as-written; san_lines = strings/comments blanked (see sanitize())."""
    findings = []
    is_client = relpath.endswith(('.html', '.js'))

    for idx, raw in enumerate(orig_lines, 1):
        line = raw.rstrip('\n')
        san = san_lines[idx - 1] if idx - 1 < len(san_lines) else ''
        stripped = line.strip()
        if stripped.startswith('//') or stripped.startswith('*'):
            continue

        # --- Hardcoded secret / passcode ---
        # Keyword must be real code (sanitized line), value must be a literal (original line).
        if SECRET_KEY_RE.search(san) and SECRET_VAL_RE.search(line):
            # Skip obvious non-secrets: HTML password inputs, element lookups, placeholders
            if not re.search(r'type\s*=\s*[\'"]password|getElementById|placeholder|autocomplete', line, re.IGNORECASE):
                findings.append(_mk(project, relpath, idx, 'security', 'high',
                                    'hardcoded-secret',
                                    'Hardcoded secret/passcode in source — move to Script Properties or a Settings sheet',
                                    line))

        # --- XSS: unescaped interpolation into innerHTML ---
        if is_client and INNERHTML_RE.search(line) and '${' in line and not ESCAPE_HINT_RE.search(line):
            findings.append(_mk(project, relpath, idx, 'security', 'medium',
                                'xss-innerhtml',
                                'Template interpolation into innerHTML without an escape helper — XSS risk if the value can contain HTML',
                                line))

        # --- Clickjacking: ALLOWALL frame embedding ---
        if 'setXFrameOptionsMode' in line and 'ALLOWALL' in line:
            findings.append(_mk(project, relpath, idx, 'security', 'low',
                                'xframe-allowall',
                                'X-Frame-Options ALLOWALL permits embedding anywhere (clickjacking) — use SAMEORIGIN unless embedding is required',
                                line))

        # --- Timezone: deriving a date string from UTC components ---
        if TZ_RE.search(line):
            findings.append(_mk(project, relpath, idx, 'correctness', 'medium',
                                'tz-utc-date',
                                'Date derived from toISOString() (UTC) — can shift a day in negative-offset timezones; format with local components',
                                line))

        # --- XSS: a variable concatenated into an HTML string without an escape helper. This is
        # the blind spot that missed the employee-name stored XSS (it used + concatenation, not ${}). ---
        if (is_client and '${' not in line and HTML_TAG_STRING_RE.search(line)
                and USER_DATA_CONCAT_RE.search(line) and not ESCAPE_HINT_RE.search(line)):
            findings.append(_mk(project, relpath, idx, 'security', 'medium',
                                'xss-concat',
                                'Variable concatenated into an HTML string without an escape helper — stored XSS if the value can contain HTML (e.g. a user-typed name)',
                                line))

        # --- Secret/token embedded in a URL (caught in code, .json and .sh alike) ---
        if URL_SECRET_RE.search(line):
            findings.append(_mk(project, relpath, idx, 'security', 'high',
                                'url-secret',
                                'Secret/token embedded in a URL — anyone with this file can use it; move it out of source/config',
                                line))

        # --- Timezone: UTC component extraction (getUTC* / toDateString) ---
        if TZ_UTC_EXTRACT_RE.search(line):
            findings.append(_mk(project, relpath, idx, 'correctness', 'low',
                                'tz-utc-extract',
                                'Date built from UTC components (getUTC*/toDateString) — can shift a day in negative-offset timezones; verify against local components',
                                line))

    return findings


def scan_functions(project, files):
    """Function-level checks: duplication, length, nesting. `files` = [(relpath, orig_lines, san_lines)]."""
    findings = []
    all_funcs = []

    for relpath, orig_lines, san_lines in files:
        funcs = extract_functions(orig_lines, san_lines)
        for fn in funcs:
            fn['file'] = relpath
            all_funcs.append(fn)

            if fn['length'] > LONG_FUNCTION_LINES:
                findings.append(_mk(project, relpath, fn['start'], 'complexity', 'low',
                                    'long-function',
                                    f"Function {fn['name']}() is {fn['length']} lines — consider splitting for maintainability",
                                    f"func:{fn['name']}"))
            if fn['max_depth'] > DEEP_NESTING:
                findings.append(_mk(project, relpath, fn['start'], 'complexity', 'low',
                                    'deep-nesting',
                                    f"Function {fn['name']}() nests {fn['max_depth']} levels deep — hard to follow/verify",
                                    f"func:{fn['name']}"))

    # --- Exact duplicate function NAMES within the same runtime. All .gs files share ONE global
    #     namespace in GAS (server), so a name defined 2+ times across .gs means the last wins and
    #     the rest are dead/load-order-dependent (the cancelUniformOrder / formatDate class). For
    #     client .html we only flag a name defined twice in the SAME file (a real same-page
    #     redefinition); cross-file client dups are normally per-page and a client wrapper sharing a
    #     name with its server function is intentional, so neither is flagged. ---
    gs_by_name = {}
    html_by_file_name = {}
    for fn in all_funcs:
        if fn['file'].endswith('.gs'):
            gs_by_name.setdefault(fn['name'], []).append(fn)
        else:
            html_by_file_name.setdefault((fn['file'], fn['name']), []).append(fn)
    for name, defs in sorted(gs_by_name.items()):
        if len(defs) < 2:
            continue
        locs = ', '.join('%s:%d' % (d['file'], d['start']) for d in defs)
        findings.append(_mk(project, defs[0]['file'], defs[0]['start'], 'correctness', 'high',
                            'duplicate-function',
                            '%s() defined %d times across .gs files (%s) — GAS runs only the last; the rest are dead and load-order dependent' % (name, len(defs), locs),
                            'dupfn:%s' % name))
    for (fname, name), defs in sorted(html_by_file_name.items()):
        if len(defs) < 2:
            continue
        lines_str = ', '.join('%d' % d['start'] for d in defs)
        findings.append(_mk(project, fname, defs[0]['start'], 'correctness', 'medium',
                            'duplicate-function',
                            '%s() defined %d times in the same file (lines %s) — the earlier copies are dead code' % (name, len(defs), lines_str),
                            'dupfn:%s:%s' % (fname, name)))

    # --- Duplication: near-identical function bodies (the '3 implementations' smell) ---
    candidates = [fn for fn in all_funcs if fn['length'] >= DUP_MIN_LINES and len(fn['tokens']) >= DUP_SHINGLE]
    shingles = []
    for fn in candidates:
        toks = fn['tokens']
        sh = set(tuple(toks[k:k + DUP_SHINGLE]) for k in range(len(toks) - DUP_SHINGLE + 1))
        shingles.append(sh)

    reported_pairs = set()
    for a in range(len(candidates)):
        for b in range(a + 1, len(candidates)):
            fa, fb = candidates[a], candidates[b]
            # Only compare similarly-sized functions (cheap pre-filter)
            if min(fa['length'], fb['length']) / max(fa['length'], fb['length']) < 0.5:
                continue
            sa, sb = shingles[a], shingles[b]
            if not sa or not sb:
                continue
            inter = len(sa & sb)
            if inter == 0:
                continue
            jac = inter / len(sa | sb)
            if jac >= DUP_SIMILARITY:
                key = tuple(sorted([f"{fa['file']}:{fa['name']}", f"{fb['file']}:{fb['name']}"]))
                if key in reported_pairs:
                    continue
                reported_pairs.add(key)
                pct = int(jac * 100)
                findings.append(_mk(project, fa['file'], fa['start'], 'duplication', 'medium',
                                    'duplicate-logic',
                                    f"{fa['name']}() ~{pct}% similar to {fb['name']}() ({fb['file']}:{fb['start']}) — likely duplicated logic; extract a shared helper",
                                    f"dup:{key[0]}|{key[1]}"))
    return findings


def _mk(project, relpath, line, category, severity, rule, message, snippet):
    return {
        'project': project,
        'file': relpath,
        'line': line,
        'category': category,
        'severity': severity,
        'rule': rule,
        'message': message,
        'fingerprint': fingerprint(rule, project, f'{relpath}', snippet),
    }


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def audit_project(project):
    project_path = os.path.join(REPO_ROOT, project['path'])
    if not os.path.isdir(project_path):
        return None

    findings = []
    file_bundles = []
    line_total = 0

    for f in sorted(os.listdir(project_path)):
        if not f.endswith(CODE_EXTENSIONS):
            continue
        try:
            with open(os.path.join(project_path, f), 'r', encoding='utf-8', errors='replace') as fh:
                text = fh.read()
        except Exception:
            continue
        orig_lines = text.split('\n')
        san_lines = sanitize(text).split('\n')
        file_bundles.append((f, orig_lines, san_lines))

        findings.extend(scan_lines(project['name'], f, orig_lines, san_lines))

        non_blank = sum(1 for ln in orig_lines if ln.strip())
        line_total += non_blank
        if non_blank > LARGE_FILE_LINES:
            findings.append(_mk(project['name'], f, 1, 'complexity', 'low',
                                'large-file',
                                f"{f} is {non_blank} lines — large single file; consider splitting by responsibility",
                                f"file:{f}"))

    # Secret-only scan of config/shell files: not code-analyzed, but they leak tokens/keys.
    for f in sorted(os.listdir(project_path)):
        if not f.endswith(('.json', '.sh')):
            continue
        try:
            with open(os.path.join(project_path, f), 'r', encoding='utf-8', errors='replace') as fh:
                text = fh.read()
        except Exception:
            continue
        findings.extend(scan_lines(project['name'], f, text.split('\n'), sanitize(text).split('\n')))

    findings.extend(scan_functions(project['name'], file_bundles))
    return {'name': project['name'], 'findings': findings}


def run_risk_audit():
    """Run risk analysis across all projects. Returns list of project results with ack state applied."""
    acks = load_acks()
    results = []
    for project in CONFIG['projects']:
        res = audit_project(project)
        if not res:
            continue
        for fnd in res['findings']:
            ack = acks.get(fnd['fingerprint'])
            fnd['acknowledged'] = bool(ack)
            fnd['ack_reason'] = ack['reason'] if ack else None
        res['findings'].sort(key=lambda x: (SEVERITY_ORDER.get(x['severity'], 9), x['file'], x['line']))

        active = [f for f in res['findings'] if not f['acknowledged']]
        res['counts'] = _counts(active)
        res['acknowledged_count'] = sum(1 for f in res['findings'] if f['acknowledged'])
        res['at_risk'] = any(f['severity'] in AT_RISK_SEVERITIES for f in active)
        results.append(res)
    return results


def _counts(findings):
    c = {'critical': 0, 'high': 0, 'medium': 0, 'low': 0}
    for f in findings:
        c[f['severity']] = c.get(f['severity'], 0) + 1
    return c


def print_report(results, show_all=False, project_filter=None):
    print("[risk-audit] Risk findings (security / correctness / duplication / complexity)\n")
    for res in results:
        if project_filter and res['name'] != project_filter:
            continue
        active = res['findings'] if show_all else [f for f in res['findings'] if not f['acknowledged']]
        c = res['counts']
        flag = 'AT RISK' if res['at_risk'] else 'ok'
        print(f"  [{flag}] {res['name']}: "
              f"{c['critical']} critical, {c['high']} high, {c['medium']} medium, {c['low']} low"
              f"  (+{res['acknowledged_count']} acknowledged)")
        for f in active:
            tag = ' [ACK]' if f.get('acknowledged') else ''
            fp = f"  {{{f['fingerprint']}}}" if show_all else ''
            print(f"      ({f['severity']}) {f['file']}:{f['line']} [{f['rule']}]{tag} — {f['message']}{fp}")
            if show_all and f.get('ack_reason'):
                print(f"           ack reason: {f['ack_reason']}")
        print()


if __name__ == '__main__':
    args = sys.argv[1:]
    if '--ack' in args:
        idx = args.index('--ack')
        fp = args[idx + 1] if idx + 1 < len(args) else None
        reason = None
        if '--reason' in args:
            ridx = args.index('--reason')
            reason = args[ridx + 1] if ridx + 1 < len(args) else None
        if not fp or not reason:
            print('Usage: risk-audit.py --ack <fingerprint> --reason "why this is acceptable"')
            sys.exit(1)
        save_ack(fp, reason)
        print(f'Acknowledged {fp}: {reason}')
        sys.exit(0)

    show_all = '--list' in args
    project_filter = None
    if '--project' in args:
        pidx = args.index('--project')
        project_filter = args[pidx + 1] if pidx + 1 < len(args) else None

    results = run_risk_audit()
    print_report(results, show_all=show_all, project_filter=project_filter)

    # Non-zero exit if any project is AT RISK (unacknowledged critical/high) so the
    # nightly orchestrator can flag it in its summary.
    if any(r['at_risk'] for r in results):
        sys.exit(1)
