#!/usr/bin/env python3
"""
analysis_common.py — Shared static-analysis primitives for the GAS tooling.

One correct parser, used everywhere, so the string-literal / comment blind spots
are fixed in ONE place instead of re-introduced per tool. Lifted and consolidated
from risk-audit.py's proven internals.

Pure stdlib. No side effects on import.
"""

import re

CODE_EXTENSIONS = ('.gs', '.html', '.js', '.css')

# Loop openers that matter for GAS batch-operation analysis (cell-by-cell smell).
_LOOP_RE = re.compile(r'\b(?:for|while)\b\s*\(|\.\s*forEach\s*\(')
_DECISION_RE = re.compile(
    r'\b(?:if|for|while|case|catch)\b|\.\s*(?:forEach|map|filter|reduce|some|every)\s*\(|&&|\|\||\?[^.]'
)


def sanitize(text):
    """Blank string/template/comment CONTENT (keep newlines + structural braces) so
    brace-matching and pattern checks aren't fooled by punctuation, the word 'for'
    inside a label, etc. Returned text is aligned line-for-line with the original."""
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
MAX_FUNCTION_SCAN = 800


def extract_functions(san_lines):
    """Return [{name, start, end, length, max_depth, tokens, complexity}] from sanitized
    lines. start/end are 1-based inclusive line numbers."""
    functions = []
    n = len(san_lines)
    i = 0
    while i < n:
        m = FUNC_DECL_RE.search(san_lines[i]) or FUNC_EXPR_RE.search(san_lines[i])
        if not m:
            i += 1
            continue
        name = m.group(1)
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
        body = '\n'.join(san_lines[i:end + 1])
        tokens = [t for t in re.split(r'[^A-Za-z0-9_]+', body.lower()) if t]
        functions.append({
            'name': name,
            'start': i + 1,
            'end': end + 1,
            'length': end - i + 1,
            'max_depth': max_depth,
            'tokens': tokens,
            'complexity': 1 + len(_DECISION_RE.findall(body)),
        })
        i = end + 1
    return functions


def loop_mask(san_lines):
    """Return a list[bool] (one per line) marking whether each sanitized line sits
    inside a for/while/forEach loop body. Approximate but multi-line aware — this is
    what lets us catch cell-by-cell sheet access that spans several lines, which a
    single-line regex cannot."""
    mask = [False] * len(san_lines)
    depth = 0
    loop_frames = []  # brace depths at which an active loop body lives
    for idx, line in enumerate(san_lines):
        # A line is "in a loop" if any loop frame is still open at the current depth.
        mask[idx] = any(depth >= f for f in loop_frames)
        opens_loop = bool(_LOOP_RE.search(line))
        for ch in line:
            if ch == '{':
                depth += 1
                if opens_loop:
                    loop_frames.append(depth)
                    opens_loop = False
            elif ch == '}':
                loop_frames = [f for f in loop_frames if f < depth]
                depth -= 1
        # Single-statement loop with no brace on the same line: treat the next
        # non-blank line as in-loop (best effort).
        if opens_loop and idx + 1 < len(san_lines):
            mask[idx] = True
    return mask


def read_code_files(project_path, os_module):
    """Read all code files in a project dir. Returns [(filename, orig_lines, san_lines)]."""
    bundles = []
    for f in sorted(os_module.listdir(project_path)):
        if not f.endswith(CODE_EXTENSIONS):
            continue
        try:
            with open(os_module.path.join(project_path, f), 'r',
                      encoding='utf-8', errors='replace') as fh:
                text = fh.read()
        except Exception:
            continue
        bundles.append((f, text.split('\n'), sanitize(text).split('\n')))
    return bundles