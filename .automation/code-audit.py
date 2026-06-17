#!/usr/bin/env python3
"""
code-audit.py — Cross-project GAS best practices audit.
Scans all .gs files for anti-patterns from gas-conventions.md.
Returns structured results for the weekly digest.
"""

import os
import re
import json
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import analysis_common as ac  # shared, string/comment-aware sanitizer

with open(os.path.join(SCRIPT_DIR, 'config.json')) as f:
    CONFIG = json.load(f)

REPO_ROOT = CONFIG['repo_root']

# All code file extensions to audit
CODE_EXTENSIONS = ('.gs', '.html', '.js', '.css')


def audit_file(filepath, filename):
    """Audit a single code file. Returns list of violation dicts.
    Server-side rules (.gs only) check GAS API patterns.
    Client-side rules (.html/.js) check browser-relevant patterns.
    """
    try:
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            code = f.read()
        lines = code.split('\n')
    except Exception:
        return []

    violations = []
    is_server = filename.endswith('.gs')
    is_client = filename.endswith(('.html', '.js'))

    def add(line_num, rule, message):
        violations.append({
            'file': filename,
            'line': line_num,
            'rule': rule,
            'message': message
        })

    for i, line in enumerate(lines, 1):

        # === Server-side rules (.gs only) ===
        if is_server:
            # getSheets()[index] instead of getSheetByName()
            if re.search(r'getSheets\(\)\s*\[', line):
                add(i, 'sheet-by-index', 'getSheets()[index] — use getSheetByName() instead')

            # console.log in trigger/event functions
            # Only flag if the enclosing function is actually wired to a trigger
            # (its name appears as the handler in ScriptApp.newTrigger('name') somewhere in the file),
            # OR it's an event handler (onEdit/onOpen/doGet/doPost).
            if 'console.log' in line and not line.strip().startswith('//'):
                enclosing_fn = None
                for j in range(i - 1, max(0, i - 200), -1):
                    fn_match = re.match(r'\s*function\s+(\w+)\s*\(', lines[j])
                    if fn_match:
                        enclosing_fn = fn_match.group(1)
                        break
                if enclosing_fn:
                    is_event_handler = re.match(r'(?:onEdit|onOpen|onInstall|onChange|doGet|doPost)$', enclosing_fn)
                    is_trigger_handler = bool(
                        re.search(r"ScriptApp\.newTrigger\(\s*['\"]" + re.escape(enclosing_fn) + r"['\"]", code)
                    )
                    if is_event_handler or is_trigger_handler:
                        add(i, 'console-in-trigger', 'console.log inside a trigger/event function — use Logger.log or Logs sheet')

            # Cell-by-cell read/write in a loop — require getRange AND a loop keyword on the same line.
            # Blank out string AND comment content first, so the word "for" inside a label like
            # 'Trainees Ready for Certification:' can't masquerade as a for-loop (real false positive).
            code_part = ac.sanitize(line)
            if re.search(r'\b(?:for|forEach|while)\b.*getRange\(', code_part) or \
               re.search(r'getRange\(.*\b(?:for|forEach|while)\b', code_part):
                add(i, 'cell-loop', 'Possible cell-by-cell read in a loop — batch with getValues()')

            # import/export statements — GAS doesn't support ES6 modules
            if re.match(r'^(?:import|export)\s', line):
                add(i, 'es6-module', 'import/export statement — GAS does not support ES6 modules')

            # getActiveSheet() — fragile in server code
            if re.search(r'getActiveSheet\(\)', line) and 'onEdit' not in code[:500] and 'onOpen' not in code[:500]:
                add(i, 'active-sheet', 'getActiveSheet() — fragile, use getSheetByName() for reliability')

            # Unprotected UrlFetchApp in a loop
            if re.search(r'UrlFetchApp\.fetch', line):
                context_start = max(0, i - 10)
                context = '\n'.join(lines[context_start:i])
                if re.search(r'\b(?:for|while|forEach)\b', context):
                    add(i, 'fetch-in-loop', 'UrlFetchApp inside a loop — risk hitting 100 calls/minute quota')

        # === Client-side rules (.html/.js) ===
        if is_client:
            # google.script.run without failure handler
            # Check a wide window (10 lines) since chained calls are often spread across lines
            if 'google.script.run' in line and 'withFailureHandler' not in line:
                context = '\n'.join(lines[max(0, i - 10):min(len(lines), i + 10)])
                if 'withFailureHandler' not in context:
                    add(i, 'no-failure-handler', 'google.script.run without withFailureHandler — errors will fail silently')

    return violations


def audit_project(project):
    """Audit all .gs files in a project. Returns dict with results."""
    project_path = os.path.join(REPO_ROOT, project['path'])
    if not os.path.isdir(project_path):
        return None

    all_violations = []
    files_scanned = 0

    for f in sorted(os.listdir(project_path)):
        if f.endswith(CODE_EXTENSIONS):
            filepath = os.path.join(project_path, f)
            violations = audit_file(filepath, f)
            all_violations.extend(violations)
            files_scanned += 1

    return {
        'name': project['name'],
        'files_scanned': files_scanned,
        'violations': all_violations,
        'violation_count': len(all_violations),
        'by_rule': {}
    }


def run_audit():
    """Run audit across all projects. Returns list of project results."""
    results = []
    for project in CONFIG['projects']:
        result = audit_project(project)
        if result:
            # Group by rule
            by_rule = {}
            for v in result['violations']:
                by_rule.setdefault(v['rule'], []).append(v)
            result['by_rule'] = by_rule
            results.append(result)

    return results


def print_report(results):
    """Print a human-readable audit report."""
    total_violations = sum(r['violation_count'] for r in results)
    print(f"[code-audit] Scanned {sum(r['files_scanned'] for r in results)} files across {len(results)} projects")
    print(f"[code-audit] Found {total_violations} total violations\n")

    for result in results:
        if result['violation_count'] == 0:
            print(f"  {result['name']}: clean")
        else:
            print(f"  {result['name']}: {result['violation_count']} violations")
            for rule, items in result['by_rule'].items():
                print(f"    - {rule}: {len(items)}")
                for item in items[:3]:
                    print(f"      {item['file']}:{item['line']} — {item['message']}")
                if len(items) > 3:
                    print(f"      ... and {len(items) - 3} more")


if __name__ == '__main__':
    results = run_audit()
    print_report(results)
