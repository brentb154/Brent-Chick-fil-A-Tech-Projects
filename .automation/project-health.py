#!/usr/bin/env python3
"""
project-health.py — Project health scores and operator readiness grades.
Combines: audit violations, TODO count, staleness, function count,
and handoff-readiness indicators into a per-project scorecard.
"""

import os
import re
import json
import subprocess
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SCRIPT_DIR, 'config.json')) as f:
    CONFIG = json.load(f)

REPO_ROOT = CONFIG['repo_root']

# All code file extensions to scan (GAS projects use .html for JS/CSS includes)
CODE_EXTENSIONS = ('.gs', '.html', '.js', '.css')

# Import audit and todo functions
import importlib.util
audit_spec = importlib.util.spec_from_file_location('code_audit', os.path.join(SCRIPT_DIR, 'code-audit.py'))
code_audit = importlib.util.module_from_spec(audit_spec)
audit_spec.loader.exec_module(code_audit)

todo_spec = importlib.util.spec_from_file_location('todo_scanner', os.path.join(SCRIPT_DIR, 'todo-scanner.py'))
todo_scanner = importlib.util.module_from_spec(todo_spec)
todo_spec.loader.exec_module(todo_scanner)

risk_spec = importlib.util.spec_from_file_location('risk_audit', os.path.join(SCRIPT_DIR, 'risk-audit.py'))
risk_audit = importlib.util.module_from_spec(risk_spec)
risk_spec.loader.exec_module(risk_audit)


def get_last_commit_date(project_path):
    """Get date of last git commit touching this project folder."""
    try:
        result = subprocess.run(
            ['git', 'log', '-1', '--format=%ct', '--', project_path],
            capture_output=True, text=True, cwd=REPO_ROOT, timeout=5
        )
        if result.returncode == 0 and result.stdout.strip():
            return datetime.fromtimestamp(int(result.stdout.strip()))
    except Exception:
        pass
    return None


def get_last_modified(project_path):
    """Fallback: most recent file modification time."""
    latest = 0
    for f in os.listdir(project_path):
        full = os.path.join(project_path, f)
        if os.path.isfile(full) and not f.startswith('.'):
            mtime = os.path.getmtime(full)
            if mtime > latest:
                latest = mtime
    return datetime.fromtimestamp(latest) if latest > 0 else None


def count_functions(project_path):
    """Count total functions across all code files."""
    count = 0
    for f in os.listdir(project_path):
        if f.endswith(CODE_EXTENSIONS):
            try:
                with open(os.path.join(project_path, f), 'r', encoding='utf-8', errors='replace') as fh:
                    code = fh.read()
                count += len(re.findall(r'^function\s+\w+\s*\(', code, re.MULTILINE))
                # Also catch var/const arrow functions and object methods in JS
                count += len(re.findall(r'(?:var|let|const)\s+\w+\s*=\s*function\s*\(', code, re.MULTILINE))
            except Exception:
                pass
    return count


def count_lines(project_path):
    """Count total lines of code across all code files."""
    total = 0
    for f in os.listdir(project_path):
        if f.endswith(CODE_EXTENSIONS):
            try:
                with open(os.path.join(project_path, f), 'r', encoding='utf-8', errors='replace') as fh:
                    total += sum(1 for line in fh if line.strip())
            except Exception:
                pass
    return total


def check_operator_readiness(project_path, all_code):
    """Check handoff-readiness indicators. Returns dict of checks."""
    checks = {}

    # Has CLAUDE.md?
    checks['claude_md'] = os.path.exists(os.path.join(project_path, 'CLAUDE.md'))

    # Has README or setup guide?
    has_readme = False
    for f in os.listdir(project_path):
        if f.lower() in ('readme.md', 'setup.md', 'setup_guide.md', 'complete_setup_guide.md'):
            has_readme = True
            break
    checks['setup_guide'] = has_readme

    # Has idempotent init function?
    checks['idempotent_init'] = bool(re.search(
        r'function\s+(?:initializeSheet|runInitialSetup|initSheets|setupAccountability|manualInit)',
        all_code
    ))

    # Has config in sheet (not hardcoded)?
    checks['config_in_sheet'] = bool(re.search(
        r"getSheetByName\(['\"](?:Settings|Config|config)['\"]",
        all_code
    ))

    # Has error alerting?
    checks['error_alerting'] = bool(re.search(
        r'(?:notifyTriggerFailure|catch.*(?:MailApp|GmailApp)\.sendEmail|catch.*sendAlert)',
        all_code, re.DOTALL
    ))

    # Has trigger cleanup utility?
    checks['trigger_cleanup'] = bool(re.search(
        r'function\s+(?:deleteTrigger|cleanupTrigger|removeTrigger|removeAll|deleteAll|remove\w*Trigger|delete\w*Trigger|cleanup\w*Trigger)',
        all_code
    ))

    return checks


def compute_health_score(violations, todo_count, days_since_change, readiness):
    """Compute a health score. Returns (score 0-100, color)."""
    score = 100

    # Violations: -3 per violation, max -30
    score -= min(violations * 3, 30)

    # TODOs: -2 per TODO, max -15
    score -= min(todo_count * 2, 15)

    # Staleness: -1 per 10 days inactive, max -20
    if days_since_change is not None:
        score -= min(days_since_change // 10, 20)

    # Readiness: -5 per missing check, max -25
    missing = sum(1 for v in readiness.values() if not v)
    score -= missing * 5

    score = max(0, min(100, score))

    if score >= 80:
        color = 'GREEN'
    elif score >= 60:
        color = 'YELLOW'
    else:
        color = 'RED'

    return score, color


def run_health_check():
    """Run health analysis across all projects."""
    # Get audit results
    audit_results = code_audit.run_audit()
    audit_by_project = {r['name']: r for r in audit_results}

    # Get TODO results
    all_todos = todo_scanner.run_scan()
    todos_by_project = {}
    for t in all_todos:
        todos_by_project.setdefault(t['project'], []).append(t)

    # Get RISK results (separate axis — never folded into the hygiene score)
    risk_results = risk_audit.run_risk_audit()
    risk_by_project = {r['name']: r for r in risk_results}

    results = []

    for project in CONFIG['projects']:
        project_path = os.path.join(REPO_ROOT, project['path'])
        if not os.path.isdir(project_path):
            continue

        # Read all code files
        all_code = ''
        for f in sorted(os.listdir(project_path)):
            if f.endswith(CODE_EXTENSIONS):
                try:
                    with open(os.path.join(project_path, f), 'r', encoding='utf-8', errors='replace') as fh:
                        all_code += fh.read() + '\n'
                except Exception:
                    pass

        # Gather metrics
        audit = audit_by_project.get(project['name'], {'violation_count': 0, 'violations': []})
        todos = todos_by_project.get(project['name'], [])
        old_todos = [t for t in todos if t['age_days'] is not None and t['age_days'] > 30]

        last_change = get_last_commit_date(project_path) or get_last_modified(project_path)
        days_since = (datetime.now() - last_change).days if last_change else None

        func_count = count_functions(project_path)
        line_count = count_lines(project_path)
        readiness = check_operator_readiness(project_path, all_code)

        score, color = compute_health_score(
            audit['violation_count'], len(todos), days_since, readiness
        )

        # Risk axis (security / correctness / duplication / complexity) — kept SEPARATE
        # from the hygiene score so cosmetic fixes can't paper over real risk.
        risk = risk_by_project.get(project['name'], {'counts': {'critical': 0, 'high': 0, 'medium': 0, 'low': 0},
                                                      'at_risk': False, 'acknowledged_count': 0, 'findings': []})

        # Staleness flag
        staleness = 'active'
        if days_since is not None:
            if days_since > 90:
                staleness = 'dormant (90+ days)'
            elif days_since > 60:
                staleness = 'stale (60+ days)'
            elif days_since > 30:
                staleness = 'aging (30+ days)'

        results.append({
            'name': project['name'],
            'score': score,
            'color': color,
            'violations': audit['violation_count'],
            'todos': len(todos),
            'old_todos': len(old_todos),
            'days_since_change': days_since,
            'staleness': staleness,
            'functions': func_count,
            'lines': line_count,
            'readiness': readiness,
            'readiness_score': sum(1 for v in readiness.values() if v),
            'readiness_total': len(readiness),
            'last_change': last_change.strftime('%Y-%m-%d') if last_change else 'unknown',
            'risk': risk['counts'],
            'at_risk': risk['at_risk'],
            'risk_acknowledged': risk['acknowledged_count'],
        })

    return results


def print_report(results):
    """Print health report. TWO axes:
      - Hygiene score (0-100): conventions + handoff readiness. Cosmetic; NOT a quality grade.
      - Risk: security/correctness/duplication findings. A project can score 98 hygiene and still
        be AT RISK. Run `risk-audit.py --list` for details and fingerprints to acknowledge."""
    print("[project-health] Project Scorecard")
    print("  Hygiene = conventions + handoff readiness (cosmetic). Risk = security/correctness (what hurts you).\n")

    # Sort: at-risk projects first, then by hygiene score
    for r in sorted(results, key=lambda x: (not x['at_risk'], x['score'])):
        icon = {'GREEN': 'OK', 'YELLOW': 'WARN', 'RED': 'ALERT'}[r['color']]
        risk_flag = '  ** AT RISK **' if r['at_risk'] else ''
        print(f"  [{icon}] {r['name']}: Hygiene {r['score']}/100{risk_flag}")
        rc = r['risk']
        ack = f" (+{r['risk_acknowledged']} ack'd)" if r['risk_acknowledged'] else ''
        print(f"       Risk: {rc['critical']} critical, {rc['high']} high, {rc['medium']} medium, {rc['low']} low{ack}")
        print(f"       Violations: {r['violations']} | TODOs: {r['todos']} ({r['old_todos']} old) | Functions: {r['functions']} | Lines: {r['lines']}")
        print(f"       Last change: {r['last_change']} ({r['staleness']})")
        print(f"       Readiness: {r['readiness_score']}/{r['readiness_total']} checks passing")
        for check, passed in r['readiness'].items():
            status = 'pass' if passed else 'MISSING'
            print(f"         {'  ' if passed else '! '}{check}: {status}")
        print()

    # At-risk summary
    at_risk = [r for r in results if r['at_risk']]
    if at_risk:
        print(f"  ** {len(at_risk)} project(s) AT RISK (unacknowledged critical/high findings):")
        for r in at_risk:
            print(f"    - {r['name']}: {r['risk']['critical']} critical, {r['risk']['high']} high  → run: risk-audit.py --project {r['name']} --list")
        print()

    # Stale projects warning
    stale = [r for r in results if r['staleness'] != 'active']
    if stale:
        print(f"  ATTENTION: {len(stale)} project(s) need attention:")
        for r in stale:
            print(f"    - {r['name']}: {r['staleness']}")


if __name__ == '__main__':
    results = run_health_check()
    print_report(results)
