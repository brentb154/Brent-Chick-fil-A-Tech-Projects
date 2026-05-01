#!/usr/bin/env python3
"""
todo-scanner.py — Find all TODO, FIXME, HACK, WORKAROUND comments across projects.
Sorts by age (oldest first) using git blame when available.
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

PATTERNS = [
    (r'\bTODO\b', 'TODO'),
    (r'\bFIXME\b', 'FIXME'),
    (r'\bHACK\b', 'HACK'),
    (r'\bWORKAROUND\b', 'WORKAROUND'),
    (r'\bXXX\b', 'XXX'),
    (r'\bBUG\b', 'BUG'),
]


def get_blame_date(filepath, line_num):
    """Get the date a line was last modified via git blame."""
    try:
        result = subprocess.run(
            ['git', 'blame', '-L', f'{line_num},{line_num}', '--porcelain', filepath],
            capture_output=True, text=True, cwd=REPO_ROOT, timeout=5
        )
        if result.returncode == 0:
            for blame_line in result.stdout.split('\n'):
                if blame_line.startswith('committer-time '):
                    timestamp = int(blame_line.split(' ')[1])
                    return datetime.fromtimestamp(timestamp)
    except Exception:
        pass
    return None


def scan_file(filepath, filename, project_name):
    """Scan a single file for TODO-like comments."""
    try:
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
    except Exception:
        return []

    findings = []
    for i, line in enumerate(lines, 1):
        stripped = line.strip()
        # Only match in comments
        if not (stripped.startswith('//') or stripped.startswith('*') or stripped.startswith('/*')):
            continue

        for pattern, tag in PATTERNS:
            if re.search(pattern, line, re.IGNORECASE):
                # Extract the comment text
                comment = stripped.lstrip('/*').lstrip('/').lstrip('*').strip()
                blame_date = get_blame_date(filepath, i)

                findings.append({
                    'project': project_name,
                    'file': filename,
                    'line': i,
                    'tag': tag,
                    'comment': comment[:120],
                    'date': blame_date,
                    'age_days': (datetime.now() - blame_date).days if blame_date else None
                })
                break  # Only match first pattern per line

    return findings


def run_scan():
    """Scan all projects for TODOs."""
    all_findings = []

    for project in CONFIG['projects']:
        project_path = os.path.join(REPO_ROOT, project['path'])
        if not os.path.isdir(project_path):
            continue

        for f in sorted(os.listdir(project_path)):
            if f.endswith(('.gs', '.html', '.py', '.js')):
                filepath = os.path.join(project_path, f)
                findings = scan_file(filepath, f, project['name'])
                all_findings.extend(findings)

    # Sort: oldest first (unknown dates last)
    all_findings.sort(key=lambda x: (x['age_days'] is None, -(x['age_days'] or 0)))

    return all_findings


def print_report(findings):
    """Print a human-readable TODO report."""
    print(f"[todo-scanner] Found {len(findings)} items across all projects\n")

    by_project = {}
    for f in findings:
        by_project.setdefault(f['project'], []).append(f)

    for project, items in sorted(by_project.items()):
        print(f"  {project}: {len(items)} items")
        for item in items:
            age = f"{item['age_days']}d ago" if item['age_days'] is not None else "unknown age"
            print(f"    [{item['tag']}] {item['file']}:{item['line']} ({age})")
            print(f"      {item['comment']}")

    # Flag old items
    old = [f for f in findings if f['age_days'] is not None and f['age_days'] > 30]
    if old:
        print(f"\n  WARNING: {len(old)} items are older than 30 days")


if __name__ == '__main__':
    findings = run_scan()
    print_report(findings)
