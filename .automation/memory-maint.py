#!/usr/bin/env python3
"""
memory-maint.py — Memory maintenance.
Checks memory files for references to files/functions that no longer exist.
Reports stale entries so they can be cleaned up. Does NOT auto-delete.
"""

import os
import re
import json
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')

with open(CONFIG_PATH) as f:
    CONFIG = json.load(f)

REPO_ROOT = CONFIG['repo_root']


def scan_all_code():
    """Build a set of all function names and file names across all projects."""
    functions = set()
    files = set()

    for project in CONFIG['projects']:
        project_path = os.path.join(REPO_ROOT, project['path'])
        if not os.path.isdir(project_path):
            continue

        for f in os.listdir(project_path):
            full = os.path.join(project_path, f)
            if not os.path.isfile(full):
                continue
            files.add(f)
            if f.endswith('.gs'):
                try:
                    with open(full, 'r', encoding='utf-8', errors='replace') as fh:
                        code = fh.read()
                    for m in re.finditer(r'function\s+(\w+)\s*\(', code):
                        functions.add(m.group(1))
                except Exception:
                    pass

    return functions, files


def check_memory_dir(memory_dir, all_functions, all_files):
    """Check memory files for stale references."""
    issues = []

    if not os.path.isdir(memory_dir):
        return issues

    for fname in os.listdir(memory_dir):
        if fname == 'MEMORY.md' or not fname.endswith('.md'):
            continue

        fpath = os.path.join(memory_dir, fname)
        try:
            with open(fpath, 'r') as f:
                content = f.read()
        except Exception:
            continue

        # Check for function references that no longer exist
        # Look for backtick-wrapped function names like `functionName()`
        for m in re.finditer(r'`(\w+)\(\)`', content):
            func_name = m.group(1)
            if func_name not in all_functions and len(func_name) > 3:
                issues.append({
                    'file': fname,
                    'type': 'stale_function',
                    'reference': func_name,
                    'message': f'Function `{func_name}()` referenced in {fname} not found in any project'
                })

        # Check for file references like `SomeFile.gs`
        for m in re.finditer(r'`(\w+\.(?:gs|html))`', content):
            file_ref = m.group(1)
            if file_ref not in all_files:
                issues.append({
                    'file': fname,
                    'type': 'stale_file',
                    'reference': file_ref,
                    'message': f'File `{file_ref}` referenced in {fname} not found in any project'
                })

    return issues


def check_memory_index(memory_dir):
    """Check MEMORY.md index for broken links."""
    issues = []
    index_path = os.path.join(memory_dir, 'MEMORY.md')
    if not os.path.exists(index_path):
        return issues

    with open(index_path, 'r') as f:
        content = f.read()

    # Find markdown links like [Title](filename.md)
    for m in re.finditer(r'\[([^\]]+)\]\(([^)]+)\)', content):
        title = m.group(1)
        link = m.group(2)
        link_path = os.path.join(memory_dir, link)
        if not os.path.exists(link_path):
            issues.append({
                'file': 'MEMORY.md',
                'type': 'broken_link',
                'reference': link,
                'message': f'MEMORY.md links to `{link}` but file does not exist'
            })

    return issues


def main():
    all_functions, all_files = scan_all_code()
    print(f"[memory-maint] Scanned {len(all_functions)} functions, {len(all_files)} files across projects")

    all_issues = []

    for memory_dir in CONFIG['memory_dirs']:
        if not os.path.isdir(memory_dir):
            print(f"[memory-maint] SKIP: {memory_dir} does not exist")
            continue

        dir_name = os.path.basename(os.path.dirname(memory_dir))
        issues = check_memory_dir(memory_dir, all_functions, all_files)
        issues += check_memory_index(memory_dir)

        if issues:
            print(f"\n[memory-maint] {len(issues)} issue(s) in {dir_name}/memory/:")
            for issue in issues:
                print(f"  - {issue['message']}")
            all_issues.extend(issues)
        else:
            print(f"[memory-maint] {dir_name}/memory/ — all clean")

    if not all_issues:
        print("\n[memory-maint] All memory files are current.")
    else:
        print(f"\n[memory-maint] {len(all_issues)} total issue(s) found. Review and update manually.")

    return len(all_issues)


if __name__ == '__main__':
    exit(main())
