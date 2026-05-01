#!/usr/bin/env python3
"""
update-claude-docs.py — Regenerate project-level CLAUDE.md files from live code.
Scans each sub-project's .gs and .html files, extracts sheet tabs, functions,
file structure, code style, and key patterns. Writes a fresh CLAUDE.md.

Run nightly or on-demand. Safe to re-run — overwrites CLAUDE.md with current state.
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


def scan_files(project_path):
    """Return sorted lists of .gs and .html files."""
    gs_files = []
    html_files = []
    for f in sorted(os.listdir(project_path)):
        full = os.path.join(project_path, f)
        if not os.path.isfile(full):
            continue
        if f.endswith('.gs'):
            gs_files.append(f)
        elif f.endswith('.html'):
            html_files.append(f)
    return gs_files, html_files


def read_file(path):
    """Read file contents, return empty string on error."""
    try:
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            return f.read()
    except Exception:
        return ''


def detect_code_style(all_code):
    """Detect whether project uses var or const/let style."""
    var_count = len(re.findall(r'^var ', all_code, re.MULTILINE))
    const_count = len(re.findall(r'^const ', all_code, re.MULTILINE))
    let_count = len(re.findall(r'^let ', all_code, re.MULTILINE))
    modern = const_count + let_count
    if var_count > modern * 2:
        return 'var', 'Uses `var` and function() style — match this, not const/let'
    elif modern > var_count * 2:
        return 'modern', 'Uses `const`/`let` style — match this'
    else:
        return 'mixed', 'Mixed `var` and `const`/`let` — check each file and match its style'


def extract_sheet_tabs(all_code):
    """Extract sheet tab names from getSheetByName calls and constant definitions."""
    tabs = set()

    # getSheetByName('Tab Name') or getSheetByName("Tab Name")
    for m in re.finditer(r"getSheetByName\(['\"]([^'\"]+)['\"]\)", all_code):
        tabs.add(m.group(1))

    # SHEET_NAMES = { KEY: 'Value', ... }
    for m in re.finditer(r":\s*['\"]([A-Za-z][A-Za-z0-9_ ]+)['\"]", all_code):
        val = m.group(1)
        # Only include if it looks like a sheet name (not a URL, email, etc.)
        if not any(c in val for c in ['@', '/', 'http', '.com', '.org']):
            # Check if it's inside a SHEET_NAMES or TAB_ context
            pass  # We'll rely on getSheetByName matches mostly

    # TAB_SOMETHING = 'Value'
    for m in re.finditer(r"(?:const|var|let)\s+TAB_\w+\s*=\s*['\"]([^'\"]+)['\"]", all_code):
        tabs.add(m.group(1))

    # TAB_NAMES array
    for m in re.finditer(r"TAB_NAMES\s*=\s*\[([\s\S]*?)\]", all_code):
        for name in re.finditer(r"['\"]([^'\"]+)['\"]", m.group(1)):
            tabs.add(name.group(1))

    # SHEET_NAMES object values
    for m in re.finditer(r"SHEET_NAMES\s*=\s*\{([\s\S]*?)\}", all_code):
        for name in re.finditer(r":\s*['\"]([^'\"]+)['\"]", m.group(1)):
            tabs.add(name.group(1))

    # var TABS = { KEY: 'Value', ... }
    for m in re.finditer(r"(?:var|const|let)\s+TABS\s*=\s*\{([\s\S]*?)\}", all_code):
        for name in re.finditer(r":\s*['\"]([^'\"]+)['\"]", m.group(1)):
            tabs.add(name.group(1))

    # Filter out obvious non-tab-names
    filtered = set()
    for t in tabs:
        if len(t) > 50 or len(t) < 2:
            continue
        if any(x in t.lower() for x in ['http', 'mailto', '.com', '@', 'function']):
            continue
        filtered.add(t)

    return sorted(filtered)


def extract_functions(all_code):
    """Extract top-level function names."""
    funcs = []
    for m in re.finditer(r'^function\s+(\w+)\s*\(', all_code, re.MULTILINE):
        funcs.append(m.group(1))
    return funcs


def detect_patterns(all_code, gs_files, html_files):
    """Detect key GAS patterns in use."""
    patterns = []

    if re.search(r'function\s+doGet\s*\(', all_code):
        patterns.append('Web app (doGet)')
    if re.search(r'function\s+doPost\s*\(', all_code):
        patterns.append('API endpoint (doPost)')
    if re.search(r'function\s+onOpen\s*\(', all_code):
        patterns.append('Spreadsheet add-on (onOpen menu)')
    if re.search(r'function\s+include\s*\(', all_code):
        patterns.append('HTML include() helper')
    if re.search(r'app_cache|writeAppCache|loadAppCache', all_code):
        patterns.append('app_cache JSON tab pattern')
    if re.search(r'newTrigger|ScriptApp\.newTrigger', all_code):
        patterns.append('Programmatic triggers')
    if re.search(r'deleteTrigger|delete.*[Tt]rigger', all_code):
        patterns.append('Trigger cleanup utility')
    if re.search(r'MailApp\.sendEmail|GmailApp\.sendEmail', all_code):
        patterns.append('Email delivery')
    if re.search(r'generatePdf|buildPdf|createPdf|Pdf', all_code, re.IGNORECASE):
        patterns.append('PDF generation')
    if re.search(r'PropertiesService', all_code):
        patterns.append('PropertiesService for config/caching')
    if re.search(r'LockService', all_code):
        patterns.append('LockService for concurrency')
    if re.search(r'notifyTriggerFailure|catch.*sendEmail|catch.*MailApp', all_code):
        patterns.append('Trigger failure alerting')
    if re.search(r'[Dd]edup|[Aa]lias|[Nn]ormalize.*[Nn]ame', all_code):
        patterns.append('Name normalization/deduplication')
    if re.search(r'initializeSheet|runInitialSetup|initSheets|setupAccountability', all_code):
        patterns.append('Idempotent setup function')

    # Numbered file convention
    if any(re.match(r'^\d{2}_', f) for f in gs_files):
        patterns.append('Numbered file convention (01_, 02_, etc.)')

    # Check for View_*.html pattern
    views = [f for f in html_files if f.startswith('View_')]
    if views:
        patterns.append(f'SPA view system ({len(views)} View_*.html files)')

    return patterns


def detect_app_type(all_code, patterns):
    """Determine if this is a web app, add-on, or standalone."""
    if 'Web app (doGet)' in patterns:
        return 'GAS web app (doGet + HTML Service)'
    elif 'Spreadsheet add-on (onOpen menu)' in patterns:
        return 'GAS spreadsheet add-on (onOpen menu)'
    else:
        return 'GAS project'


def describe_file_structure(gs_files, html_files, project_path):
    """Build a description of the file structure."""
    lines = []

    # Check for subdirectories
    subdirs = []
    for item in sorted(os.listdir(project_path)):
        full = os.path.join(project_path, item)
        if os.path.isdir(full) and not item.startswith('.'):
            subdirs.append(item)

    for f in gs_files:
        lines.append(f'  - {f}')

    if html_files:
        ui_files = [f for f in html_files if not f.startswith('View_')]
        view_files = [f for f in html_files if f.startswith('View_')]
        if ui_files:
            lines.append(f'  - UI: {", ".join(ui_files)}')
        if view_files:
            view_names = [f.replace('View_', '').replace('.html', '') for f in view_files]
            lines.append(f'  - Views ({len(view_files)}): {", ".join(view_names)}')

    if subdirs:
        lines.append(f'  - Subdirectories: {", ".join(subdirs)}/')

    return '\n'.join(lines)


def generate_claude_md(project):
    """Generate CLAUDE.md content for a project."""
    project_path = os.path.join(REPO_ROOT, project['path'])
    if not os.path.isdir(project_path):
        return None

    gs_files, html_files = scan_files(project_path)
    if not gs_files:
        return None  # Not a GAS project, skip

    # Read all code
    all_code = ''
    for f in gs_files:
        all_code += read_file(os.path.join(project_path, f)) + '\n'

    # Extract everything
    style_key, style_note = detect_code_style(all_code)
    tabs = extract_sheet_tabs(all_code)
    functions = extract_functions(all_code)
    patterns = detect_patterns(all_code, gs_files, html_files)
    app_type = detect_app_type(all_code, patterns)
    file_structure = describe_file_structure(gs_files, html_files, project_path)

    # Build the CLAUDE.md
    lines = []
    lines.append(f'# {project["name"]}')
    lines.append('')
    lines.append(project['description'])
    lines.append('')
    lines.append(f'*Auto-generated by update-claude-docs.py on {datetime.now().strftime("%Y-%m-%d")}. Do not edit manually — changes will be overwritten.*')
    lines.append('')

    # Architecture
    lines.append('## Architecture')
    lines.append(f'- **Type:** {app_type}')
    lines.append(f'- **Code style:** {style_note}')
    lines.append(f'- **GS files ({len(gs_files)}):** {", ".join(gs_files)}')
    lines.append(f'- **HTML files ({len(html_files)}):** {", ".join(html_files)}')
    lines.append('')

    # Sheet tabs
    if tabs:
        lines.append(f'## Sheet Tabs ({len(tabs)})')
        for tab in tabs:
            lines.append(f'- `{tab}`')
        lines.append('')

    # Key patterns
    if patterns:
        lines.append('## Key Patterns')
        for p in patterns:
            lines.append(f'- {p}')
        lines.append('')

    # Top-level functions (grouped by file)
    lines.append(f'## Functions ({len(functions)} total)')
    for gs_file in gs_files:
        code = read_file(os.path.join(project_path, gs_file))
        file_funcs = re.findall(r'^function\s+(\w+)\s*\(', code, re.MULTILINE)
        if file_funcs:
            # Show first 8 per file, summarize rest
            shown = file_funcs[:8]
            remaining = len(file_funcs) - len(shown)
            func_list = ', '.join(f'`{f}()`' for f in shown)
            suffix = f' + {remaining} more' if remaining > 0 else ''
            lines.append(f'- **{gs_file}:** {func_list}{suffix}')
    lines.append('')

    return '\n'.join(lines)


def main():
    updated = 0
    skipped = 0

    for project in CONFIG['projects']:
        content = generate_claude_md(project)
        if content is None:
            print(f"[update-docs] SKIP: {project['name']} — no .gs files found")
            skipped += 1
            continue

        claude_path = os.path.join(REPO_ROOT, project['path'], 'CLAUDE.md')

        # Check if content actually changed
        existing = ''
        if os.path.exists(claude_path):
            existing = read_file(claude_path)

        # Compare ignoring the date line (which changes daily)
        def strip_date(s):
            return re.sub(r'Auto-generated by.*\n', '', s)

        if strip_date(existing) == strip_date(content):
            print(f"[update-docs] UNCHANGED: {project['name']}")
        else:
            with open(claude_path, 'w') as f:
                f.write(content)
            print(f"[update-docs] UPDATED: {project['name']}")
            updated += 1

    print(f"[update-docs] Done. {updated} updated, {skipped} skipped.")


if __name__ == '__main__':
    main()
