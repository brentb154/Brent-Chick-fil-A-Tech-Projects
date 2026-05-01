#!/usr/bin/env python3
"""
weekly-digest.py — Weekly email digest combining all automation results.
Sends a single HTML email with: health scores, code audit, TODOs,
staleness alerts, operator readiness, and AI/tech news.
"""

import os
import sys
import json
import smtplib
import subprocess
import importlib.util
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SCRIPT_DIR, 'config.json')) as f:
    CONFIG = json.load(f)

# Import modules
def load_module(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(SCRIPT_DIR, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

code_audit = load_module('code_audit', 'code-audit.py')
todo_scanner = load_module('todo_scanner', 'todo-scanner.py')
project_health = load_module('project_health', 'project-health.py')
news_fetcher = load_module('news_fetcher', 'news-fetcher.py')


def get_git_activity(days=7):
    """Get git commit summary for the past week."""
    since = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
    try:
        result = subprocess.run(
            ['git', 'log', f'--since={since}', '--oneline', '--stat'],
            capture_output=True, text=True, cwd=CONFIG['repo_root'], timeout=10
        )
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            # Count commits (lines that don't start with space and aren't stat summaries)
            commits = [l for l in lines if l and not l.startswith(' ') and 'file' not in l and 'insertion' not in l]
            return len(commits), result.stdout[:500]
    except Exception:
        pass
    return 0, ''


def health_color(color):
    """Return CSS color for health status."""
    return {'GREEN': '#22c55e', 'YELLOW': '#eab308', 'RED': '#ef4444'}[color]


def readiness_icon(passed):
    return '&#10003;' if passed else '&#10007;'


def readiness_color(passed):
    return '#22c55e' if passed else '#ef4444'


def build_html(health_results, audit_results, todo_results, news_items, git_commits, git_log):
    """Build the HTML email body."""
    now = datetime.now()
    week_of = now.strftime('%B %d, %Y')

    # --- Health Scorecard ---
    health_rows = ''
    for r in sorted(health_results, key=lambda x: x['score']):
        health_rows += f'''
        <tr>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;">
                <span style="display:inline-block;width:12px;height:12px;border-radius:50%;background:{health_color(r['color'])};margin-right:8px;"></span>
                <strong>{r['name']}</strong>
            </td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;text-align:center;font-size:1.3em;font-weight:700;color:{health_color(r['color'])};">{r['score']}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;text-align:center;">{r['violations']}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;text-align:center;">{r['todos']}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;text-align:center;">{r['functions']}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;text-align:center;">{r['last_change']}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;text-align:center;">{r['readiness_score']}/{r['readiness_total']}</td>
        </tr>'''

    # --- Stale Projects ---
    stale_section = ''
    stale = [r for r in health_results if r['staleness'] != 'active']
    if stale:
        stale_items = ''.join(
            f'<li style="margin-bottom:4px;"><strong>{r["name"]}</strong> — {r["staleness"]} (last change: {r["last_change"]})</li>'
            for r in stale
        )
        stale_section = f'''
        <div style="background:#451a03;border:1px solid #92400e;border-radius:8px;padding:16px;margin:16px 0;">
            <h3 style="color:#fbbf24;margin:0 0 8px;">&#9888; Stale Projects</h3>
            <ul style="margin:0;padding-left:20px;color:#fde68a;">{stale_items}</ul>
        </div>'''

    # --- Operator Readiness ---
    readiness_rows = ''
    for r in health_results:
        checks_html = ''
        for check, passed in r['readiness'].items():
            label = check.replace('_', ' ').title()
            checks_html += f'<span style="color:{readiness_color(passed)};margin-right:12px;">{readiness_icon(passed)} {label}</span>'
        readiness_rows += f'''
        <tr>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;"><strong>{r['name']}</strong></td>
            <td style="padding:8px 12px;border-bottom:1px solid #334155;">{checks_html}</td>
        </tr>'''

    # --- Code Audit Violations ---
    audit_section = ''
    total_violations = sum(r['violation_count'] for r in audit_results)
    if total_violations > 0:
        audit_items = ''
        for r in audit_results:
            if r['violation_count'] == 0:
                continue
            rules_summary = ', '.join(f'{rule} ({len(items)})' for rule, items in r['by_rule'].items())
            audit_items += f'<li style="margin-bottom:4px;"><strong>{r["name"]}</strong>: {r["violation_count"]} — {rules_summary}</li>'
        audit_section = f'''
        <h3 style="color:#f1f5f9;margin:24px 0 8px;">Code Audit ({total_violations} violations)</h3>
        <ul style="color:#cbd5e1;padding-left:20px;">{audit_items}</ul>'''
    else:
        audit_section = '<h3 style="color:#f1f5f9;margin:24px 0 8px;">Code Audit</h3><p style="color:#22c55e;">All projects clean — no violations found.</p>'

    # --- TODOs ---
    todo_section = ''
    if todo_results:
        old_todos = [t for t in todo_results if t['age_days'] is not None and t['age_days'] > 30]
        todo_items = ''
        # Show oldest 10
        shown = todo_results[:10]
        for t in shown:
            age = f"{t['age_days']}d" if t['age_days'] is not None else '?'
            color = '#ef4444' if t.get('age_days', 0) and t['age_days'] > 30 else '#94a3b8'
            todo_items += f'''
            <tr>
                <td style="padding:4px 8px;border-bottom:1px solid #1e293b;color:{color};">{t['tag']}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #1e293b;">{t['project']}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #1e293b;">{t['file']}:{t['line']}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #1e293b;color:#94a3b8;font-size:0.85em;">{t['comment'][:80]}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #1e293b;text-align:right;color:{color};">{age}</td>
            </tr>'''
        remaining = len(todo_results) - len(shown)
        todo_section = f'''
        <h3 style="color:#f1f5f9;margin:24px 0 8px;">TODOs &amp; FIXMEs ({len(todo_results)} total, {len(old_todos)} older than 30 days)</h3>
        <table style="width:100%;border-collapse:collapse;font-size:0.9em;">
            <tr style="color:#64748b;text-align:left;">
                <th style="padding:4px 8px;">Type</th><th style="padding:4px 8px;">Project</th>
                <th style="padding:4px 8px;">Location</th><th style="padding:4px 8px;">Comment</th>
                <th style="padding:4px 8px;text-align:right;">Age</th>
            </tr>
            {todo_items}
        </table>'''
        if remaining > 0:
            todo_section += f'<p style="color:#64748b;font-size:0.85em;">... and {remaining} more</p>'
    else:
        todo_section = '<h3 style="color:#f1f5f9;margin:24px 0 8px;">TODOs &amp; FIXMEs</h3><p style="color:#22c55e;">None found.</p>'

    # --- AI & Tech News ---
    news_section = ''
    if news_items:
        by_category = {}
        for item in news_items:
            by_category.setdefault(item['category'], []).append(item)

        for category in ['AI', 'GAS', 'Tools']:
            cat_items = by_category.get(category, [])
            if not cat_items:
                continue
            cat_label = {'AI': 'AI Industry', 'GAS': 'Google Workspace & Apps Script', 'Tools': 'Developer Tools'}[category]
            news_section += f'<h4 style="color:#94a3b8;margin:16px 0 8px;">{cat_label}</h4>'
            for item in cat_items[:6]:
                date_str = item['date'].strftime('%b %d') if item.get('date') else ''
                news_section += f'''
                <div style="margin-bottom:8px;padding:8px 12px;background:#1e293b;border-radius:6px;">
                    <a href="{item['link']}" style="color:#60a5fa;text-decoration:none;font-weight:600;">{item['title']}</a>
                    <div style="color:#64748b;font-size:0.8em;margin-top:2px;">{item['source']} &middot; {date_str}</div>
                </div>'''
    else:
        news_section = '<p style="color:#64748b;">No relevant news this week.</p>'

    # --- Assemble ---
    html = f'''<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="background:#0f172a;color:#e2e8f0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;padding:0;margin:0;">
<div style="max-width:700px;margin:0 auto;padding:24px;">

    <h1 style="font-size:1.5em;margin:0;background:linear-gradient(90deg,#E8363A,#DD2476);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">
        CFA Tech Hub — Weekly Digest
    </h1>
    <p style="color:#64748b;margin:4px 0 24px;">Week of {week_of} &middot; {git_commits} commits this week</p>

    <!-- Health Scorecard -->
    <h2 style="color:#f1f5f9;font-size:1.2em;margin:0 0 12px;">Project Health</h2>
    <table style="width:100%;border-collapse:collapse;font-size:0.9em;">
        <tr style="color:#64748b;text-align:left;">
            <th style="padding:8px 12px;">Project</th>
            <th style="padding:8px 12px;text-align:center;">Score</th>
            <th style="padding:8px 12px;text-align:center;">Violations</th>
            <th style="padding:8px 12px;text-align:center;">TODOs</th>
            <th style="padding:8px 12px;text-align:center;">Functions</th>
            <th style="padding:8px 12px;text-align:center;">Last Change</th>
            <th style="padding:8px 12px;text-align:center;">Ready</th>
        </tr>
        {health_rows}
    </table>

    {stale_section}

    <!-- Operator Readiness -->
    <h3 style="color:#f1f5f9;margin:24px 0 8px;">Operator Readiness</h3>
    <table style="width:100%;border-collapse:collapse;font-size:0.85em;">
        {readiness_rows}
    </table>

    {audit_section}

    {todo_section}

    <!-- News -->
    <h2 style="color:#f1f5f9;font-size:1.2em;margin:32px 0 12px;">AI &amp; Tech News</h2>
    {news_section}

    <hr style="border:none;border-top:1px solid #334155;margin:32px 0 16px;">
    <p style="color:#475569;font-size:0.8em;text-align:center;">
        Generated by CFA Tech Hub Automation &middot; {now.strftime('%Y-%m-%d %H:%M')}
    </p>

</div>
</body>
</html>'''

    return html


def send_via_osascript(to_email, subject, html_body):
    """Send email via macOS Mail.app using osascript (AppleScript).
    Works without SMTP credentials — uses whatever email account is configured in Mail.app.
    """
    # Write HTML to temp file for the body
    tmp_html = os.path.join(SCRIPT_DIR, '_digest_temp.html')
    with open(tmp_html, 'w') as f:
        f.write(html_body)

    # Plain text fallback for Mail.app
    import re
    plain_text = re.sub(r'<[^>]+>', '', html_body)
    plain_text = re.sub(r'\n\s*\n', '\n\n', plain_text)[:5000]

    script = f'''
    tell application "Mail"
        set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{plain_text.replace('"', '\\"').replace(chr(10), '\\n')[:3000]}", visible:false}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{address:"{to_email}"}}
        end tell
        send newMessage
    end tell
    '''

    try:
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            return True, 'Sent via Mail.app'
        else:
            return False, f'Mail.app error: {result.stderr}'
    except Exception as e:
        return False, str(e)


def save_html_report(html_body):
    """Always save the HTML report locally as a fallback."""
    report_path = os.path.join(SCRIPT_DIR, 'latest-digest.html')
    with open(report_path, 'w') as f:
        f.write(html_body)
    return report_path


def main():
    print("[weekly-digest] Generating weekly digest...\n")

    # Gather all data
    print("  Collecting project health data...")
    health_results = project_health.run_health_check()

    print("  Running code audit...")
    audit_results = code_audit.run_audit()

    print("  Scanning TODOs...")
    todo_results = todo_scanner.run_scan()

    print("  Fetching news...")
    news_items = news_fetcher.fetch_all_news(days_back=7)

    print("  Checking git activity...")
    git_commits, git_log = get_git_activity(days=7)

    # Build HTML
    print("  Building email...")
    html = build_html(health_results, audit_results, todo_results, news_items, git_commits, git_log)

    # Save locally (always)
    report_path = save_html_report(html)
    print(f"  Saved report: {report_path}")

    # Try to send email
    to_email = CONFIG.get('alert_email', '')
    subject = f"CFA Tech Hub Weekly Digest — {datetime.now().strftime('%b %d, %Y')}"

    if to_email:
        print(f"  Sending to {to_email}...")
        success, msg = send_via_osascript(to_email, subject, html)
        if success:
            print(f"  Email sent: {msg}")
        else:
            print(f"  Email failed: {msg}")
            print(f"  Report saved locally at: {report_path}")
            print(f"  Open it with: open '{report_path}'")
    else:
        print("  No alert_email configured — report saved locally only.")

    print("\n[weekly-digest] Done.")


if __name__ == '__main__':
    main()
