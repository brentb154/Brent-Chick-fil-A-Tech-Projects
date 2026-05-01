#!/usr/bin/env python3
"""
news-fetcher.py — Fetch AI industry news and GAS updates from RSS/Atom feeds.
Filters by relevance keywords, returns structured results for the weekly digest.
No API key required — pure RSS parsing.
"""

import os
import re
import json
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from html import unescape

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# RSS/Atom feeds to monitor
FEEDS = [
    {
        'name': 'Anthropic Blog',
        'url': 'https://www.anthropic.com/feed.xml',
        'category': 'AI',
        'priority': 'high'
    },
    {
        'name': 'OpenAI Blog',
        'url': 'https://openai.com/blog/rss.xml',
        'category': 'AI',
        'priority': 'high'
    },
    {
        'name': 'Google AI Blog',
        'url': 'https://blog.google/technology/ai/rss/',
        'category': 'AI',
        'priority': 'high'
    },
    {
        'name': 'Google Workspace Updates',
        'url': 'https://workspaceupdates.googleblog.com/atom.xml',
        'category': 'GAS',
        'priority': 'high'
    },
    {
        'name': 'Hacker News (AI)',
        'url': 'https://hnrss.org/newest?q=AI+OR+Claude+OR+GPT+OR+Gemini+OR+LLM&points=100',
        'category': 'AI',
        'priority': 'medium'
    },
    {
        'name': 'The Verge - AI',
        'url': 'https://www.theverge.com/rss/ai-artificial-intelligence/index.xml',
        'category': 'AI',
        'priority': 'medium'
    },
]

# GitHub repos to monitor for releases
GITHUB_REPOS = [
    'anthropics/claude-code',
    'anthropics/anthropic-sdk-python',
    'anthropics/anthropic-sdk-typescript',
    'google/clasp',
]

# Keywords that boost relevance
HIGH_RELEVANCE = [
    'claude', 'anthropic', 'apps script', 'google sheets', 'workspace',
    'coding agent', 'code generation', 'ai agent', 'tool use',
    'claude code', 'api change', 'deprecat', 'breaking change',
    'gemini', 'gpt-5', 'gpt-4', 'o3', 'sonnet', 'opus', 'haiku',
]

MEDIUM_RELEVANCE = [
    'llm', 'large language model', 'ai coding', 'copilot', 'cursor',
    'kiro', 'windsurf', 'vscode', 'ide', 'developer tool',
    'prompt engineering', 'context window', 'fine-tun', 'rag',
    'model release', 'benchmark', 'safety', 'alignment',
]


def fetch_feed(feed_info):
    """Fetch and parse an RSS/Atom feed. Returns list of items."""
    items = []
    try:
        req = urllib.request.Request(feed_info['url'], headers={
            'User-Agent': 'CFA-Tech-Hub-NewsBot/1.0'
        })
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = resp.read()

        root = ET.fromstring(data)

        # Handle both RSS and Atom feeds
        ns = {'atom': 'http://www.w3.org/2005/Atom'}

        # Try RSS format
        for item in root.findall('.//item'):
            title = item.findtext('title', '')
            link = item.findtext('link', '')
            pub_date = item.findtext('pubDate', '')
            description = item.findtext('description', '')
            items.append({
                'title': unescape(title).strip(),
                'link': link.strip(),
                'date_str': pub_date,
                'description': unescape(re.sub(r'<[^>]+>', '', description))[:200].strip(),
                'source': feed_info['name'],
                'category': feed_info['category'],
                'priority': feed_info['priority'],
            })

        # Try Atom format
        for entry in root.findall('.//atom:entry', ns) or root.findall('.//entry'):
            title_el = entry.find('atom:title', ns)
            if title_el is None:
                title_el = entry.find('title')
            link_el = entry.find('atom:link', ns)
            if link_el is None:
                link_el = entry.find('link')
            updated_el = entry.find('atom:updated', ns)
            if updated_el is None:
                updated_el = entry.find('updated')
            if updated_el is None:
                updated_el = entry.find('atom:published', ns)
            if updated_el is None:
                updated_el = entry.find('published')
            summary_el = entry.find('atom:summary', ns)
            if summary_el is None:
                summary_el = entry.find('summary')
            if summary_el is None:
                summary_el = entry.find('atom:content', ns)
            if summary_el is None:
                summary_el = entry.find('content')

            title = title_el.text if title_el is not None and title_el.text else ''
            link = link_el.get('href', '') if link_el is not None else ''
            date_str = updated_el.text if updated_el is not None and updated_el.text else ''
            summary = summary_el.text if summary_el is not None and summary_el.text else ''

            items.append({
                'title': unescape(title).strip(),
                'link': link.strip(),
                'date_str': date_str,
                'description': unescape(re.sub(r'<[^>]+>', '', summary))[:200].strip(),
                'source': feed_info['name'],
                'category': feed_info['category'],
                'priority': feed_info['priority'],
            })

    except Exception as e:
        print(f"[news] Failed to fetch {feed_info['name']}: {e}")

    return items


def fetch_github_releases():
    """Check GitHub repos for recent releases."""
    items = []
    for repo in GITHUB_REPOS:
        try:
            url = f'https://api.github.com/repos/{repo}/releases?per_page=3'
            req = urllib.request.Request(url, headers={
                'User-Agent': 'CFA-Tech-Hub-NewsBot/1.0',
                'Accept': 'application/vnd.github.v3+json'
            })
            with urllib.request.urlopen(req, timeout=10) as resp:
                releases = json.loads(resp.read())

            for release in releases:
                items.append({
                    'title': f"{repo}: {release.get('name', release.get('tag_name', 'new release'))}",
                    'link': release.get('html_url', ''),
                    'date_str': release.get('published_at', ''),
                    'description': (release.get('body', '') or '')[:200],
                    'source': 'GitHub Releases',
                    'category': 'Tools',
                    'priority': 'high',
                })
        except Exception as e:
            print(f"[news] Failed to fetch releases for {repo}: {e}")

    return items


def score_relevance(item):
    """Score an item's relevance. Higher = more relevant."""
    text = (item['title'] + ' ' + item['description']).lower()
    score = 0

    for keyword in HIGH_RELEVANCE:
        if keyword in text:
            score += 3

    for keyword in MEDIUM_RELEVANCE:
        if keyword in text:
            score += 1

    # Boost high-priority sources
    if item['priority'] == 'high':
        score += 2

    return score


def parse_date(date_str):
    """Try to parse a date string. Returns datetime or None."""
    formats = [
        '%a, %d %b %Y %H:%M:%S %z',
        '%a, %d %b %Y %H:%M:%S %Z',
        '%Y-%m-%dT%H:%M:%S%z',
        '%Y-%m-%dT%H:%M:%SZ',
        '%Y-%m-%dT%H:%M:%S.%f%z',
        '%Y-%m-%dT%H:%M:%S.%fZ',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d',
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            return dt.replace(tzinfo=None) if dt.tzinfo else dt
        except ValueError:
            continue
    return None


def fetch_all_news(days_back=7):
    """Fetch news from all sources, filter to last N days, score and rank."""
    cutoff = datetime.now() - timedelta(days=days_back)
    all_items = []

    # Fetch RSS feeds
    for feed in FEEDS:
        items = fetch_feed(feed)
        all_items.extend(items)

    # Fetch GitHub releases
    all_items.extend(fetch_github_releases())

    # Parse dates and filter to recent
    dated_items = []
    for item in all_items:
        dt = parse_date(item['date_str'])
        item['date'] = dt
        if dt and dt < cutoff:
            continue
        item['relevance'] = score_relevance(item)
        if item['relevance'] > 0:  # Only include relevant items
            dated_items.append(item)

    # Deduplicate by title similarity
    seen_titles = set()
    unique_items = []
    for item in dated_items:
        title_key = re.sub(r'[^a-z0-9]', '', item['title'].lower())[:50]
        if title_key not in seen_titles:
            seen_titles.add(title_key)
            unique_items.append(item)

    # Sort by relevance then date
    unique_items.sort(key=lambda x: (-x['relevance'], x['date'] or datetime.min), reverse=False)
    unique_items.sort(key=lambda x: -x['relevance'])

    return unique_items[:20]  # Top 20


def print_report(items):
    """Print news report."""
    print(f"[news] {len(items)} relevant items from the past week\n")

    by_category = {}
    for item in items:
        by_category.setdefault(item['category'], []).append(item)

    for category in ['AI', 'GAS', 'Tools']:
        cat_items = by_category.get(category, [])
        if not cat_items:
            continue
        print(f"  {category}:")
        for item in cat_items[:8]:
            date_str = item['date'].strftime('%m/%d') if item['date'] else '??/??'
            print(f"    [{date_str}] {item['title']}")
            print(f"           {item['source']} | {item['link'][:80]}")
        print()


if __name__ == '__main__':
    items = fetch_all_news()
    print_report(items)
