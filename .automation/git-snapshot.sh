#!/bin/bash
# Git auto-commit — nightly safety snapshot
# Only commits if there are actual changes. Never pushes automatically.

REPO="/Users/brentjustworking/Desktop/Brent-CFA-Tech-Hub"
cd "$REPO" || exit 1

# Check for changes (tracked + untracked, excluding .automation/nightly.log)
CHANGES=$(git status --porcelain | grep -v 'nightly.log' | head -1)

if [ -z "$CHANGES" ]; then
  echo "[git-snapshot] No changes detected. Skipping."
  exit 0
fi

# Stage everything except the log file
git add -A
git reset -- .automation/nightly.log 2>/dev/null

# Count what's staged
STAGED=$(git diff --cached --stat | tail -1)
if [ -z "$STAGED" ]; then
  echo "[git-snapshot] Nothing staged after filtering. Skipping."
  exit 0
fi

# Commit with timestamp
DATE=$(date '+%Y-%m-%d %H:%M')
git commit -m "nightly snapshot: $DATE

Auto-committed by .automation/git-snapshot.sh"

echo "[git-snapshot] Committed: $STAGED"
