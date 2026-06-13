#!/bin/bash
# nightly.sh — Orchestrator for all nightly automation
# Runs: update CLAUDE.md docs → memory maintenance → health checks → git snapshot
# Order matters: doc updates and memory checks happen BEFORE git snapshot so changes get committed.

DIR="$(cd "$(dirname "$0")" && pwd)"
LOG="$DIR/nightly.log"
TIMESTAMP=$(date '+%Y-%m-%d %H:%M:%S')

echo "========================================" >> "$LOG"
echo "Nightly run: $TIMESTAMP" >> "$LOG"
echo "========================================" >> "$LOG"

# 1. Update CLAUDE.md files from live code
echo "" >> "$LOG"
echo "--- Update CLAUDE.md docs ---" >> "$LOG"
python3 "$DIR/update-claude-docs.py" >> "$LOG" 2>&1
DOC_STATUS=$?

# 2. Memory maintenance
echo "" >> "$LOG"
echo "--- Memory maintenance ---" >> "$LOG"
python3 "$DIR/memory-maint.py" >> "$LOG" 2>&1
MEM_STATUS=$?

# 3. Code audit
echo "" >> "$LOG"
echo "--- Code audit ---" >> "$LOG"
python3 "$DIR/code-audit.py" >> "$LOG" 2>&1
AUDIT_STATUS=$?

# 4. TODO scanner
echo "" >> "$LOG"
echo "--- TODO scanner ---" >> "$LOG"
python3 "$DIR/todo-scanner.py" >> "$LOG" 2>&1
TODO_STATUS=$?

# 5. Project health scores
echo "" >> "$LOG"
echo "--- Project health ---" >> "$LOG"
python3 "$DIR/project-health.py" >> "$LOG" 2>&1
HEALTH_SCORE_STATUS=$?

# 5b. Risk audit (security / correctness / duplication) — exits non-zero if anything is AT RISK
echo "" >> "$LOG"
echo "--- Risk audit ---" >> "$LOG"
python3 "$DIR/risk-audit.py" >> "$LOG" 2>&1
RISK_STATUS=$?

# 6. Health checks — ping deployed web apps
echo "" >> "$LOG"
echo "--- Web app health checks ---" >> "$LOG"
bash "$DIR/health-check.sh" >> "$LOG" 2>&1
HEALTH_STATUS=$?

# 7. Weekly digest (only on Sundays)
DAY_OF_WEEK=$(date '+%u')  # 7 = Sunday
DIGEST_STATUS="-"
if [ "$DAY_OF_WEEK" -eq 7 ]; then
  echo "" >> "$LOG"
  echo "--- Weekly digest ---" >> "$LOG"
  python3 "$DIR/weekly-digest.py" >> "$LOG" 2>&1
  DIGEST_STATUS=$?
fi

# 8. Git snapshot (commits any changes including updated CLAUDE.md files) — always last
echo "" >> "$LOG"
echo "--- Git snapshot ---" >> "$LOG"
bash "$DIR/git-snapshot.sh" >> "$LOG" 2>&1
GIT_STATUS=$?

# Summary
echo "" >> "$LOG"
echo "--- Summary ---" >> "$LOG"
echo "  Docs:        $([ $DOC_STATUS -eq 0 ] && echo 'OK' || echo 'ISSUES')" >> "$LOG"
echo "  Memory:      $([ $MEM_STATUS -eq 0 ] && echo 'OK' || echo 'STALE ENTRIES')" >> "$LOG"
echo "  Code audit:  $([ $AUDIT_STATUS -eq 0 ] && echo 'OK' || echo 'VIOLATIONS')" >> "$LOG"
echo "  TODOs:       $([ $TODO_STATUS -eq 0 ] && echo 'OK' || echo 'ITEMS FOUND')" >> "$LOG"
echo "  Proj health: $([ $HEALTH_SCORE_STATUS -eq 0 ] && echo 'OK' || echo 'ISSUES')" >> "$LOG"
echo "  Risk audit:  $([ $RISK_STATUS -eq 0 ] && echo 'OK' || echo 'AT RISK')" >> "$LOG"
echo "  Web apps:    $([ $HEALTH_STATUS -eq 0 ] && echo 'OK' || echo 'FAILURES')" >> "$LOG"
echo "  Digest:      $( [ "$DIGEST_STATUS" = "-" ] && echo 'SKIPPED (not Sunday)' || ([ $DIGEST_STATUS -eq 0 ] && echo 'SENT' || echo 'FAILED'))" >> "$LOG"
echo "  Git:         $([ $GIT_STATUS -eq 0 ] && echo 'OK' || echo 'FAILED')" >> "$LOG"
echo "" >> "$LOG"

# Keep log file from growing forever — trim to last 500 lines
tail -500 "$LOG" > "$LOG.tmp" && mv "$LOG.tmp" "$LOG"
