#!/bin/bash
# Health check — ping deployed GAS web apps
# Only reports failures. Silence = healthy.

CONFIG="/Users/brentjustworking/Desktop/Brent-CFA-Tech-Hub/.automation/config.json"
FAILURES=""

# Read web app URLs from config
# Uses python3 since jq may not be installed
URLS=$(python3 -c "
import json, sys
with open('$CONFIG') as f:
    cfg = json.load(f)
for name, url in cfg.get('web_apps', {}).items():
    if url and 'PASTE_YOUR' not in url:
        print(f'{name}|{url}')
")

if [ -z "$URLS" ]; then
  echo "[health-check] No web app URLs configured. Edit config.json to add them."
  exit 0
fi

while IFS='|' read -r NAME URL; do
  # Curl with 15 second timeout, follow redirects, check for HTTP 200
  HTTP_CODE=$(curl -s -o /dev/null -w "%{http_code}" -L --max-time 15 "$URL" 2>/dev/null)

  if [ "$HTTP_CODE" -ge 200 ] && [ "$HTTP_CODE" -lt 400 ]; then
    : # healthy, say nothing
  else
    FAILURES="$FAILURES\n  - $NAME: HTTP $HTTP_CODE ($URL)"
    echo "[health-check] FAIL: $NAME returned HTTP $HTTP_CODE"
  fi
done <<< "$URLS"

if [ -n "$FAILURES" ]; then
  echo "[health-check] Failures detected:$FAILURES"
  exit 1
else
  echo "[health-check] All web apps healthy."
  exit 0
fi
