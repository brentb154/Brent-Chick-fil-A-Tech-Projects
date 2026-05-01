# Shared Table Email Summary

Automated weekly waste summary emails from Google Form responses.

## Setup

1. Open the Google Sheet that receives Shared Table form responses
2. Go to Extensions > Apps Script
3. Copy `Code.gs` and `index.html` into the script editor
4. Refresh the spreadsheet — a "Shared Table" menu appears
5. Use the menu to open Settings and configure email recipients, schedule, and format

## Triggers

- **Daily trigger** — `onDailyTrigger()` checks if it's time to send the summary email
- Use the Settings sidebar to configure which day and hour the summary sends
- To remove triggers: run `deleteTrigger()` from the script editor

## Key Files

- `Code.gs` — all server-side logic (settings, email generation, trigger management)
- `index.html` — settings sidebar UI
