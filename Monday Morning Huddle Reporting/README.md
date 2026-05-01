# Monday Morning Huddle Reporting

Schedule variance analysis — two locations, planned vs actual hours, weekly reports.

## Setup

1. Open the target Google Sheet
2. Go to Extensions > Apps Script
3. Copy all `.gs` and `.html` files into the script editor
4. Run `initializeAllSheets()` to create required sheet tabs
5. Refresh the spreadsheet — the custom menu appears

## Triggers

- **Weekly archive trigger** — `archiveWeeklySnapshot()` runs weekly to save historical data
- To set up: run `setupWeeklyArchiveTrigger()`
- To remove: run `removeWeeklyArchiveTrigger()`

## Key Files

- `Code.gs` — menu setup, initialization, core utilities
- `Parser.gs` — schedule file parsing logic
- `Analysis.gs` — variance calculations and analysis
- `Report.gs` / `Reports.gs` — report generation
- `Email.gs` — email delivery
- `Archive.gs` — weekly snapshot archiving
- `History.gs` — historical data queries
- `MondaySheet.gs` — Monday sheet management
- `Upload.html` — schedule upload UI
