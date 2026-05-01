# Schedule Counter GAS

Labor scheduling tool — schedule upload, sales curve weighting, weekly snapshots, productivity tracking.

## Setup

1. Open the target Google Sheet
2. Go to Extensions > Apps Script
3. Copy all `.gs` and `.html` files into the script editor
4. Run `runInitialSetup()` — this creates all sheet tabs and installs triggers in one step
5. Deploy as web app (Execute as: Me, Access: Anyone within domain)

## Triggers

Three time-driven triggers are installed automatically by setup:

- **Sunday 10 PM** — `sundayAlertCheck` (schedule upload reminder)
- **Monday 6 AM** — `mondayPipeline` (weekly processing pipeline)
- **Thursday 7 PM** — `thursdayScheduleReminder` (schedule reminder)

To reinstall triggers: run `installTriggers()`
To remove all triggers: run `deleteAllTriggers()`

## Sheet Tabs

- **config** — operator-editable settings (store name, email recipients, etc.)
- **sales_curves** — sales distribution curves by day of week
- **sales_history** — historical sales data for forecasting
- **productivity_tracker** — weekly productivity scores
- **app_cache** — pre-computed JSON cache (do not edit manually)

## Key Files

- `Code.gs` — all server-side logic (setup, triggers, data processing, web app)
- `Index.html` — main web app UI
- `JavaScript.html` — client-side logic
- `Stylesheet.html` — CSS
