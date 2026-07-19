# Schedule Counter — Operator Guide

**This turns your posted schedule and your sales history into a weekly labor picture — automatically — so you can see whether you're staffed right instead of eyeballing it.** Upload the schedule; it does the weighting and the math and keeps a weekly record.

## Why it matters
Labor is the biggest cost you actually control, and "does this schedule match the day" is easy to get wrong by feel. This weights the schedule against real sales patterns for you, every week, and tracks productivity over time — so the conversation is about the numbers, not a guess.

## What it does
- **Schedule upload.** Drop in the week's schedule and it processes it.
- **Sales-curve weighting.** It weights the day against your sales curves — how the day actually flows — not a flat average.
- **Weekly snapshots.** Each week is saved, so you see the trend, not just today.
- **Productivity tracking.** Scores that show how you're trending week to week.
- **Reminders.** It nudges you (and whoever you list) to upload the schedule on time.

## How it works (the plain version)
A web app on a Google Sheet, with a weekly autopilot.
- **The web page** is where you upload and look at the numbers.
- **The `config` tab** in the Sheet is your control panel — store name, who gets the reminder emails, and the other settings. Edit it there; no code.
- **A Monday-morning job** does the heavy lifting on its own: it checks the week, files it, runs the forecasting math, and caches the result so the page loads fast.
- **Three reminders** run on a schedule — Sunday night and Thursday to upload, Monday morning to process.

## Everyday tasks
- **Upload the schedule:** open the app → upload the week. (Forget, and the Sunday/Thursday reminders will catch you.)
- **Check where you stand:** the dashboard shows the week's picture and the productivity trend.
- **Change who gets reminders, or the store name:** the `config` tab.
- **Re-run setup after a change:** `initSheets()` is safe to re-run; `installTriggers()` re-installs the scheduled jobs.

## When something looks broken
- **Stuck on "Loading…" or a blank page.** Pasted new code but didn't redeploy — **Deploy → Manage deployments → Edit → New version.**
- **The reminders stopped.** The triggers aren't installed (or got cleared). Run **`installTriggers()`** once; `deleteAllTriggers()` resets them.
- **An upload didn't seem to process.** The weekly job runs Monday morning — if you uploaded after it ran, it picks up next cycle, or you can run the pipeline manually.
- **The numbers look stale.** The page reads a cached snapshot for speed (the `app_cache` tab). Re-running the weekly pipeline refreshes it. Don't hand-edit `app_cache`.

Apps Script platform breakage is rare; a new deployment version plus a re-run of `installTriggers()` clears up almost everything. When in doubt: `initSheets()` → `installTriggers()` → deploy a new version.

## The one rule
**Your settings live in the `config` tab; the code doesn't.** Store name, reminder recipients, thresholds — all editable there. After any code change, publish a new version, and re-run `installTriggers()` if you changed the scheduled jobs.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next. You don't need this to use it week to week.*

### The Monday pipeline is the engine
Everything real happens in one weekly job, `mondayPipeline` (Monday 6 AM), which runs a numbered sequence with a check at each step and one big try/catch that emails an alert if anything throws:
1. **Validate `Sheet1`** — the raw upload landing tab. No sales data → alert and stop.
2. **Read `Sheet1`** into a payload.
3. **Detect the week start** from the payload's dates.
4. **Archive** the week to `sales_history` (tagged `sales_source = 'api'`).
5. **Smoothing update** — fold the new week into `sales_curves`.
6. **Bias detection** — recompute `bias_flags` in `config`.
7. **Cleanup** — delete history rows older than 90 days.
8. **Clear `Sheet1`** so it's ready for the next upload.
8.5. **Refresh actuals calibration** — a hook that's present but dormant; it's a silent no-op unless wired, and if it ever fails it degrades to the configured goal rather than aborting the run.
9. **Build and write `app_cache`.**

That "validate → process → archive → cleanup → cache, alert on failure" shape is the same pattern across these tools.

### How the forecast is actually computed
Not a flat average. `sales_curves` describe how each day of the week actually flows, and the prediction is a **linearly-weighted moving average with outlier exclusion and a trend adjustment** — recent weeks count more, a genuinely weird day is dropped before it can skew the curve (`curve_outlier_pct`, default 30%), and a trend nudge is applied but **capped at ±10%** of the LWMA so it can't run away. This replaced an earlier exponential smoothing (`smoothing_alpha`, default 0.3, still in config); both live in `config` so you can tune them without code. Bias detection watches whether the forecast is consistently over or under and flags it.

### The cache tab
`app_cache` holds a pre-computed JSON blob in cell A1 — the web page reads that one cell instead of recomputing the whole forecast on every load (the app_cache pattern: seconds → milliseconds). The pipeline rebuilds it at step 9; the reader falls back to `getConfig()` if the cache is empty or unparseable, so a missing cache degrades instead of breaking. Never hand-edit it.

### The tabs
`Sheet1` (the raw weekly upload), `config` (settings + `bias_flags`), `sales_curves` (the day-of-week shapes the forecast leans on), `sales_history` (the archived weeks that feed it, pruned to 90 days), `productivity_tracker` (the weekly scores), `app_cache` (the JSON snapshot). `initSheets()` builds them all and is safe to re-run.

### The three triggers
Sunday 10 PM `sundayAlertCheck` (upload reminder), Monday 6 AM `mondayPipeline` (the engine above), Thursday 7 PM `thursdayScheduleReminder`. Installed by `installTriggers()`, removed by `deleteAllTriggers()` — both check for existing triggers so they don't duplicate.

### Deploy model
No auto-sync. Paste `Code.gs` + the HTML files, set the `ALERT_EMAIL` script property (so failures reach you), run `initSheets()` then `installTriggers()`, then publish. The link serves the last *published* version — republish after edits.
