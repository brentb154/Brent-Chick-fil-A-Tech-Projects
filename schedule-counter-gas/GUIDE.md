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

**The Monday pipeline is the engine.** Once a week (Monday 6 AM) one job runs the whole sequence in order, with a check at each step: validate the upload → read it → archive the week → run the smoothing/forecast → check for bias → clean up → cache the result. If a step fails, it emails an alert instead of silently half-finishing. That "validate → process → archive → cache → alert" shape is the same pattern used across these tools.

**How the forecast is weighted.** It doesn't use a flat average. Sales curves describe how a given day actually flows, and the week is weighted against those — with smoothing so one weird day doesn't whipsaw the number, and with recent weeks counting more than old ones. The point is a forecast that reflects real patterns, not a straight line.

**The cache tab.** `app_cache` holds a pre-computed JSON snapshot so the web page reads one cell instead of recomputing everything on every load — the difference between a slow page and an instant one. The pipeline rebuilds it; never edit it by hand.

**The tabs.** `config` (settings), `sales_curves` (the day-of-week shapes), `sales_history` (what feeds the forecast), `productivity_tracker` (the weekly scores), `app_cache` (the snapshot).

**Deploy model.** No auto-sync. Paste `Code.gs` + the HTML files, set the `ALERT_EMAIL` script property, run `initSheets()` then `installTriggers()`, then publish. The link serves the last *published* version — republish after edits.
