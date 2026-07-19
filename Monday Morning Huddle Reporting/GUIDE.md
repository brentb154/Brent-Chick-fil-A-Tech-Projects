# Monday Morning Huddle — Operator Guide

**This builds the Monday huddle report for you: last week's schedule-vs-actual review and this week's projected hours in one place — with the cross-location overtime that per-store reports quietly miss — then it emails the summary and archives the presentation.** Upload the schedule, and the meeting's numbers are ready.

## Why it matters
The Monday huddle lives or dies on accurate numbers, and pulling "planned vs actual, OT, and attendance" together by hand every week is slow and easy to get wrong — especially overtime, which hides when one employee works both locations and each store only sees its own half. This does the math in one pass, catches that cross-location OT, and hands you a report you can present and email.

## What it does
- **Last Week Review** — schedule vs actual hours, variance, overtime, absences, missed clock-outs, and attendance flags.
- **This Week Plan** — projected scheduled hours and scheduled OT, so you catch problems before the week starts.
- **Cross-location OT detection** — combines an employee's hours across both locations, catching overtime a single-store view misses.
- **Two-tier email summaries** — the right level of detail for directors vs. managers, sent automatically.
- **Auto-archive** — every week's huddle presentation is saved, so you keep the history.

## How it works (the plain version)
It lives inside Google Sheets — a custom menu, not a separate app.
- **You upload the schedule.** It parses the file and pulls out the hours.
- **The Sheet does the analysis.** It compares planned vs actual, finds the variances and the OT (including across locations), and builds the two views — last week and this week.
- **It emails and archives.** The two-tier summaries go out, and the week's report is filed automatically.
- **Settings live in the Sheet.** Recipients and the like are config you edit, not code.

## Everyday tasks
- **Run the weekly huddle:** upload the week's schedule from the menu → the report builds → present it, and the emails go out.
- **Look back:** the archive keeps every week; the history is queryable.
- **Change who gets the emails:** the recipient settings in the Sheet.
- **Re-run setup after a change:** `initializeSheet()` rebuilds the tabs safely; `setupWeeklyArchiveTrigger()` (re)installs the weekly save.

## When something looks broken
- **Hours look doubled or too high.** That's the classic one: uploading more than one week at a time double-counts. Upload a single week — the preflight check now warns you if a file looks like it spans multiple weeks.
- **The custom menu isn't there.** Refresh the spreadsheet after setup — the menu is added on open. Still missing? Re-run `initializeSheet()`.
- **The weekly archive didn't save.** The trigger isn't installed — run `setupWeeklyArchiveTrigger()` once. `removeWeeklyArchiveTrigger()` clears it.
- **Emails didn't go out.** Check the recipient settings and that the script is authorized to send mail from the right account.
- **A number disagrees with a store's own report.** Usually the cross-location combine — this tool sums an employee's hours across both stores on purpose, which is exactly the OT a single-store report won't show.

Real breakage is rare; most "wrong numbers" trace back to a multi-week upload, and most "nothing happened" traces back to the archive trigger not being installed.

## The one rule
**Upload one week at a time, and keep the recipient settings current in the Sheet.** Everything else — the math, the views, the archive — runs itself.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

### The pipeline
Upload lands the schedules and punches into **six input tabs**. From there: parse the punches (`parsePunchesFromSheet`) → analyze (planned vs actual, variance, OT, attendance flags) → write the two views onto the Monday sheet (last-week review in one block of rows, this-week projection in another) → append the week to `History` → send the two-tier emails → (later) the input tabs get wiped for the next week. Each stage is its own file — `Parser.gs`, `Analysis.gs`, `Report(s).gs`, `Email.gs`, `Archive.gs`, `History.gs` — so a change to one stays contained.

### `History` is the spine — and it's protected
`History` is the permanent backend data store. It powers two things nothing else can: the **chronic flags** (an employee who's flagged week after week) and every historical report. Because so much depends on it, the tab is **protected** (`protectHistoryTab_`) — automated writes only. **Editing `History` by hand corrupts the chronic flags and the historical reports.** If a `History` row is wrong, re-run the week rather than editing the cell. Rows carry a `weekLabel`, which is how reports pull one week without bleeding in the last.

### Why cross-location OT matters here
Overtime is a per-*employee* number, not a per-*store* number. Someone who works 25 hours at each location has 10 hours of OT that neither store's own report will ever show. This tool **combines an employee's hours across both CH and DBU before it computes OT** — that's the whole reason it exists alongside the standard analytics. The corollary is a real gotcha: **if you only upload one location, the combined-hours OT is invisible** — you have to upload both for the number to be right.

### The multi-week trap
The one recurring failure mode is uploading a file that spans more than one week, which double-counts hours. The **upload dialog checks `History` and warns you** before it processes — read the warning. Combined with the `weekLabel` filtering, that's what keeps old weeks from stacking onto the current one. If numbers ever look doubled, this is the first thing to check. (Related: closing the upload dialog early can leave the input tabs half-written — finish the run.)

### The two-tier email
`sendTwoTierEmails` builds two levels from the same analysis: a director-level summary (staffing and labor conversations) and a manager-level one (more detail). There's also a re-send that emails the latest analysis again **without** re-running it — handy if a recipient was wrong or the first send failed. Recipients live on the **Settings** tab; update them when a manager leaves or a director joins.

### It's a Sheets add-on, not a web app
No `/exec` URL — the "deploy" is the Sheet, its bound script, and the custom menu. A weekly trigger (`setupWeeklyArchiveTrigger`) snapshots each huddle so the history keeps accumulating; keep it installed. Re-running `initializeSheet()` is safe and rebuilds the tabs without touching your data.
