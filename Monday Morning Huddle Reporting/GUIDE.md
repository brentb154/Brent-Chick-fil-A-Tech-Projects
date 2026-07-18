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

**The pipeline.** Upload → parse the schedule file → analyze (planned vs actual, variance, OT, attendance flags) → build the two views (last week's review, this week's projection) → email the two-tier summaries → archive the presentation. Each stage is its own file (parsing, analysis, reporting, email, archive, history), so a change to one stays contained.

**Why cross-location OT matters here.** Overtime is a per-*employee* number, not a per-*store* number. Someone who works 25 hours at each location has 10 hours of OT that neither store's own report will ever show. This tool combines an employee's hours across both locations before it computes OT — that's the whole reason it exists alongside the standard analytics.

**The multi-week trap.** The one recurring failure mode is uploading a file that spans more than one week, which double-counts hours. There's a preflight check that inspects the upload and flags a multi-week file before it processes, plus week-labeled history filters so old weeks don't bleed into the current one. If numbers ever look doubled, check this first.

**Archive + history.** A weekly trigger snapshots each huddle so you build a season of history; the history queries read from those snapshots. Keep the trigger installed (`setupWeeklyArchiveTrigger`) or the history stops accumulating.

**It's a Sheets add-on, not a web app.** No `/exec` URL — the "deploy" is the Sheet, its bound script, and the custom menu. Re-running `initializeSheet()` is safe and rebuilds the tabs without touching your data.
