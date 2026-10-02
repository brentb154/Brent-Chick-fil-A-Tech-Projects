# Checklist System (FOH, BOH in phase 2)

Replaces the eight separate FOH checklist Google Forms with one master sheet, three forms that rebuild themselves each morning with that day's tasks, and an Apps Script project that logs every submission, tracks what's due vs. done, sends late alerts, and emails a daily summary.

**Master sheet:** Checklist Master (FOH + BOH). The script is container-bound to it.
**Today tab:** first tab in the sheet. It shows tonight's checklists and positions with live status (green = complete, red = late/missed). It's a single formula, so there's no code behind it.
**Setup:** see [SETUP_GUIDE.md](SETUP_GUIDE.md).

## What managers change (no code)

| To... | Edit |
|---|---|
| Add or remove a task, change its days or photo | **Items** row. New tasks need a new, never-reused Item ID. |
| Change a due / late time | **Checklists** Due by / Late after. Applies from the next morning's rows. |
| Change who gets late alerts | **Settings** "FOH escalation email" (read on every run) |
| Change summary recipients | **Settings** "Daily summary recipients" |
| Change rebuild / summary time | **Settings**. Takes effect on the next 15-minute check. |
| Turn on reminders + escalations | **Settings** "Alerts mode" → `Live` |
| Turn on photo uploads | **Settings** "Photo upload page" (one-time setup, see SETUP_GUIDE step 7). Clear "Photo email" once the team is using it. |
| Retire an old photo link / QR | **Settings** "Photo upload key": change it, then reprint the QR from *Show form links*. |
| Change photo retention or alert sensitivity | **Settings** "Keep photos for (days)", "Photo alert sensitivity (std devs)" |
| Change the scorecard day | **Settings** "Weekly scorecard day" (default Tue) |
| Close for a holiday or weather | **Settings** "Closed dates" (e.g. `11/26/2026, 12/25/2026`). Forms close and nothing is tracked. Adding today's date mid-day stops today's alerts. |
| Add a checklist (Restroom 2.0, BOH Closing) | **Checklists** row (Active = TRUE, due times, days) + its **Items**, then *Checklists > Rebuild forms now*. The form and its trigger are created automatically. |

Edits made during the day show up in the next morning's forms. Use *Rebuild forms now* for same-day changes, but not while someone is filling out a checklist: rebuilding wipes their answers. When the store account clicks it, the rebuild runs right away. When anyone else clicks it (say, a director), it's queued and runs within 15 minutes as the store account, because only the store account can edit the forms.

## Files

| File | Job |
|---|---|
| `01_Menu_and_Triggers.gs` | Checklists menu, the 15-minute trigger, install/remove triggers |
| `02_Helpers.gs` | Read tabs by header name, Settings, **all date/time handling** |
| `03_Model.gs` | Checklists / Positions / Items → what's scheduled on a date |
| `04_Validate.gs` | Sheet checks; blocking problems stop the rebuild |
| `05_Rebuild_Forms.gs` | Daily form rebuild, section routing, Form Map |
| `06_Form_Submit.gs` | onFormSubmit → Submissions, Item Results, Daily Status |
| `07_Daily_Status.gs` | Daily rows, Late / Missed |
| `08_Alerts.gs` | Heads-ups, escalations, error alerts, Alert Log |
| `09_Summary.gs` | Morning problems email + weekly scorecard |
| `10_Photos.gs` | Photo upload page (web app), Drive folders, 60-day cleanup, photo checks |
| `PhotoPage.html`, `PhotoJavaScript.html`, `PhotoStylesheet.html` | The upload page people see on their phone |

## How it runs

One time-driven trigger, `quarterHourTick`, runs every 15 minutes:

1. After **Rebuild forms at** (once a day), it runs Validate, rebuilds the forms, creates today's Daily Status rows, and recreates any missing form submit trigger. Only one rebuild can run at a time. If Validate finds a blocking problem, yesterday's forms stay up, the FOH escalation email gets one warning that day, and every later run retries, so fixing the sheet is enough.
2. Every run, it marks past days Missed and flips overdue rows to Late. In `Live` mode it also sends one late alert per row once Late after passes. There's no heads-up before the due time.
3. After **Daily summary at** (once a day), it checks yesterday. It emails only if something needs attention: anything Missed, a task marked "Could not complete", a note, a position with photo tasks that uploaded no photos, or a photo count far below usual. A clean night sends nothing; Completed late still counts as done. The weekly scorecard goes out on **Weekly scorecard day** (Tuesday).

**Photo check:** for each checklist, it compares last night's photo count with the last 28 nights it was submitted. It flags the night if the count is below mean − k × std dev, where k is "Photo alert sensitivity", default 2. The std dev is floored at 1, so one missing photo on a very steady checklist isn't flagged. It needs 10 nights of history first; until then the email shows "building history".

Each form also has an `onChecklistSubmit` trigger. Triggers do work only for the account that installed them (`TRIGGER_OWNER` in Script Properties), so leftover triggers from another account can't double-log.

## Time rules (read before editing date code)

1. **Clock times are read as displayed text** ("11:00 PM") with `getDisplayValues()` and parsed into minutes after midnight. `getValues()` returns time-only cells as a Date on 12/30/1899, whose hour depends on the script and sheet time zones matching.
2. **Dates are written as text** (`yyyy-MM-dd`) into cells formatted as plain text (`@`). Otherwise Sheets converts "2026-10-06" and "11:00 PM" into date/time values, and lookups stop matching.
3. **Business clock:** anything before *Business day ends at* (4:00 AM) belongs to the night before and gets +1440 minutes. 12:30 AM = 1470, which sorts after 11:30 PM = 1410. A Late after of 12:30 AM works.
4. **No millisecond math between days.** DST nights are 23 or 25 hours. Day math uses UTC date parts (`addDays_`). Timestamps go through `Utilities.formatDate` with the spreadsheet's time zone.

## Known limits

- Leader names are free text, so a *missed* checklist can't be charged to a leader. The scorecard shows submissions, on-time %, and task completion %.
- The daily rebuild deletes questions, and Google deletes their stored answers with them. Submissions and Item Results are the only record.
- The photo upload page runs as the store account. After a code change to the page, the store account has to publish a new version (Deploy → Manage deployments → Edit → New version); `clasp push` alone doesn't update it.
- Photos use the store account's 15 GB of free storage. At about 400 KB each, 60 days is roughly 1.5 GB.
- Free Gmail quotas: 100 email recipients a day and 90 minutes of trigger runtime a day. Normal use is well under both.
