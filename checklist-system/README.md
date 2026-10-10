# Checklist System (FOH, BOH in phase 2)

Replaces the eight separate FOH checklist Google Forms with one master sheet, three forms that rebuild themselves each morning with that day's tasks, and an Apps Script project that logs every submission, tracks what's due vs. done, sends late alerts, and emails a daily summary.

**Master sheet:** Checklist Master (FOH + BOH). The script is container-bound to it.
**Today tab:** first tab in the sheet. It shows tonight's checklists and positions with live status (green = complete, red = late/missed). It's a single formula, so there's no code behind it.
**Setup:** see [SETUP_GUIDE.md](SETUP_GUIDE.md).

## What managers change (no code)

| To... | Edit |
|---|---|
| Add a task | **Checklists > Add a task**. Pick the checklist, position, days (one, several, or every day), rotation weeks (rotating checklists only), photo, and where it goes in the list. It fills in the Item ID and Order and puts the row with that position's other tasks. |
| Change or turn off a task | **Checklists > Edit a task**. Change the wording, days, rotation weeks, photo, reference, on/off, or its place in the list. The Item ID never changes, so its history stays together. Tasks are turned off, not deleted. To move a task to another position, turn it off and add it there. A row added by hand needs a new, never-reused Item ID. |
| Check off stations by QR code (Daily Facilities Walk) | **Checklists** "Station QR codes" checked (needs one submission per position). Each position is a station. **Checklists > Print station QR codes** prints one code per station; post each at its station. A scan opens today's form with the station, its code and the person's name filled in. A station only counts when the form came from that station's own QR code; anything else is logged with a "Not counted" note and the station stays missed. Reprint after renaming, adding or removing a station. Changing "Photo upload key" retires the codes. |
| Ask for a typed answer (e.g. a temperature) | **Items** "Answer" = `Type in`: the form shows a text box instead of Complete / Could not complete, and the answer is saved in Item Results. |
| Pick one of several each day | Put the options in braces in the task: `Type in the current temperature of the {Walk In Cooler\|Fry Freezer\|Prep Table}.` (in the sheet, without the backslashes) One is picked for each day (the same all day). Edit the list to add or remove options. |
| Put a form in Spanish (BOH Closing) | **Checklists** "Language" = `Spanish`. Fill each task's **Items** "Spanish task"; a blank one shows in English (Validate lists them). The questions, answers, photo tag, photo lines, date and closed message on the form are in Spanish; the form title (the checklist name) and the photo upload page stay in English. Submissions, Item Results, the morning email and the scorecard stay in English. Add/Edit a task only change the English task, so update the Spanish task when the English wording changes. Takes effect at the next morning update or *Rebuild forms now*. |
| Rotate tasks week to week (Sunday Rotation) | **Checklists** "Rotation length (weeks)" (e.g. 6) and "Rotation start" (the date Week 1 starts). Each task's **Items** "Rotation weeks" says which weeks it's on: `3`, or `1, 5`; blank = every week. The Add/Edit pop-up shows which week the next day falls in. |
| Change a due / late time | **Checklists** Due by / Late after. Applies from the next morning's rows. |
| Change who gets late alerts | **Settings** "FOH escalation email" (read on every run) |
| Change who hears about system errors | **Settings** "System error emails" (blank = the FOH escalation email). The **Alert Log** "Details" column keeps each error's text. |
| Change summary recipients | **Settings** "Daily summary recipients" |
| Change rebuild / summary time | **Settings**. Takes effect on the next 15-minute check. |
| Turn on reminders + escalations | **Settings** "Alerts mode" → `Live` |
| Turn on photo uploads | **Settings** "Photo upload page" (one-time setup, see SETUP_GUIDE step 7). Clear "Photo email" once the team is using it. |
| Retire an old photo link / QR | **Settings** "Photo upload key": change it, then reprint the QR from *Show form links*. |
| Change photo retention or alert sensitivity | **Settings** "Keep photos for (days)", "Photo alert sensitivity (std devs)" |
| Look at any day's photos | **Checklists > View photos**. Flip days with the arrows. Each position shows its status, who submitted it, "Could not complete" tasks, notes, and photos (click to enlarge). Managers need the **Checklist Photos** Drive folder shared with them (Viewer). |
| Change the scorecard day | **Settings** "Weekly scorecard day" (default Tue) |
| Close for a holiday or weather | **Settings** "Closed dates" (e.g. `11/26/2026, 12/25/2026`). Forms close and nothing is tracked. Adding today's date mid-day stops today's alerts. |
| Add a checklist (Restroom 2.0, BOH Closing) | **Checklists** row (Active = TRUE, due times, days) + its **Items**, then *Checklists > Rebuild forms now*. The form and its trigger are created automatically. |

Edits made during the day show up in the next morning's forms. Use *Rebuild forms now* for same-day changes. It only replaces questions that changed, so it usually takes seconds, but anyone in the middle of a checklist with a changed question may have to start over. When the store account clicks it, the rebuild runs right away. When anyone else clicks it (say, a director), it's queued and runs within 15 minutes as the store account, because only the store account can edit the forms.

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
| `11_Photo_Viewer.gs` | Checklists > View photos: one day's statuses, could-not-complete tasks, notes and photos |
| `PhotoViewer.html`, `PhotoViewerJavaScript.html`, `PhotoViewerStylesheet.html` | The View photos pop-up |
| `12_Add_Item.gs` | Checklists > Add a task: next Item ID, row placement, Order renumbering (shared with Edit) |
| `13_Edit_Item.gs` | Checklists > Edit a task: saves changes only if the row hasn't changed since the pop-up opened |
| `AddItem.html`, `AddItemJavaScript.html`, `AddItemStylesheet.html` | The Add / Edit a task pop-up |
| `15_Form_Sync.gs` | Updates each form to today's questions, changing only what differs; remembers each form's layout |
| `16_Languages.gs` | Spanish forms: the form's fixed wording in each language, the Spanish date, and turning Spanish answers back into English for the logs |
| `14_Station_QR.gs` | Station QR codes: per-station codes, the page a QR opens, the printable sheet, today's form entry IDs |
| `StationPage.html`, `QrSheet.html` | What a station QR opens; the printable QR sheet (uses qrcodejs from cdnjs) |

## How it runs

One time-driven trigger, `quarterHourTick`, runs every 15 minutes:

1. After **Rebuild forms at** (once a day), it runs Validate, updates the forms, creates today's Daily Status rows, and recreates any missing form submit trigger. Today's rows are created even if a form fails to rebuild, so one broken form can't leave the day untracked. The failure is emailed once and retried every 15 minutes. Only one rebuild can run at a time. If Validate finds a blocking problem, yesterday's forms stay up, the FOH escalation email gets one warning that day, and every later run retries, so fixing the sheet is enough.
2. Every run, it marks past days Missed and flips overdue rows to Late. In `Live` mode it also sends one late alert per row once Late after passes. There's no heads-up before the due time.
3. After **Daily summary at** (once a day), it checks yesterday. It emails only if something needs attention: anything Missed, a task marked "Could not complete", a note, a position with photo tasks that uploaded no photos, or a photo count far below usual. Its photo section also lists photos uploaded under a checklist or position that was never submitted (usually the wrong position picked on the photo page). A clean night sends nothing; Completed late still counts as done. The weekly scorecard goes out on **Weekly scorecard day** (Tuesday).

**Rotating checklists:** weeks count in 7-day blocks from "Rotation start" (Week 1), then repeat after "Rotation length" weeks. Each day's form gets only the tasks for that day and that week. Validate blocks a rotation length or start that isn't a number or date, and a task week outside 1 to the length. It warns if a week would have no tasks.

**Photo check:** for each checklist, it compares last night's photo count with the last 28 nights it was submitted. It flags the night if the count is below mean − k × std dev, where k is "Photo alert sensitivity", default 2. The std dev is floored at 1, so one missing photo on a very steady checklist isn't flagged. It needs 10 nights of history first; until then the email shows "building history".

Each form also has an `onChecklistSubmit` trigger. Triggers do work only for the account that installed them (`TRIGGER_OWNER` in Script Properties), so leftover triggers from another account can't double-log.

## Time rules (read before editing date code)

1. **Clock times are read as displayed text** ("11:00 PM") with `getDisplayValues()` and parsed into minutes after midnight. `getValues()` returns time-only cells as a Date on 12/30/1899, whose hour depends on the script and sheet time zones matching.
2. **Dates are written as text** (`yyyy-MM-dd`) into cells formatted as plain text (`@`). Otherwise Sheets converts "2026-10-06" and "11:00 PM" into date/time values, and lookups stop matching.
3. **Business clock:** anything before *Business day ends at* (4:00 AM) belongs to the night before and gets +1440 minutes. 12:30 AM = 1470, which sorts after 11:30 PM = 1410. A Late after of 12:30 AM works.
4. **No millisecond math between days.** DST nights are 23 or 25 hours. Day math uses UTC date parts (`addDays_`). Timestamps go through `Utilities.formatDate` with the spreadsheet's time zone.

## Known limits

- Leader names are free text, so a *missed* checklist can't be charged to a leader. The scorecard shows submissions, on-time %, and task completion %.
- **Forms are updated, not rebuilt** (`15_Form_Sync.gs`). Google Forms makes every change a separate slow call, so each form's layout is remembered and only questions that changed are deleted or added (about 20–40 calls on a normal day instead of ~800 for the closing form). Each run stops at 4 minutes, well inside Google's 6, and the next run carries on. A replaced question loses its stored answers in Google; Submissions and Item Results are the record.
- Forms are driven by the sheet: don't edit them by hand. A question added or deleted by hand is noticed on the next update and put back to match the sheet. A question retitled by hand is caught by the weekly full re-read of each form, within 7 days.
- Photo uploads: the position list shows what's already been submitted today ("Chutes – submitted 10:09 PM"), and picking one that hasn't been submitted shows a reminder to submit it too. Picked photos wait in a tray (each can be removed with ✕) and nothing is sent until **Upload** is tapped. Photos are shrunk while they're being picked, so the upload itself is quick. Phones pause web pages that are closed or locked, so the page asks people to keep it open until ✓ and warns if they leave with photos not sent. Failed photos retry on their own 3 times, then show a Try again button.
- The photo upload page runs as the store account. After a code change to the page, the store account has to publish a new version (Deploy → Manage deployments → Edit → New version); `clasp push` alone doesn't update it.
- On days a checklist doesn't run, its form is closed. Google currently rejects the custom "No checklist today." message on these forms, so people see Google's standard "no longer accepting responses" page instead.
- Photos use the store account's 15 GB of free storage. At about 400 KB each, 60 days is roughly 1.5 GB.
- Free Gmail quotas: 100 email recipients a day and 90 minutes of trigger runtime a day. Normal use is well under both.
