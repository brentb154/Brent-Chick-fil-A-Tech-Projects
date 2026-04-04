# Catering Quote Tool — Setup Instructions for New Features

These instructions cover deploying three new features:
1. **Calendar Event Creation** — auto-creates a Google Calendar event when a quote is saved
2. **Quote Edit Persistence** — editing an existing quote overwrites the original row instead of creating a duplicate
3. **Reminder Email Fix** — fixes silent failure in the follow-up reminder trigger

---

## Pre-Flight Checklist

- [ ] You have editor access to the Google Apps Script project
- [ ] You have editor access to the linked Google Sheet
- [ ] The Google account running the script has access to Google Calendar

---

## Step 1: Back Up Your Sheet

1. Open the linked Google Sheet
2. **File > Make a copy** — save it somewhere safe
3. This is your rollback if anything goes wrong

---

## Step 2: Copy the Updated Code into Apps Script

Open the Apps Script editor (Extensions > Apps Script from the Sheet, or script.google.com).

Replace these files with the updated versions:

| File | What changed |
|------|-------------|
| `Code.gs` | Calendar event creation, quote edit-in-place, reminder fix, new columns (17-19) |
| `App.html` | Edit & Reuse now preserves quote number, time field persists, calendar status in detail modal |

**How to replace:**
1. Click on `Code.gs` in the left sidebar
2. Select all (Cmd+A), delete
3. Paste the full contents of the updated `Code.gs` from this repo
4. Repeat for `App.html`

`Pipeline.gs` and `PipelineView.html` are **unchanged** — do not touch them.

---

## Step 3: Run `initializeSheet()` to Migrate Columns

This adds the three new column headers to your existing Quotes sheet. It will NOT delete or modify any existing data.

1. In the Apps Script editor, select `initializeSheet` from the function dropdown (top toolbar)
2. Click **Run**
3. If prompted for permissions, click **Review Permissions > Allow**
   - The script needs Calendar access now (it didn't before)
4. Check the Execution Log — you should see no errors

**Verify in the Sheet:**
Open the Quotes tab. You should now see three new columns at the end:
- Column Q: `Calendar Event ID`
- Column R: `Last Modified`
- Column S: `Event Time`

These columns are auto-populated by the app. Nobody needs to manually enter anything.

---

## Step 4: Authorize Calendar Access

The first time a quote with a date and time is saved, the script will try to use `CalendarApp`. If you haven't authorized Calendar access yet:

1. In the Apps Script editor, select `createCalendarEvent` from the function dropdown
2. Click **Run** — it will fail (no data), but this triggers the OAuth consent screen
3. Click **Review Permissions > Allow** and approve Calendar access
4. You're done — future saves will work automatically

---

## Step 5: (Optional) Set a Specific Calendar

By default, events are created on the **default calendar** of the Google account running the script. If you want events on a different shared calendar:

1. Open Google Calendar
2. Find the target calendar in the left sidebar
3. Click the three dots > **Settings and sharing**
4. Scroll to **Integrate calendar** > copy the **Calendar ID** (looks like `abc123@group.calendar.google.com`)
5. In your linked Google Sheet, go to the **Settings** tab
6. Add a new row:
   - Column A: `Calendar ID`
   - Column B: paste the Calendar ID
7. Save — all future events will be created on that calendar

If this row is blank or missing, events go to the default calendar. Either way works.

---

## Step 6: Deploy the Web App

1. In Apps Script, click **Deploy > Manage deployments**
2. Click the **pencil icon** on your existing deployment
3. Under "Version", select **New version**
4. Click **Deploy**
5. The web app URL stays the same — no need to update any links

---

## Step 7: Test Each Feature

### Test 1: Calendar Event Creation
1. Open the web app
2. Create a new quote with a **date AND time** filled in
3. Save it (any method — Save, Save & Email, or Print)
4. Open Google Calendar and navigate to that date
5. **Expected:** A 30-minute calendar event appears, starting 30 minutes before the order time
6. Title should read: `🔴 NEEDS PO — [Customer Name] — [Date]`
7. Click the event — description should contain full quote details

### Test 2: Calendar PO Update
1. Open Quote History in the web app
2. Click on the quote you just created
3. Enter a PO number and click **Save PO**
4. Go back to Google Calendar
5. **Expected:** The event title now reads `🟢 HAVE PO — [Customer Name] — [Date]`

### Test 3: No Time = No Calendar Event
1. Create a new quote with a date but **no time**
2. Save it
3. **Expected:** Quote saves normally, but no calendar event is created (the Calendar Event ID cell in the Sheet will be blank)

### Test 4: Quote Edit Persistence
1. Open Quote History, click a quote, click **Edit & Reuse**
2. Change something (add an item, change quantity, etc.)
3. Click **Save**
4. **Expected:** Toast says "Quote Q-XXXX-XXXX updated!" (not "saved")
5. Open the Sheet — the original row is updated in place, same quote number, `Last Modified` column has a timestamp
6. No new row was created

### Test 5: Time Survives Edit & Reuse
1. Open Quote History, click a quote that had a time set
2. Click **Edit & Reuse**
3. **Expected:** The Time field on the form is pre-filled with the original time

### Test 6: Reminder Fix (if reminders are enabled)
1. In Settings, ensure reminders are enabled
2. Check the `Reminders_Sent` sheet
3. If any quotes were marked as reminded but never actually received an email (the bug that was fixed), delete those rows from `Reminders_Sent` so they get re-processed on the next trigger run

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "CalendarApp is not defined" or permission error | Run `initializeSheet()` or `createCalendarEvent` manually once to trigger the OAuth consent screen. Approve Calendar permissions. |
| No calendar event created but quote saved fine | Check that both Date AND Time are filled in on the form. No time = no event. |
| Events on wrong calendar | Add or update the `Calendar ID` row in the Settings tab (see Step 5) |
| New columns not appearing | Run `initializeSheet()` again. Check that the Quotes tab has headers through column S. |
| Edit & Reuse creates a new quote instead of updating | Make sure you're using the updated `App.html`. The `editingQuoteId` variable must be set. |
| Reminder emails still not sending | Check `Reminders_Sent` sheet for stale entries from the old bug. Delete rows for quotes that never actually received an email. |
| "Cannot read property of undefined" on old quotes | Old quotes won't have columns Q-S populated. This is fine — the code handles empty values gracefully. |
