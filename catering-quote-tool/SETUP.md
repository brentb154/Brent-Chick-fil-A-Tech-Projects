# Catering Quote Tool — Setup Instructions for New Features

## July 2026 Batch 4 (newest) — two-restaurant switching

Run one location or two from the same tool. Store 1 keeps everything it has today; Store 2 gets its own menu and cheat sheet.

**To deploy:** paste the updated `Code.gs` and `Index` (`App.html`), run `initializeSheet()` once, then redeploy (**Manage deployments → Edit → New version**). No new files or permissions.

The tabs are **store-labeled** so nobody edits the wrong restaurant's prices. `initializeSheet()` **renames your existing tabs in place** — `Menu` → **CH Menu**, `Off_Menu` → **CH Off-Menu** — so all your current Cockrell Hill prices move with them (nothing is lost). It then creates **DBU Menu** and **DBU Off-Menu** (empty) for Dallas Baptist University; fill DBU Menu the same way you filled the old Menu (CFA Home export → paste). Safe to re-run — the rename only happens once, and it never touches Cockrell Hill's data. It also adds a per-store `Last Price Verification 2` Settings row (written automatically the first time you verify DBU's prices).

**How it works:** the quote form has a **Restaurant** toggle above Order Type — it switches the menu, the cheat sheet, and the pickup address to that store, and re-prices any items already on the quote (with a confirm). The **Menu Catalog** editor and **Cheat Sheet** each have their own Store 1 / Store 2 switch so you edit the right menu. **Settings → Default store for new quotes** sets which store the form opens on the first time; after that it remembers whichever store you last used. Editing or reordering a past quote automatically switches to that quote's store. Tax rate stays shared (still editable per quote). The quarterly price check now runs per store — each store's menu is verified on its own 90-day clock, and an empty/unconfigured Store 2 never triggers the lock.

**Also in this deploy:** the **Tax Forms** tab now shows a copyable **Guest Upload Link** (the public tax-exempt upload page) you can hand to anyone directly; the **Get Directions** link was removed from the quote PDF (it stays on the internal quote popup and the delivery runsheet); and menu lookups on the quote form now use an in-memory index so repricing and edit/reorder are instant on large menus.

## July 2026 Batch 3 — tax-exempt registry + guest upload page

This adds a guest-facing tax-form upload page and an organization-level registry, and **changes how the app is deployed.**

**What changed in deployment (important):**
1. Add a **third HTML file** named exactly `TaxForm` holding `TaxForm.html` (alongside `Code` and `Index`).
2. Redeploy with **Access = "Anyone"** (not "Anyone within your domain") so external guests can reach the upload page. Edit the existing deployment → Access → Anyone → new version.
3. The team's URL is now `…/exec?view=app` (bookmark it). Guests get `…/exec?view=taxform&quote=…` automatically inside the request email. The bare `…/exec` shows a harmless landing page.
4. Run `initializeSheet()` — it creates the visible **Tax_Exempt_Registry** tab and hidden **Tax_Form_Uploads** tab. First run prompts a Drive permission — Allow.

**How it works:** on a tax-exempt quote, "Look it up" opens the registry; "Request from Guest" emails an upload link. Guests upload a PDF (no login) → it lands in the Drive folder as a **pending review** on the **Tax Forms** tab → a team member confirms it into the registry under the right organization. **Add Existing Form** backfills what you already have. The Dec year-end reminder now summarizes the registry.

## July 2026 Batch 2 — customer memory, runsheet, automation, Pipeline retired

This update adds: customer memory (org + contact autofill from past quotes), a **Reorder** button, a printable/emailable daily delivery runsheet, day-before confirmation emails, a daily missing-PO alert digest, a "valid through" date on PDFs, Settings-tunable delivery warning, a collapsible menu editor with CFA Home import instructions — and **removes the Pipeline feature entirely**.

**To deploy it:**
1. Paste the new `Code.gs` into `Code.gs` and the new `App.html` into the `Index` HTML file.
2. **Delete the `Pipeline` and `PipelineView` files** from the Apps Script project (click the ⋮ next to each → Remove). The old `Pipeline` sheet tab keeps its data — delete the tab by hand whenever you like.
3. Run `initializeSheet()` once. It creates the hidden `Confirmations_Sent` / `PO_Alerts_Sent` log tabs and seeds the new Settings rows (`Quote Valid For (Days)`, `Delivery Warning Count`, `Delivery Warning Window (Minutes)`, `Confirmation Enabled/Subject/Body`, `PO Alert Enabled/Days Before/Email`). Safe to re-run; touches no data.
4. Deploy → Manage deployments → Edit → **New version**.
5. In the app: **Settings → Daily Automation** — flip on the jobs you want and click **⏰ Enable (3pm daily)** once. Confirmations default **off**; PO alerts default **on** (recipient = the deploying account until you set `PO Alert Email`).

## July 2026 Batch 1 — PDF details, Calendar tab, revisions, price check

Added: customer phone + email on the PDF, a quote history date filter, a Calendar tab (Month/Week/Day with 🔴/🟢 PO status), a busy-delivery-window warning, Google Maps directions links on deliveries, a 15-minute time dropdown, quote revision history (view + restore), and a quarterly price-check lock.

Its migration is covered by the same `initializeSheet()` run as Batch 2 (adds the `Customer Phone` column W, the hidden `Quote_Revisions` tab, and the price-check Settings rows). **Expect a one-time lock on first open:** the app asks for the current POS prices of 3 menu items before it unlocks; it repeats every `Price Check Interval (Days)` (default 90). Revision history starts at this deploy — earlier edits aren't recoverable.

---

## Earlier Feature Deployment (calendar events, edit persistence, reminder fix)

These instructions cover deploying three earlier features:
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

(`Pipeline.gs` and `PipelineView.html` existed at the time of this deployment but were retired in July 2026 Batch 2 — a current install has neither.)

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
5. **Expected:** A calendar event appears, starting a set number of minutes before the order time. That lead time is the **Calendar Lead Time (Minutes)** value in the Settings tab (default 30) — change it there to make events start earlier or later.
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
