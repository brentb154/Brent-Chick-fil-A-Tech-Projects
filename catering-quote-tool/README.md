# Chick-fil-A Catering Quote Generator

A Google Apps Script web application for creating, managing, emailing, and printing professional catering quotes. Uses a Google Sheet as both the database and settings layer.

---

## Setup Instructions

### Step 1: Create the Google Sheet
Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet. Name it **"CFA Catering Quotes"**.

### Step 2: Open the Apps Script Editor
In your Google Sheet: **Extensions → Apps Script**.

### Step 3: Add the Code Files

You need **four files** in the Apps Script editor:

| GAS Filename | Source File |
|---|---|
| `Code` (auto-created) | `Code.gs` |
| `Index` (HTML) | `App.html` |
| `Pipeline` (GS) | `Pipeline.gs` |
| `PipelineView` (HTML) | `Pipelineview.html` |

**Code.gs:** Replace all code in the existing `Code.gs` with the provided `Code.gs` file.

**Index.html:** Click **+** next to Files → select **HTML** → name it exactly `Index` → paste the contents of `App.html`.

**Pipeline.gs:** Click **+** next to Files → select **Script** → name it `Pipeline` → paste the contents of `Pipeline.gs`.

**PipelineView.html:** Click **+** next to Files → select **HTML** → name it `PipelineView` → paste the contents of `Pipelineview.html`.

> **Important:** `doGet()` uses `HtmlService.createTemplateFromFile('Index')` — the `Index` file name must match exactly.

### Step 4: Initialize the Spreadsheet
Select `initializeSheet` from the function dropdown → click **▶ Run** → authorize when prompted.

Your sheet will have five visible tabs — **Settings**, **Menu**, **Quotes**, **Quote_Sequence**, **Pipeline** — plus a hidden **Quote_Revisions** tab that stores prior versions of edited quotes.

### Step 5: Deploy as a Web App
**Deploy → New deployment** → Web app → Execute as "Me" → Access "Anyone within [org]" (or "Anyone") → Deploy → copy the `/exec` URL.

> After any code change, create a **new deployment version** — the `/exec` URL always serves the last deployed version, not the latest saved code. Use `/dev` during testing to always see the latest save.

### Step 6: Set Up Nightly Archive (Optional)
Triggers (clock icon) → **+ Add Trigger** → `cleanOldQuotes` → Time-driven → Day timer → Midnight to 1am. This moves quotes older than the **Archive After Days** setting (Settings tab, default 120) into a hidden archive sheet — nothing is deleted, and you can change the threshold anytime.

### Step 7: Add Menu Items

The Menu tab has **4 columns**: Category | Item Name | Pickup Price | Delivery Price.

Use `N/A` in the Delivery Price column for items that cannot be delivered — they will automatically be hidden from the item picker when Delivery mode is selected.

**Example entries:**

| Category | Item Name | Pickup Price | Delivery Price |
|---|---|---|---|
| Trays | Chick-fil-A Nuggets Tray - Small | $32.00 | $38.00 |
| Trays | Chick-fil-A Nuggets Tray - Large | $58.00 | $66.00 |
| Box Meals | Chick-fil-A Deluxe Meal | $8.99 | $10.49 |
| Sides | Mac & Cheese | $3.35 | $3.95 |
| Drinks | Gallon Freshly-Brewed Iced Tea | $11.00 | N/A |

Categories are **freeform** — type any category name. Items with the same category text automatically group together in the dropdown.

### Step 8: Configure Settings
Fill in store names, addresses, phone numbers, contact name, tax rate, logo, and email template in the **Settings** tab of the app.

---

## Golden Rules (so it doesn't break)

- **Don't hand-create or rename tabs.** `initializeSheet` builds `Settings`, `Menu`, `Quotes`, `Quote_Sequence`, `Pipeline`, and the hidden `Quote_Revisions` with the exact names the code expects. Renaming one silently breaks that feature.
- **The `Index` file must be named exactly `Index`** (it holds the contents of `App.html`). `doGet` loads it by that name.
- **Redeploy a new version after any code change** — the `/exec` URL serves the last *deployed* version, not your latest save.

## Setup Troubleshooting

| Problem | Fix |
|---------|-----|
| Web app shows a blank page or "script function not found" | Confirm the HTML file holding `App.html` is named exactly `Index` (Step 3). |
| "getSheetByName(...) is null" / a tab is missing | Run `initializeSheet` again (Step 4). It re-creates any missing tab without touching existing data. |
| Menu items don't appear in the picker | Check the **Menu** tab has all 4 columns (Category, Item Name, Pickup Price, Delivery Price). Use `N/A` for non-deliverable items' delivery price. |
| Quote saves but no email/PDF | Re-authorize: run `initializeSheet` once from the editor and click **Allow** (the script needs Gmail + Drive access). |
| Calendar events or reminders not working | See **[SETUP.md](SETUP.md)** — those are the feature-specific deployment steps and their troubleshooting. |
| Changes to code don't show up | You edited but didn't redeploy. **Deploy ▸ Manage deployments ▸ Edit ▸ New version ▸ Deploy.** |

---

## Features

### Searchable Item Picker
- Click or focus the field to see all items grouped by category
- Start typing to filter by item name or category
- Use arrow keys + Enter for keyboard navigation
- Items with `N/A` delivery price are hidden when Delivery is selected
- "Custom Item" option at the bottom for freeform entries

### PO Number
- Optional PO Number field on every new quote
- Shows on the PDF directly below the Quote ID
- Can be added or edited after the fact from the History modal

### Email Integration
- **Save & Email** button — one click saves + generates PDF + emails it
- **Email Quote** button in History — send any saved quote to any email
- PDF is generated server-side and attached automatically
- Optional BCC to receive a copy of every sent quote

### Quote Follow-Up Reminders
Configurable in **Settings → Quote Follow-Up Reminders**:
- Enable/disable toggle
- Set how many days after the quote to send a reminder
- Send to customer (uses their email on the quote), internal staff, or both
- Fully customizable subject and body templates with `{{placeholders}}`
- **Enable Daily Auto-Send** button installs a GAS time-based trigger that runs at 9am — no manual action needed after setup
- Each quote is only reminded once (tracked in the `Reminders_Sent` sheet)

### Sales Pipeline
The **Pipeline** tab shows all emailed quotes organized by sales stage:
- **Quoted / Sent** → **Awaiting Response** → **Confirmed / Booked**
- Stats cards show counts, totals, and overdue follow-ups per stage
- Update status, follow-up date, and notes from the table
- Entries are added automatically when a quote is emailed

### Quote Management
- Sequential IDs (Q-2026-0001, Q-2026-0002, …) that never repeat
- Prices frozen at creation time — menu changes don't affect old quotes
- Edit & Reuse updates the quote in place — and the outgoing version is kept in **revision history**, viewable and restorable from the quote detail popup
- History view: text search plus an event-date filter to pull up any day's orders
- Auto-archive moves quotes older than 120 days to a hidden archive sheet — nothing is deleted (requires trigger setup)

### Calendar Tab
- Month / Week / Day views of every order by event date
- Chips are color-coded pickup vs delivery and prefixed 🔴 (needs PO) / 🟢 (PO handled), matching the Google Calendar event titles
- Click any order to open its detail popup

### Delivery Helpers
- Google Maps directions link on delivery quotes — on the PDF, in the detail popup, and in the calendar event description (no API key needed)
- "Busy delivery window" warning before saving a 4th delivery within ±60 minutes of an existing one

### Quarterly Price Check
- Every **Price Check Interval (Days)** (default 90) the app locks on open until the operator types the current POS prices for 3 randomly chosen menu items (pickup and delivery both checked)
- Items are drawn from the categories listed in **Price Check Categories** — matched by category name, so menu rows can move freely
- A mismatch means the Menu tab is stale: fix the item's price there, then re-verify

---

## Spreadsheet Architecture

### Menu Tab (4 columns)
| Column | Description |
|---|---|
| **Category** | Freeform grouping label (e.g., "Box Meals", "Sides") |
| **Item Name** | The menu item name shown in the picker |
| **Pickup Price** | Price for pickup orders; use `N/A` if pickup is not available |
| **Delivery Price** | Price for delivery orders; use `N/A` if delivery is not available |

### Quotes Tab (23 columns)
Columns A–V cover the quote fields (customer, contact, order type, line items JSON, totals, tax, PO, event date/time, calendar event ID, discounts, notes); column W is **Customer Phone**. All columns are written by the app — nobody edits this tab by hand. The hidden **Quote_Revisions** tab mirrors these columns plus a **Revised At** timestamp.

### Settings Tab
| Label | Description |
|---|---|
| Store Name (Active) | Which location appears on quotes |
| Location 1/2 Name/Address/Phone | Store details |
| Quote Contact Name | Shown on PDF "questions?" line |
| Default Tax Rate (%) | Pre-filled on new quotes |
| Archive After Days | How old a quote gets before `cleanOldQuotes` archives it (default 120) |
| Calendar Lead Time (Minutes) | How many minutes before the order time the calendar event starts (default 30) |
| Price Check Interval (Days) | How often the quarterly price check locks the tool (default 90) |
| Price Check Categories | Comma-separated menu categories the price spot-check draws from |
| Last Price Verification | Written automatically when a price check passes — don't edit |
| Logo (Base64) | Uploaded via the app Settings tab |
| Email Subject | Template with `{{placeholders}}` |
| Email Body | Template with `{{placeholders}}` |
| BCC Email | Optional copy recipient for every sent quote |
| Reminder Enabled | `TRUE` / `FALSE` |
| Reminder After Days | Number of days after quote before sending reminder |
| Reminder Send To Customer | `TRUE` / `FALSE` |
| Reminder Send To Internal | `TRUE` / `FALSE` |
| Reminder Internal Email | Staff email for internal reminder copies |
| Reminder Subject | Template with `{{placeholders}}` |
| Reminder Body | Template with `{{placeholders}}` |

### Email & Reminder Placeholders
**Email:** `{{customer}}` `{{contact}}` `{{location}}` `{{phone}}` `{{quoteId}}` `{{total}}` `{{date}}`

**Reminder (all above +):** `{{daysSince}}`

---

## Troubleshooting

| Issue | Fix |
|---|---|
| App redirects to wrong page on deploy | Make sure the GAS `Index` HTML file contains `App.html` content, not a GitHub page |
| Pipeline view has white space at top | Ensure `<?!= include('PipelineView'); ?>` is inside `<main class="main-content">` in `Index.html` |
| `Cannot read properties of null` on Pipeline tab | `doGet()` must use `createTemplateFromFile().evaluate()`, not `createHtmlOutputFromFile()` |
| Pop-up blocked on PDF | Allow pop-ups for the Apps Script domain |
| Email not sending | Check daily quota (`MailApp.getRemainingDailyQuota()`); re-authorize if needed |
| Items not showing in delivery picker | Those items have `N/A` in the Delivery Price column — this is intentional |
| Reminder trigger not running | Click "Enable (9am daily)" in Settings → Reminders to install the trigger; verify in Apps Script → Triggers |
| Old Menu tab had 3 columns | Add a `Category` column A and shift existing data right one column |
| Settings not saving reminder fields | The Settings tab reads up to row 50 — ensure rows aren't beyond that |
