# FOH Links Menu

A Google Sheets add-on that gives front-of-house team members one-click access to every link they need (training docs, forms, schedules) from a custom menu and sidebar — plus an automatic Sunday reset that wipes the week's day sheets back to clean templates.

Everything is configured in the spreadsheet itself: links live in the **Links** tab, behavior in the **Settings** tab. No code edits needed to maintain it.

> 📖 **[Go deeper: what it is & how it works →](GUIDE.md)** — the plain-language operator guide (what it does, how it works, and how to fix the common stuff).

---

## Setup Instructions

### Step 1: Open the Apps Script Editor
In your Google Sheet: **Extensions → Apps Script**.

### Step 2: Add the Code Files

Create **four script files** and **two HTML files**:

| GAS Filename | Source File | Type |
|---|---|---|
| `01_Menu_and_Setup` | `01_Menu_and_Setup.gs` | Script |
| `02_Link_Functions` | `02_Link_Functions.gs` | Script |
| `03_Weekly_Reset` | `03_Weekly_Reset.gs` | Script |
| `04_Settings` | `04_Settings.gs` | Script |
| `Sidebar` | `Sidebar.html` | HTML |
| `DonationDialog` | `DonationDialog.html` | HTML |

Click **+** next to Files, pick Script or HTML, name it exactly as shown, paste the contents.

### Step 3: Run Initial Setup
Select `runInitialSetup` from the function dropdown → **▶ Run** → authorize when prompted. This creates the **Links** tab (with example rows) and the **Settings** tab. Safe to re-run — it never overwrites existing data.

### Step 4: Add Your Links
On the **Links** tab, one row per link: **Category | Name | URL** (plus optional Pin column managed from the sidebar). Reload the sheet and the **FOH Links** menu appears with everything grouped by category.

### Step 5 (Optional): Weekly Reset
If you use the day tabs (Monday–Saturday with matching "(Reset)" templates):

1. Set **Reset Day** and **Reset Hour** on the **Settings** tab.
2. Run `testResetMonday` from the menu first to verify the template copy looks right.
3. Run `setupResetTrigger` — installs the weekly trigger (removes any existing reset trigger first, so re-running is safe).

If a reset run has problems, it emails the **Alert Email** from Settings (falls back to the sheet owner). Remove the automation anytime with `deleteResetTriggers`.

---

## Day-to-Day Use

- **FOH Links menu** — every link, grouped by category, opens in a new tab.
- **Open Links Panel** — sidebar with search, pinning, and add/edit/delete for links.
- **Donation dialog** — pre-filled donation request email (defaults from Settings).

## Settings Tab Reference

| Setting | What it does |
|---|---|
| Reset Day | Day of week the automatic reset runs (e.g. `Sunday`) |
| Reset Hour | Hour it runs (e.g. `9 AM`) |
| Alert Email | Where reset errors/failures get emailed |
| Donation defaults | Pre-filled fields for the donation dialog |

## Troubleshooting

- **Menu doesn't appear** — reload the spreadsheet; `onOpen` only runs on open.
- **Reset didn't run** — check the trigger exists (Apps Script → Triggers). Re-run `setupResetTrigger`. Failures email the Alert Email, so silence + no reset usually means the trigger was never installed.
- **A day tab was renamed** — the reset needs both `Monday` and `Monday (Reset)` (etc.) to exist with exact names.
