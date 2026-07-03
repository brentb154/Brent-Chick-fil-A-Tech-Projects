# Schedule Counter — Setup Guide

This guide takes you from zero to a working labor-scheduling tool in about **20 minutes**. No coding experience needed — you copy files, set one setting, and run two functions. Follow the steps **in order**.

Schedule Counter uploads your published schedule, weights it against sales curves, takes weekly snapshots, and tracks productivity. It lives **inside one Google Sheet** — there are no IDs, URLs, or passwords to paste.

---

## Before You Start

- A Google account that can create Google Sheets (your CFA Workspace account is ideal).
- The project files (all the `.gs` and `.html` files in this folder).
- ~20 minutes.

> ⚠️ **The single most important rule:** Do **not** create or rename any tabs by hand. The setup function builds every tab with the exact names and headers the tool needs. If you make a tab called "Config" instead of `config`, or rename one later, the tool silently stops working.

---

## Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and click **Blank spreadsheet**.
2. Rename it something clear, like **"Schedule Counter — [Your Store]"**.
3. Leave the default `Sheet1` tab alone. Don't add any other tabs — Step 4 does that.

---

## Step 2 — Open the Apps Script Editor

1. In your new Sheet, click **Extensions ▸ Apps Script**. A code editor opens in a new tab.
2. This script is now **bound** to your Sheet — that's why you never paste a Sheet ID.

---

## Step 3 — Paste the Files

For each file in this project, create a matching file in the editor and paste the contents:

- **`.gs` files** (like `Code.gs`): use the **+ ▸ Script** button, name it to match, paste.
- **`.html` files** (like `Index.html`, `JavaScript.html`, `Stylesheet.html`): use **+ ▸ HTML**, name it to match (no `.html` extension in the name), paste.

Delete the empty starter `myFunction` if it's in the way. Click **💾 Save** when done.

> File names must match **exactly**, including capitalization (`Index`, not `index`) — the app loads them by name.

---

## Step 4 — Set the One Required Setting

The tool emails you if a scheduled background job ever fails. Tell it where to send that:

1. In the Apps Script editor, click the **⚙️ Project Settings** (gear, left sidebar).
2. Scroll to **Script Properties ▸ Add script property**.
3. Property: `ALERT_EMAIL`  ·  Value: your email address.
4. Click **Save script properties**.

> Skipping this won't break the app, but you won't get alerted if a weekly job fails. Set it.

---

## Step 5 — Build the Tabs

1. At the top of the editor, open the function dropdown and choose **`initSheets`**.
2. Click **▶ Run**.
3. The first time, Google asks for permission:
   - **Review permissions** ▸ pick your account
   - "Google hasn't verified this app" ▸ **Advanced ▸ Go to (project)** ▸ **Allow**
4. Open the **Execution log** at the bottom. You should see `initSheets complete. All tabs verified.`

This creates every tab — `config`, `sales_curves`, `sales_history`, `productivity_tracker`, `app_cache`, and the snapshot tabs — and seeds default settings and sales curves.

`initSheets` is **safe to re-run** — it skips tabs that already exist and never deletes data.

---

## Step 6 — Install the Background Jobs

This is a **separate** step from Step 5 (the old README wrongly said setup did both — it doesn't).

1. In the function dropdown, choose **`installTriggers`**.
2. Click **▶ Run**.

This installs three time-driven jobs:
- **Sunday 10 PM** — schedule-upload reminder
- **Monday 6 AM** — the weekly processing pipeline
- **Thursday 7 PM** — schedule reminder

To reinstall later, run `installTriggers` again. To remove them all, run `deleteTriggers`.

---

## Step 7 — Deploy as a Web App

This gives your team a link to open the tool.

1. Click **Deploy ▸ New deployment** (top-right).
2. Click the gear ⚙️ next to "Select type" ▸ **Web app**.
3. Set:
   - **Description:** `Schedule Counter` (anything)
   - **Execute as:** **Me**
   - **Who has access:** **Anyone within [your domain]** ← keeps it to your CFA Workspace
4. Click **Deploy**, authorize if asked, and **copy the Web app URL**.

> **Updating later:** after any code change, click **Deploy ▸ Manage deployments ▸ ✏️ Edit ▸ Version: New version ▸ Deploy**. The URL stays the same. (Editing code alone does nothing until you redeploy.)

---

## Step 8 — Open It and Verify

1. Open the **Web app URL** from Step 7.
2. Confirm the tool loads (no blank white screen).
3. Open the **config** tab in your Sheet and set your store name and email recipients.

✅ If the page loads and the `config` tab is full of settings, you're done.

---

## Configure for Your Store

Everything operator-editable lives in the **`config`** tab — store name, email recipients, thresholds. Edit it in the Sheet; the code reads it at runtime. You never edit code to change settings.

---

## Troubleshooting

**Blank white screen when opening the web app**
The browser libraries must be version-pinned (they are in this copy). If you ever paste an older `Index.html`, make sure the React and Babel `<script>` tags have specific version numbers (`react@18.3.1`, `@babel/standalone@7.26.4`) — an unpinned Babel floats to v8 and white-screens the page. Hard-refresh (Ctrl/Cmd+Shift+R) after redeploying.

**"Tab "X" not found. Run initSheets() first."**
You deployed or used the app before running `initSheets` (Step 5), or a tab got renamed/deleted. Run `initSheets` from the editor and try again.

**"app_cache tab not found" / "sales_curves is empty"**
Same fix — run `initSheets` (Step 5). It seeds these.

**The weekly numbers never update**
The triggers aren't installed. Run `installTriggers` (Step 6). Confirm under **Triggers** (clock icon, left sidebar) that three triggers exist.

**"You do not have permission to call …"**
Re-authorize: run `initSheets` once from the editor and click **Allow**.

**A background job failed**
You'll get an email at your `ALERT_EMAIL` (Step 4). Open the Apps Script **Executions** log to see which job and why.

---

## Quick Reference

| What | Where |
|------|-------|
| Failure-alert email | Script Property `ALERT_EMAIL` (Step 4) |
| Build all tabs | run `initSheets` (Step 5) — safe to re-run |
| Install weekly jobs | run `installTriggers` (Step 6) |
| Remove weekly jobs | run `deleteTriggers` |
| Web app URL | Deploy ▸ Manage deployments |
| Operator settings | the `config` tab in the Sheet |

> **Golden rules:** never hand-create or rename tabs (let `initSheets` do it), run `initSheets` **and** `installTriggers` (two steps), and redeploy a **New version** after any code change.

---
Built with care for Cockrell Hill Chick-fil-A
