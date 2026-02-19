# Chick-fil-A Catering Quote Generator

A Google Apps Script web application for creating, managing, emailing, and printing professional catering quotes. Uses a Google Sheet as both the database and settings layer.

---

## Setup Instructions

### Step 1: Create the Google Sheet
Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet. Name it **"CFA Catering Quotes"**.

### Step 2: Open the Apps Script Editor
In your Google Sheet: **Extensions → Apps Script**.

### Step 3: Add the Code Files

**Code.gs:** Replace all code in the existing `Code.gs` with the provided `Code.gs` file.

**Index.html:** Click **+** next to Files → select **HTML** → name it `Index` → paste the provided `Index.html` contents.

### Step 4: Initialize the Spreadsheet
Select `initializeSheet` from the function dropdown → click **▶ Run** → authorize when prompted.

Your sheet will have four tabs: **Settings**, **Menu**, **Quotes**, **Quote_Sequence**.

### Step 5: Deploy as a Web App
**Deploy → New deployment** → Web app → Execute as "Me" → Access "Anyone within [org]" → Deploy → copy the URL.

### Step 6: Set Up Nightly Cleanup
Triggers (clock icon) → **+ Add Trigger** → `cleanOldQuotes` → Time-driven → Day timer → Midnight.

### Step 7: Add Menu Items

The Menu tab now has **4 columns**: Category | Item Name | Pickup Price | Delivery Price.

**Example entries:**

| Category | Item Name | Pickup Price | Delivery Price |
|---|---|---|---|
| Trays | Chick-fil-A Nuggets Tray - Small | $32.00 | $38.00 |
| Trays | Chick-fil-A Nuggets Tray - Large | $58.00 | $66.00 |
| Box Meals | Chick-fil-A Deluxe Meal | $8.99 | $10.49 |
| Box Meals | Spicy Deluxe Meal | $8.99 | $10.49 |
| Hot Entrees | Chick-fil-A Nuggets - 12ct | $4.65 | $5.50 |
| Sides | Mac & Cheese | $3.35 | $3.95 |
| Sides | Fruit Cup | $3.85 | $4.45 |
| Drinks | Gallon Freshly-Brewed Iced Tea | $11.00 | $13.00 |
| Desserts | Chocolate Chunk Cookie | $1.65 | $1.95 |

Categories are **freeform** — just type whatever category name you want. Items with the same category text automatically group together in the dropdown. No separate category list to maintain.

### Step 8: Configure Settings
Fill in store names, addresses, phone numbers, contact name, tax rate, logo, and email template. All editable in the app or directly in the sheet.

---

## Features

### Searchable Item Picker
When adding line items, the dropdown is a **searchable, categorized selector**:
- Click or focus the field to see all items grouped by category
- Start typing to filter — matches against both item name and category
- Use arrow keys + Enter for keyboard navigation
- Category headers (e.g., "BOX MEALS", "SIDES") are sticky as you scroll
- Each option shows the current price (Pickup or Delivery) on the right
- "Custom Item" option at the bottom for freeform entries

### Email Integration
- **Save & Email** button on new quotes — one click saves + generates PDF + emails it
- **Email Quote** button in History — send any saved quote to any email
- Email subject and body are fully customizable templates in Settings
- PDF is generated server-side and attached automatically
- Optional BCC to get a copy of every sent quote

### Quote Management
- Sequential IDs (Q-2026-0001, Q-2026-0002, …) that never repeat
- Prices frozen at creation time — menu changes don't affect old quotes
- Edit & Reuse creates a new quote; originals are never modified
- Auto-cleanup deletes quotes older than 30 days

---

## Spreadsheet Architecture

### Menu Tab (4 columns)
| Column | Description |
|---|---|
| **Category** | Freeform grouping label (e.g., "Box Meals", "Sides", "Drinks") |
| **Item Name** | The menu item name shown in the picker |
| **Pickup Price** | Price for pickup orders |
| **Delivery Price** | Price for delivery orders |

### Settings Tab
| Label | Description |
|---|---|
| Store Name (Active) | Which location appears on quotes |
| Location 1/2 Name/Address/Phone | Store details (per-location phone numbers) |
| Quote Contact Name | Shown on PDF "questions?" line |
| Default Tax Rate (%) | Pre-filled on new quotes |
| Logo (Base64) | Uploaded via app |
| Email Subject | Template with {{placeholders}} |
| Email Body | Template with {{placeholders}} |
| BCC Email | Optional copy recipient |

### Email Placeholders
`{{customer}}` `{{contact}}` `{{location}}` `{{phone}}` `{{quoteId}}` `{{total}}` `{{date}}`

---

## Customization

**For other restaurants:** Change location info, upload your logo, fill in your menu, customize the email template. Categories adapt to whatever you type — no code changes needed.

**Adding more locations:** Add "Location 3 Name/Address/Phone" rows to Settings and update the code to read them.

---

## Troubleshooting

| Issue | Fix |
|---|---|
| Pop-up blocked on PDF | Allow pop-ups for the Apps Script domain |
| Email not sending | Check daily quota; re-authorize if needed |
| Items not showing in picker | Ensure Menu tab has all 4 columns filled starting at row 2 |
| Categories not grouping | Make sure the Category text matches exactly (case-sensitive) |
| Old menu had 3 columns | Add a "Category" column A and shift existing data right |
