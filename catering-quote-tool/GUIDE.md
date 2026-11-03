# Catering Quote Tool — Operator Guide

**This tool builds, sends, and tracks catering quotes for both restaurants — no spreadsheets to wrestle with, no math to redo.** You fill out a short form, it makes a clean PDF, emails or prints it, and remembers everything for next time.

## Why it matters
Catering quotes used to live in scattered spreadsheets and people's heads. This puts every quote in one place with consistent pricing, a professional PDF, and a paper trail for POs, tax-exempt forms, and delivery days — so anyone on the team can pick it up and look sharp in front of a customer.

## What it does
- **Quotes in a couple of minutes.** Pick items off your menu; it prices pickup vs. delivery automatically.
- **Two restaurants, one tool.** A Restaurant switch at the top of the form flips between your two stores — menu, cheat sheet, and pickup address all follow.
- **One-click PDF.** Email it to the customer or print it.
- **Remembers customers.** Past organizations and contacts autofill; "Reorder" copies a previous order.
- **Tracks tax-exempt forms.** See who has a form on file, and send guests a link to upload one.
- **Keeps prices honest.** A quarterly check makes you confirm a few menu prices against the POS.
- **Tells you what to pull.** Pick a date and the **Supplies** tab turns that day's orders into raw counts — filets, nuggets, buns, mac bags — split by restaurant, with a rough case count.
- **Hands off the far ones.** Mark an order as passed to a closer store; it drops out of your prep numbers and you get a draft email to send them.
- **Names the driver.** Assign a driver and their name shows up right on the Google Calendar event.

## How it works (the plain version)
Think of it as a form sitting on top of a Google Sheet.
- **The web page** is where the team works — make quotes, view history, edit menus, change settings.
- **The Google Sheet** is the filing cabinet behind it. Every quote, menu item, and setting lives on a tab there. You rarely open it — but when you want to bulk-edit the menu or change a setting, that's where it is.
- **The team link** ends in `?view=app`. Bookmark that one. The plain link is just a harmless landing page, and guests get a separate upload link.

## Everyday tasks
- **New quote:** pick the restaurant → Pickup or Delivery → add the customer, items, and date → **Save & Email** (or Save, then Print).
- **Repeat customer:** start typing their name and contacts autofill — or open their last quote and hit **Reorder**.
- **Add or change a menu item:** open the **Menu Catalog** (in Settings), pick the store, edit inline. For big changes, paste into that store's menu tab of the Sheet — by default `CH Menu` for Main Street and `DBU Menu` for DBU. (The tabs are labeled by store on purpose, so nobody edits the wrong restaurant's prices. Different restaurants? You pick the labels — see Go deeper.)
- **Tax-exempt customer:** mark the quote tax-exempt, then **Look it up** to check the registry or **Request from Guest** to email them an upload link.
- **Planning a prep or order day:** the **Supplies** tab → pick the date → read the totals per restaurant. Red rows are raw proteins to pull; yellow rows are mac & cheese bags to thaw the day before.
- **An order is too far to run:** open it → **Hand off** → pick the partner store → **✉️ Hand-off email** to send them the details (you review it before it goes).
- **Assign who's driving:** open the order (it needs a PO first) → **Assign Driver**.
- **Change a price, tax rate, or email wording:** the **Settings** tab.

## When something looks broken
Most "breakage" isn't the tool — it's usually one of these:
- **Stuck on "Loading…" or a blank page.** You probably pasted a new version but didn't redeploy. In Apps Script: **Deploy → Manage deployments → Edit → New version.** The link always serves the last *deployed* version, not the latest save.
- **"It just takes me to GitHub" / shows a plain landing page.** The wrong file is in the `App` slot. `App.html` goes in `App` — not the repo's `index.html`.
- **It asks you to verify 3 prices and won't let you in.** That's the quarterly price check, on purpose. Enter the current POS prices for those items and you're good for another 90 days.
- **A quote didn't email.** Check the customer's email address on the quote, and that you're signed into the right Google account.
- **The calendar event didn't show up.** The calendar is best-effort — re-save the quote or add it by hand. It never blocks a quote from saving.
- **A guest can't reach the upload page.** The deployment has to be set to **Anyone** access, not "Anyone within your domain."

Google does occasionally change Apps Script under the hood, but real platform breakage is rare — and redeploying (which you've been doing) is exactly the right habit. When in doubt: **re-run `initializeSheet` once** (it's safe to repeat and never deletes data), **then deploy a new version.**

## The one rule
**Settings and menus live in the Sheet; the code doesn't.** Anyone can keep this running by editing tabs — no coding required. Just remember: after any code change, deploy a **new version**, or the link keeps serving the old one.

---

## Go deeper
*How the machine actually works, for whoever maintains it next. You don't need any of this to use the tool day to day — it's here so the next person can change something without spelunking through the code first.*

### The data model — it's all in the Google Sheet
There's no separate database. Every tab is a table the code reads and writes.
- **`Quotes`** — one row per quote, **26 columns**. Most fields get their own column: customer, contact, phone, email, order type, delivery address, subtotal, tax rate used, tax amount, total, tax-exempt flag, `Location Name`, PO number, event date, event time, calendar event id, last-modified, the order-level discount value/type, and quote notes. **The line items are stored as JSON text in a single cell** — an array of `{description, quantity, price, discount}` — so one quote is always exactly one row no matter how many items it has. `Location Name` is how a saved quote remembers which restaurant it was for; the PDF later maps that name back to the store's address and phone — and it's also what the Supplies tab groups by.
  The last three columns were added over time and are all **blank-safe by design** — an old row with an empty cell reads as a sensible default, so no backfill was ever needed: **24 `Status`** (blank = Active), **25 `Handoff Store`** (blank = not handed off), **26 `Driver`** (blank = unassigned). None of them are written by the main save path — `saveQuote` still writes a fixed 23-column block, and each of these has its own targeted single-cell writer (`setQuoteStatus`, `setHandoffStore`, `assignDriver`). That's deliberate: it keeps the save path from having to know about every column that gets bolted on later.
- **`Recipe_BOM`** — Menu Item | Component | Qty per Unit. One row per component. The Supplies engine matches on the **full item name, lowercased** — no fuzzy matching, on purpose, so a wrong match can never quietly inflate an order.
- **`Case_Conversions`** — Component | Per Case | Per Bag. Turns raw counts into "≈ 2.4 cases". Per Bag is optional; blank hides the bag figure. Note the cases/bags numbers are computed **at runtime**, not stored — change a conversion and every past date recalculates.
- **`Supplies_Reference`** — Section | Item | Detail. The editable notes shown on the Reference sub-tab; a blank Item makes a full-width note.
- **`Drivers`** — one name per row, team-editable.
- **`Partner Store`** — Store Location | Operator Name | Store Phone | Store Email | Catering POC. One row per store you might hand off to.
- **`<prefix> Menu`** (default `S1 Menu` / `S2 Menu`) — Category | Item Name | Pickup Price | Delivery Price. `N/A` or blank in a price column hides that item from that order type. Category groups items in the picker.
- **`<prefix> Off-Menu`** (default `S1 Off-Menu` / `S2 Off-Menu`) — Item | Base Price. The delivery price you see is computed live as base × (1 + markup%), rounded to the nearest tenth — never stored, so changing the markup reprices the whole list at once.
- **`Settings`** — every tunable value as a key/value pair, read at runtime, so a change takes effect with no redeploy.
- **`Quote_Sequence`** — a single running counter, so quote IDs never repeat even across edits.
- **`Tax_Exempt_Registry`** — one row per organization: a status and a link to the form PDF in Drive.
- **Hidden tabs** — `Quotes_Archive` (aged-out quotes), `Quote_Revisions` (prior versions of edited quotes), `Tax_Forms` (per-quote status), `Tax_Form_Uploads` (the guest pending queue), `Geocode_Cache` (address → lat/lng, so repeat addresses never re-hit the Maps quota), and one append-only log per automation so nothing ever double-sends: `Confirmations_Sent`, `PO_Alerts_Sent`, `PO_Warn_DayBefore_Sent`, `PO_Warn_DayOf_Sent`, `DISD_Weekly_Sent`. `Reminders_Sent` (visible) does the same job for follow-ups.
  The dedup pattern is worth knowing if you add an automation: each log is append-only, keyed on `quoteId` or `quoteId|eventDate`, and read into a lookup before the send loop. Keying on quote+date rather than quote alone is what lets a re-quote for a *new* date legitimately alert again while the same order never alerts twice.

### The life of a quote
1. **Draft.** The form validates the required fields and asks the server for the next ID — the running counter, with the contact's initials appended (`Q-2026-0042-KD`).
2. **Save.** The server writes the row. Line items are JSON-encoded into one cell; totals, discounts, and tax are computed and stored so history and the PDF never have to recompute.
3. **PDF.** The quote is built as an HTML document and converted to PDF — the *same* HTML the on-screen preview and the print window use, so what you see is what sends.
4. **Email.** "Save & Email" sends the PDF with your Settings email template (subject/body with `{{placeholders}}`), optionally BCC'ing whoever you set.
5. **Calendar.** If the quote has a date and time, a color-coded Google Calendar event is created and its id is stored on the row — so a later edit updates that event instead of creating a duplicate.
6. **Edit.** Editing copies the outgoing version to `Quote_Revisions` first — nothing is ever truly overwritten — then writes the new values. You can view and restore any prior version from the quote's popup.
7. **Expire / archive.** Every PDF shows a "valid through" date (Settings-tunable). Quotes past the archive threshold are moved to the hidden archive by the nightly job — moved, never deleted.

### How pricing works
A line item's price comes from the selected store's menu, keyed by the current order type — pickup vs delivery — with `N/A`/blank hiding unavailable items from that mode. Off-menu items price live off base × markup. On top of line items you can apply per-line discounts and one order-level discount (percent or dollar). Tax uses the Settings default rate (editable per quote); marking a quote tax-exempt zeroes the tax and ties into the registry. Switching order type *or* store re-prices every line against the new context (with a confirm), and leaves any item that isn't on the new menu at its typed price.

### Two restaurants, in detail
- **Tab names are operator-chosen.** `Store 1/2 Tab Prefix` in Settings (default S1/S2) feed `<prefix> Menu` / `<prefix> Off-Menu`. `initializeSheet` renames the tabs *in place* when the prefix changes — it tracks the last-applied name in Script Properties — so a rename never loses data or leaves an orphaned empty tab.
- **Both menus preload.** On load the app pulls both stores' menus into memory, so switching restaurants (and editing/reordering a past quote into its store) is instant, no round-trip.
- **Fast lookups.** Menu items are indexed by name (a hash map, first-row-wins on duplicates), so repricing a whole quote or matching a reordered quote's items is one lookup per line instead of a scan of the whole menu.
- **The quote remembers its store** via `Location Name`; edit/reorder reads that back to switch stores; the PDF resolves it to the store's address and phone.

### The quarterly price check
On load, if it's been longer than `Price Check Interval (Days)` (default 90) since the active store's last verification, the app locks until the operator types the current POS prices for three menu items and they match. It's **per store** — separate `Last Price Verification` dates — and it **skips a store whose menu is empty**, so an unconfigured second store never wedges the lock. The three items are drawn from the categories listed in Settings (falling back to the whole menu if that's too narrow), always including a delivery-priced item; a typed price passes if it matches *any* menu row with that name, so a duplicate item name can't jam it.

### Tax-exempt tracking
A tax-exempt quote can look the org up in the registry or send the guest an upload link. The guest page is public (`?view=taxform`), no login; their PDF or photo lands in a Drive folder and shows up as a *pending* row on the Tax Forms tab. A human confirms it into the registry under the canonical org name — the app never auto-matches names, on purpose, so a slightly different spelling can't create a silent duplicate. The Guest Upload Link on that tab is that same public page, copyable to hand out directly. The In-Store QR Code card below it is the same page again with `&quote=IN-STORE` added. Print it for the manager stand so walk-in uploads are tagged `IN-STORE` for the monthly POS match. Access is **Anyone** so outside guests can reach it; the team tool is URL-gated (`?view=app`), not login-walled — so don't post the team link publicly.

### The automations
Each is installed separately — **`initializeSheet` creates no triggers at all** — each has its own on/off switch, and each is wrapped so a failure emails an alert instead of dying silently. Every installer deletes its own old trigger before creating a new one, so re-running never doubles up, and each has a matching `remove…Trigger()`.

| When | Handler | Does |
|---|---|---|
| Daily 3pm | `dailyCateringAutomation` | Day-before confirmations, missing-PO digest, and (once, at year-end) the tax-exempt review |
| Daily 2pm + 6am | `poWarningDayBefore` / `poWarningDayOf` | Urgent "still no PO" nudges for tomorrow's and today's orders |
| Daily 9am | `processReminders` | Follow-up on quotes that haven't been acted on |
| Saturday 12pm | `disdWeeklyReminder` | Weekly DISD recap — guest + total only, no line items |
| Nightly | `cleanOldQuotes` | Archives quotes past the age threshold (moved, never deleted) |

The urgent PO warnings are deliberately **separate from** the 7-day sweep and keep their own logs, so an order that already got the early heads-up still gets the day-before and day-of nudges.

### The Supplies Calc, in detail
`getSuppliesForDate(dateStr)` is the whole engine:
1. Pull every quote on that date whose status isn't `Inactive`.
2. Any quote with a non-blank `Handoff Store` is pulled out into a separate `handedOff` list — shown for visibility, **never counted**.
3. The rest are bucketed by `Location Name`, so each restaurant gets its own totals. The bucketing is list-driven off whatever store names it actually finds — nothing hardcodes "two stores", so a third would just appear.
4. Per bucket, each line item is matched against `Recipe_BOM` by lowercased full name and exploded into components × quantity. **No match = no guess** — it goes to that store's `noRecipe` list with its count intact.
5. Components are joined to `Case_Conversions` for the "≈ cases" figure, and tagged for highlighting.

The red/yellow highlighting is driven by a small **component → tag** map (`supplyTagMap_`), not by hardcoding rows in the UI: raw proteins get `protein` (red), mac bags get `thaw` (yellow). To flag something new, add it to that map — the client just reads the tag off each row.

**The gotcha:** because matching is exact, a menu rename silently moves an item into the ⚠️ no-recipe list rather than breaking anything. That list is the health check — if it's growing, `Recipe_BOM` has drifted from the menu.

**The other gotcha:** `initializeSheet` will *not* reseed a non-empty `Recipe_BOM` or `Case_Conversions`. New recipe data only lands via `resetRecipeBom()` / `resetCaseConversions()`, which overwrite the whole tab. That's a deliberate trade — it protects hand-edits from a routine setup re-run, at the cost of making recipe updates an explicit, destructive step.

### Hand-off and driver
- **Hand-off** is just column 25 plus a contacts tab. Setting it removes the order from that store's supplies totals immediately — that's the entire mechanism, and it's why the totals can be trusted the morning of.
- **The hand-off email is a draft, never auto-sent.** `buildHandoffEmailDraft` fills a Settings template and hands subject/body back to the same preview modal the quote email uses; a human edits and sends. Outbound mail to another operator is exactly the kind of thing that shouldn't fire on a checkbox. It sends plain (no quote PDF attached).
- **The distance nudge** reuses the geocoding already built for delivery-route suggestions: `geocodeAddress_` (cached in `Geocode_Cache`) + `haversineMiles_`. It's straight-line distance, not driving distance — a hint, not a routing engine — and it never blocks a save.
- **The driver goes into the calendar title**, and *how* matters. Titles are `<prefix> — <customer> — <date>`, and `updateCalendarEventPO` recovers the customer and date by splitting on that ` — `. So the driver is embedded **inside the customer segment** as `Smith Elementary-"Marcus"`, not as a fourth segment — the split still yields three parts and the driver rides along untouched when the PO changes. If you ever extend the title again, do the same thing: enrich a segment, don't add a delimiter.
- A normal quote edit rebuilds the title from form data, which has no driver on it — so `updateCalendarEvent` re-reads the existing title and carries the driver forward. Without that, editing an order would silently drop the driver.

### Gotchas and hard-won lessons
- **`google.script.run` nulls any return value that contains a Date object.** It once nulled the whole settings payload and hung the app on "Loading…". Every server function that returns sheet values keeps dates as `yyyy-MM-dd` **strings** — if you add one, do the same.
- **Sheets auto-converts date-like strings into date cells.** Writing a date-shaped setting uses `setNumberFormat('@')` (the `asText` flag on `updateSetting`) to stop that.
- **The `index.html` trap.** The repo's `index.html` is the GitHub Pages landing page, NOT the app. `App.html` is what goes in the Apps Script `App` file — paste the landing page there and the web app "takes you to GitHub."
- **The link serves the last *published* version**, not your latest paste — always publish a new version.
- **Config lives in the Sheet, on purpose**, so a non-technical person can run this forever without touching code.
- **Adding a Quotes column is a four-touch change**, and skipping one is the usual bug: widen the read cap in `getQuotes`, map the new index with a blank-safe default, add it to *both* the fresh-header list and the migration map in `initializeSheet`, and give it its own targeted writer. Don't widen `saveQuote`'s fixed 23-column write — that's what keeps old rows and new columns from fighting.
- **Anything appended to a calendar title has to survive `updateCalendarEventPO`'s ` — ` split.** Enrich an existing segment; never add a delimiter.
- **SMS was considered and dropped.** Apps Script has no native SMS, so it would mean a paid third party (Twilio) or unreliable carrier email gateways. Every notification here is email on purpose — if someone asks "can it text me", that's the answer and the reason.

### Deploy model
No auto-sync. Three files: `Code.gs` → `Code.gs`; `App.html` → the Apps Script file named **`App`**; `TaxForm.html` → **`TaxForm`**. Run `initializeSheet()` once, then **Deploy → Manage deployments → Edit → New version**, access **Anyone**. The `?view=` tag routes the one URL: `app` = team, `taxform` = guest, bare = landing.

**What `initializeSheet` does and doesn't do** — worth being precise, because "just re-run setup" is the standard fix for a lot of things:

*It is genuinely idempotent.* It creates any missing tab, adds Settings keys that don't exist yet (never overwriting ones that do), adds missing `Quotes` column headers in place, and leaves every row of data alone. Re-run it as often as you like.

*It does not:*
- **install any triggers** — every automation is its own `setup…Trigger()` call
- **reseed a non-empty `Recipe_BOM` / `Case_Conversions`** — that's `resetRecipeBom()` / `resetCaseConversions()`, which overwrite
- **fill `Drivers` or `Partner Store`** — it creates them empty for you to fill
- **redeploy anything** — the `/exec` URL still serves the last published version
