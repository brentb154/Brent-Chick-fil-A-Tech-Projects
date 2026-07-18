# Catering Quote Tool — Operator Guide

**This tool builds, sends, and tracks catering quotes for both restaurants — no spreadsheets to wrestle with, no math to redo.** You fill out a short form, it makes a clean PDF, emails or prints it, and remembers everything for next time.

## Why it matters
Catering quotes used to live in scattered spreadsheets and people's heads. This puts every quote in one place with consistent pricing, a professional PDF, and a paper trail for POs, tax-exempt forms, and delivery days — so anyone on the team can pick it up and look sharp in front of a customer.

## What it does
- **Quotes in a couple of minutes.** Pick items off your menu; it prices pickup vs. delivery automatically.
- **Two restaurants, one tool.** A Restaurant switch at the top of the form flips between Cockrell Hill and DBU — menu, cheat sheet, and pickup address all follow.
- **One-click PDF.** Email it to the customer or print it.
- **Remembers customers.** Past organizations and contacts autofill; "Reorder" copies a previous order.
- **Tracks tax-exempt forms.** See who has a form on file, and send guests a link to upload one.
- **Keeps prices honest.** A quarterly check makes you confirm a few menu prices against the POS.

## How it works (the plain version)
Think of it as a form sitting on top of a Google Sheet.
- **The web page** is where the team works — make quotes, view history, edit menus, change settings.
- **The Google Sheet** is the filing cabinet behind it. Every quote, menu item, and setting lives on a tab there. You rarely open it — but when you want to bulk-edit the menu or change a setting, that's where it is.
- **The team link** ends in `?view=app`. Bookmark that one. The plain link is just a harmless landing page, and guests get a separate upload link.

## Everyday tasks
- **New quote:** pick the restaurant → Pickup or Delivery → add the customer, items, and date → **Save & Email** (or Save, then Print).
- **Repeat customer:** start typing their name and contacts autofill — or open their last quote and hit **Reorder**.
- **Add or change a menu item:** open the **Menu Catalog** (in Settings), pick the store, edit inline. For big changes, paste into that store's menu tab of the Sheet — by default `CH Menu` for Cockrell Hill and `DBU Menu` for DBU. (The tabs are labeled by store on purpose, so nobody edits the wrong restaurant's prices. Different restaurants? You pick the labels — see Go deeper.)
- **Tax-exempt customer:** mark the quote tax-exempt, then **Look it up** to check the registry or **Request from Guest** to email them an upload link.
- **Change a price, tax rate, or email wording:** the **Settings** tab.

## When something looks broken
Most "breakage" isn't the tool — it's usually one of these:
- **Stuck on "Loading…" or a blank page.** You probably pasted a new version but didn't redeploy. In Apps Script: **Deploy → Manage deployments → Edit → New version.** The link always serves the last *deployed* version, not the latest save.
- **"It just takes me to GitHub" / shows a plain landing page.** The wrong file is in the `Index` slot. `App.html` goes in `Index` — not the repo's `index.html`.
- **It asks you to verify 3 prices and won't let you in.** That's the quarterly price check, on purpose. Enter the current POS prices for those items and you're good for another 90 days.
- **A quote didn't email.** Check the customer's email address on the quote, and that you're signed into the right Google account.
- **The calendar event didn't show up.** The calendar is best-effort — re-save the quote or add it by hand. It never blocks a quote from saving.
- **A guest can't reach the upload page.** The deployment has to be set to **Anyone** access, not "Anyone within your domain."

Google does occasionally change Apps Script under the hood, but real platform breakage is rare — and redeploying (which you've been doing) is exactly the right habit. When in doubt: **re-run `initializeSheet` once** (it's safe to repeat and never deletes data), **then deploy a new version.**

## The one rule
**Settings and menus live in the Sheet; the code doesn't.** Anyone can keep this running by editing tabs — no coding required. Just remember: after any code change, deploy a **new version**, or the link keeps serving the old one.

---

## Go deeper
*The 1,000-foot view — how the machine actually works, for whoever maintains it next. You don't need any of this to use the tool day to day.*

**It's all in the Google Sheet.** Every tab is a table the code reads and writes:
- **`Quotes`** — one row per quote (~23 columns). The line items are stored as JSON text in a single cell; everything else is its own column: customer, totals, tax, PO, event date/time, and `Location Name` (how a saved quote remembers which restaurant it was for).
- **`CH Menu` / `DBU Menu`** — Category | Item Name | Pickup Price | Delivery Price. `N/A` in delivery price hides an item from delivery quotes.
- **`CH Off-Menu` / `DBU Off-Menu`** — Item | Base Price. The delivery price you see is computed live (base price × the markup in Settings), never stored.
- **`Tax_Exempt_Registry`** — one row per organization: a status and a link to the form PDF in Drive.
- **`Settings`** — every tunable value (tax rate, store details, email wording, warning thresholds, the price-check clock). The code reads these at runtime, so changing a setting changes behavior with no code edit.
- **Hidden tabs** (`Quotes_Archive`, `Quote_Revisions`, `Confirmations_Sent`, `PO_Alerts_Sent`, `Tax_Forms`, `Tax_Form_Uploads`) are logs and history — they keep the visible tabs clean.

**What happens when you save a quote.** The form gathers your inputs, checks the required fields, and asks the server for the next quote ID — a running counter that never repeats, with the contact's initials tacked on. The server writes the row and, if the quote has a date and time, creates a Google Calendar event. If you're *editing* an existing quote, the old version is copied to `Quote_Revisions` first (nothing is ever truly overwritten), and the calendar event is updated instead of duplicated. "Save & Email" builds the PDF as HTML, converts it, and sends it with your Settings email template; "Print" opens the same HTML in a print window.

**How the two restaurants stay separate.** The Restaurant toggle sets which store you're on. Both menus load up front, so switching is instant — it just re-points the item picker and re-prices the lines. The saved quote records the store's name, and the PDF looks up that store's address and phone from Settings by matching that name. The price check runs per store, each on its own 90-day clock, and an empty or unconfigured store never triggers the lock.

**The automation (scheduled jobs you install once from Settings).**
- **Daily at 3pm** — sends day-before confirmation emails, sweeps for orders whose event is coming up without a PO and emails a digest, and once a year emails a tax-exempt review. Each piece has its own on/off switch in Settings, and a failure emails an alert instead of dying silently.
- **Daily at 9am** — the follow-up reminder for quotes that haven't been acted on.
- **Nightly (optional)** — `cleanOldQuotes` moves quotes older than the "Archive After Days" setting into the hidden archive. Nothing is deleted.

**The tax-exempt flow, end to end.** On a tax-exempt quote you either look the organization up in the registry or hit "Request from Guest," which emails them a link to a public upload page (no login). Their PDF lands in a Drive folder and shows up as a *pending* item on the Tax Forms tab. A team member confirms it into the registry under the correct organization name — you pick the name, so a slightly different spelling never silently creates a duplicate. The Guest Upload Link on that tab is that same public page, copyable to hand out directly.

**The deploy model (why "new version" matters).** There's no automatic sync. The code lives in a repo; the live app only changes when someone pastes the files into Apps Script and publishes a **new version**. Three files: `Code.gs` (the server), `App.html` → the Apps Script file named **`Index`** (the team UI), and `TaxForm.html` → **`TaxForm`** (the guest page). The URL routes by a `?view=` tag: `?view=app` is the team, `?view=taxform` is the guest page, a bare URL is a harmless landing page. Access is set to **Anyone** so outside guests can reach the upload page — the team tool isn't login-walled, just link-obscured, so don't post the `?view=app` link publicly.

**Three things that will bite you if you don't know them.**
- **Dates and the Sheet fight each other.** Sheets auto-converts anything that looks like a date into a real date cell, and the page↔server bridge throws away an entire response if it contains one. The code stores dates as plain `yyyy-MM-dd` text on purpose — if you add a server function that returns sheet values, keep dates as strings or the page will silently hang on "Loading…".
- **The link serves the last *published* version, not your latest paste.** Always publish a new version after editing.
- **The tab names are labeled by store, and you choose the labels.** They default to `CH` (Cockrell Hill) and `DBU` — so nobody edits the wrong restaurant's prices. Running different restaurants? Set `Store 1 Tab Prefix` and `Store 2 Tab Prefix` in the Settings tab to your own short codes, then re-run `initializeSheet` — it renames the tabs in place and keeps every price. (First-time upgrade also renames the old `Menu`/`Off_Menu` tabs for you.)
