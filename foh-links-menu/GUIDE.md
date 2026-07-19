# FOH Links Menu — Operator Guide

**This gives your front-of-house team one-click access to every link they need — training docs, forms, schedules — right from a menu in the Sheet, and it wipes the week's day tabs back to clean templates every week on its own.** No more hunting for URLs or resetting sheets by hand.

## Why it matters
The links a shift leader needs are scattered across bookmarks, texts, and memory, and the daily tabs have to be cleared and reset every week. Small annoyances, but they add up to wasted time and links nobody can find. This puts every link one click away and makes the weekly reset automatic.

## What it does
- **One-click links** — every link, grouped by category, in an "FOH Links" menu and a searchable sidebar.
- **Pin the important ones** — the sidebar lets anyone pin, add, edit, or delete links.
- **Automatic weekly reset** — on the day and hour you set, it copies your clean templates back over the week's day tabs.
- **Donation dialog** — a pre-filled donation-request email, with defaults from Settings.

## How it works (the plain version)
It's a Google Sheets add-on — a menu and a sidebar, no separate app.
- **Links live on the `Links` tab** — one row each: Category, Name, URL. Add a row, reload, and it shows up in the menu.
- **Behavior lives on the `Settings` tab** — reset day/hour, alert email, donation defaults.
- **Nothing you'd need is in the code** — it's all those two tabs.

## Everyday tasks
- **Add or change a link:** add a row on the `Links` tab (Category | Name | URL), or use the sidebar's add/edit. Reload to see it.
- **Change the reset day/time:** the `Settings` tab.
- **Test the reset before trusting it:** run `testResetMonday` from the menu to preview the template copy.
- **Turn the weekly reset on:** run `setupResetTrigger` once (safe to re-run — it clears the old one first).

## When something looks broken
- **The menu isn't there.** Reload the spreadsheet — the menu is added on open.
- **The weekly reset didn't run.** The trigger probably isn't installed — re-run `setupResetTrigger`. Failures email the Alert Email, so silence *and* no reset usually means the trigger was never set up.
- **The reset copied the wrong thing (or nothing).** It needs both the day tab and its `(Reset)` template to exist with exact names — e.g. `Monday` and `Monday (Reset)`. A renamed tab breaks the match.

## The one rule
**Everything is the `Links` tab and the `Settings` tab — no code.** Add links, set the reset, change the alert email, all right there. `runInitialSetup` is safe to re-run and never overwrites your data.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

### Two tabs run the whole thing
`Links` (Category | Name | URL, plus a Pin column the sidebar manages) drives the menu and the sidebar; `Settings` drives the reset schedule, the alert email, and the donation defaults. `runInitialSetup` creates both with example rows and is idempotent — re-running rebuilds anything missing without touching what's there. (There's also a hidden `DonationLog` tab where the donation dialog records each request.)

### How the menu gets built
`onOpen` runs when the spreadsheet opens: it reads the `Links` tab, groups rows by Category, and builds the **FOH Links** menu on the fly — which is exactly why the menu is missing until you reload after setup. The sidebar (`Sidebar.html`) is the richer version: search, add/edit/delete, and `togglePin(row)` to pin a link to the top.

### The weekly reset, mechanically
Each day tab (`Monday`…`Saturday`) has a matching `Monday (Reset)`…`Saturday (Reset)` **template**. `resetSingleDay_(ss, day)` copies the template over the live day tab, wiping that day's names back to a clean slate. On the configured day/hour a trigger (`resetWeeklySheets`) does all of them; `manualReset` does it on demand with a confirm, and `testResetMonday` previews just Monday so you can eyeball the copy before trusting the automation. `setupResetTrigger` removes any existing reset trigger before installing a new one (safe to re-run); `deleteResetTriggers` removes it. A failed run emails the Alert Email (falling back to the sheet owner) — so a silent failure almost always means the trigger was never installed.

### Why exact tab names matter
The reset matches `Monday` to `Monday (Reset)` **by name**. Rename either — or delete a `(Reset)` template — and the pair no longer matches, and that day silently stops resetting. It's the one thing that quietly breaks this tool.

### It's a menu add-on, not a web app
No deployment or `/exec` URL — the "install" is pasting the four numbered script files and two HTML files, running `runInitialSetup`, and reloading so `onOpen` builds the menu. (If `ADOPTION_SHEET_ID` is set in Script Properties, it also pings the adoption tracker once a day; no property means it's a silent no-op.)
