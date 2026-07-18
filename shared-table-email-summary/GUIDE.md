# Shared Table Email Summary — Operator Guide

**This turns your Shared Table waste log into a weekly summary email — automatically — so leadership sees what's being tossed without anyone building a report.** The team logs on a form; the summary sends itself on the day and time you pick.

## Why it matters
Waste data piles up in a form-response sheet that nobody opens. The numbers only matter if someone sees them regularly. This reads the week's responses and emails a clean summary on a schedule, so waste stays visible and actionable instead of buried in a tab.

## What it does
- **Weekly summary email** — totals the week's Shared Table form responses and emails them out.
- **You choose who, when, and how** — recipients, the send day and hour, and the format, all from a Settings sidebar.
- **Runs itself** — a daily check fires the email on the day/time you set; no one has to remember.

## How it works (the plain version)
It's a small Google Sheets add-on sitting on your form-response sheet.
- **The team logs waste on the Shared Table Google Form.** Responses land in the Sheet.
- **This reads those responses** and builds the summary.
- **A daily trigger checks the clock** — when it's your configured send day and hour, it emails the summary.
- **You configure everything in the Settings sidebar** (from the "Shared Table" menu) — recipients, schedule, format.

## Everyday tasks
- **Nothing, most weeks** — it sends on its own; the team just keeps logging on the form.
- **Change recipients, day, hour, or format:** open the "Shared Table" menu → Settings.
- **Send a test / check it's working:** use the menu options to preview or trigger a run.

## When something looks broken
- **The menu isn't there.** Reload the spreadsheet — the "Shared Table" menu is added on open.
- **The summary didn't send.** The daily trigger checks the time — make sure a trigger is installed (Apps Script → Triggers) and that the send day/hour in Settings is what you expect. `deleteTrigger()` clears them for a reset.
- **The email's empty or wrong.** Check that the form responses are landing in this Sheet, and that the send day lines up with when the week's data is complete.
- **Wrong people getting it (or not).** Recipients live in the Settings sidebar.

## The one rule
**Recipients, schedule, and format are all in the Settings sidebar — not the code.** Set them once; the daily trigger does the rest.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

**How the send decides to fire.** There's no weekly trigger — instead a **daily** trigger runs and asks "is today the configured day, and is it past the configured hour?" If yes, it builds and sends the summary; if not, it does nothing. That's a robust pattern: a cheap daily check is easier to reason about than juggling exact weekly schedules, and it survives daylight-saving shifts and missed runs better.

**Where the data comes from.** The tool reads the Google Form response rows already in the Sheet — it doesn't own the form, it summarizes it. So the whole thing depends on the form staying connected to this Sheet.

**Config lives in the sidebar.** Recipients, send day/hour, and format are settings the sidebar reads and writes; `Code.gs` holds the email generation and trigger management. Nothing operator-facing requires editing code.

**It's a menu/sidebar add-on, not a web app.** No `/exec` URL — paste `Code.gs` and the sidebar HTML, reload so the menu appears, set your Settings, and make sure the daily trigger is installed.
