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
- **You configure everything in the Settings sidebar** (from the "Waste Summary" menu) — recipients, schedule, format.

## Everyday tasks
- **Nothing, most weeks** — it sends on its own; the team just keeps logging on the form.
- **Change recipients, day, hour, or format:** open the "Waste Summary" menu → Settings.
- **Send a test / check it's working:** "Waste Summary" menu → **Send Test Email Now**.

## When something looks broken
- **The menu isn't there.** Reload the spreadsheet — the "Waste Summary" menu is added on open.
- **The summary didn't send.** The daily trigger checks the time — make sure a trigger is installed (Apps Script → Triggers) and that the send day/hour in Settings is what you expect. `deleteTrigger()` clears them for a reset.
- **The email's empty or wrong.** Check that the form responses are landing in this Sheet, and that the send day lines up with when the week's data is complete.
- **Wrong people getting it (or not).** Recipients live in the Settings sidebar.

## The one rule
**Recipients, schedule, and format are all in the Settings sidebar — not the code.** Set them once; the daily trigger does the rest.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

### How the send decides to fire
There's no weekly trigger. Instead `onDailyTrigger` runs every day and walks a short checklist: (1) is today's day-of-week the configured **Send Day**? If not, exit. (2) does **Last Sent** already equal today? If so, exit — this is the duplicate-send guard. (3) compute last week's date range, (4) filter the form rows into it, (5) sum them, (6) send the email, then stamp **Last Sent**. A cheap daily check that no-ops most days is easier to reason about than an exact weekly schedule, and the Last-Sent stamp means even if the trigger fires twice, only one email goes out.

### Where the data comes from
It reads the Google Form response rows already in the Sheet and filters them to last week's range — it doesn't own the form, it summarizes it. So the whole thing depends on the form staying connected to this Sheet; if responses stop landing here, the summary goes empty.

### Config lives in a config sheet, edited from the sidebar
`Send Day`, `Send Time`, `Last Sent`, recipients, and format are stored as config rows (`getConfigValue` / `ensureConfigRow`) and edited through the "Waste Summary → Settings" sidebar. Notably, **saving settings re-installs the trigger** to match the new time — so changing the hour in the sidebar actually reschedules the job, you don't touch triggers by hand. There's also a **"Send Test Email Now"** menu item to fire the summary immediately without waiting for the send day.

### It's a menu/sidebar add-on, not a web app
No `/exec` URL — paste `Code.gs` and the sidebar HTML, reload so the "Waste Summary" menu appears, open Settings (which installs the daily trigger when you save), and you're done. `deleteTrigger()` clears the schedule for a reset.
