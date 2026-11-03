# H.E.A.R.D. Log — Operator Guide

**This is one fast place to log every guest complaint and how you recovered it — from a phone or tablet, in under a minute — and it automatically flags the guest who's called in three times this month.** Front-of-house leaders log; managers get the dashboard and the alerts.

## Why it matters
Complaints get handled in the moment and then forgotten, so nobody notices the patterns: the repeat caller, the daypart that generates most of the issues, the "free item" handed out more than it should be. Paper logs and one-off texts don't add up. This makes the record shared and searchable, so the trends are visible and the recovery conversation — Hear, Empathize, Apologize, Resolve, Delight — has a memory behind it.

## What it does
- **Log a complaint in seconds.** Guest, order, issue type, resolution, daypart, and which manager handled it.
- **Catch repeat complainers automatically.** Type a phone number and it checks the recent window (30 days), and badges the guest if there's a history.
- **Show the whole picture.** A dashboard of totals, flagged guests, and trends, plus a searchable, editable log.
- **Alert leadership.** An email when a guest crosses the repeat threshold you set.

## How it works (the plain version)
Everything lives in one Google Sheet, and your team opens it through a single web link:
- **The app** (the page leaders open) comes from that link. Put it on a tablet's home screen or bookmark it. It needs internet but no install.
- **The backend** is a Google Apps Script app attached to the sheet. It serves the app, checks the passcode, and does every read and write, so the sheet itself stays private.
- **The Google Sheet** is the database — one tab, with the columns created automatically the first time you submit.

## Everyday tasks
- **Log an issue:** open the form → fill it in → submit. Watch for the repeat-guest badge when you enter the phone.
- **Review trends:** the dashboard shows totals, flagged guests, and the log; you can search and edit past entries.
- **Change who gets alerted:** Settings in the app (settings passcode). The repeat threshold is a number at the top of `Code.gs`.
- **Change a passcode:** in the sheet, **👂 H.E.A.R.D. Setup → Start Setup**. It takes effect right away.

## When something looks broken
- **The app loads but won't save.** You changed the code and didn't publish a new version. Re-publish (**Deploy → Manage deployments → Edit → New version**).
- **"Incorrect passcode" for everyone.** Someone changed it with Start Setup. Run Start Setup again to see the current one.
- **Blank / white page.** The form loads its framework from the web — check the device has internet. (If you host your own copy and it white-screens after an update, it's almost always that front-end loading step, not your data.)
- **The repeat badge didn't show.** It keys off the phone number over the recent window (30 days) — it only flags once the phone is entered and there's prior history.
- **Alerts aren't arriving.** Check the alert email and threshold settings, and that the backend is authorized to send mail from the right Google account.

Real platform breakage is rare. For the backend, a new deployment version fixes most "it stopped saving" issues; for the form, it's almost always the link or the internet.

## The one rule
**Settings live in the sheet's script settings, not in the code.** Passcodes and alert emails are saved by Start Setup and the app's Settings screen, so the code stays the same in every copy. Never type a passcode or your app link into the code.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

### One link, two jobs
The frontend is a self-contained React app (`index.html`, loaded from a CDN, no build step). The backend is an Apps Script web app bound to the Sheet. Opening the web app link with no `action` serves `index.html` with the link filled in; calls with an `action` are the API. The frontend never has direct access to the Sheet. (You can still host `index.html` elsewhere by pasting the link into `APPS_SCRIPT_URL`.)

### How the two pieces actually talk
The backend has a `doGet` and a `doPost`, both gated by a token (`requireAuth_`). The page holds no secrets: the `login` action checks the passcode on the server and only then hands back the token, and `checkSettingsPasscode` does the same for Settings. Passcodes, the token and alert emails live in Script Properties. Reads go through `doGet` with an `action` (e.g. `checkPhone` for the repeat lookup, `getEmailSettings`); writes go through `doPost` with an `action` (`add` / edit / delete a complaint). There's a small **field-alias** layer so the frontend can speak its own names (`phone`, `timestamp`) while the sheet stores its columns (`phoneNumber`, `dateSubmitted`) — handy to know if a field ever looks like it "didn't save," it may just be an alias mismatch.

### The data
One tab (`Sheet1`) is the database; the header row is `id, phone, guestName, orderNumber, orderType, issueType, resolution, notes, dateOccurred, timestamp, managerName, dayPart, resolved`, built on first-time setup (and auto-created on first submit, so there's no schema step). `runFirstTimeSetup()` wires the tab, headers, and alert storage; it's idempotent.

### Repeat detection + alerts, by the numbers
Two constants at the top of `Code.gs` drive it: `REPEAT_WINDOW_DAYS` (30) and `ALERT_THRESHOLD` (4). When a phone is entered, the backend counts that number's complaints inside the rolling window and badges the guest; when a new complaint pushes a guest to the threshold *within* the window, it emails the alert list. **These two numbers must match the wording the frontend shows guests** — a code comment flags this, and it's the one spot where the number in the UI and the number in the logic can silently drift apart.

### One copy of the code
Because the sheet ID, passcodes and alert emails are stored in Script Properties (set by Start Setup), the code is identical in every store's copy and in the public template. `SHEET_ID` stays `'YOUR_SHEET_ID'`; the script finds the sheet it's attached to.

### The frontend's one dependency
It pulls React/Babel from a CDN at load. If you ever fork or update it and it white-screens, check the CDN script versions are pinned to known-good ones — an unpinned major version has broken it before.

### Deploy model
Publish as a **new version** after any code change (the link serves the last published version). Deploy as **Execute as: Me** and **Who has access: Anyone** so managers can open it without a Google login; the passcode is what keeps the data private.
