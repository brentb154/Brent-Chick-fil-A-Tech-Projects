# H.E.A.R.D. Log — Operator Guide

**This is one fast place to log every guest complaint and how you recovered it — from a phone or tablet, in under a minute — and it automatically flags the guest who's called in three times this month.** Front-of-house leaders log; managers get the dashboard and the alerts.

## Why it matters
Complaints get handled in the moment and then forgotten, so nobody notices the patterns: the repeat caller, the daypart that generates most of the issues, the "free item" handed out more than it should be. Paper logs and one-off texts don't add up. This makes the record shared and searchable, so the trends are visible and the recovery conversation — Hear, Empathize, Apologize, Resolve, Delight — has a memory behind it.

## What it does
- **Log a complaint in seconds.** Guest, order, issue type, resolution, daypart, and which manager handled it.
- **Catch repeat complainers automatically.** Type a phone number and it checks the last 60 days, and badges the guest if there's a history.
- **Show the whole picture.** A dashboard of totals, flagged guests, and trends, plus a searchable, editable log.
- **Alert leadership.** An email when a guest crosses the repeat threshold you set.

## How it works (the plain version)
This one's in two pieces that talk over a single web link:
- **The form** (the page leaders open) is self-contained — put it on a tablet's home screen, a bookmark, anywhere. It needs internet (it loads from the web) but no install.
- **The backend** is a Google Apps Script app tied to your Google Sheet. The form never touches the sheet directly — everything goes through the backend link, so the sheet stays private and only your deployed app can read or write it.
- **The Google Sheet** is the database — one tab, with the columns created automatically the first time you submit.

## Everyday tasks
- **Log an issue:** open the form → fill it in → submit. Watch for the repeat-guest badge when you enter the phone.
- **Review trends:** the dashboard shows totals, flagged guests, and the log; you can search and edit past entries.
- **Change who gets alerted, or the repeat threshold:** those are settings — set them once and they hold.

## When something looks broken
- **The form loads but won't save.** It's pointed at the wrong backend link, or you changed the backend code and didn't publish a new version. Re-publish (**Deploy → Manage deployments → Edit → New version**) and make sure the form's link matches the deployed app.
- **Blank / white page.** The form loads its framework from the web — check the device has internet. (If you host your own copy and it white-screens after an update, it's almost always that front-end loading step, not your data.)
- **The repeat badge didn't show.** It keys off the phone number over the last 60 days — it only flags once the phone is entered and there's prior history.
- **Alerts aren't arriving.** Check the alert email and threshold settings, and that the backend is authorized to send mail from the right Google account.

Real platform breakage is rare. For the backend, a new deployment version fixes most "it stopped saving" issues; for the form, it's almost always the link or the internet.

## The one rule
**Keep your live copy and the public template separate.** The version you actually run has your real backend link and settings baked into it — don't overwrite it with the shared/public copy, and don't paste your live link into a public one. Config (alert email, threshold) is a setting, not a code rewrite.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

**Two pieces, one link.** The frontend is a self-contained React app (loaded from a CDN, no build step) — it can live anywhere. The backend is an Apps Script web app bound to the Sheet. The frontend calls the backend URL for every read and write; it never has direct access to the Sheet. That's the privacy boundary — only the deployed endpoint can touch the data, so hosting the form publicly doesn't expose the log.

**The data.** One tab is the database; the columns are created on first submit, so there's no schema to set up. Repeat-complainer detection is a lookback over recent entries by phone number (a rolling ~60-day window), surfaced the instant the phone is typed.

**Two copies, on purpose.** There's a public/template version and your private live version, and they are NOT meant to be identical — the live one carries your real backend link and alert config. Never force them into sync; that's how a private link or setting leaks into the public copy.

**The frontend's one dependency.** It pulls its framework from a CDN at load. If you ever fork or update it and it white-screens, check that the CDN script versions are pinned to known-good ones — an unpinned major version has broken it before.

**Deploy model.** Publish the backend as a **new version** after any `Code.gs` change (the URL serves the last published version). The form just needs to point at that URL, and the deployment access has to let whoever's logging reach it.
