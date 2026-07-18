# Payroll System — Operator Guide

**This is the hub for the parts of payroll that used to live on paper and in scattered spreadsheets — overtime, PTO requests, and uniform orders with their deductions, across your locations, in one access-code-protected app.** Managers run it; employees submit their requests through simple form links.

## Why it matters
Overtime creep, uniform deductions, and time-off requests are easy to lose track of when they're on sticky notes and separate sheets. This keeps them in one place: OT you can actually see and trend, uniform orders that flow straight into payroll deductions, and PTO requests with a clean record — so payroll runs go faster and nothing slips.

## What it does
- **Overtime monitoring.** Upload the OT report and see it by employee and location, with history and trends.
- **PTO requests.** Employees request time off (English or Spanish); managers see the records. (Actual PTO *balances* live in your HR system — this tracks the requests, not the balance math.)
- **Uniform ordering + deductions.** An editable catalog, employee orders, and the payroll deductions that come out of them, all tied together.
- **Payroll calendar + year-end.** The pay-period calendar, and a wizard for the once-a-year cleanup.
- **Dashboards + system health.** The manager view, plus a health page that flags anything off.

## How it works (the plain version)
Same shape as the other tools: a web app sitting on a Google Sheet.
- **The web page is access-code protected.** Managers and admins get in with a code — it's auto-generated on first use and emailed to your admins, and you can change it in the `Payroll_Settings` sheet.
- **Employees don't log in.** They submit PTO and uniform requests through their own simple form links.
- **The Google Sheet behind it** holds everything: employees, OT, PTO records, the uniform catalog and orders, and your settings. Config — locations, admin email, wage, login hint — lives on the Settings page, no code editing.

## Everyday tasks
- **Log overtime:** OT Upload → paste in the report → it lands in history and the trends update.
- **Handle a PTO request:** it shows up in PTO Records — review it there.
- **Run uniform deductions:** Uniform Orders shows what employees ordered; Uniform Deductions is what comes out of pay; the catalog is editable.
- **A pay period / year-end:** the Payroll Calendar shows the periods; the Year-End Wizard walks the annual steps.
- **Add or remove an employee:** manage them in the app. A removed employee is archived (not deleted) and comes back if you re-add them.
- **Change a setting:** the Settings page — locations, admin email, wage, login hint, codes.

## When something looks broken
- **Stuck on "Loading…" or a blank page.** New code was pasted but not redeployed — **Deploy → Manage deployments → Edit → New version.** The link always serves the last *published* version.
- **The access code doesn't work.** It's in the **`Payroll_Settings`** sheet (auto-generated and emailed to admins on first use). Reset it there.
- **An employee's form link asks them to log in.** They shouldn't need to — check the deployment access, and that they're using the *form* link, not the manager app link.
- **OT numbers look off after an upload.** Re-check the report you uploaded — history keeps prior uploads, so you can compare.
- **A removed employee reappeared.** That's the archive doing its job: removed employees are archived, not deleted, and restored if they come back. Nothing's lost.

Real Apps Script breakage is rare; publishing a new version fixes the large majority of "it stopped working" moments. When in doubt: re-run the setup function once (safe to repeat), then deploy a new version.

## The one rule
**Everything you'd want to change lives on the Settings page or in the Sheet — not in the code.** Locations, admin email, wage, codes, the login hint: all operator-editable. After any code change, publish a new version.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next. You don't need this to run payroll day to day.*

**Who gets in, and how.** The app is gated by an access code (managers/admins), checked when the page loads; codes live in `Payroll_Settings`, auto-generated on first use and emailed to admins. Employees never log in — PTO and uniform requests come through separate public form links. Inside, the app routes by a `?view=` tag to each screen (dashboard, OT upload / history / trends / reconciliation, PTO records / summary, uniform catalog / orders / deductions / summary, payroll calendar, year-end wizard, settings, system health, help).

**The data model.** Google Sheet tabs back everything: an employees tab, overtime records, PTO records, the uniform catalog + orders + deductions, and `Payroll_Settings` for config. Removed employees move to an **archive** tab instead of being deleted, and are restored automatically if re-added — so history and prior deductions are never lost.

**PTO balances are deliberately NOT here.** Your HR system owns the actual balance math; this tool records the *requests* and their status. Don't rebuild balance tracking here — it would only drift from HR.

**The money math has one source.** Payday dates and uniform-deduction amounts are computed by shared helpers, so every screen agrees. If you change that math, change the helper — not each screen.

**Deploy model.** No auto-sync. Paste the files into Apps Script and publish a **new version**; the numbered code files and the `View_*` screen files go into their matching slots. Multi-location is built in — the OT monitor spans every location you set up.

**Two gotchas.** The link serves the last *published* version, not your latest paste — always publish a new one. And config lives in the Sheet/Settings on purpose, so a non-technical manager can keep it running without touching code.
