# Training Tracker — Operator Guide

**This turns training a new hire from a paper guessing game into a live picture: log hours through a form your trainers already use, and it builds each person's timeline, tracks every position, and tells you the moment someone's ready to certify.** Trainers log; managers watch the dashboard and get the alerts.

## Why it matters
A new team member has dozens of positions, each with an hour requirement, across a moving schedule and several trainers. On paper, people fall through the cracks and nobody's sure who's ready for what. This keeps it all in one Sheet that updates itself, so "who's training where today" and "who's ready to certify" are answers, not guesses.

## What it does
- **Builds a personalized timeline** for each new hire from their start date, schedule, and position requirements.
- **Tracks hours from a Google Form** your trainers already know how to use — no new app to learn.
- **Live dashboard** of every active trainee: hours, position-by-position progress, certification status.
- **Email alerts** when someone finishes a position, goes inactive, or is ready to certify.
- **Handles name messiness** automatically — "Tim" vs "Timothy," typos — so one person doesn't become three.
- **Fills the weekly Training Needs sheet** so managers know who's training where each day.

## How it works (the plain version)
This one lives inside Google Sheets and Forms — no separate app.
- **Trainers log on a Google Form.** Each entry lands in the Sheet as raw training data.
- **The Sheet does the work.** Apps Script turns those entries into the timeline, the dashboard, the weekly needs list, and the certification workflow — automatically.
- **You configure it in the Sheet.** Position hour requirements and which alerts fire (and to whom) are just tabs you edit: Position Requirements and Alert Settings.

## Everyday tasks
- **Log training:** trainers fill the Google Form — that's it.
- **See who's where:** the Master Dashboard (live) and the weekly Training Needs sheet.
- **Adjust a position's required hours:** the Position Requirements tab.
- **Change who gets alerts, or which fire:** the Alert Settings tab.
- **Certify someone:** the system flags cert-ready trainees and archives them to the Certification Log when done.

## When something looks broken
- **Form entries aren't showing up.** Make sure the Google Form is still connected to this Sheet — responses have to land in the Daily Training Log.
- **The same person shows up twice.** The dedup catches most variants, but a very different spelling can slip through — merge them on the Name Deduplication tab.
- **Alerts aren't sending.** Check the Alert Settings tab (which alerts are on, and the recipient), and that the script is authorized to send mail from the right account.
- **The dashboard looks stale.** It updates on new entries — give it a beat after a form submit, or re-run the refresh from the custom menu.
- **Hours look wrong for a position.** Check that position's target/min hours on the Position Requirements tab — that's what everything measures against.

Real breakage is rare, and usually comes down to the Form losing its link to the Sheet, or an alert recipient that changed. Re-running the setup is safe and rebuilds the tabs without touching your data.

## The one rule
**The whole thing is configured in Sheet tabs — Position Requirements and Alert Settings — not in the code.** Change what a position requires, or who hears about it, right there. The trainers never leave the Google Form.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

**Form in, everything else out.** The one thing humans do is submit the Google Form; it lands in the Daily Training Log. From that raw log, the script generates the personalized timeline, keeps the Master Dashboard current, populates the weekly Training Needs view, runs the certification workflow, and archives finished trainees to the Certification Log. Everything downstream is derived — the log is the source of truth.

**The tabs.** Daily Training Log (raw form data), Position Requirements (FOH/BOH positions with min/target/max hours — editable), Master Dashboard (live overview), Training Schedule (generated day-by-day plan), Training Needs (weekly manager view), Certification Log (archive), Name Deduplication (merge map), Alert Settings (which alerts, who gets them). Setup creates them all.

**Name normalization.** New-hire names come in inconsistent — nicknames, typos, spacing. The dedup layer normalizes and matches them so one trainee's hours don't scatter across three near-identical names. When it can't be sure, it surfaces the pair for a human to merge rather than guessing.

**Setup is idempotent.** Re-running it rebuilds any missing tabs and re-wires things without wiping your data — safe to run again if something looks off.

**It's a Sheets/Forms system, not a hosted web app.** There's no `/exec` URL to redeploy — the "deploy" is the Sheet plus its bound script and the connected Form. Keep the Form linked to the Sheet, keep the position list and alert recipients current, and it runs itself.
