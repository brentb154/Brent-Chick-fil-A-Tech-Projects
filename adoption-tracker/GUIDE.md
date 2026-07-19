# Tool Adoption Tracker — Operator Guide

**This answers one question every Monday: which of your tools did people actually open last week?** Every tool quietly logs when it's used; this collects those pings and emails you a weekly digest — so you invest in what's used and stop polishing what nobody touches.

## Why it matters
It's easy to keep building features for a tool nobody opens, and easy to under-invest in the one everybody relies on. Without usage data, that's gut feel. This gives you the real number — days used this week vs. the recent average — so the "keep, cut, or push harder" call is backed by evidence.

## What it does
- **Counts real usage** — one row per tool per day, the first time someone actually uses it that day.
- **Weekly digest email** — every Monday: days used this week, the prior 4-week average, and last opened.
- **Measures use, not just opens** — web apps ping on load; sheet tools ping when someone runs a real action from the menu, not just opening the file.

## How it works (the plain version)
Three moving parts, all in Google:
- **Each tool has a tiny "ping" built in.** The first time it's used on a given day, it drops one row into a shared **Tool Adoption** sheet.
- **This script lives on that shared sheet.** It reads the pings and, on a Monday-morning trigger, emails the digest.
- **The ping is silent until you turn it on.** A tool only phones home once you give it the adoption sheet's ID (a Script Property) — so a tool works fine whether or not it's tracked, and copies installed by other people never report back.

## Everyday tasks
- **Read the Monday email** — that's the whole job. A dash means nobody opened that tool; a number well below its average is worth asking about.
- **Add a new tool to the tracking:** set the `ADOPTION_SHEET_ID` Script Property in that tool's Apps Script project to the Tool Adoption sheet's ID.
- **Change where the digest goes:** the `Digest Email` on the Settings tab.

## When something looks broken
- **A tool always shows a dash / never reports.** Its ping isn't wired up — set `ADOPTION_SHEET_ID` in that tool's Script Properties. No property = silent no-op, by design.
- **No digest arrived Monday.** The weekly trigger isn't installed — run `createWeeklyTrigger()` once. Check the `Digest Email` on Settings too.
- **`shared-table-email-summary` looks quiet.** That's normal — it runs as an automated email, so it only pings when someone opens its settings.
- **HEARD isn't tracked.** The repo copy is the public template with no ping; only the private live copy can report, and only after you add the ping and the sheet ID to it.

## The one rule
**A tool is invisible to this until you set its `ADOPTION_SHEET_ID`.** That one Script Property per tool is the on-switch. Everything else — the pings, the digest — runs itself.

---

## Go deeper
*The 1,000-foot view for whoever maintains it next.*

### Why a ping instead of a dashboard
Rather than each tool reporting rich analytics, every tool drops a single dated row (`[date, toolName]`) the first time it's used that day. That's deliberately dumb: one Script Properties read (~5 ms) plus at most one sheet write per tool per day, so it's invisible to the tool's performance, and the `Pings` tab grows ~8 rows/day worst case — years of headroom. All the intelligence lives in one place — the digest — not scattered across every tool.

### The de-duplication (why it's "days," not "opens")
A tool remembers the last day it pinged in its own Script Properties (`ADOPTION_LAST_PING`) and skips if it already pinged today. So the raw data is already one-row-per-tool-per-day. That's the whole reason "days used" is a clean count of *distinct days* and not a noisy hit count — it measures adoption, not traffic.

### How the digest is computed
`sendWeeklyAdoptionDigest` reads every `Pings` row and buckets it per tool into three numbers: **this week** (pings since the start of the current week), **prior 4** (pings in the preceding four weeks), and **last seen** (the most recent date). The email table then shows *Days used this week*, *Avg days/week (prior 4)* — literally `prior4 / 4` — and *Last opened*, so a tool sitting well under its own average jumps out. A dash means zero this week. The digest goes to the `Digest Email` from Settings, falling back to the script owner's address if that's blank.

### Fail-silent by design
The ping is wrapped in a try/catch that can never throw — if the adoption sheet is missing, the ID isn't set, or anything else goes wrong, the tool keeps working and just doesn't log. That's what makes it safe to ship in a public template: no ID, no phone-home, no risk to the host tool.

### The pieces
The shared `Tool Adoption` sheet has a `Pings` tab (the raw `[date, tool]` rows) and a `Settings` tab (`Digest Email`, seeded to the owner's address). This bound `Code.gs` builds and sends the digest on a Monday 6 AM trigger (`createWeeklyTrigger`, which checks for an existing one so re-running won't duplicate). `runInitialSetup` builds the tabs and is safe to re-run.
