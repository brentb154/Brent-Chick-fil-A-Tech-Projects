# Tool Adoption Tracker

Answers one question every Monday morning: **which tools did people actually open last week?**

Every hub tool logs one row per day to a shared "Tool Adoption" spreadsheet the first time someone opens or uses it that day. This script (bound to that spreadsheet) emails a weekly digest: days used this week, 4-week average, last opened. Tools nobody opens stop getting investment; tools people rely on get more.

> 📖 **[Go deeper: what it is & how it works →](GUIDE.md)** — the plain-language operator guide (what it does, how it works, and how to fix the common stuff).

## One-time setup

1. **Create a new Google Sheet** named `Tool Adoption`. Copy its ID from the URL (the long string between `/d/` and `/edit`).
2. **Extensions → Apps Script**, paste in `Code.gs`, save.
3. Run `runInitialSetup()` once (authorize when prompted). This creates the `Pings` and `Settings` tabs — safe to re-run, never touches existing data.
4. On the `Settings` tab, set **Digest Email** to wherever the weekly report should go.
5. Run `createWeeklyTrigger()` once — installs the Monday 6 AM digest (checks for an existing trigger first, so re-running won't duplicate).

## Wiring up each tool

Each tool already contains a `logAdoptionPing_()` helper. It does **nothing** until you give it the adoption sheet's ID:

1. Open the tool's Apps Script project.
2. Project Settings (gear icon) → **Script Properties** → Add property:
   - Property: `ADOPTION_SHEET_ID`
   - Value: the Tool Adoption sheet ID from step 1 above.

That's it. No property = silent no-op, so tools keep working fine if this is never set up (and copies of a tool installed by someone else never phone home).

Tools wired in this repo: payroll-system, catering-quote-tool, schedule-counter-gas, monday-huddle, training-tracker, foh-links-menu, shared-table-email-summary.

**HEARD (guest-recovery-log):** the repo copy is the public template, so it does NOT include the ping. To track the private live copy, paste this into its Code.gs and call `logAdoptionPing_('guest-recovery-log')` at the top of `doGet`:

```javascript
// Adoption ping: one row per day to the shared adoption sheet.
// No-op unless ADOPTION_SHEET_ID is set in Script Properties. Never throws.
function logAdoptionPing_(toolName) {
  try {
    var props = PropertiesService.getScriptProperties();
    var sheetId = props.getProperty('ADOPTION_SHEET_ID');
    if (!sheetId) return;
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (props.getProperty('ADOPTION_LAST_PING') === today) return;
    var tab = SpreadsheetApp.openById(sheetId).getSheetByName('Pings');
    if (!tab) return;
    tab.appendRow([today, toolName]);
    props.setProperty('ADOPTION_LAST_PING', today);
  } catch (err) {
    // Never let adoption logging break the tool.
  }
}
```

## Reading the digest

- **Days used this week** — how many distinct days someone opened the tool (max 7). A dash means nobody touched it.
- **Avg days/week (prior 4)** — the trend baseline. This week well below the average = worth asking why.
- Web apps ping on page load; sheet-bound tools ping when someone runs a real action from the menu (not just opening the spreadsheet), so it measures actual use.
- `shared-table-email-summary` works via an automated daily email, so it only pings when someone opens its settings — quiet is normal for that one.

## Notes

- One row per tool per day, so the Pings tab grows ~8 rows/day worst case. Years of headroom.
- The ping adds one Script Properties read (~5 ms) to each open; the sheet write happens at most once a day per tool.
