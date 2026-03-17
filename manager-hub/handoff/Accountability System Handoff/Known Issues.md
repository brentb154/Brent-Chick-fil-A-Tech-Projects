# Known Issues

## Current
- Health checks may warn about missing triggers if not scheduled.
- MailApp permission warning appears if not authorized in Apps Script.

## Workarounds
- Run `scheduleHealthChecks()` and `scheduleLogCleanupTrigger()`.
- In Apps Script editor, run a MailApp function once to authorize.
