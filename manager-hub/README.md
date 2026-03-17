# CFA Accountability System

Internal system for tracking employee infractions, accountability thresholds, reporting, and system monitoring.

## Quick Links
- Web app URL: provided by director
- Docs: `docs/`
- Handoff materials: `handoff/Accountability System Handoff/`

## Running / Deployment
Deploy as a Google Apps Script web app:
1. Apps Script → Deploy → New deployment
2. Type: Web app
3. Execute as: Me
4. Who has access: Anyone with link (or domain)

## Key Sheets
- `Infractions`
- `Settings`
- `Email_Log`
- `System_Log`
- `Health_Check_History`
- `Backup_Log`

## Maintenance
- `scheduleHealthChecks()` (every 6 hours)
- `scheduleLogCleanupTrigger()` (daily cleanup)
- `scheduleQuarterlyBackup()` (quarterly backup)

## Documentation
See:
- `docs/manager-guide.md`
- `docs/director-guide.md`
- `docs/technical-docs.md`
- `docs/troubleshooting.md`
