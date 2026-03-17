# CFA Accountability System - Technical Docs

Version: 2026-01-24  
Audience: Developers / System Administrators

---

## 1) System Architecture

### High-Level Components
- **Google Apps Script** backend (logic, auth, data access)
- **Google Sheets** as database
- **HTML/CSS/JS** front-end served by Apps Script web app
- **Google Drive** for reports and backups
- **Email** via Gmail/MailApp

### Component Diagram (Mermaid)
```
graph TD
  UI[HTML/JS Web UI] -->|google.script.run| GAS[Apps Script Backend]
  GAS --> Sheets[(Google Sheets)]
  GAS --> Drive[(Google Drive)]
  GAS --> Mail[(Gmail/MailApp)]
  GAS --> Triggers[(Time-based Triggers)]
```

### Data Flow (Infractions)
1. UI submits infraction data.
2. Backend validates, writes to `Infractions` sheet.
3. Points recalculated.
4. Thresholds detected.
5. Emails sent and logged.
6. System log entry created.

### Integration Points
- Payroll Tracker sheet for employee roster
- Drive folder(s) for reports/backups
- Gmail for notifications

---

## 2) Technology Stack
- Google Apps Script (V8)
- Google Sheets
- HTML/CSS/Vanilla JS
- Chart.js (analytics dashboard)

---

## 3) Database Schema (Sheets)

### Core Sheets
**Infractions**
- Employee_ID (string)
- Full_Name (string)
- Date (date)
- Infraction_Type (string)
- Points_Assigned (number)
- Description (string)
- Location (string)
- Entered_By (string)
- Status (Active/Expired/Deleted)
- Expiration_Date (date)

**Settings**
- Key-value configuration
- Passwords (manager/director/operator)
- Email recipients and thresholds

**Email_Log**
- Log_ID, Timestamp, Recipient_Email
- Email_Type, Thresholds_Crossed
- Status, Retry_Count, Error_Message

**System_Log**
- Event_Type, Severity, User, Function_Name
- Error_Stack, Resolved status/notes

**Health_Check_History**
- Overall_Status
- Issues_JSON
- Metrics_JSON

**Backup_Log**
- backup_id, backup_date, status, file_url

---

## 4) Function Reference (Summary)

### Core Workflow
- `processInfractionWithNotifications()`
- `addInfraction()`
- `calculatePoints()`
- `detectThresholds()`
- `sendThresholdEmail()`

### System Monitoring
- `logSystemEvent()`
- `checkSystemHealth()`
- `sendHealthAlert()`
- `scheduleHealthChecks()`
- `getSystemMetrics()`
- `cleanupOldLogs()`

### Reporting
- `generateDetailedReport()`
- `getReportHistory()`
- `scheduleRecurringReport()`

### User/Auth
- `handleDashboardLogin()`
- `getCurrentRole()`
- `validateSessionToken()`
- `logoutRole()`

### Backups
- `createSystemBackup()`
- `scheduleQuarterlyBackup()`
- `runAutomaticBackup()`

> **Full Function Index**  
> See source files in the repository. Update this section when new functions are added or renamed.

---

## 5) API Documentation

The app exposes Apps Script functions for client calls via `google.script.run`.
No public REST API is exposed.

Common client calls:
- `getDashboardStatistics(token)`
- `getSystemStatusData(token)`
- `runHealthCheckNow(token)`
- `getReportHistory(token)`

Authentication: token-based session (stored in `sessionStorage`).

---

## 6) Calculation Logic

### Point Calculation
- Active infractions only.
- Points expire after 90 days.
- Negative points reduce totals.

### Threshold Detection
Typically:
- 6 points: warning
- 9 points: probation
- 15 points: termination review

---

## 7) Email System

### Template Structure
Email templates are HTML strings built in Apps Script:
`buildThresholdEmail()` and related helpers.

### Send Logic
`sendThresholdEmail()`:
- attempts send
- retries once on failure
- logs to `Email_Log`
- logs system event

---

## 8) Authentication & Security

### Session Model
- `validateSessionToken()` checks active session
- `getCurrentRole()` extends session on use
- Token stored in `sessionStorage` and `localStorage`

### Permissions
Role-based gating in Apps Script:
- Manager: standard operations
- Director/Operator: admin actions, settings, reports

---

## 9) Deployment

### Web App Deployment
1. Apps Script → Deploy → New deployment
2. Type: Web app
3. Execute as: Me
4. Who has access: Anyone with link (or domain)

### Permissions
Reauthorize when:
- MailApp APIs added
- Drive APIs added
- New triggers created

---

## 10) Maintenance Tasks

### Daily
- `cleanupOldLogs()` (via trigger)
- `checkPasswordExpiration()`

### Weekly
- Review System Status
- Check Email_Log for failures

### Monthly
- Report exports review
- Threshold policy audit

### Quarterly
- Verify backup creation
- Test restore workflow

---

## 11) Troubleshooting

See `docs/troubleshooting.md` for detailed guide.

---

## 12) Future Development

Known limitations:
- Sheet size limits
- Trigger quota limits
- Email quota limits

Enhancement ideas:
- Role-based UI personalization
- Audit dashboards
- Performance caching improvements

Migration paths:
- Move data to a database (BigQuery/Firestore)
- Use Cloud Functions for heavy processing
