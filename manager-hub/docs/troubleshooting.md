# CFA Accountability Troubleshooting Guide

Version: 2026-01-24

---

## Login Issues

### Forgot Password
- Contact a director to reset the shared role password.
- Log out and log back in with the updated password.

### Account Locked / Session Expired
- Refresh the page.
- If error persists, clear browser cache and retry.
- Contact a director if the issue continues.

### Wrong Role Access
- Verify role selection at login.
- Directors can adjust user roles in User Management.

---

## Adding Infractions

### Cannot Find Employee
- Confirm the employee is active in Payroll Tracker.
- Ask a director to verify the Payroll Tracker sheet.

### Validation Errors
Common causes:
- Missing required fields
- Description < 240 characters
- Invalid date (future or > 7 days old)
- Invalid location

### Submission Fails
- Check network connection.
- Refresh the page and re-submit.
- If repeated, notify a director with screenshot/time.

---

## Email Issues

### Notifications Not Received
- Check spam/junk folder.
- Verify recipient list in Settings.
- Ensure email quota not exceeded.

### Wrong Recipients
- Review email settings in the Settings page.

### Email Quota Exceeded
- Wait for daily quota reset.
- Reduce non-critical emails.

---

## Performance Issues

### Slow Loading
- Refresh page.
- Clear browser cache.
- Check internet connection.

### Timeout Errors
- Reduce report date range.
- Avoid running multiple exports at once.

### Mobile Lag
- Close other browser tabs.
- Try a different browser.

---

## Data Issues

### Point Totals Wrong
- Check for expired infractions.
- Verify positive credits in history.
- Ask a director to run recalculation.

### Missing Infractions
- Confirm filters on Employee List.
- Check date range on reports.

### Incorrect Dates
- Verify system timezone in Apps Script.
- Confirm correct date input at submission.

---

## System Errors

### Error Codes / Meanings
- **SessionExpired**: login required
- **ValidationError**: missing or invalid field
- **PermissionError**: access or API permission issue

### When to Retry
- Temporary network error
- Google Apps Script timeout

### When to Escalate
- Data integrity errors
- Payroll Tracker access failure
- Repeated email failures

---

## How to Report Bugs

Include:
- What happened
- Steps to reproduce
- Timestamp
- Screenshots
- User role
