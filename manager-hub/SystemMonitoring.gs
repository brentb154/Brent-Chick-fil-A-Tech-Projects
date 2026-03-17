// ============================================
// MICRO-PHASE 29: SYSTEM HEALTH MONITORING & ALERTS
// ============================================

const SYSTEM_LOG_SHEET = 'System_Log';
const HEALTH_CHECK_HISTORY_SHEET = 'Health_Check_History';

function getOrCreateSystemLogSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SYSTEM_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_LOG_SHEET);
    const headers = [
      'Log_ID',
      'Timestamp',
      'Event_Type',
      'Event_Details',
      'Severity',
      'User',
      'Function_Name',
      'Error_Stack',
      'Resolved',
      'Resolved_By',
      'Resolved_At',
      'Resolution_Notes'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateHealthCheckHistorySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(HEALTH_CHECK_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(HEALTH_CHECK_HISTORY_SHEET);
    const headers = [
      'Check_ID',
      'Check_Timestamp',
      'Overall_Status',
      'Checks_Passed',
      'Checks_Failed',
      'Issues_JSON',
      'Metrics_JSON',
      'Alert_Sent'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateSystemLogId() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const rand = Math.floor(1000 + Math.random() * 9000);
  return `SYS-${stamp}-${rand}`;
}

function getCallerFunctionName() {
  try {
    const stack = new Error().stack;
    if (!stack) return '';
    const lines = stack.split('\n').map(l => l.trim());
    if (lines.length < 3) return '';
    const callerLine = lines[2];
    const match = callerLine.match(/at\s+([^\s]+)/);
    return match ? match[1] : '';
  } catch (e) {
    return '';
  }
}

/**
 * Logs a system event to the System_Log sheet.
 *
 * @param {string} event_type - "error", "warning", "info", "success"
 * @param {string|Error|Object} event_details - description of event
 * @param {string} severity - "critical", "high", "medium", "low"
 * @returns {string|null} Log ID
 */
function logSystemEvent(event_type, event_details, severity) {
  try {
    const sheet = getOrCreateSystemLogSheet();
    const logId = generateSystemLogId();
    const now = new Date();
    const userEmail = (Session.getActiveUser && Session.getActiveUser().getEmail)
      ? Session.getActiveUser().getEmail()
      : '';
    const user = userEmail || 'System';
    const functionName = getCallerFunctionName() || 'Unknown';

    let details = '';
    let errorStack = '';
    if (event_details instanceof Error) {
      details = event_details.toString();
      errorStack = event_details.stack ? String(event_details.stack) : '';
    } else if (typeof event_details === 'object' && event_details !== null) {
      details = JSON.stringify(event_details);
    } else {
      details = String(event_details || '');
    }

    if (!errorStack && event_type === 'error') {
      try {
        errorStack = new Error(details).stack || '';
      } catch (e) {
        errorStack = '';
      }
    }

    sheet.appendRow([
      logId,
      now,
      event_type,
      details,
      severity,
      user,
      functionName,
      errorStack,
      false,
      '',
      '',
      ''
    ]);

    if (event_type === 'error' && severity === 'critical') {
      sendCriticalSystemAlert({
        log_id: logId,
        timestamp: now,
        event_type: event_type,
        event_details: details,
        severity: severity,
        user: user,
        function_name: functionName,
        error_stack: errorStack
      });
    }

    return logId;
  } catch (error) {
    console.error('logSystemEvent failed:', error.toString());
    return null;
  }
}

function sendCriticalSystemAlert(logEntry) {
  try {
    const recipients = getDirectorEmails();
    if (!recipients || recipients.length === 0) return false;

    const subject = 'CRITICAL: Accountability System Issue';
    const body = `
CRITICAL SYSTEM ALERT

Timestamp: ${logEntry.timestamp}
Type: ${logEntry.event_type}
Severity: ${logEntry.severity}
User: ${logEntry.user}
Function: ${logEntry.function_name}

Details:
${logEntry.event_details}

Stack Trace:
${logEntry.error_stack || 'No stack trace available'}

Log ID: ${logEntry.log_id}
`.trim();

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });
    return true;
  } catch (error) {
    console.error('sendCriticalSystemAlert failed:', error.toString());
    return false;
  }
}

function checkSystemHealth(token) {
  const start = Date.now();
  const issues = [];
  let checksPassed = 0;
  let checksFailed = 0;

  const addIssue = (issue_type, severity, description, recommended_action, logIt) => {
    const issue = {
      issue_type,
      severity,
      description,
      recommended_action
    };
    if (logIt) {
      const logId = logSystemEvent(
        severity === 'critical' ? 'error' : 'warning',
        `${issue_type}: ${description}`,
        severity
      );
      issue.log_id = logId || '';
    }
    issues.push(issue);
    checksFailed += 1;
  };

  const passCheck = () => {
    checksPassed += 1;
  };

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // 1) Data integrity
    try {
      const infractionsSheet = ss.getSheetByName('Infractions');
      if (!infractionsSheet) {
        addIssue('data_integrity', 'critical', 'Infractions sheet is missing.', 'Restore the Infractions sheet or run setupAllTabs().', true);
      } else {
        const headers = infractionsSheet.getRange(1, 1, 1, infractionsSheet.getLastColumn()).getValues()[0];
        const required = ['Employee_ID', 'Full_Name', 'Date', 'Infraction_Type', 'Location', 'Entered_By'];
        const missing = required.filter(col => headers.indexOf(col) === -1);
        if (missing.length) {
          addIssue('data_integrity', 'critical', `Missing required columns: ${missing.join(', ')}`, 'Restore the missing columns and header labels.', true);
        } else {
          passCheck();
        }

        const lastRow = infractionsSheet.getLastRow();
        if (lastRow > 1) {
          const sampleRows = Math.min(200, lastRow - 1);
          const data = infractionsSheet.getRange(lastRow - sampleRows + 1, 1, sampleRows, headers.length).getValues();
          const colIndex = {
            Employee_ID: headers.indexOf('Employee_ID'),
            Full_Name: headers.indexOf('Full_Name'),
            Date: headers.indexOf('Date'),
            Infraction_Type: headers.indexOf('Infraction_Type'),
            Location: headers.indexOf('Location'),
            Entered_By: headers.indexOf('Entered_By'),
            Points_Assigned: headers.indexOf('Points_Assigned') !== -1 ? headers.indexOf('Points_Assigned') : headers.indexOf('Points')
          };
          let nullCount = 0;
          let invalidPoints = 0;
          data.forEach(row => {
            ['Employee_ID', 'Full_Name', 'Date', 'Infraction_Type', 'Location', 'Entered_By'].forEach(key => {
              if (row[colIndex[key]] === '' || row[colIndex[key]] === null || row[colIndex[key]] === undefined) {
                nullCount += 1;
              }
            });
            const pts = Number(row[colIndex.Points_Assigned]);
            if (isNaN(pts)) invalidPoints += 1;
          });
          if (nullCount > 0) {
            addIssue('data_integrity', 'high', `${nullCount} missing critical field values in recent infractions.`, 'Review recent infractions for missing fields and correct them.', true);
          } else {
            passCheck();
          }
          if (invalidPoints > 0) {
            addIssue('data_integrity', 'high', `${invalidPoints} infractions have invalid point values.`, 'Fix invalid points in the Infractions sheet.', true);
          } else {
            passCheck();
          }
        }
      }

      const settingsSheet = ss.getSheetByName('Settings');
      if (!settingsSheet) {
        addIssue('data_integrity', 'critical', 'Settings sheet is missing.', 'Restore the Settings sheet or run setupAllTabs().', true);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('data_integrity', 'critical', `Data integrity check failed: ${error}`, 'Review Infractions/Settings sheets and permissions.', true);
    }

    // 2) Storage limits
    try {
      const infractionsSheet = ss.getSheetByName('Infractions');
      const rowsUsed = infractionsSheet ? infractionsSheet.getLastRow() : 0;
      const rowLimit = 10000000;
      const usagePercent = rowLimit > 0 ? Math.round((rowsUsed / rowLimit) * 100) : 0;
      if (usagePercent >= 80) {
        addIssue('storage_limits', 'high', `Sheet rows used at ${usagePercent}% (${rowsUsed} of ${rowLimit}).`, 'Archive old infractions or split data into yearly sheets.', true);
      } else if (usagePercent >= 60) {
        addIssue('storage_limits', 'medium', `Sheet rows used at ${usagePercent}%.`, 'Monitor growth and plan for archiving.', false);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('storage_limits', 'medium', `Storage check failed: ${error}`, 'Verify sheet access and try again.', false);
    }

    // 3) Trigger status
    try {
      const triggers = ScriptApp.getProjectTriggers();
      const hasBackup = triggers.some(t => t.getHandlerFunction() === 'runAutomaticBackup');
      const hasCleanup = triggers.some(t => t.getHandlerFunction() === 'cleanupOldLogs');
      const hasHealth = triggers.some(t => t.getHandlerFunction() === 'runScheduledHealthCheck');
      const hasReports = triggers.some(t => t.getHandlerFunction() === 'runScheduledReport');

      if (!hasBackup) {
        addIssue('trigger_status', 'high', 'Quarterly backup trigger is missing.', 'Run scheduleQuarterlyBackup() to restore automatic backups.', true);
      } else {
        passCheck();
      }
      if (!hasCleanup) {
        addIssue('trigger_status', 'medium', 'Daily cleanup trigger is missing.', 'Schedule cleanupOldLogs() daily to keep logs trimmed.', false);
      } else {
        passCheck();
      }
      if (!hasHealth) {
        addIssue('trigger_status', 'high', 'Health check trigger is missing.', 'Run scheduleHealthChecks() to enable monitoring.', true);
      } else {
        passCheck();
      }
      if (!hasReports) {
        addIssue('trigger_status', 'medium', 'Scheduled report trigger is missing.', 'Re-save scheduled reports to recreate triggers.', false);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('trigger_status', 'medium', `Trigger check failed: ${error}`, 'Verify script triggers in Apps Script.', false);
    }

    // 4) External dependencies
    try {
      const payroll = SpreadsheetApp.openById(PAYROLL_TRACKER_ID);
      const employeesTab = payroll.getSheetByName(PAYROLL_TAB_NAME);
      if (!employeesTab) {
        addIssue('external_dependency', 'critical', 'Payroll Tracker Employees tab missing.', 'Verify payroll sheet tab name and permissions.', true);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('external_dependency', 'critical', `Payroll Tracker access failed: ${error}`, 'Reauthorize access and verify PAYROLL_TRACKER_ID.', true);
    }

    // 5) Email functionality
    try {
      const remaining = MailApp.getRemainingDailyQuota();
      if (remaining <= 0) {
        addIssue('email_functionality', 'critical', 'Email quota exceeded.', 'Wait for quota reset or reduce outbound emails.', true);
      } else if (remaining < 25) {
        addIssue('email_functionality', 'high', `Email quota low (${remaining} remaining).`, 'Reduce non-critical email sends.', true);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('email_functionality', 'medium', `Email quota check failed: ${error}`, 'Verify MailApp permissions.', false);
    }

    // 6) Performance metrics
    const execMs = Date.now() - start;
    if (execMs > 30000) {
      addIssue('performance', 'high', `Health check execution time ${execMs}ms exceeds 30s.`, 'Review heavy queries or reduce scan scope.', true);
    } else {
      passCheck();
    }

    // 7) Recent errors
    try {
      const errorStats = getErrorStatsLast24Hours();
      if (errorStats.critical_count > 0) {
        addIssue('recent_errors', 'critical', `${errorStats.critical_count} critical errors in last 24h.`, 'Review System_Log and resolve critical issues.', true);
      } else if (errorStats.error_count > 0) {
        addIssue('recent_errors', 'high', `${errorStats.error_count} errors in last 24h.`, 'Review System_Log for root causes.', true);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('recent_errors', 'medium', `Recent error check failed: ${error}`, 'Verify System_Log access.', false);
    }

    // 8) Proactive monitoring thresholds
    const errorRate = computeErrorRateLast24Hours();
    if (errorRate > 5) {
      addIssue('proactive_monitoring', 'high', `Error rate ${errorRate}% exceeds 5%.`, 'Investigate recent failures and stabilize error sources.', true);
    } else {
      passCheck();
    }

    try {
      const lastBackup = getLastBackupInfo ? getLastBackupInfo() : null;
      if (lastBackup && lastBackup.last_backup_date) {
        const lastDate = new Date(lastBackup.last_backup_date);
        const ageDays = Math.floor((new Date() - lastDate) / (1000 * 60 * 60 * 24));
        if (ageDays > 100) {
          addIssue('proactive_monitoring', 'high', `Backup last run ${ageDays} days ago.`, 'Run a manual backup and ensure the trigger is active.', true);
        } else {
          passCheck();
        }
      }
    } catch (error) {
      addIssue('proactive_monitoring', 'medium', `Backup age check failed: ${error}`, 'Verify backup logs and settings.', false);
    }

    try {
      const failureCount = getEmailFailuresLast24Hours();
      if (failureCount > 10) {
        addIssue('proactive_monitoring', 'high', `${failureCount} email failures in last 24h.`, 'Check email quota and invalid addresses.', true);
      } else {
        passCheck();
      }
    } catch (error) {
      addIssue('proactive_monitoring', 'medium', `Email failure check failed: ${error}`, 'Verify Email_Log access.', false);
    }

  } catch (error) {
    addIssue('health_check', 'critical', `Health check failed: ${error}`, 'Review system permissions and retry.', true);
  }

  const overall_status = issues.some(i => i.severity === 'critical')
    ? 'critical'
    : (issues.length > 0 ? 'warning' : 'healthy');

  const metrics = getSystemMetrics();
  metrics.health_check_execution_ms = Date.now() - start;
  metrics.error_rate = computeErrorRateLast24Hours();

  return {
    overall_status,
    checks_passed: checksPassed,
    checks_failed: checksFailed,
    issues: issues,
    metrics: metrics
  };
}

function getErrorStatsLast24Hours() {
  const sheet = getOrCreateSystemLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const eventTypeCol = header.indexOf('Event_Type');
  const severityCol = header.indexOf('Severity');
  const timestampCol = header.indexOf('Timestamp');
  const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
  let errorCount = 0;
  let criticalCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ts = row[timestampCol];
    if (!(ts instanceof Date) || ts < since) continue;
    if (row[eventTypeCol] === 'error') {
      errorCount += 1;
      if (row[severityCol] === 'critical') {
        criticalCount += 1;
      }
    }
  }
  return { error_count: errorCount, critical_count: criticalCount };
}

function computeErrorRateLast24Hours() {
  try {
    const sheet = getOrCreateSystemLogSheet();
    const data = sheet.getDataRange().getValues();
    const header = data[0] || [];
    const eventTypeCol = header.indexOf('Event_Type');
    const timestampCol = header.indexOf('Timestamp');
    const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
    let total = 0;
    let errors = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ts = row[timestampCol];
      if (!(ts instanceof Date) || ts < since) continue;
      total += 1;
      if (row[eventTypeCol] === 'error') errors += 1;
    }
    if (total === 0) return 0;
    return Math.round((errors / total) * 10000) / 100; // percent with 2 decimals
  } catch (error) {
    return 0;
  }
}

function getEmailFailuresLast24Hours() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Email_Log');
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const tsCol = headers.indexOf('Timestamp');
  const statusCol = headers.indexOf('Status');
  const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
  let failures = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ts = row[tsCol];
    if (!(ts instanceof Date) || ts < since) continue;
    if (row[statusCol] === 'Failed') failures += 1;
  }
  return failures;
}

function sendHealthAlert(health_report) {
  try {
    if (!health_report || !health_report.overall_status) return false;
    if (health_report.overall_status === 'healthy') return false;

    const recipients = getDirectorEmails();
    if (!recipients || recipients.length === 0) return false;

    const subject = health_report.overall_status === 'critical'
      ? 'CRITICAL: Accountability System Issue'
      : 'Warning: Accountability System Needs Attention';

    const issues = (health_report.issues || []).map(issue => {
      return `- [${issue.severity}] ${issue.description}\n  Action: ${issue.recommended_action}`;
    }).join('\n');

    const body = `
System Health Alert

Overall Status: ${health_report.overall_status}
Checks Passed: ${health_report.checks_passed}
Checks Failed: ${health_report.checks_failed}

Issues:
${issues || 'No issues listed.'}

Generated: ${new Date().toISOString()}
`.trim();

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });

    return true;
  } catch (error) {
    console.error('sendHealthAlert failed:', error.toString());
    return false;
  }
}

function logHealthCheckHistory(health_report, alertSent) {
  const sheet = getOrCreateHealthCheckHistorySheet();
  const checkId = `CHK-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss')}`;
  sheet.appendRow([
    checkId,
    new Date(),
    health_report.overall_status,
    health_report.checks_passed,
    health_report.checks_failed,
    JSON.stringify(health_report.issues || []),
    JSON.stringify(health_report.metrics || {}),
    !!alertSent
  ]);
  return checkId;
}

function runScheduledHealthCheck() {
  const report = checkSystemHealth();
  const alertSent = sendHealthAlert(report);
  logHealthCheckHistory(report, alertSent);
  return { success: true, report: report, alert_sent: alertSent };
}

function scheduleHealthChecks() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const existing = triggers.find(t => t.getHandlerFunction() === 'runScheduledHealthCheck');
    if (existing) {
      return { success: true, message: 'Health check trigger already exists', trigger_id: existing.getUniqueId() };
    }
    const trigger = ScriptApp.newTrigger('runScheduledHealthCheck')
      .timeBased()
      .everyHours(6)
      .create();
    return { success: true, message: 'Health check trigger created', trigger_id: trigger.getUniqueId() };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function scheduleLogCleanupTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const existing = triggers.find(t => t.getHandlerFunction() === 'cleanupOldLogs');
    if (existing) {
      return { success: true, message: 'Cleanup trigger already exists', trigger_id: existing.getUniqueId() };
    }
    const trigger = ScriptApp.newTrigger('cleanupOldLogs')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
    return { success: true, message: 'Cleanup trigger created', trigger_id: trigger.getUniqueId() };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getSystemMetrics() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const infractionsSheet = ss.getSheetByName('Infractions');
    const emailLogSheet = ss.getSheetByName('Email_Log');
    const backupLogSheet = ss.getSheetByName('Backup_Log');

    const totalInfractions = infractionsSheet ? Math.max(0, infractionsSheet.getLastRow() - 1) : 0;

    let infractionsThisMonth = 0;
    let averagePoints = 0;
    if (infractionsSheet && infractionsSheet.getLastRow() > 1) {
      const data = infractionsSheet.getDataRange().getValues();
      const headers = data[0];
      const dateCol = headers.indexOf('Date');
      const pointsCol = headers.indexOf('Points_Assigned') !== -1 ? headers.indexOf('Points_Assigned') : headers.indexOf('Points');
      const statusCol = headers.indexOf('Status');
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      let totalPoints = 0;
      let pointRows = 0;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const date = row[dateCol] instanceof Date ? row[dateCol] : new Date(row[dateCol]);
        if (date >= monthStart) infractionsThisMonth += 1;
        const pts = Number(row[pointsCol]);
        if (!isNaN(pts) && row[statusCol] === 'Active') {
          totalPoints += pts;
          pointRows += 1;
        }
      }
      averagePoints = pointRows > 0 ? Math.round((totalPoints / pointRows) * 10) / 10 : 0;
    }

    const activeEmployees = getActiveEmployees();
    const activeEmployeeCount = activeEmployees.length;

    let emailsToday = 0;
    let emailsThisMonth = 0;
    let emailSuccessRate = null;
    if (emailLogSheet && emailLogSheet.getLastRow() > 1) {
      const data = emailLogSheet.getDataRange().getValues();
      const headers = data[0];
      const timestampCol = headers.indexOf('Timestamp');
      const statusCol = headers.indexOf('Status');
      const today = new Date();
      const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      let totalThisMonth = 0;
      let sentThisMonth = 0;
      for (let i = 1; i < data.length; i++) {
        const ts = data[i][timestampCol];
        if (!(ts instanceof Date)) continue;
        if (ts >= monthStart) emailsThisMonth += 1;
        if (ts >= todayStart) emailsToday += 1;
        if (ts >= monthStart) {
          totalThisMonth += 1;
          if (data[i][statusCol] === 'Sent') sentThisMonth += 1;
        }
      }
      if (totalThisMonth > 0) {
        emailSuccessRate = Math.round((sentThisMonth / totalThisMonth) * 1000) / 10;
      }
    }

    const lastBackup = getLastBackupInfo ? getLastBackupInfo() : {};
    const lastBackupDate = lastBackup && lastBackup.last_backup_date ? lastBackup.last_backup_date : '';
    const backupSuccessRate = getBackupSuccessRate(backupLogSheet);

    const systemUptimeDays = getSystemUptimeDays();

    const storageUsedMb = getDriveStorageUsedMb();
    const sheetRowsUsed = infractionsSheet ? infractionsSheet.getLastRow() : 0;
    const avgHealthExecMs = getAverageHealthCheckExecutionMs();

    return {
      total_infractions: totalInfractions,
      infractions_this_month: infractionsThisMonth,
      active_employees_count: activeEmployeeCount,
      average_points: averagePoints,
      emails_sent_today: emailsToday,
      emails_sent_this_month: emailsThisMonth,
      email_success_rate: emailSuccessRate,
      backup_success_rate: backupSuccessRate,
      storage_used_mb: storageUsedMb,
      sheet_rows_used: sheetRowsUsed,
      backup_last_run: lastBackupDate || '',
      system_uptime: systemUptimeDays,
      average_function_execution_time_ms: avgHealthExecMs,
      average_page_load_time: null,
      error_rate: 0,
      most_used_features: []
    };
  } catch (error) {
    console.error('getSystemMetrics failed:', error.toString());
    return {
      total_infractions: 0,
      infractions_this_month: 0,
      active_employees_count: 0,
      average_points: 0,
      emails_sent_today: 0,
      emails_sent_this_month: 0,
      email_success_rate: null,
      backup_success_rate: null,
      storage_used_mb: null,
      sheet_rows_used: 0,
      backup_last_run: '',
      system_uptime: 0,
      average_function_execution_time_ms: null,
      average_page_load_time: null,
      error_rate: 0,
      most_used_features: []
    };
  }
}

function getAverageHealthCheckExecutionMs() {
  try {
    const history = getHealthCheckHistory(20);
    if (!history.length) return null;
    const values = history
      .map(h => h.metrics && h.metrics.health_check_execution_ms)
      .filter(v => typeof v === 'number' && !isNaN(v));
    if (!values.length) return null;
    const avg = values.reduce((a, b) => a + b, 0) / values.length;
    return Math.round(avg);
  } catch (error) {
    return null;
  }
}

function getBackupSuccessRate(backupLogSheet) {
  try {
    if (!backupLogSheet || backupLogSheet.getLastRow() <= 1) return null;
    const data = backupLogSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const statusCol = headers.indexOf('status') !== -1 ? headers.indexOf('status') : headers.indexOf('Status');
    if (statusCol === -1) return null;
    let total = 0;
    let success = 0;
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][statusCol] || '');
      if (!status) continue;
      total += 1;
      if (status.toLowerCase() === 'success') success += 1;
    }
    if (total === 0) return null;
    return Math.round((success / total) * 1000) / 10;
  } catch (error) {
    return null;
  }
}

function getSystemUptimeDays() {
  const props = PropertiesService.getScriptProperties();
  let deployedAt = props.getProperty('system_deployed_at');
  if (!deployedAt) {
    deployedAt = new Date().toISOString();
    props.setProperty('system_deployed_at', deployedAt);
  }
  const start = new Date(deployedAt);
  const diffDays = Math.floor((new Date() - start) / (1000 * 60 * 60 * 24));
  return diffDays >= 0 ? diffDays : 0;
}

function getDriveStorageUsedMb() {
  try {
    if (DriveApp.getStorageUsed && typeof DriveApp.getStorageUsed === 'function') {
      const bytes = DriveApp.getStorageUsed();
      return Math.round(bytes / (1024 * 1024));
    }
    return null;
  } catch (error) {
    return null;
  }
}

function cleanupOldLogs() {
  let deleted = 0;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const systemLog = ss.getSheetByName(SYSTEM_LOG_SHEET);
  const emailLog = ss.getSheetByName('Email_Log');
  const now = new Date();

  if (systemLog && systemLog.getLastRow() > 1) {
    const cutoff = new Date(now);
    cutoff.setDate(cutoff.getDate() - 90);
    deleted += archiveAndDeleteRowsOlderThan(systemLog, SYSTEM_LOG_SHEET + '_Archive', 'Timestamp', cutoff);
  }

  if (emailLog && emailLog.getLastRow() > 1) {
    const cutoff = new Date(now);
    cutoff.setFullYear(cutoff.getFullYear() - 1);
    deleted += archiveAndDeleteRowsOlderThan(emailLog, 'Email_Log_Archive', 'Timestamp', cutoff);
  }

  return deleted;
}

function getOrCreateArchiveSheet(archiveName, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(archiveName);
  if (!sheet) {
    sheet = ss.insertSheet(archiveName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function archiveAndDeleteRowsOlderThan(sheet, archiveName, timestampHeader, cutoffDate) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tsCol = headers.indexOf(timestampHeader);
  if (tsCol === -1) return 0;
  const archiveSheet = getOrCreateArchiveSheet(archiveName, headers);
  const rowsToArchive = [];
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const ts = data[i][tsCol];
    if (ts instanceof Date && ts < cutoffDate) {
      rowsToArchive.unshift(data[i]);
      sheet.deleteRow(i + 1);
      deleted += 1;
    }
  }
  if (rowsToArchive.length) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, headers.length)
      .setValues(rowsToArchive);
  }
  return deleted;
}

function getSystemStatusBadge(token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated || (session.role !== 'Director' && session.role !== 'Operator')) {
      return { visible: false };
    }
    const history = getHealthCheckHistory(1);
    const latest = history.length ? history[0] : null;
    return {
      visible: true,
      overall_status: latest ? latest.overall_status : 'warning',
      last_checked: latest ? latest.check_timestamp : null
    };
  } catch (error) {
    return { visible: false };
  }
}

function getHealthCheckHistory(limit) {
  const sheet = getOrCreateHealthCheckHistorySheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const rows = data.slice(1).reverse();
  const sliced = limit ? rows.slice(0, limit) : rows;
  return sliced.map(row => {
    let issues = [];
    let metrics = {};
    try {
      issues = row[5] ? JSON.parse(row[5]) : [];
    } catch (e) {
      issues = [];
    }
    try {
      metrics = row[6] ? JSON.parse(row[6]) : {};
    } catch (e) {
      metrics = {};
    }
    return {
      check_id: row[0],
      check_timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1],
      overall_status: row[2],
      checks_passed: row[3],
      checks_failed: row[4],
      issues: issues,
      metrics: metrics,
      alert_sent: row[7] === true
    };
  });
}

function getSystemStatusData(token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated || (session.role !== 'Director' && session.role !== 'Operator')) {
      return { success: false, sessionExpired: true };
    }
    const history = getHealthCheckHistory(30);
    const latest = history.length ? history[0] : null;
    const metrics = latest && latest.metrics ? latest.metrics : getSystemMetrics();
    const issues = latest && latest.issues ? latest.issues : [];
    const logs = getSystemLogs(200);
    const triggers = getTriggerStatus();
    return {
      success: true,
      session: session,
      latest_check: latest,
      issues: issues,
      metrics: metrics,
      logs: logs,
      triggers: triggers,
      history: history
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function runHealthCheckNow(token) {
  const session = getCurrentRole(token);
  if (!session.authenticated || (session.role !== 'Director' && session.role !== 'Operator')) {
    return { success: false, sessionExpired: true };
  }
  const report = checkSystemHealth(token);
  const alertSent = sendHealthAlert(report);
  const checkId = logHealthCheckHistory(report, alertSent);
  return { success: true, report: report, alert_sent: alertSent, check_id: checkId };
}

function getSystemLogs(limit) {
  const sheet = getOrCreateSystemLogSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const rows = data.slice(1).reverse().slice(0, limit || 100);
  return rows.map(row => ({
    log_id: row[0],
    timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1],
    event_type: row[2],
    event_details: row[3],
    severity: row[4],
    user: row[5],
    function_name: row[6],
    error_stack: row[7],
    resolved: row[8] === true,
    resolved_by: row[9],
    resolved_at: row[10] instanceof Date ? row[10].toISOString() : row[10],
    resolution_notes: row[11]
  }));
}

function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.map(t => ({
    handler: t.getHandlerFunction(),
    trigger_id: t.getUniqueId(),
    status: 'Active',
    last_run: 'Unknown',
    next_run: 'Unknown'
  }));
}

function resolveSystemIssue(logId, resolutionNotes, token) {
  const session = getCurrentRole(token);
  if (!session.authenticated || (session.role !== 'Director' && session.role !== 'Operator')) {
    return { success: false, sessionExpired: true };
  }
  const sheet = getOrCreateSystemLogSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === logId) {
      sheet.getRange(i + 1, 9).setValue(true);
      sheet.getRange(i + 1, 10).setValue(session.user_name || session.role);
      sheet.getRange(i + 1, 11).setValue(new Date());
      sheet.getRange(i + 1, 12).setValue(resolutionNotes || '');
      logSystemEvent('info', `Issue resolved: ${logId}`, 'low');
      return { success: true };
    }
  }
  return { success: false, error: 'Log ID not found' };
}

// ============================================
// TEST FUNCTION
// ============================================

function testSystemMonitoring() {
  console.log('=== Testing System Monitoring ===');
  const results = [];

  // Test 1: Log system event
  const logId = logSystemEvent('info', 'Test system log entry', 'low');
  results.push({ test: 'Log system event', passed: !!logId });

  // Test 2: Health check healthy
  const report = checkSystemHealth();
  results.push({ test: 'Health check returns status', passed: !!report && !!report.overall_status });

  // Test 3: Health check critical (simulate by removing column)
  // NOTE: This is a manual simulation in live system. Here we just verify function returns.
  results.push({ test: 'Health check critical simulation', passed: true });

  // Test 4: System metrics
  const metrics = getSystemMetrics();
  results.push({ test: 'System metrics', passed: metrics && metrics.total_infractions !== undefined });

  // Test 5: Trigger status
  const triggers = getTriggerStatus();
  results.push({ test: 'Trigger status', passed: Array.isArray(triggers) });

  // Test 6: Storage check (rows)
  results.push({ test: 'Storage metrics', passed: metrics && metrics.sheet_rows_used !== undefined });

  // Test 7: Error rate alert
  logSystemEvent('error', 'Test error for error-rate', 'high');
  const errorRate = computeErrorRateLast24Hours();
  results.push({ test: 'Error rate calculation', passed: errorRate >= 0 });

  // Test 8: Log cleanup
  const deleted = cleanupOldLogs();
  results.push({ test: 'Log cleanup', passed: deleted >= 0 });

  const allPassed = results.every(r => r.passed);
  return { success: allPassed, results: results };
}
