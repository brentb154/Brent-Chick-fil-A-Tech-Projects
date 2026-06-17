/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Alerts & Daily Triggers
 * ============================================================
 * Email alert system with configurable recipients,
 * daily reminders, inactive trainee detection, and
 * milestone notifications.
 */

// -- Core Alert Sender ----------------------------------------

/**
 * Sends an email alert based on the Alert Settings sheet.
 * Checks if the alert type is enabled and has valid recipients.
 *
 * @param {string} alertType  Must match a row in Alert Settings
 * @param {string} message    Body text for the email
 */
function sendAlert(alertType, message) {
  try {
    var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Alert Settings');
    if (!settingsSheet || settingsSheet.getLastRow() < 2) return;

    var data = settingsSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== alertType) continue;

      var enabled    = data[i][1];
      var recipient1 = String(data[i][2] || '').trim();
      var recipient2 = String(data[i][3] || '').trim();
      var recipient3 = String(data[i][4] || '').trim();

      if (enabled !== true && enabled !== 'TRUE') return;

      var recipients = [recipient1, recipient2, recipient3]
        .filter(function (r) { return r && r.indexOf('@') > -1; })
        .join(',');

      if (!recipients) {
        Logger.log('sendAlert: "' + alertType + '" is enabled but has no recipient emails configured. Add emails in Alert Settings.');
        return;
      }

      var subject = ' Training Alert: ' + alertType;
      var body    = message + '\n\n' +
        '---------------------------\n' +
        'View Training Tracker: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '\n' +
        'To adjust alerts: Training Tools -> Alert Settings';

      MailApp.sendEmail({
        to:      recipients,
        subject: subject,
        body:    body
      });

      return;
    }
  } catch (err) {
    Logger.log('sendAlert error (' + alertType + '): ' + err.message);
  }
}


// -- Milestone Alerts -----------------------------------------

/**
 * Called after each form submission to check whether the
 * new entry pushed a trainee past a position minimum or
 * made them certification-ready.
 */
function checkMilestoneAlerts(traineeName, position, newHours) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Daily Training Log');
  var reqSheet = ss.getSheetByName('Position Requirements');

  // Total hours on this position BEFORE the new entry
  var positionsMap     = getTraineePositions(traineeName, logSheet);
  var totalPosHours    = positionsMap[position] || 0;
  var previousPosHours = totalPosHours - newHours;

  // Check if they just crossed the minimum for this position
  var reqData = reqSheet.getDataRange().getValues();
  for (var i = 1; i < reqData.length; i++) {
    if (reqData[i][1] === position) {
      var minHours = reqData[i][2];

      if (totalPosHours >= minHours && previousPosHours < minHours) {
        sendAlert('Position Completion Milestone',
          traineeName + ' completed ' + position + '! (' + totalPosHours.toFixed(1) + ' hours logged, ' + minHours + ' required)');
      }
      break;
    }
  }

  // Check if now certification-ready
  var house = inferHouse(position);
  if (isCertificationReady(traineeName, house, positionsMap)) {
    // Were they ready BEFORE this entry?
    var prevPositions     = Object.assign({}, positionsMap);
    prevPositions[position] = previousPosHours;

    if (!isCertificationReady(traineeName, house, prevPositions)) {
      sendAlert('Trainee Ready for Certification',
        ' ' + traineeName + ' has completed ALL ' + house + ' position minimums and is ready for certification!');
    }
  }
}


// -- Daily Triggers -------------------------------------------

/**
 * Returns training rows from Daily Training Log, or Form_Responses as a
 * fallback, so alerts see the SAME data the dashboard does.
 * Each row: [timestamp, date, rawName, position, hours, onTrack, notes, name].
 */
function getActiveLogRows(ss) {
  var log = ss.getSheetByName('Daily Training Log');
  if (log && log.getLastRow() > 1) {
    return log.getRange(2, 1, log.getLastRow() - 1, 8).getValues();
  }
  var names = ['Form_Responses', 'Form Responses 1', 'Form responses 1', 'Form_Responses_1'];
  for (var i = 0; i < names.length; i++) {
    var fs = ss.getSheetByName(names[i]);
    if (fs && fs.getLastRow() > 1) {
      return fs.getRange(2, 1, fs.getLastRow() - 1, 7).getValues().map(function (r) {
        return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], String(r[2] || '').trim()];
      });
    }
  }
  return [];
}

/**
 * Sends a reminder if no training was logged today.
 * TRIGGER: Time-driven -> Day timer -> 8-9 PM
 */
function dailyReminderCheck() {
  try {
    var ss           = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Alert Settings');

    // Check if alert is enabled
    var data = settingsSheet.getDataRange().getValues();
    var enabled = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'Daily Training Log Reminder') {
        enabled = data[i][1];
        break;
      }
    }

    if (enabled !== true && enabled !== 'TRUE') return;

    // Use the same data source the dashboard does (log, or Form_Responses fallback)
    var rows = getActiveLogRows(ss);
    if (rows.length === 0) {
      sendAlert('Daily Training Log Reminder',
        'No training has been logged today. Please ensure all training activities are recorded.');
      return;
    }

    // Check for entries today
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var hasEntriesToday = false;

    rows.forEach(function (row) {
      var entryDate = new Date(row[1]);
      entryDate.setHours(0, 0, 0, 0);

      if (entryDate.getTime() === today.getTime()) {
        hasEntriesToday = true;
      }
    });

    if (!hasEntriesToday) {
      sendAlert('Daily Training Log Reminder',
        'No training has been logged today (' +
        Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMM d') +
        '). Please ensure all training activities are recorded.');
    }
  } catch (err) {
    Logger.log('dailyReminderCheck error: ' + err.message);
  }
}

/**
 * Flags trainees who haven't trained in 3+ days.
 * TRIGGER: Time-driven -> Day timer -> 6-7 AM
 */
function checkInactiveTrainees() {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var certSheet = ss.getSheetByName('Certification Log');

    // Same data source as the dashboard (log, or Form_Responses fallback)
    var rows = getActiveLogRows(ss);
    if (rows.length === 0) return;

    // Certified names to exclude (normalized for case/spacing)
    var certifiedKeys = {};
    if (certSheet && certSheet.getLastRow() > 1) {
      certSheet.getRange(2, 1, certSheet.getLastRow() - 1, 1)
        .getValues().forEach(function (r) {
          var k = String(r[0]).trim().toLowerCase();
          if (k) certifiedKeys[k] = true;
        });
    }

    // Build last-training-date map
    var lastDates = {};

    rows.forEach(function (row) {
      var name = String(row[7]).trim();
      if (!name) name = String(row[2]).trim();
      if (!name) return;

      if (certifiedKeys[name.toLowerCase()]) return;

      var date = new Date(row[1]);
      if (!lastDates[name] || date > lastDates[name]) {
        lastDates[name] = date;
      }
    });

    // Find inactive trainees
    var today            = new Date();
    var inactiveTrainees = [];

    Object.keys(lastDates).forEach(function (name) {
      var daysSince = Math.round((today - lastDates[name]) / (1000 * 60 * 60 * 24));

      if (daysSince >= 3) {
        inactiveTrainees.push('  • ' + name + ' - ' + daysSince + ' days since last training');
      }
    });

    if (inactiveTrainees.length > 0) {
      sendAlert('Trainee Inactive 3+ Days',
        'The following trainees have not trained recently:\n\n' +
        inactiveTrainees.join('\n') +
        '\n\nPlease follow up to ensure they stay on track.');
    }
  } catch (err) {
    Logger.log('checkInactiveTrainees error: ' + err.message);
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: 'Training Tracker: checkInactiveTrainees FAILED',
        body: 'Error: ' + err.message + '\n\nStack: ' + err.stack
      });
    } catch (mailErr) { Logger.log('Could not send failure alert: ' + mailErr.message); }
  }
}


// -- Trigger Setup ------------------------------------------------

/**
 * Sets up daily time-driven triggers for reminder and inactive checks.
 * Safe to re-run — removes existing triggers before creating new ones.
 */
function setupDailyTriggers() {
  var functions = ['dailyReminderCheck', 'checkInactiveTrainees'];

  // Remove existing triggers for these functions
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (functions.indexOf(trigger.getHandlerFunction()) > -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Daily reminder at 8-9 PM
  ScriptApp.newTrigger('dailyReminderCheck')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();

  // Inactive check at 6-7 AM
  ScriptApp.newTrigger('checkInactiveTrainees')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert(
    'Daily alert triggers set!\n\n' +
    '• Daily Reminder Check: ~8 PM each day\n' +
    '• Inactive Trainee Check: ~6 AM each day\n\n' +
    'Make sure recipient emails are configured in\n' +
    'Training Tools -> Alert Settings.');
}

/**
 * Sets up ALL automation triggers (daily alerts + Monday populate).
 * Safe to re-run.
 */
function setupAllTriggers() {
  var functions = ['dailyReminderCheck', 'checkInactiveTrainees', 'mondayAutoPopulate'];

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (functions.indexOf(trigger.getHandlerFunction()) > -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Daily reminder at 8-9 PM
  ScriptApp.newTrigger('dailyReminderCheck')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();

  // Inactive check at 6-7 AM
  ScriptApp.newTrigger('checkInactiveTrainees')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  // Monday auto-populate at 5 AM
  ScriptApp.newTrigger('mondayAutoPopulate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'All triggers set!\n\n' +
    '• Daily Reminder Check: ~8 PM\n' +
    '• Inactive Trainee Check: ~6 AM\n' +
    '• Monday Training Needs: ~5 AM Monday\n\n' +
    'Make sure recipient emails are configured in\n' +
    'Training Tools -> Alert Settings.');
}
