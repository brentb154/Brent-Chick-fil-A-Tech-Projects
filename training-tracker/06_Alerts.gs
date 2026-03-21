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

      if (enabled !== true && enabled !== 'TRUE' && enabled !== true) return;

      var recipients = [recipient1, recipient2, recipient3]
        .filter(function (r) { return r && r.indexOf('@') > -1; })
        .join(',');

      if (!recipients) return;

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
 * Sends a reminder if no training was logged today.
 * TRIGGER: Time-driven -> Day timer -> 8-9 PM
 */
function dailyReminderCheck() {
  try {
    var ss           = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Alert Settings');
    var logSheet      = ss.getSheetByName('Daily Training Log');

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
    if (logSheet.getLastRow() < 2) {
      sendAlert('Daily Training Log Reminder',
        'No training has been logged today. Please ensure all training activities are recorded.');
      return;
    }

    // Check for entries today
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var logData        = logSheet.getDataRange().getValues();
    var hasEntriesToday = false;

    logData.slice(1).forEach(function (row) {
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
    var logSheet  = ss.getSheetByName('Daily Training Log');
    var certSheet = ss.getSheetByName('Certification Log');

    if (logSheet.getLastRow() < 2) return;

    // Certified names to exclude
    var certifiedNames = [];
    if (certSheet.getLastRow() > 1) {
      certifiedNames = certSheet.getRange(2, 1, certSheet.getLastRow() - 1, 1)
        .getValues().map(function (r) { return String(r[0]).trim(); });
    }

    // Build last-training-date map
    var lastDates = {};
    var logData   = logSheet.getDataRange().getValues();

    logData.slice(1).forEach(function (row) {
      var name = String(row[7]).trim();
      if (!name) name = String(row[2]).trim();
      if (!name) return;

      if (certifiedNames.indexOf(name) > -1) return;

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
  }
}
