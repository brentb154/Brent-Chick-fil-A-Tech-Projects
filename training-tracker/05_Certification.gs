/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Certification
 * ============================================================
 * Handles certification form launch, form response processing,
 * and archiving certified trainees.
 */

// -- Open Certification Form ----------------------------------

/**
 * Opens the correct certification form (FOH or BOH) with
 * the trainee's name pre-filled.
 *
 * NOTE: You must update the entry.XXXXX parameter to match
 * the actual entry ID from your Google Form. To find it:
 *   1. Open the form
 *   2. Click "Get pre-filled link"
 *   3. Fill in the name field -> Generate link
 *   4. Copy the entry.XXXXXXXXX value from the URL
 */
function openCertificationForm(traineeName, house) {
  // -- UPDATE THESE URLs WITH YOUR ACTUAL FORM URLs --------
  var fohFormUrl = 'https://docs.google.com/forms/d/e/1FAIpQLSc9GspJgkpwivBT-G_pTOmdj0CnkC_RhJPGXeWBiBlmkQAs7w/viewform';
  var bohFormUrl = 'https://docs.google.com/forms/d/e/1FAIpQLSe2Plx8TXue29gAv_ohFCqrKt9QtKYeygvu_tGnxKnKUO6DKQ/viewform';

  var formUrl = (house === 'FOH') ? fohFormUrl : bohFormUrl;

  // -- UPDATE THIS ENTRY ID TO MATCH YOUR FORM -------------
  // Replace 123456789 with the actual entry ID for the name field
  formUrl += '?entry.123456789=' + encodeURIComponent(traineeName);

  var html = '<script>window.open("' + formUrl + '", "_blank"); google.script.host.close();</script>';
  var ui   = HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Opening Certification Form...');
}


// -- Certification Form Response ------------------------------

/**
 * Triggered when the certification form is submitted.
 * Archives the trainee in Certification Log and removes
 * them from the active dashboard.
 *
 * TRIGGER SETUP: Installable trigger on certification form -> onCertificationFormSubmit
 *
 * NOTE: Adjust the e.values indices to match your form's question order.
 */
function onCertificationFormSubmit(e) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var certSheet = ss.getSheetByName('Certification Log');
    var logSheet  = ss.getSheetByName('Daily Training Log');

    // Parse form response - adjust indices as needed
    var traineeName = String(e.values[1]).trim(); // Adjust index
    var house       = String(e.values[2]).trim(); // FOH or BOH

    var certDate = new Date();
    var stats    = getTraineeStats(traineeName, logSheet);

    certSheet.appendRow([
      traineeName,
      house,
      Utilities.formatDate(certDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      stats.totalHours.toFixed(1),
      stats.durationDays,
      e.values[0] || '',    // Form response ID / timestamp
      'Certified via form'
    ]);

    // Refresh dashboard (will now exclude this trainee)
    updateDashboard();

    // Alert
    sendAlert('Trainee Ready for Certification',
      traineeName + ' completed ' + house + ' certification! ');

  } catch (err) {
    Logger.log('onCertificationFormSubmit error: ' + err.message);
  }
}


/**
 * Manually certify a trainee (in case you're not using a form).
 * Adds them to the Certification Log directly.
 */
function manuallyCertifyTrainee() {
  var ui = SpreadsheetApp.getUi();

  var nameResponse = ui.prompt('Certify Trainee', 'Enter the trainee\'s full name:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;

  var name = nameResponse.getResponseText().trim();
  if (!name) return;

  var houseResponse = ui.prompt('House', 'Enter FOH or BOH:', ui.ButtonSet.OK_CANCEL);
  if (houseResponse.getSelectedButton() !== ui.Button.OK) return;

  var house = houseResponse.getResponseText().trim().toUpperCase();
  if (house !== 'FOH' && house !== 'BOH') {
    ui.alert('Invalid house. Please enter FOH or BOH.');
    return;
  }

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var certSheet = ss.getSheetByName('Certification Log');
  var logSheet  = ss.getSheetByName('Daily Training Log');

  var stats = getTraineeStats(name, logSheet);

  certSheet.appendRow([
    name,
    house,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy'),
    stats.totalHours.toFixed(1),
    stats.durationDays,
    'Manual',
    'Manually certified'
  ]);

  updateDashboard();

  ui.alert(name + ' has been certified and archived.');
}


// -- Stats Helper ---------------------------------------------

/**
 * Returns total hours and duration in days for a trainee.
 */
function getTraineeStats(traineeName, logSheet) {
  if (logSheet.getLastRow() < 2) {
    return { totalHours: 0, durationDays: 0 };
  }

  var data       = logSheet.getDataRange().getValues();
  var firstDate  = null;
  var lastDate   = null;
  var totalHours = 0;

  data.slice(1).forEach(function (row) {
    var name = String(row[7]).trim(); // Canonical name
    if (!name) name = String(row[2]).trim();

    if (name === traineeName) {
      var date  = new Date(row[1]);
      var hours = parseFloat(row[4]) || 0;

      if (!firstDate || date < firstDate) firstDate = date;
      if (!lastDate  || date > lastDate)  lastDate  = date;
      totalHours += hours;
    }
  });

  var durationDays = (firstDate && lastDate)
    ? Math.round((lastDate - firstDate) / (1000 * 60 * 60 * 24))
    : 0;

  return { totalHours: totalHours, durationDays: durationDays };
}
