/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Data Sync & Backfill
 * ============================================================
 * Imports data from Form_Responses into Daily Training Log,
 * backfills canonical names, and syncs to Training Needs tabs.
 */

// -- Import Form Responses ------------------------------------

/**
 * Copies all form response data into Daily Training Log,
 * backfills canonical names for every row, and deduplicates.
 * Safe to run multiple times - skips rows already imported.
 *
 * Accessible via Training Tools -> Sync Form Data
 */
function syncFormData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Daily Training Log');

  if (!logSheet) {
    SpreadsheetApp.getUi().alert('Error: "Daily Training Log" sheet not found. Run Initial Setup first.');
    return;
  }

  // Find the form responses sheet (try common names)
  var formSheet = null;
  var tryNames = ['Form_Responses', 'Form Responses 1', 'Form responses 1', 'Form_Responses_1'];
  for (var i = 0; i < tryNames.length; i++) {
    formSheet = ss.getSheetByName(tryNames[i]);
    if (formSheet && formSheet.getLastRow() > 1) break;
    formSheet = null;
  }

  if (!formSheet) {
    SpreadsheetApp.getUi().alert(
      'No form response sheet found.\n\n' +
      'Looked for: ' + tryNames.join(', ') + '\n\n' +
      'If your sheet has a different name, rename it to "Form_Responses" and try again.'
    );
    return;
  }

  // Get existing timestamps in Daily Training Log to avoid duplicates
  var existingTimestamps = {};
  if (logSheet.getLastRow() > 1) {
    var existingData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues();
    existingData.forEach(function (row) {
      if (row[0]) existingTimestamps[String(row[0])] = true;
    });
  }

  // Read form response data
  var formData = formSheet.getRange(2, 1, formSheet.getLastRow() - 1, formSheet.getLastColumn()).getValues();
  var imported = 0;
  var skipped = 0;

  formData.forEach(function (row) {
    var timestamp = String(row[0]);
    if (!timestamp || timestamp === 'undefined') return;

    // Skip if already imported
    if (existingTimestamps[timestamp]) {
      skipped++;
      return;
    }

    // Map form columns to Daily Training Log columns
    // Form: A=Timestamp, B=Date, C=Name, D=Position, E=Hours, F=OnTrack, G=Notes
    var traineeName  = String(row[2] || '').trim();
    var position     = String(row[3] || '').trim();
    var hours        = row[4] || '';
    var onTrack      = String(row[5] || '').trim();
    var notes        = String(row[6] || '').trim();

    if (!traineeName) return;

    // Handle multi-position entries (e.g., "iPOS, Register")
    // Store as-is but note it
    var canonicalName = getCanonicalName(traineeName);

    // If getCanonicalName didn't resolve to a different name,
    // apply basic title-case to normalize "bald" -> "Bald"
    if (canonicalName === traineeName) {
      canonicalName = traineeName.split(' ').map(function(word) {
        if (!word) return '';
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      }).join(' ');
    }

    var newRow = [
      row[0],         // A: Timestamp
      row[1],         // B: Date
      traineeName,    // C: Trainee Name
      position,       // D: Position Trained
      hours,          // E: Hours
      onTrack,        // F: On Track
      notes,          // G: Notes
      canonicalName,  // H: Canonical Name
      false           // I: Synced to Dashboard
    ];

    logSheet.appendRow(newRow);
    imported++;
  });

  // Now backfill any rows in Daily Training Log that are missing canonical names
  backfillCanonicalNames();

  // Refresh dashboard
  updateDashboard();

  SpreadsheetApp.getUi().alert(
    'Sync Complete!\n\n' +
    'Imported: ' + imported + ' new entries\n' +
    'Skipped (already imported): ' + skipped + '\n' +
    'Source sheet: "' + formSheet.getName() + '"\n\n' +
    'Dashboard has been refreshed.'
  );
}


// -- Backfill Canonical Names ---------------------------------

/**
 * Scans Daily Training Log and fills in Column H (Canonical Name)
 * for any row where it's missing.
 *
 * Accessible via Training Tools -> Backfill Names
 */
function backfillCanonicalNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('Daily Training Log');

  if (!logSheet || logSheet.getLastRow() < 2) return;

  var data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues();
  var updates = 0;

  for (var i = 0; i < data.length; i++) {
    var traineeName   = String(data[i][2] || '').trim();  // Col C
    var canonicalName = String(data[i][7] || '').trim();  // Col H

    if (traineeName && !canonicalName) {
      var resolved = getCanonicalName(traineeName);
      logSheet.getRange(i + 2, 8).setValue(resolved);  // Write to Col H
      updates++;
    }
  }

  if (updates > 0) {
    Logger.log('Backfilled ' + updates + ' canonical names');
  }

  return updates;
}
