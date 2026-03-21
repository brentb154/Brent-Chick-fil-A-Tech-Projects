/**
 * ============================================================
 * TRAINING TRACKING SYSTEM - Form & Deduplication
 * ============================================================
 * Handles Google Form submissions, canonical name resolution,
 * fuzzy name matching, and merge operations.
 */

// -- Form Submission Handler ----------------------------------

/**
 * Triggered when Google Form is submitted.
 * Processes new training entry, resolves canonical name,
 * checks for milestones, and refreshes the dashboard.
 *
 * TRIGGER SETUP: Installable trigger -> On form submit
 */
function onFormSubmit(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Daily Training Log');
    var lastRow = sheet.getLastRow();

    // Get form response values
    // Indices depend on your form question order - adjust if needed
    var timestamp   = e.values[0];
    var date        = e.values[1];
    var traineeName = e.values[2];
    var position    = e.values[3];
    var hours       = e.values[4];
    var onTrack     = e.values[5];
    var notes       = e.values[6] || '';

    // Resolve canonical name
    var canonicalName = getCanonicalName(traineeName);
    sheet.getRange(lastRow, 8).setValue(canonicalName);   // Column H
    sheet.getRange(lastRow, 9).insertCheckboxes();
    sheet.getRange(lastRow, 9).setValue(false);            // Column I - not yet synced

    // Check for milestone alerts
    checkMilestoneAlerts(canonicalName, position, parseFloat(hours));

    // Refresh dashboard
    updateDashboard();

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.message);
  }
}


// -- Canonical Name Resolution --------------------------------

/**
 * Returns the canonical (standardized) name for a trainee.
 * Checks completed merges in the Name Deduplication sheet.
 * If no match, returns the original name trimmed.
 */
function getCanonicalName(inputName) {
  var dedupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Deduplication');
  if (!dedupSheet || dedupSheet.getLastRow() < 2) return inputName.trim();

  var data = dedupSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var canonicalName = data[i][0];
    var variants      = String(data[i][1]).split(',').map(function (v) { return v.trim(); });
    var action        = data[i][3];
    var status        = data[i][4];

    // Only use approved merges
    if (action === 'Merge' && status === 'Completed') {
      for (var j = 0; j < variants.length; j++) {
        if (inputName.trim().toLowerCase() === variants[j].toLowerCase()) {
          return canonicalName;
        }
      }
      // Also match canonical itself
      if (inputName.trim().toLowerCase() === canonicalName.toLowerCase()) {
        return canonicalName;
      }
    }
  }

  return inputName.trim();
}


// -- Duplicate Detection --------------------------------------

/**
 * Scans Daily Training Log for similar names and populates
 * the Name Deduplication sheet with suggestions.
 * Called: Daily trigger + manual via Training Tools menu.
 */
function checkForDuplicates() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet  = ss.getSheetByName('Daily Training Log');
  var dedupSheet = ss.getSheetByName('Name Deduplication');

  // Collect names from Daily Training Log, fall back to Form_Responses
  var nameValues = [];

  if (logSheet && logSheet.getLastRow() > 1) {
    nameValues = logSheet.getRange('C2:C' + logSheet.getLastRow()).getValues();
  }

  // If no data, try Form_Responses
  if (nameValues.length === 0) {
    var formSheetNames = ['Form_Responses', 'Form Responses 1', 'Form responses 1', 'Form_Responses_1'];
    for (var fi = 0; fi < formSheetNames.length; fi++) {
      var formSheet = ss.getSheetByName(formSheetNames[fi]);
      if (formSheet && formSheet.getLastRow() > 1) {
        nameValues = formSheet.getRange('C2:C' + formSheet.getLastRow()).getValues();
        break;
      }
    }
  }

  if (nameValues.length === 0) return;

  // Collect unique names and counts
  var uniqueNames = {};

  nameValues.forEach(function (row) {
    var name = String(row[0]).trim();
    if (name) {
      uniqueNames[name] = (uniqueNames[name] || 0) + 1;
    }
  });

  var nameList    = Object.keys(uniqueNames);
  var suggestions = [];

  // Compare every pair
  for (var i = 0; i < nameList.length; i++) {
    for (var j = i + 1; j < nameList.length; j++) {
      var name1 = nameList[i];
      var name2 = nameList[j];

      if (areSimilar(name1, name2)) {
        if (!isDuplicateSuggestionExists(name1, name2, dedupSheet)) {
          suggestions.push({
            canonical: name1,
            variants:  name2,
            count:     uniqueNames[name1] + uniqueNames[name2]
          });
        }
      }
    }
  }

  // Write new suggestions
  if (suggestions.length > 0) {
    var lastRow = dedupSheet.getLastRow();
    suggestions.forEach(function (sug, index) {
      dedupSheet.getRange(lastRow + index + 1, 1, 1, 5).setValues([[
        sug.canonical,
        sug.variants,
        sug.count,
        'Pending',
        'Pending'
      ]]);
    });

    // UI calls only work when a user has the spreadsheet open (not from time-based triggers)
    try {
      showDeduplicationSidebar();
    } catch (uiErr) {
      Logger.log('Skipping sidebar (running from trigger): ' + uiErr.message);
    }
    sendAlert('Duplicate Name Detected',
      suggestions.length + ' potential duplicate(s) found. Review in Training Tools -> Review Duplicate Names.');
  } else {
    try {
      SpreadsheetApp.getUi().alert('No new duplicate names detected.');
    } catch (uiErr) {
      Logger.log('No new duplicate names detected. (silent — running from trigger)');
    }
  }
}


// -- Similarity Helpers ---------------------------------------

/**
 * Checks if two names are "similar" using multiple heuristics:
 *  1. Levenshtein distance <= 2
 *  2. Substring containment (Tim / Timothy)
 *  3. Common nickname table
 */
function areSimilar(name1, name2) {
  var n1 = name1.toLowerCase().trim();
  var n2 = name2.toLowerCase().trim();

  if (n1 === n2) return true;
  if (levenshteinDistance(n1, n2) <= 2) return true;
  if (n1.includes(n2) || n2.includes(n1)) return true;

  // Common nickname -> full name pairs
  var nicknames = {
    'tim':   'timothy',
    'mike':  'michael',
    'dave':  'david',
    'chris': 'christopher',
    'rob':   'robert',
    'bob':   'robert',
    'bill':  'william',
    'jim':   'james',
    'joe':   'joseph',
    'dan':   'daniel',
    'alex':  'alexander',
    'matt':  'matthew',
    'jon':   'jonathan',
    'tom':   'thomas',
    'ben':   'benjamin',
    'nick':  'nicholas',
    'jake':  'jacob',
    'sam':   'samuel',
    'tony':  'anthony',
    'ed':    'edward',
    'pat':   'patrick',
    'kate':  'katherine',
    'liz':   'elizabeth',
    'jen':   'jennifer',
    'meg':   'megan'
  };

  for (var nick in nicknames) {
    if ((n1 === nick && n2 === nicknames[nick]) ||
        (n2 === nick && n1 === nicknames[nick])) {
      return true;
    }
  }

  return false;
}

/**
 * Levenshtein distance between two strings.
 */
function levenshteinDistance(str1, str2) {
  var matrix = [];

  for (var i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  for (var j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  for (var i = 1; i <= str2.length; i++) {
    for (var j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

/**
 * Returns true if this pair already exists in the dedup sheet.
 */
function isDuplicateSuggestionExists(name1, name2, dedupSheet) {
  if (dedupSheet.getLastRow() < 2) return false;

  var data = dedupSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var canonical = String(data[i][0]).toLowerCase();
    var variants  = String(data[i][1]).split(',').map(function (v) { return v.trim().toLowerCase(); });

    var n1 = name1.toLowerCase();
    var n2 = name2.toLowerCase();

    if ((canonical === n1 && variants.indexOf(n2) > -1) ||
        (canonical === n2 && variants.indexOf(n1) > -1)) {
      return true;
    }
  }

  return false;
}


// -- Merge Duplicates -----------------------------------------

/**
 * Merges a variant name into the canonical name across
 * all historical entries in the Daily Training Log.
 */
function mergeDuplicateNames(canonicalName, variantName) {
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily Training Log');
  if (logSheet.getLastRow() < 2) return 0;

  var data = logSheet.getDataRange().getValues();
  var updateCount = 0;

  for (var i = 1; i < data.length; i++) {
    var currentName = String(data[i][2]).trim(); // Column C

    if (currentName.toLowerCase() === variantName.toLowerCase()) {
      logSheet.getRange(i + 1, 3).setValue(canonicalName); // Col C
      logSheet.getRange(i + 1, 8).setValue(canonicalName); // Col H
      updateCount++;
    }
  }

  // Mark dedup entry as completed
  var dedupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Deduplication');
  var dedupData  = dedupSheet.getDataRange().getValues();

  for (var i = 1; i < dedupData.length; i++) {
    if (dedupData[i][0] === canonicalName && String(dedupData[i][1]).includes(variantName)) {
      dedupSheet.getRange(i + 1, 4).setValue('Merge');
      dedupSheet.getRange(i + 1, 5).setValue('Completed');
      break;
    }
  }

  // Refresh dashboard
  updateDashboard();

  return updateCount;
}
