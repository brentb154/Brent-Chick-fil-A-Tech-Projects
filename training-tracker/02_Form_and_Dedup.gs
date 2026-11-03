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
    if (!sheet) return;

    // Form responses land in the form's OWN "Form Responses" sheet, not here.
    // So append a normalized row to the Daily Training Log (our canonical store)
    // instead of stamping the log's last (unrelated) row.
    // Form question order: Timestamp, Date, Name, Position, Hours, On Track?, Notes
    var timestamp   = e.values[0];
    var date        = e.values[1];
    var traineeName = String(e.values[2] || '').trim();
    var position    = e.values[3];
    var hours       = e.values[4];
    var onTrack     = e.values[5];
    var notes       = e.values[6] || '';

    if (!traineeName) return;

    // Skip if this timestamp was already imported (avoids double-entry with syncFormData)
    if (timestamp && sheet.getLastRow() > 1) {
      var stamps = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < stamps.length; i++) {
        if (String(stamps[i][0]) === String(timestamp)) return;
      }
    }

    // Resolve canonical name and append a proper new row
    var canonicalName = getCanonicalName(traineeName);
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 9).setValues([[
      timestamp, date, traineeName, position, hours, onTrack, notes, canonicalName, false
    ]]);

    // Check for milestone alerts, then refresh dashboard
    checkMilestoneAlerts(canonicalName, position, parseFloat(hours));
    updateDashboard();

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.message);
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: 'Training Tracker: onFormSubmit FAILED',
        body: 'Error: ' + err.message + '\n\nStack: ' + err.stack
      });
    } catch (mailErr) { Logger.log('Could not send failure alert: ' + mailErr.message); }
  }
}


// -- Canonical Name Resolution --------------------------------

/**
 * Builds a hash map of variant -> canonical name from the
 * Name Deduplication sheet. One sheet read, O(1) lookups.
 * Call once, pass to getCanonicalName() in loops.
 */
function buildCanonicalMap() {
  var dedupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Deduplication');
  if (!dedupSheet || dedupSheet.getLastRow() < 2) return {};

  var data = dedupSheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][3] !== 'Merge' || data[i][4] !== 'Completed') continue;

    var canonical = data[i][0];
    map[canonical.toLowerCase()] = canonical;

    var variants = String(data[i][1]).split(',');
    for (var j = 0; j < variants.length; j++) {
      var key = variants[j].trim().toLowerCase();
      if (key) map[key] = canonical;
    }
  }

  return map;
}

/**
 * Returns the canonical (standardized) name for a trainee.
 * Pass a pre-built canonicalMap for batch operations (avoids repeated sheet reads).
 * If no map provided, builds one on the fly (fine for single calls).
 */
function getCanonicalName(inputName, canonicalMap) {
  var trimmed = inputName.trim();
  if (!canonicalMap) canonicalMap = buildCanonicalMap();

  var result = canonicalMap[trimmed.toLowerCase()];
  return result || trimmed;
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

  // Collect unique names and counts
  var uniqueNames = {};

  nameValues.forEach(function (row) {
    var name = String(row[0]).trim();
    if (name) {
      uniqueNames[name] = (uniqueNames[name] || 0) + 1;
    }
  });

  // Also fold in scheduled-only trainees the dashboard shows but the log
  // doesn't contain yet: Timeline sheet names + Training Schedule trainees.
  // Without this, a scheduled duplicate (e.g. "jada tadros" with a timeline
  // but no log entries) is invisible to the scanner.
  ss.getSheets().forEach(function (sheet) {
    var sName = sheet.getName();
    if (sName.indexOf('Timeline - ') === 0) {
      var tName = sName.replace('Timeline - ', '').trim();
      if (tName) uniqueNames[tName] = uniqueNames[tName] || 1;
    }
  });

  var schedSheet = ss.getSheetByName('Training Schedule');
  if (schedSheet && schedSheet.getLastRow() > 1) {
    schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 1).getValues().forEach(function (row) {
      var name = String(row[0]).trim();
      if (name) uniqueNames[name] = uniqueNames[name] || 1;
    });
  }

  var nameList = Object.keys(uniqueNames);
  if (nameList.length === 0) return;
  var suggestions = [];

  // Build hash set of existing dedup pairs (one read instead of per-pair)
  var existingPairs = {};
  if (dedupSheet.getLastRow() >= 2) {
    var dedupData = dedupSheet.getDataRange().getValues();
    for (var d = 1; d < dedupData.length; d++) {
      var canonical = String(dedupData[d][0]).toLowerCase();
      var variants  = String(dedupData[d][1]).split(',');
      for (var v = 0; v < variants.length; v++) {
        var variant = variants[v].trim().toLowerCase();
        if (variant) {
          existingPairs[canonical + '|' + variant] = true;
          existingPairs[variant + '|' + canonical] = true;
        }
      }
    }
  }

  // Compare every pair
  for (var i = 0; i < nameList.length; i++) {
    for (var j = i + 1; j < nameList.length; j++) {
      var name1 = nameList[i];
      var name2 = nameList[j];

      if (areSimilar(name1, name2)) {
        var pairKey = name1.toLowerCase() + '|' + name2.toLowerCase();
        if (!existingPairs[pairKey]) {
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

  // Length-aware edit distance: short names must match almost exactly (distance 1);
  // only longer names tolerate distance 2. Avoids flagging Maria/Mario, Ben/Ken, Sara/Cara.
  var maxLen = Math.max(n1.length, n2.length);
  var dist   = levenshteinDistance(n1, n2);
  if (maxLen >= 6 ? dist <= 2 : dist <= 1) return true;

  // Same person logged by first name vs. full name ("Sam" vs "Sam Smith"):
  // the shorter must be a whole LEADING word of the longer, min 3 chars.
  var shortN = n1.length <= n2.length ? n1 : n2;
  var longN  = n1.length <= n2.length ? n2 : n1;
  if (shortN.length >= 3 && longN.indexOf(shortN + ' ') === 0) return true;

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

// -- Manual Merge (menu) --------------------------------------

/**
 * Operator-driven merge for names the scanner never flagged as similar
 * (e.g. a nickname with no overlap: "JJ" -> "Jada Tadros").
 * Prompts for the wrong name and the correct name, records the alias in
 * Name Deduplication so future form entries auto-resolve, then runs the
 * same merge used by the review sidebar.
 */
function manuallyMergeNames() {
  var ui = SpreadsheetApp.getUi();

  var variantResp = ui.prompt('Merge Names (1 of 2)',
    'Enter the WRONG / duplicate name exactly as it appears:', ui.ButtonSet.OK_CANCEL);
  if (variantResp.getSelectedButton() !== ui.Button.OK) return;
  var variantName = variantResp.getResponseText().trim();
  if (!variantName) { ui.alert('No name entered.'); return; }

  var canonResp = ui.prompt('Merge Names (2 of 2)',
    'Enter the CORRECT name to keep (everything above merges into this):', ui.ButtonSet.OK_CANCEL);
  if (canonResp.getSelectedButton() !== ui.Button.OK) return;
  var canonicalName = canonResp.getResponseText().trim();
  if (!canonicalName) { ui.alert('No name entered.'); return; }

  // Reject only an exact-string match. Case-only differences (Carmen / carmen)
  // are distinct dashboard rows and ARE a valid merge.
  if (variantName === canonicalName) {
    ui.alert('Those are identical. Enter the names with the casing/spelling you want to merge.');
    return;
  }

  var confirm = ui.alert('Confirm Merge',
    'Merge "' + variantName + '"  ->  "' + canonicalName + '"?\n\n' +
    'This rewrites log entries, renames/removes the duplicate Timeline sheet,\n' +
    'and updates the Training Schedule. It cannot be undone.',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  // Record the alias as a completed dedup row so buildCanonicalMap() picks it up.
  // mergeDuplicateNames() will find this row and stamp Merge/Completed.
  var dedupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Deduplication');
  dedupSheet.getRange(dedupSheet.getLastRow() + 1, 1, 1, 5)
    .setValues([[canonicalName, variantName, '', 'Pending', 'Pending']]);

  var count = mergeDuplicateNames(canonicalName, variantName);

  ui.alert('Merge Complete',
    '"' + variantName + '" merged into "' + canonicalName + '".\n\n' +
    count + ' log entries updated; timeline & schedule reconciled.',
    ui.ButtonSet.OK);
}


// -- Merge Duplicates -----------------------------------------

/**
 * Merges a variant name into the canonical name across
 * all historical entries in the Daily Training Log.
 */
function mergeDuplicateNames(canonicalName, variantName) {
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily Training Log');
  var updateCount = 0;

  // Merge log entries (skip if the log is empty — a scheduled-only trainee
  // still needs the Timeline/Schedule reconciliation below).
  if (logSheet && logSheet.getLastRow() >= 2) {
    var data = logSheet.getDataRange().getValues();

    // Batch: update the in-memory arrays, then write columns C and H back
    var colC = logSheet.getRange(2, 3, data.length - 1, 1).getValues();
    var colH = logSheet.getRange(2, 8, data.length - 1, 1).getValues();

    for (var i = 0; i < colC.length; i++) {
      var currentName = String(colC[i][0]).trim();
      if (currentName.toLowerCase() === variantName.toLowerCase()) {
        colC[i][0] = canonicalName;
        colH[i][0] = canonicalName;
        updateCount++;
      }
    }

    if (updateCount > 0) {
      logSheet.getRange(2, 3, colC.length, 1).setValues(colC);
      logSheet.getRange(2, 8, colH.length, 1).setValues(colH);
      SpreadsheetApp.flush();
    }
  }

  // -- Reconcile scheduled-only sources -----------------------
  // The log merge above doesn't touch scheduled trainees (no log rows).
  // Rename their Timeline sheet and Training Schedule rows to canonical so
  // the dashboard collapses to a single row per person.
  // variantName may be a comma-separated list of variants.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var variants = String(variantName).split(',').map(function (v) { return v.trim(); })
    .filter(function (v) { return v; });

  variants.forEach(function (variant) {
    // -- Timeline sheet: rename to canonical, or delete if canonical exists --
    var variantSheet  = ss.getSheetByName('Timeline - ' + variant);
    if (variantSheet) {
      var canonSheet = ss.getSheetByName('Timeline - ' + canonicalName);
      if (canonSheet) {
        ss.deleteSheet(variantSheet); // canonical timeline already there; drop the dup
      } else {
        variantSheet.setName('Timeline - ' + canonicalName);
      }
    }
  });

  // -- Training Schedule col A: variant -> canonical (batch) --
  var schedSheet = ss.getSheetByName('Training Schedule');
  if (schedSheet && schedSheet.getLastRow() > 1) {
    var colA = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 1).getValues();
    var schedChanged = false;
    for (var s = 0; s < colA.length; s++) {
      var schedName = String(colA[s][0]).trim().toLowerCase();
      for (var vi = 0; vi < variants.length; vi++) {
        if (schedName === variants[vi].toLowerCase()) {
          colA[s][0] = canonicalName;
          schedChanged = true;
          break;
        }
      }
    }
    if (schedChanged) {
      schedSheet.getRange(2, 1, colA.length, 1).setValues(colA);
      SpreadsheetApp.flush();
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
