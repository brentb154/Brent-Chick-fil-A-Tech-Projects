// ============================================================
// CATERING PIPELINE — Pipeline.gs
// ============================================================
// Add this as a new file in your Apps Script project.
// (Click + next to "Files" → Script → name it "Pipeline")
//
// Pipeline Sheet Columns:
//   A: Quote ID        B: Customer Name     C: Email
//   D: Location        E: Total             F: Event Date
//   G: Date Sent       H: Status            I: Follow-Up Date
//   J: Notes           K: Last Updated
// ============================================================

const TAB_PIPELINE = 'Pipeline';

const PIPELINE_STATUSES = ['Quoted / Sent', 'Awaiting Response', 'Confirmed / Booked'];

// Column index map (1-based for getRange, 0-based for arrays)
const COL = {
  QUOTE_ID:    1,
  CUSTOMER:    2,
  EMAIL:       3,
  LOCATION:    4,
  TOTAL:       5,
  EVENT_DATE:  6,
  DATE_SENT:   7,
  STATUS:      8,
  FOLLOWUP:    9,
  NOTES:       10,
  UPDATED:     11
};

// ── INIT ──────────────────────────────────────────────────────
// Run this once manually to create the Pipeline tab.
// Safe to run again — skips if tab already exists.

function initPipelineSheet() {
  var ss = getSpreadsheet();
  var existing = ss.getSheetByName(TAB_PIPELINE);
  if (existing) return 'Pipeline tab already exists.';

  var sheet = ss.insertSheet(TAB_PIPELINE);
  var headers = [
    'Quote ID', 'Customer Name', 'Email', 'Location', 'Total',
    'Event Date', 'Date Sent', 'Status', 'Follow-Up Date', 'Notes', 'Last Updated'
  ];

  // Header row styling
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#DD0033');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);

  // Column widths
  sheet.setColumnWidth(1, 120);  // Quote ID
  sheet.setColumnWidth(2, 180);  // Customer
  sheet.setColumnWidth(3, 200);  // Email
  sheet.setColumnWidth(4, 160);  // Location
  sheet.setColumnWidth(5, 90);   // Total
  sheet.setColumnWidth(6, 110);  // Event Date
  sheet.setColumnWidth(7, 110);  // Date Sent
  sheet.setColumnWidth(8, 140);  // Status
  sheet.setColumnWidth(9, 120);  // Follow-Up Date
  sheet.setColumnWidth(10, 280); // Notes
  sheet.setColumnWidth(11, 140); // Last Updated

  // Freeze header row
  sheet.setFrozenRows(1);

  // Status dropdown validation on column H
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PIPELINE_STATUSES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('H2:H1000').setDataValidation(statusRule);

  return 'Pipeline tab created successfully.';
}


// ── ADD ENTRY ─────────────────────────────────────────────────
// Called automatically when a quote email is sent.
// quoteData must include: quoteId, customerName, recipientEmail,
//   locationName, total, eventDate (optional)

function addPipelineEntry(quoteData) {
  var sheet = getSpreadsheet().getSheetByName(TAB_PIPELINE);
  if (!sheet) {
    initPipelineSheet();
    sheet = getSpreadsheet().getSheetByName(TAB_PIPELINE);
  }

  // Check if this quote is already in the pipeline (prevent duplicates on re-send)
  var existing = findPipelineRow(quoteData.quoteId);
  if (existing) {
    // Update the date sent and last updated, keep everything else
    sheet.getRange(existing, COL.DATE_SENT).setValue(new Date());
    sheet.getRange(existing, COL.UPDATED).setValue(new Date());
    return { success: true, message: 'Pipeline entry updated (re-sent).' };
  }

  var now = new Date();
  var followUpDate = new Date(now);
  followUpDate.setDate(followUpDate.getDate() + 3); // Default: follow up in 3 days

  var row = [
    quoteData.quoteId        || '',
    quoteData.customerName   || '',
    quoteData.recipientEmail || '',
    quoteData.locationName   || '',
    parseFloat(quoteData.total) || 0,
    quoteData.eventDate      || '',
    now,
    'Quoted / Sent',
    followUpDate,
    '',  // Notes — blank to start
    now
  ];

  sheet.appendRow(row);

  // Format the new row's date columns
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL.DATE_SENT).setNumberFormat('M/d/yyyy');
  sheet.getRange(lastRow, COL.FOLLOWUP).setNumberFormat('M/d/yyyy');
  sheet.getRange(lastRow, COL.UPDATED).setNumberFormat('M/d/yyyy h:mm am/pm');
  sheet.getRange(lastRow, COL.EVENT_DATE).setNumberFormat('M/d/yyyy');
  sheet.getRange(lastRow, COL.TOTAL).setNumberFormat('$#,##0.00');

  return { success: true, message: 'Pipeline entry created.' };
}


// ── GET ENTRIES ───────────────────────────────────────────────
// Returns all pipeline entries as an array of objects.
// Sorted: follow-ups due today first, then by event date ascending.

function getPipelineEntries() {
  var sheet = getSpreadsheet().getSheetByName(TAB_PIPELINE);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  var entries = data
    .filter(function(row) { return row[0] && row[0].toString().trim() !== ''; })
    .map(function(row, i) {
      var followUpDate = row[COL.FOLLOWUP - 1] ? new Date(row[COL.FOLLOWUP - 1]) : null;
      var eventDate    = row[COL.EVENT_DATE - 1] ? new Date(row[COL.EVENT_DATE - 1]) : null;
      if (followUpDate) followUpDate.setHours(0, 0, 0, 0);
      if (eventDate)    eventDate.setHours(0, 0, 0, 0);

      var followUpDue = followUpDate && followUpDate <= today;
      var followUpSoon = followUpDate && followUpDate >= today && followUpDate < tomorrow;

      return {
        row:          i + 2,
        quoteId:      row[COL.QUOTE_ID   - 1] ? row[COL.QUOTE_ID - 1].toString() : '',
        customer:     row[COL.CUSTOMER   - 1] ? row[COL.CUSTOMER - 1].toString() : '',
        email:        row[COL.EMAIL      - 1] ? row[COL.EMAIL - 1].toString() : '',
        location:     row[COL.LOCATION   - 1] ? row[COL.LOCATION - 1].toString() : '',
        total:        parseFloat(row[COL.TOTAL - 1]) || 0,
        eventDate:    eventDate    ? eventDate.toLocaleDateString('en-US')    : '',
        dateSent:     row[COL.DATE_SENT - 1] ? new Date(row[COL.DATE_SENT - 1]).toLocaleDateString('en-US') : '',
        status:       row[COL.STATUS - 1]    ? row[COL.STATUS - 1].toString() : 'Quoted / Sent',
        followUpDate: followUpDate ? followUpDate.toLocaleDateString('en-US') : '',
        followUpRaw:  followUpDate ? followUpDate.getTime() : null,
        notes:        row[COL.NOTES - 1]     ? row[COL.NOTES - 1].toString() : '',
        updated:      row[COL.UPDATED - 1]   ? new Date(row[COL.UPDATED - 1]).toLocaleDateString('en-US') : '',
        followUpDue:  followUpDue,
        followUpSoon: followUpSoon
      };
    });

  // Sort: overdue first → due today → upcoming by event date
  entries.sort(function(a, b) {
    if (a.followUpDue && !b.followUpDue) return -1;
    if (!a.followUpDue && b.followUpDue) return 1;
    if (a.followUpSoon && !b.followUpSoon) return -1;
    if (!a.followUpSoon && b.followUpSoon) return 1;
    return (a.followUpRaw || 0) - (b.followUpRaw || 0);
  });

  return entries;
}


// ── UPDATE ENTRY ──────────────────────────────────────────────
// Called from the UI when a director updates status, follow-up date, or notes.

function updatePipelineEntry(quoteId, updates) {
  var sheet = getSpreadsheet().getSheetByName(TAB_PIPELINE);
  if (!sheet) return { success: false, message: 'Pipeline sheet not found.' };

  var rowNum = findPipelineRow(quoteId);
  if (!rowNum) return { success: false, message: 'Quote not found in pipeline.' };

  if (updates.status !== undefined) {
    sheet.getRange(rowNum, COL.STATUS).setValue(updates.status);
  }
  if (updates.followUpDate !== undefined) {
    var fd = updates.followUpDate ? new Date(updates.followUpDate) : '';
    sheet.getRange(rowNum, COL.FOLLOWUP).setValue(fd);
    if (fd) sheet.getRange(rowNum, COL.FOLLOWUP).setNumberFormat('M/d/yyyy');
  }
  if (updates.notes !== undefined) {
    sheet.getRange(rowNum, COL.NOTES).setValue(updates.notes);
  }

  sheet.getRange(rowNum, COL.UPDATED).setValue(new Date());
  sheet.getRange(rowNum, COL.UPDATED).setNumberFormat('M/d/yyyy h:mm am/pm');

  return { success: true, message: 'Entry updated.' };
}


// ── DELETE ENTRY ──────────────────────────────────────────────
// Removes a completed or lost quote from the pipeline.

function deletePipelineEntry(quoteId) {
  var sheet = getSpreadsheet().getSheetByName(TAB_PIPELINE);
  if (!sheet) return { success: false, message: 'Pipeline sheet not found.' };

  var rowNum = findPipelineRow(quoteId);
  if (!rowNum) return { success: false, message: 'Quote not found.' };

  sheet.deleteRow(rowNum);
  return { success: true, message: 'Entry removed from pipeline.' };
}


// ── HELPER: FIND ROW ──────────────────────────────────────────

function findPipelineRow(quoteId) {
  var sheet = getSpreadsheet().getSheetByName(TAB_PIPELINE);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var ids = sheet.getRange(2, COL.QUOTE_ID, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0].toString().trim() === quoteId.toString().trim()) {
      return i + 2;
    }
  }
  return null;
}


// ── DAILY FOLLOW-UP DIGEST ────────────────────────────────────
// Set this up as a time-driven trigger: daily, morning (7-8am).
// Goes to the BCC email from Settings (same as quote emails).
// Setup: Apps Script → Triggers → + Add Trigger
//   Function: sendFollowUpDigest | Time-driven | Day timer | 7am–8am

function sendFollowUpDigest() {
  var settings = getSettings();
  var digestEmail = (settings['Pipeline Digest Email'] || settings['BCC Email'] || '').toString().trim();
  if (!digestEmail) return;

  var entries = getPipelineEntries();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var overdue = entries.filter(function(e) { return e.followUpDue && e.status !== 'Confirmed / Booked'; });
  var upcoming = entries.filter(function(e) {
    if (e.followUpDue || e.status === 'Confirmed / Booked') return false;
    if (!e.followUpRaw) return false;
    var daysOut = Math.ceil((e.followUpRaw - today.getTime()) / (1000 * 60 * 60 * 24));
    return daysOut <= 3;
  });

  if (overdue.length === 0 && upcoming.length === 0) return; // Nothing to report

  var subject = '📋 Catering Follow-Up Digest — ' + today.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });

  var body = 'Good morning!\n\nHere\'s your catering pipeline follow-up summary:\n\n';

  if (overdue.length > 0) {
    body += '🔴 OVERDUE FOLLOW-UPS (' + overdue.length + ')\n';
    body += '─'.repeat(40) + '\n';
    overdue.forEach(function(e) {
      body += '• ' + e.customer + ' (' + e.quoteId + ') — $' + e.total.toFixed(2);
      if (e.eventDate) body += ' | Event: ' + e.eventDate;
      body += ' | Status: ' + e.status;
      body += ' | Follow-up was due: ' + e.followUpDate;
      if (e.notes) body += '\n  Notes: ' + e.notes;
      body += '\n';
    });
    body += '\n';
  }

  if (upcoming.length > 0) {
    body += '🟡 COMING UP (next 3 days)\n';
    body += '─'.repeat(40) + '\n';
    upcoming.forEach(function(e) {
      body += '• ' + e.customer + ' (' + e.quoteId + ') — $' + e.total.toFixed(2);
      if (e.eventDate) body += ' | Event: ' + e.eventDate;
      body += ' | Follow-up: ' + e.followUpDate;
      if (e.notes) body += '\n  Notes: ' + e.notes;
      body += '\n';
    });
    body += '\n';
  }

  body += '─'.repeat(40) + '\n';
  body += 'Open the catering app to update statuses and reschedule follow-ups.\n\n';
  body += 'Chick-fil-A Catering Pipeline';

  MailApp.sendEmail(digestEmail, subject, body);
}


// ── PIPELINE STATS ────────────────────────────────────────────
// Returns summary stats for the pipeline header card.

function getPipelineStats() {
  var entries = getPipelineEntries();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var stats = {
    total:       entries.length,
    quoted:      0,
    awaiting:    0,
    confirmed:   0,
    totalValue:  0,
    confirmedValue: 0,
    overdueFollowUps: 0
  };

  entries.forEach(function(e) {
    stats.totalValue += e.total;
    if (e.status === 'Quoted / Sent')      stats.quoted++;
    if (e.status === 'Awaiting Response')  stats.awaiting++;
    if (e.status === 'Confirmed / Booked') { stats.confirmed++; stats.confirmedValue += e.total; }
    if (e.followUpDue && e.status !== 'Confirmed / Booked') stats.overdueFollowUps++;
  });

  return stats;
}