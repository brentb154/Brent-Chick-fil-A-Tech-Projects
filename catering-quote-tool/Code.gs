// ============================================================
// CHICK-FIL-A CATERING QUOTE GENERATOR — Server-Side (Code.gs)
// ============================================================
// Handles all communication between the web app and the Google
// Sheet: settings, menu items (with categories), quotes,
// server-side PDF generation, and email sending.
// ============================================================

// Run this once from the dropdown to force Calendar permission, then delete it
function testCalendarAccess() {
  var cal = CalendarApp.getDefaultCalendar();
  Logger.log('Calendar access OK: ' + cal.getName());
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

const TAB_SETTINGS       = 'Settings';
const TAB_MENU           = 'Menu';
const TAB_QUOTES         = 'Quotes';
const TAB_QUOTES_ARCHIVE = 'Quotes_Archive';
const TAB_QUOTE_SEQUENCE = 'Quote_Sequence';
const TAB_QUOTE_REVISIONS = 'Quote_Revisions';
const TAB_OFF_MENU = 'Off_Menu';


// ── WEB APP ENTRY POINT ──────────────────────────────────────

// Adoption ping: one row per day to the shared adoption sheet.
// No-op unless ADOPTION_SHEET_ID is set in Script Properties. Never throws.
function logAdoptionPing_(toolName) {
  try {
    var props = PropertiesService.getScriptProperties();
    var sheetId = props.getProperty('ADOPTION_SHEET_ID');
    if (!sheetId) return;
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (props.getProperty('ADOPTION_LAST_PING') === today) return;
    var tab = SpreadsheetApp.openById(sheetId).getSheetByName('Pings');
    if (!tab) return;
    tab.appendRow([today, toolName]);
    props.setProperty('ADOPTION_LAST_PING', today);
  } catch (err) {
    // Never let adoption logging break the tool.
  }
}

// The app is deployed "Anyone" so external guests can reach the tax-form upload
// page. Routing keeps the internal tool behind an obscure key:
//   ?view=taxform&quote=Q-...  → public guest upload page (no gate)
//   ?view=app                  → internal quote tool
//   anything else / bare URL   → bland landing (don't expose Index to the public)
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) ? e.parameter.view : '';

  if (view === 'taxform') {
    var tpl = HtmlService.createTemplateFromFile('TaxForm');
    tpl.quoteParam = (e && e.parameter && e.parameter.quote) ? e.parameter.quote : '';
    return tpl.evaluate()
      .setTitle('Upload Your Tax-Exempt Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  if (view === 'app') {
    logAdoptionPing_('catering-quote-tool');
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('CFA Catering Quotes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutput(
      '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:460px;margin:80px auto;padding:0 24px;text-align:center;color:#374151;">'
    + '<div style="font-size:40px;">🍽️</div>'
    + '<h2 style="color:#E51636;">Chick-fil-A Catering</h2>'
    + '<p style="line-height:1.6;">This link isn\'t valid on its own. If a team member sent you here to upload a tax-exempt form, please use the exact link from your email.</p>'
    + '</div>')
    .setTitle('Chick-fil-A Catering')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── SETTINGS ─────────────────────────────────────────────────

function getSettings() {
  var sheet = getSpreadsheet().getSheetByName(TAB_SETTINGS);
  var data  = sheet.getRange('A1:B100').getValues();
  var tz = Session.getScriptTimeZone();
  var settings = {};
  data.forEach(function(row) {
    if (row[0] && row[0].toString().trim() !== '') {
      var v = row[1];
      // Sheets auto-converts date-like strings (e.g. "2026-07-16") into Date cells,
      // and google.script.run nulls the ENTIRE return value if it contains a Date.
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      }
      settings[row[0].toString().trim()] = v;
    }
  });
  return settings;
}

function updateSetting(label, value, asText) {
  var sheet = getSpreadsheet().getSheetByName(TAB_SETTINGS);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var data = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === label) {
      var cell = sheet.getRange(i + 1, 2);
      if (asText) cell.setNumberFormat('@'); // stop Sheets auto-converting date-like strings
      cell.setValue(value);
      return true;
    }
  }
  // Label not found — append a new row instead of silently dropping the save
  if (asText) sheet.getRange(lastRow + 1, 2).setNumberFormat('@');
  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[label, value]]);
  return true;
}


// ── MENU (now with Category column) ──────────────────────────
// Menu tab columns: A=Category | B=Item Name | C=Pickup Price | D=Delivery Price

function getMenuItems() {
  var sheet = getSpreadsheet().getSheetByName(TAB_MENU);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var items = [];
  data.forEach(function(row, index) {
    if (row[1] && row[1].toString().trim() !== '') {
      var rawPickup   = (row[2] !== null && row[2] !== undefined) ? row[2].toString().trim() : '';
      var rawDelivery = (row[3] !== null && row[3] !== undefined) ? row[3].toString().trim() : '';
      items.push({
        category:         (row[0] || '').toString().trim(),
        name:             row[1].toString().trim(),
        pickupPrice:      parseFloat(rawPickup)   || 0,
        deliveryPrice:    parseFloat(rawDelivery) || 0,
        pickupAvailable:  rawPickup   !== '' && !isNaN(parseFloat(rawPickup)),
        deliveryAvailable: rawDelivery !== '' && !isNaN(parseFloat(rawDelivery)),
        row: index + 2
      });
    }
  });
  return items;
}

// ── QUARTERLY PRICE VERIFICATION ─────────────────────────────
// Every "Price Check Interval (Days)" (default 90) the app locks on load until
// the operator confirms 3 random menu prices against the POS. The client decides
// "due" from Settings; these functions serve the items and check the answers.

// 3 random items to spot-check — always includes a delivery-priced item when one exists.
// Draws only from the categories listed in Settings "Price Check Categories" (comma-separated),
// matched by category name — never by row — so items can move around freely.
// Falls back to the whole menu if the list is blank or matches fewer than 3 items.
function getPriceCheckItems() {
  var pool = getMenuItems();
  if (!pool.length) return [];
  var raw = (getSettings()['Price Check Categories'] || '').toString();
  var cats = raw.split(',').map(function(c) { return c.trim().toLowerCase(); }).filter(function(c) { return c !== ''; });
  if (cats.length) {
    var inCats = pool.filter(function(it) { return cats.indexOf((it.category || '').toLowerCase()) >= 0; });
    if (inCats.length >= 3) pool = inCats;
  }
  var picked = [];
  var withDelivery = pool.filter(function(it) { return it.deliveryAvailable; });
  if (withDelivery.length) {
    var d = withDelivery[Math.floor(Math.random() * withDelivery.length)];
    picked.push(d);
    pool = pool.filter(function(it) { return it.name !== d.name; });
  }
  while (picked.length < 3 && pool.length) {
    picked.push(pool.splice(Math.floor(Math.random() * pool.length), 1)[0]);
  }
  // Names only — answers are checked server-side against the Menu tab
  return picked.map(function(it) {
    return { name: it.name, category: it.category, pickupAvailable: it.pickupAvailable, deliveryAvailable: it.deliveryAvailable };
  });
}

// entries: [{name, pickup, delivery}] typed from the POS. Match within a cent.
// An item name can appear on multiple menu rows (same item, two categories) —
// the typed prices pass if they match ANY row with that name, so duplicates can't wedge the lock.
function submitPriceVerification(entries) {
  var rowsByName = {};
  getMenuItems().forEach(function(it) {
    if (!rowsByName[it.name]) rowsByName[it.name] = [];
    rowsByName[it.name].push(it);
  });
  function priceMatches(typed, expected) {
    var n = parseFloat(typed);
    if (isNaN(n)) return false; // blank never passes — even for $0 items, they have to type it
    return Math.abs(n - expected) <= 0.005;
  }
  var mismatches = [];
  (entries || []).forEach(function(e) {
    var rows = rowsByName[e.name];
    if (!rows) { mismatches.push(e.name); return; }
    var ok = rows.some(function(it) {
      if (it.pickupAvailable && !priceMatches(e.pickup, it.pickupPrice)) return false;
      if (it.deliveryAvailable && !priceMatches(e.delivery, it.deliveryPrice)) return false;
      return true;
    });
    if (!ok) mismatches.push(e.name);
  });
  if (mismatches.length) return { ok: false, mismatches: mismatches };
  markPricesVerified_();
  return { ok: true };
}

function markPricesVerified_() {
  // asText keeps Sheets from converting the yyyy-MM-dd string into a Date cell
  updateSetting('Last Price Verification', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), true);
}

// ── OFF-MENU CHEAT SHEET ─────────────────────────────────────
// Off_Menu tab: A=Item Name | B=Base Price. Delivery price is computed
// in the app as base × (1 + "Off-Menu Markup (%)" / 100), delivery only.

function getOffMenuItems() {
  var sheet = getSpreadsheet().getSheetByName(TAB_OFF_MENU);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var items = [];
  data.forEach(function(row, index) {
    if (row[0] && row[0].toString().trim() !== '') {
      var raw = (row[1] !== null && row[1] !== undefined) ? row[1].toString().trim() : '';
      items.push({
        name: row[0].toString().trim(),
        basePrice: parseFloat(raw) || 0,
        hasPrice: raw !== '' && !isNaN(parseFloat(raw)),
        row: index + 2
      });
    }
  });
  return items;
}

function addOffMenuItem(name, basePrice) {
  var sheet = getSpreadsheet().getSheetByName(TAB_OFF_MENU);
  if (!sheet) sheet = createOffMenuSheet_();
  var r = sheet.getLastRow() + 1;
  sheet.getRange(r, 1, 1, 2).setValues([[name, parseFloat(basePrice) || 0]]);
  sheet.getRange(r, 2).setNumberFormat('$#,##0.00');
  return true;
}

// Fallback creator if someone deleted the tab — same shape as initializeSheet builds.
function createOffMenuSheet_() {
  var sheet = getSpreadsheet().insertSheet(TAB_OFF_MENU);
  sheet.getRange(1, 1, 1, 2).setValues([['Item Name', 'Base Price']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function updateOffMenuItem(row, name, basePrice) {
  var sheet = getSpreadsheet().getSheetByName(TAB_OFF_MENU);
  sheet.getRange(row, 1, 1, 2).setValues([[name, parseFloat(basePrice) || 0]]);
  return true;
}

function deleteOffMenuItem(row) {
  getSpreadsheet().getSheetByName(TAB_OFF_MENU).deleteRow(row);
  return true;
}

function addMenuItem(category, name, pickupPrice, deliveryPrice) {
  var sheet = getSpreadsheet().getSheetByName(TAB_MENU);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([
    [category || '', name, parseFloat(pickupPrice) || 0, parseFloat(deliveryPrice) || 0]
  ]);
  return true;
}

function updateMenuItem(row, category, name, pickupPrice, deliveryPrice) {
  var sheet = getSpreadsheet().getSheetByName(TAB_MENU);
  sheet.getRange(row, 1, 1, 4).setValues([
    [category || '', name, parseFloat(pickupPrice) || 0, parseFloat(deliveryPrice) || 0]
  ]);
  return true;
}

function deleteMenuItem(row) {
  var sheet = getSpreadsheet().getSheetByName(TAB_MENU);
  sheet.deleteRow(row);
  return true;
}


// ── QUOTES ───────────────────────────────────────────────────

function getQuotes() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var width = Math.max(19, Math.min(23, lastCol));
  var data = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  var tz = Session.getScriptTimeZone();
  var quotes = [];
  data.forEach(function(row, index) {
    if (row[0] && row[0].toString().trim() !== '') {
      quotes.push({
        quoteId: row[0], createdDate: row[1] ? new Date(row[1]).toISOString() : '',
        customerName: row[2], contactName: row[3], orderType: row[4],
        deliveryAddress: row[5], lineItems: row[6], subtotal: row[7],
        taxRateUsed: row[8], taxAmount: row[9], total: row[10],
        taxExempt: row[11], locationName: row[12], customerEmail: row[13] || '',
        poNumber: row[14] ? row[14].toString().trim() : '',
        eventDate: normalizeEventDate_(row[15], tz),
        calendarEventId: row[16] ? row[16].toString().trim() : '',
        lastModified: row[17] ? new Date(row[17]).toISOString() : '',
        eventTime: normalizeEventTime_(row[18], tz),
        orderDiscountValue: row[19] != null ? row[19] : 0,
        orderDiscountType: row[20] ? row[20].toString().trim() : 'percent',
        quoteNotes: row[21] ? row[21].toString() : '',
        customerPhone: row[22] ? row[22].toString().trim() : '',
        sheetRow: index + 2
      });
    }
  });
  quotes.sort(function(a, b) { return new Date(b.createdDate) - new Date(a.createdDate); });
  return quotes;
}

// Sheets often auto-converts the date/time inputs into Date objects.
// Normalize both back to simple round-trippable strings so the UI inputs
// accept them and the display layer can format them cleanly.
function normalizeEventDate_(val, tz) {
  if (val === '' || val == null) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  }
  return val.toString().trim();
}
function normalizeEventTime_(val, tz) {
  if (val === '' || val == null) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  return val.toString().trim();
}

// Reads the Settings tab "Calendar Lead Time (Minutes)" — clamps to a sane
// range and falls back to 30 if blank or non-numeric.
function getCalendarLeadMinutes_(settings) {
  var n = parseInt(settings['Calendar Lead Time (Minutes)'], 10);
  if (isNaN(n) || n < 0) return 30;
  if (n > 480) return 480; // hard cap at 8 hours
  return n;
}

// Reads the Settings tab "Archive After Days" — how old a quote gets before
// cleanOldQuotes() moves it to the archive. Falls back to 120 if blank/invalid.
function getArchiveAfterDays_(settings) {
  var n = parseInt(settings['Archive After Days'], 10);
  if (isNaN(n) || n < 1) return 120;
  if (n > 3650) return 3650; // hard cap at ~10 years
  return n;
}

// "Jane Smith" → "JS". First letter of the first and last words, letters only.
function contactInitials_(contactName) {
  var words = (contactName || '').toString().trim().split(/\s+/).filter(function(w) { return /[A-Za-z]/.test(w); });
  if (!words.length) return '';
  var first = words[0].replace(/[^A-Za-z]/g, '').charAt(0);
  var last = words.length > 1 ? words[words.length - 1].replace(/[^A-Za-z]/g, '').charAt(0) : '';
  return (first + last).toUpperCase();
}

function getNextQuoteId(contactName) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTE_SEQUENCE);
  var current = parseInt(sheet.getRange('B1').getValue()) || 0;
  var next = current + 1;
  sheet.getRange('B1').setValue(next);
  var initials = contactInitials_(contactName);
  // Initials are a label riding on the serial — uniqueness comes from the counter alone
  return 'Q-' + new Date().getFullYear() + '-' + ('0000' + next).slice(-4) + (initials ? '-' + initials : '');
}

function saveQuote(quoteData) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var existingQuoteId = (quoteData.quoteId || '').toString().trim();
  var existingRow = 0;

  // Check if this is an edit of an existing quote
  if (existingQuoteId) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0].toString().trim() === existingQuoteId) {
          existingRow = i + 2;
          break;
        }
      }
    }
  }

  if (existingRow > 0) {
    // Edit: overwrite existing row, preserve original created date
    var existing = sheet.getRange(existingRow, 1, 1, 23).getValues()[0];
    var eventId = existing[16] ? existing[16].toString().trim() : '';

    // Keep the outgoing version in Quote_Revisions before overwriting
    try { appendQuoteRevision_(existing); } catch(e) {}

    // Update existing calendar event, or create one if it doesn't exist yet
    if (eventId) {
      try { eventId = updateCalendarEvent(quoteData, existingQuoteId, eventId) || eventId; } catch(e) {}
    } else if (quoteData.date && quoteData.time) {
      try { eventId = createCalendarEvent(quoteData, existingQuoteId) || ''; } catch(e) { eventId = ''; }
    }

    sheet.getRange(existingRow, 1, 1, 23).setValues([[
      existingQuoteId, existing[1], quoteData.customerName, quoteData.contactName,
      quoteData.orderType, quoteData.deliveryAddress || '',
      JSON.stringify(quoteData.lineItems),
      parseFloat(quoteData.subtotal) || 0, parseFloat(quoteData.taxRate) || 0,
      parseFloat(quoteData.taxAmount) || 0, parseFloat(quoteData.total) || 0,
      quoteData.taxExempt ? 'TRUE' : 'FALSE', quoteData.locationName || '',
      quoteData.customerEmail || '', quoteData.poNumber || '', quoteData.date || '',
      eventId, new Date(), quoteData.time || '',
      parseFloat(quoteData.orderDiscountValue) || 0, quoteData.orderDiscountType || 'percent', quoteData.quoteNotes || '',
      quoteData.customerPhone || ''
    ]]);
    return existingQuoteId;
  }

  // New quote
  var quoteId = getNextQuoteId(quoteData.contactName);
  var eventId = '';
  try { eventId = createCalendarEvent(quoteData, quoteId) || ''; } catch(e) { eventId = ''; }
  sheet.appendRow([
    quoteId, new Date(), quoteData.customerName, quoteData.contactName,
    quoteData.orderType, quoteData.deliveryAddress || '',
    JSON.stringify(quoteData.lineItems),
    parseFloat(quoteData.subtotal) || 0, parseFloat(quoteData.taxRate) || 0,
    parseFloat(quoteData.taxAmount) || 0, parseFloat(quoteData.total) || 0,
    quoteData.taxExempt ? 'TRUE' : 'FALSE', quoteData.locationName || '',
    quoteData.customerEmail || '', quoteData.poNumber || '', quoteData.date || '',
    eventId, '', quoteData.time || '',
    parseFloat(quoteData.orderDiscountValue) || 0, quoteData.orderDiscountType || 'percent', quoteData.quoteNotes || '',
    quoteData.customerPhone || ''
  ]);
  return quoteId;
}

// ── QUOTE REVISIONS ──────────────────────────────────────────

// One source of truth for the revisions tab: Quotes' 23 columns + Revised At.
// Creates the hidden tab if missing — used by both saveQuote and initializeSheet.
function getRevisionsSheet_() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_QUOTE_REVISIONS);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_QUOTE_REVISIONS);
    var rh = ['Quote ID','Created Date','Customer Name','Contact Name','Order Type','Delivery Address','Line Items (JSON)','Subtotal','Tax Rate Used','Tax Amount','Total','Tax Exempt','Location Name','Customer Email','PO Number','Event Date','Calendar Event ID','Last Modified','Event Time','Order Discount Value','Order Discount Type','Quote Notes','Customer Phone','Revised At'];
    sheet.getRange(1, 1, 1, rh.length).setValues([rh]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

// Append a pre-edit Quotes row (23 values) + timestamp to the hidden revisions tab.
function appendQuoteRevision_(rowValues) {
  var row = rowValues.slice(0, 23);
  while (row.length < 23) row.push('');
  row.push(new Date());
  getRevisionsSheet_().appendRow(row);
}

// All prior versions of one quote, newest first. Same field names as getQuotes plus revisedAt.
function getQuoteRevisions(quoteId) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTE_REVISIONS);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
  var tz = Session.getScriptTimeZone();
  var wanted = (quoteId || '').toString().trim();
  var revisions = [];
  data.forEach(function(row) {
    if (!row[0] || row[0].toString().trim() !== wanted) return;
    revisions.push({
      quoteId: row[0], createdDate: row[1] ? new Date(row[1]).toISOString() : '',
      customerName: row[2], contactName: row[3], orderType: row[4],
      deliveryAddress: row[5], lineItems: row[6], subtotal: row[7],
      taxRateUsed: row[8], taxAmount: row[9], total: row[10],
      taxExempt: row[11], locationName: row[12], customerEmail: row[13] || '',
      poNumber: row[14] ? row[14].toString().trim() : '',
      eventDate: normalizeEventDate_(row[15], tz),
      eventTime: normalizeEventTime_(row[18], tz),
      orderDiscountValue: row[19] != null ? row[19] : 0,
      orderDiscountType: row[20] ? row[20].toString().trim() : 'percent',
      quoteNotes: row[21] ? row[21].toString() : '',
      customerPhone: row[22] ? row[22].toString().trim() : '',
      revisedAt: row[23] ? new Date(row[23]).toISOString() : ''
    });
  });
  revisions.sort(function(a, b) { return new Date(b.revisedAt) - new Date(a.revisedAt); });
  return revisions;
}

function deleteQuote(sheetRow) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var quoteId = sheet.getRange(sheetRow, 1).getValue().toString().trim();
  sheet.deleteRow(sheetRow);
  return true;
}

function updateQuotePO(sheetRow, poNumber, calendarEventId) {
  getSpreadsheet().getSheetByName(TAB_QUOTES).getRange(sheetRow, 15).setValue(poNumber || '');
  var status = { poSaved: true, calendarUpdated: false, calendarReason: '' };
  if (!calendarEventId) {
    status.calendarReason = 'No calendar event linked to this quote';
    return status;
  }
  try {
    var result = updateCalendarEventPO(sheetRow, poNumber, calendarEventId);
    if (result === true) { status.calendarUpdated = true; }
    else { status.calendarReason = (typeof result === 'string') ? result : 'Event not found on calendar'; }
  } catch(e) {
    status.calendarReason = e.message || String(e);
  }
  return status;
}

// Move quotes older than the "Archive After Days" setting (default 120) to a
// hidden archive sheet instead of deleting.
// Nothing is lost — the archive keeps the full row so old quotes can always be reviewed.
function cleanOldQuotes() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_QUOTES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var archiveAfterDays = getArchiveAfterDays_(getSettings());
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - archiveAfterDays);

  var toArchive = [];   // full rows to copy into the archive
  var rowsToDelete = []; // sheet row numbers, deleted bottom-up
  for (var i = data.length - 1; i >= 0; i--) {
    if (new Date(data[i][1]) < cutoff) {
      toArchive.push(data[i]);
      rowsToDelete.push(i + 2);
    }
  }
  if (!toArchive.length) return;

  // Archive sheet mirrors the Quotes header and stays hidden.
  var archive = ss.getSheetByName(TAB_QUOTES_ARCHIVE);
  if (!archive) {
    archive = ss.insertSheet(TAB_QUOTES_ARCHIVE);
    var header = sheet.getRange(1, 1, 1, lastCol).getValues();
    archive.getRange(1, 1, 1, lastCol).setValues(header).setFontWeight('bold');
    archive.setFrozenRows(1);
    archive.hideSheet();
  }

  archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, lastCol).setValues(toArchive);
  SpreadsheetApp.flush();

  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]); // already sorted descending, safe to delete in order
  }
}


// ── SHEET INITIALIZATION ─────────────────────────────────────

function initializeSheet() {
  var ss = getSpreadsheet();

  // Settings
  var sSheet = ss.getSheetByName(TAB_SETTINGS);
  if (!sSheet) sSheet = ss.insertSheet(TAB_SETTINGS);
  var seedSettings = [
    ['Store Name (Active)',         'Cockrell Hill DTO'],
    ['Location 1 Name',             'Cockrell Hill DTO'],
    ['Location 1 Address',          '1535 N Cockrell Hill Rd., Dallas, Texas 75211'],
    ['Location 1 Phone',            '214-331-2400'],
    ['Location 2 Name',             'Dallas Baptist University OCV'],
    ['Location 2 Address',          ''],
    ['Location 2 Phone',            ''],
    ['Quote Contact Name',          ''],
    ['Default Tax Rate (%)',         8.25],
    ['Calendar Lead Time (Minutes)', 30],
    ['Archive After Days',           120],
    ['Price Check Interval (Days)',  90],
    ['Price Check Categories',      'Catering Items, Catering - Bulk Drinks'],
    ['Quote Valid For (Days)',       30],
    ['Delivery Warning Count',       3],
    ['Delivery Warning Window (Minutes)', 60],
    ['Confirmation Enabled',        'FALSE'],
    ['Confirmation Subject',        'See you tomorrow — your Chick-fil-A catering order {{quoteId}}'],
    ['Confirmation Body',           'Hi {{contactPerson}},\n\nJust confirming your catering order for {{customer}} — we\'ll see you {{when}}.\n\nOrder total: {{total}}\n\nIf anything has changed, reply to this email or call us at {{phone}} and we\'ll take care of it.\n\nSee you soon!\n{{contact}}\nChick-fil-A {{location}}'],
    ['PO Alert Enabled',            'TRUE'],
    ['PO Alert Days Before',         7],
    ['PO Alert Email',              ''],
    ['Off-Menu Markup (%)',          30],
    ['Tax Form Request Subject',    'Tax-exempt form needed for your catering quote {{quoteId}}'],
    ['Tax Form Request Body',       'Hi {{contactPerson}},\n\nThanks for your catering order with Chick-fil-A {{location}}! To honor the tax exemption on quote {{quoteId}}, we need a copy of your organization\'s tax-exempt form on file.\n\nPlease upload a PDF of the form using this secure link:\n{{uploadLink}}\n\nIt takes about a minute and asks for your name and quote number ({{quoteId}}) so we can match it up.\n\nThank you!\n{{contact}}\nChick-fil-A {{location}}\n{{phone}}'],
    ['Logo (Base64)',                ''],
    ['Email Subject',               'Your Catering Quote from Chick-fil-A {{location}}'],
    ['Email Body',                  'Hi {{customer}},\n\nThank you for considering Chick-fil-A {{location}} for your catering needs! We appreciate you reaching out and would love to help make your event something special.\n\nPlease find your catering quote attached. If you have any questions or would like to make any changes, don\'t hesitate to reach out — we\'re happy to help.\n\nWe look forward to serving you!\n\nWarm regards,\n{{contact}}\nChick-fil-A {{location}}\n{{phone}}'],
    ['BCC Email',                   ''],
    // Custom display names for the calendar color picker (e.g. a client name, "Delivery", "Pickup")
    ['Calendar Color: Lavender',    'Lavender'],
    ['Calendar Color: Sage',        'Sage'],
    ['Calendar Color: Grape',       'Grape'],
    ['Calendar Color: Flamingo',    'Flamingo'],
    ['Calendar Color: Banana',      'Banana'],
    ['Calendar Color: Tangerine',   'Tangerine'],
    ['Calendar Color: Peacock',     'Peacock'],
    ['Calendar Color: Graphite',    'Graphite'],
    ['Calendar Color: Blueberry',   'Blueberry'],
    ['Calendar Color: Basil',       'Basil'],
    ['Calendar Color: Tomato',      'Tomato']
  ];
  if (!sSheet.getRange('A1').getValue()) {
    sSheet.getRange(1, 1, seedSettings.length, 2).setValues(seedSettings);
    sSheet.setColumnWidth(1, 220);
    sSheet.setColumnWidth(2, 500);
  } else {
    // Migrate: append any seed keys that don't already exist (preserves user edits)
    var existingKeys = {};
    var existing = sSheet.getRange('A1:A' + Math.max(sSheet.getLastRow(), 1)).getValues();
    existing.forEach(function(r) { if (r[0]) existingKeys[r[0].toString().trim()] = true; });
    seedSettings.forEach(function(pair) {
      if (!existingKeys[pair[0]]) {
        sSheet.getRange(sSheet.getLastRow() + 1, 1, 1, 2).setValues([pair]);
      }
    });
  }

  // Menu — now 4 columns: Category, Item Name, Pickup Price, Delivery Price
  var mSheet = ss.getSheetByName(TAB_MENU);
  if (!mSheet) mSheet = ss.insertSheet(TAB_MENU);
  if (!mSheet.getRange('A1').getValue()) {
    mSheet.getRange(1, 1, 1, 4).setValues([['Category', 'Item Name', 'Pickup Price', 'Delivery Price']]);
    mSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    mSheet.setFrozenRows(1);
    mSheet.setColumnWidth(1, 160);
    mSheet.setColumnWidth(2, 300);
    mSheet.setColumnWidth(3, 120);
    mSheet.setColumnWidth(4, 120);
    mSheet.getRange(2, 3, 30, 2).setNumberFormat('$#,##0.00');
  }

  // Quotes
  var qSheet = ss.getSheetByName(TAB_QUOTES);
  if (!qSheet) qSheet = ss.insertSheet(TAB_QUOTES);
  if (!qSheet.getRange('A1').getValue()) {
    var h = ['Quote ID','Created Date','Customer Name','Contact Name','Order Type','Delivery Address','Line Items (JSON)','Subtotal','Tax Rate Used','Tax Amount','Total','Tax Exempt','Location Name','Customer Email','PO Number','Event Date','Calendar Event ID','Last Modified','Event Time','Order Discount Value','Order Discount Type','Quote Notes','Customer Phone'];
    qSheet.getRange(1, 1, 1, h.length).setValues([h]);
    qSheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    qSheet.setFrozenRows(1);
  } else {
    // Migrate: add new column headers if missing
    if (qSheet.getMaxColumns() < 23) qSheet.insertColumnsAfter(qSheet.getMaxColumns(), 23 - qSheet.getMaxColumns()); // a trimmed grid would make the col-23 write throw
    var lastCol = qSheet.getLastColumn();
    var newHeaders = {'Calendar Event ID': 17, 'Last Modified': 18, 'Event Time': 19, 'Order Discount Value': 20, 'Order Discount Type': 21, 'Quote Notes': 22, 'Customer Phone': 23};
    for (var label in newHeaders) {
      if (lastCol < newHeaders[label]) {
        qSheet.getRange(1, newHeaders[label]).setValue(label).setFontWeight('bold');
      }
    }
  }

  // Revisions — prior versions of edited quotes, hidden. Created by the shared helper.
  getRevisionsSheet_();

  // Hidden automation logs (day-before confirmations, missing-PO alerts)
  initLogSheet_(TAB_CONFIRMATIONS, ['Quote ID', 'Event Date', 'Sent At']);
  initLogSheet_(TAB_PO_ALERTS, ['Quote ID', 'Alerted At']);
  initLogSheet_(TAB_TAX_FORMS, ['Quote ID', 'Status', 'Updated At']);
  initLogSheet_(TAB_TAX_PENDING, ['Submitted At', 'Organization', 'Contact Name', 'Quote ID', 'PDF Link', 'PDF File ID', 'Status']);
  ensureTaxRegistrySheet_();

  // Off-menu cheat sheet — visible tab, team-editable. Base prices start blank
  // on purpose: fill them from the POS; the app computes delivery price (+markup).
  var oSheet = ss.getSheetByName(TAB_OFF_MENU);
  if (!oSheet) oSheet = ss.insertSheet(TAB_OFF_MENU);
  if (!oSheet.getRange('A1').getValue()) {
    oSheet.getRange(1, 1, 1, 2).setValues([['Item Name', 'Base Price']]).setFontWeight('bold');
    oSheet.setFrozenRows(1);
    oSheet.setColumnWidth(1, 300);
    oSheet.setColumnWidth(2, 120);
    var offMenuSeed = ['Sliced Tomatoes', 'Lettuce', 'Sliced Cheese', 'Folded Yellow Egg', 'Folded White Egg', 'English Muffins', 'Shredded Cheese', 'Blue Cheese Crumbles', 'Bacon Crumbles', 'Sliced Hardboiled Egg', '12ct Regular Nugget Package Meal (chips & cookie)', 'Strawberries', 'Blueberries', 'Apples'];
    oSheet.getRange(2, 1, offMenuSeed.length, 2).setValues(offMenuSeed.map(function(n) { return [n, '']; }));
    oSheet.getRange(2, 2, offMenuSeed.length, 1).setNumberFormat('$#,##0.00');
  }

  // Sequence
  var seqSheet = ss.getSheetByName(TAB_QUOTE_SEQUENCE);
  if (!seqSheet) seqSheet = ss.insertSheet(TAB_QUOTE_SEQUENCE);
  if (!seqSheet.getRange('A1').getValue()) {
    seqSheet.getRange('A1').setValue('Last Quote Number Used');
    seqSheet.getRange('A1').setFontWeight('bold');
    seqSheet.getRange('B1').setValue(0);
    seqSheet.setColumnWidth(1, 220);
  }

  try { var d = ss.getSheetByName('Sheet1'); if (d && d.getLastRow() <= 1 && d.getLastColumn() <= 1) ss.deleteSheet(d); } catch(e) {}
  return 'Sheet initialized successfully!';
}


// ── PRINT DATA ───────────────────────────────────────────────

function getPrintData(quoteData) {
  var settings = getSettings();
  // "Quoted on" + "valid through" — both anchored to the quote's created date
  // (today for brand-new quotes). Valid-through hides if the setting is blank/0.
  var tz = Session.getScriptTimeZone();
  var createdBase = quoteData.createdDate ? new Date(quoteData.createdDate) : new Date();
  if (isNaN(createdBase.getTime())) createdBase = new Date();
  var quotedOn = Utilities.formatDate(createdBase, tz, 'M/d/yyyy');
  var validThrough = '';
  var validDays = parseInt(settings['Quote Valid For (Days)'], 10);
  if (!isNaN(validDays) && validDays > 0) {
    var expiry = new Date(createdBase);
    expiry.setDate(expiry.getDate() + validDays);
    validThrough = Utilities.formatDate(expiry, tz, 'M/d/yyyy');
  }
  var storeName = quoteData.locationName || settings['Store Name (Active)'] || '';
  var storeAddress = '', storePhone = '';
  if (storeName === (settings['Location 1 Name'] || '')) {
    storeAddress = settings['Location 1 Address'] || ''; storePhone = settings['Location 1 Phone'] || '';
  } else if (storeName === (settings['Location 2 Name'] || '')) {
    storeAddress = settings['Location 2 Address'] || ''; storePhone = settings['Location 2 Phone'] || '';
  } else {
    storeAddress = settings['Location 1 Address'] || ''; storePhone = settings['Location 1 Phone'] || '';
  }
  return {
    logo: settings['Logo (Base64)'] || '', storeName: storeName,
    storeAddress: storeAddress, storePhone: storePhone,
    contactName: settings['Quote Contact Name'] || '',
    customerName: quoteData.customerName || '', contactPerson: quoteData.contactName || '',
    customerEmail: quoteData.customerEmail || '', customerPhone: quoteData.customerPhone || '',
    validThrough: validThrough, quotedOn: quotedOn,
    orderType: quoteData.orderType || 'Pickup', deliveryAddress: quoteData.deliveryAddress || '',
    directionsUrl: (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress) ? mapsDirectionsUrl_(quoteData.deliveryAddress) : '',
    date: quoteData.date || new Date().toLocaleDateString(), time: quoteData.time || '',
    lineItems: quoteData.lineItems || [], subtotal: quoteData.subtotal || 0,
    taxRate: quoteData.taxRate || 0, taxAmount: quoteData.taxAmount || 0,
    total: quoteData.total || 0, taxExempt: quoteData.taxExempt || false,
    quoteId: quoteData.quoteId || '',
    poNumber: quoteData.poNumber === 'NO_PO_NEEDED' ? '' : (quoteData.poNumber || ''),
    orderDiscountValue: parseFloat(quoteData.orderDiscountValue) || 0,
    orderDiscountType: quoteData.orderDiscountType || 'percent',
    orderDiscountAmount: parseFloat(quoteData.orderDiscountAmount) || 0,
    quoteNotes: quoteData.quoteNotes || ''
  };
}


// Google Maps directions link for a delivery address — no API key needed.
function mapsDirectionsUrl_(addr) {
  return 'https://www.google.com/maps/dir/?api=1&destination=' + encodeURIComponent(addr);
}


// ── SERVER-SIDE PDF ──────────────────────────────────────────

function buildPdfHtml(pd) {
  var items = pd.lineItems || [];
  if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
  var liHtml = '';
  var itemDiscTotal = 0;
  items.forEach(function(it) {
    var da = parseFloat(it.discountAmount) || 0;
    itemDiscTotal += da;
    var net = (parseFloat(it.amount)||0) - da;
    var descCell = esc(it.description) + (da > 0 ? '<div style="font-size:12px;color:#6B7280;font-style:italic;margin-top:2px;">Discount: −$' + da.toFixed(2) + '</div>' : '');
    liHtml += '<tr><td style="text-align:center;padding:10px 12px;border-bottom:1px solid #E5E7EB;font-size:14px;">' + it.quantity + '</td><td style="padding:10px 12px;border-bottom:1px solid #E5E7EB;font-size:14px;">' + descCell + '</td><td style="text-align:right;padding:10px 12px;border-bottom:1px solid #E5E7EB;font-family:Courier New,monospace;font-size:14px;">$' + (it.price||0).toFixed(2) + '</td><td style="text-align:right;padding:10px 12px;border-bottom:1px solid #E5E7EB;font-family:Courier New,monospace;font-size:14px;font-weight:600;">$' + net.toFixed(2) + '</td></tr>';
  });
  var taxEx = pd.taxExempt === true || pd.taxExempt === 'TRUE';
  var taxLine = taxEx ? '<tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">TAX EXEMPT</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">$0.00</td></tr>' : '<tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">SALES TAX (' + (pd.taxRate||0) + '%)</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">$' + (pd.taxAmount||0).toFixed(2) + '</td></tr>';
  var orderDiscAmt = parseFloat(pd.orderDiscountAmount) || 0;
  if (!orderDiscAmt) {
    var afterItem = (parseFloat(pd.subtotal) || 0) - itemDiscTotal;
    var odv = parseFloat(pd.orderDiscountValue) || 0;
    var odt = pd.orderDiscountType || 'percent';
    orderDiscAmt = odt === 'percent' ? afterItem * (odv/100) : Math.min(odv, afterItem);
    if (orderDiscAmt < 0) orderDiscAmt = 0;
  }
  var discRows = '';
  if (itemDiscTotal > 0) discRows += '<tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">ITEM DISCOUNTS</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">−$' + itemDiscTotal.toFixed(2) + '</td></tr>';
  if (orderDiscAmt > 0) {
    var odLbl = 'ORDER DISCOUNT' + (pd.orderDiscountType === 'percent' && pd.orderDiscountValue ? ' (' + pd.orderDiscountValue + '%)' : '');
    discRows += '<tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">' + odLbl + '</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">−$' + orderDiscAmt.toFixed(2) + '</td></tr>';
  }
  var notesHtml = pd.quoteNotes ? '<div style="background:#FFFBEB;border-left:3px solid #D97706;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#374151;border-radius:4px;"><div style="font-weight:700;font-size:11px;text-transform:uppercase;color:#6B7280;margin-bottom:4px;">Notes</div>' + esc(pd.quoteNotes).split('\n').join('<br>') + '</div>' : '';
  var addr = pd.orderType === 'Delivery' ? (pd.deliveryAddress||'') : (pd.storeAddress||'');
  var logo = pd.logo ? '<img src="'+pd.logo+'" style="max-width:140px;max-height:80px;margin-bottom:8px;">' : '<div style="font-size:24px;font-weight:800;color:#E51636;margin-bottom:8px;">Chick-fil-A</div>';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>@page{size:letter;margin:0.6in;}body{font-family:Helvetica,Arial,sans-serif;color:#1F2937;margin:0;padding:40px;}</style></head><body>' +
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;"><div>'+logo+'<div style="font-size:16px;font-weight:700;">'+esc(pd.storeName)+'</div><div style="font-size:13px;color:#6B7280;margin-top:2px;">'+esc(pd.storeAddress)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.storePhone)+'</div></div><div style="text-align:right;"><div style="font-size:36px;font-weight:300;color:#D1D5DB;letter-spacing:2px;">QUOTE</div><div style="font-size:13px;color:#6B7280;margin-top:4px;font-family:Courier New,monospace;">'+esc(pd.quoteId)+'</div>'+(pd.poNumber?'<div style="font-size:13px;color:#6B7280;margin-top:2px;">PO: '+esc(pd.poNumber)+'</div>':'')+'<div style="font-size:13px;color:#6B7280;margin-top:8px;">Date: '+esc(pd.date)+'</div>'+(pd.time?'<div style="font-size:13px;color:#6B7280;">Time: '+esc(pd.time)+'</div>':'')+(pd.quotedOn?'<div style="font-size:12px;color:#9CA3AF;margin-top:2px;">Quoted on '+esc(pd.quotedOn)+'</div>':'')+(pd.validThrough?'<div style="font-size:12px;color:#9CA3AF;">Quote valid through '+esc(pd.validThrough)+'</div>':'')+'<div style="font-size:14px;font-weight:700;margin-top:8px;">For: '+esc(pd.customerName)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.orderType)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(addr)+'</div>'+(pd.directionsUrl?'<div style="font-size:12px;"><a href="'+pd.directionsUrl+'" style="color:#2563EB;">Get Directions (Google Maps)</a></div>':'')+'<div style="font-size:13px;font-weight:600;margin-top:4px;">'+esc(pd.contactPerson)+'</div>'+(pd.customerEmail?'<div style="font-size:13px;color:#6B7280;">'+esc(pd.customerEmail)+'</div>':'')+(pd.customerPhone?'<div style="font-size:13px;color:#6B7280;">'+esc(pd.customerPhone)+'</div>':'')+'</div></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;"><thead><tr style="background:#F3F4F6;"><th style="padding:10px 12px;text-align:center;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:10%;">QTY</th><th style="padding:10px 12px;text-align:left;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:50%;">DESCRIPTION</th><th style="padding:10px 12px;text-align:right;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:18%;">PRICE/ITEM</th><th style="padding:10px 12px;text-align:right;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:22%;">AMOUNT</th></tr></thead><tbody>'+liHtml+'</tbody></table>' +
    notesHtml +
    '<div style="display:flex;justify-content:space-between;align-items:flex-end;"><div style="font-size:13px;color:#6B7280;max-width:320px;">If you have any questions concerning this Quote, contact:<br><strong style="color:#1F2937;">'+esc(pd.contactName)+' at '+esc(pd.storePhone)+'</strong></div><table style="width:280px;background:#F9FAFB;border-radius:8px;overflow:hidden;"><tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">SUBTOTAL</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">$'+(pd.subtotal||0).toFixed(2)+'</td></tr>'+discRows+taxLine+'<tr style="background:#E5E7EB;"><td style="padding:10px 16px;font-size:15px;font-weight:700;">TOTAL</td><td style="padding:10px 16px;text-align:right;font-family:Courier New,monospace;font-size:16px;font-weight:700;">$'+(pd.total||0).toFixed(2)+'</td></tr></table></div></body></html>';
}

function generatePdfBlob(quoteData) {
  var pd = getPrintData(quoteData);
  pd.quoteId = quoteData.quoteId || '';
  return HtmlService.createHtmlOutput(buildPdfHtml(pd)).getBlob().setName((quoteData.quoteId||'Quote')+'.pdf').getAs('application/pdf');
}


// ── EMAIL ────────────────────────────────────────────────────

function sendQuoteEmail(quoteData, recipientEmail) {
  var settings = getSettings();
  var storeName = quoteData.locationName || settings['Store Name (Active)'] || '';
  var storePhone = '';
  if (storeName === (settings['Location 1 Name']||'')) storePhone = settings['Location 1 Phone']||'';
  else if (storeName === (settings['Location 2 Name']||'')) storePhone = settings['Location 2 Phone']||'';
  else storePhone = settings['Location 1 Phone']||'';

  var reps = { '{{customer}}': quoteData.customerName||'', '{{contact}}': settings['Quote Contact Name']||'', '{{location}}': storeName, '{{phone}}': storePhone, '{{quoteId}}': quoteData.quoteId||'', '{{total}}': '$'+(parseFloat(quoteData.total)||0).toFixed(2), '{{date}}': quoteData.date||new Date().toLocaleDateString() };
  var subj = settings['Email Subject'] || 'Your Catering Quote from Chick-fil-A {{location}}';
  var body = settings['Email Body'] || 'Hi {{customer}},\n\nPlease find your catering quote attached.\n\nThank you!\n{{contact}}';
  var bcc = (settings['BCC Email']||'').toString().trim();
  for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
  var opts = { attachments: [generatePdfBlob(quoteData)], name: 'Chick-fil-A ' + storeName };
  if (bcc) opts.bcc = bcc;
  MailApp.sendEmail(recipientEmail, subj, body, opts);
  return { success: true, message: 'Email sent to ' + recipientEmail };
}

function sendQuoteEmailDirect(quoteData, recipientEmail, subject, body) {
  var settings = getSettings();
  var storeName = quoteData.locationName || settings['Store Name (Active)'] || '';
  var bcc = (settings['BCC Email']||'').toString().trim();
  var opts = { attachments: [generatePdfBlob(quoteData)], name: 'Chick-fil-A ' + storeName };
  if (bcc) opts.bcc = bcc;
  MailApp.sendEmail(recipientEmail, subject, body, opts);
  return { success: true };
}

function getEmailQuota() { return MailApp.getRemainingDailyQuota(); }

function esc(s) { return s ? s.toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;') : ''; }


// ── CALENDAR EVENTS ──────────────────────────────────────────

function buildCalendarTitle(customerName, dateStr, poNumber) {
  var prefix = '\uD83D\uDD34 NEEDS PO';
  if (poNumber === 'NO_PO_NEEDED') prefix = '\uD83D\uDFE2 NO PO NEEDED';
  else if (poNumber && poNumber.toString().trim() !== '') prefix = '\uD83D\uDFE2 HAVE PO';
  return prefix + ' \u2014 ' + customerName + ' \u2014 ' + dateStr;
}

function createCalendarEvent(quoteData, quoteId) {
  if (!quoteData.date) return '';
  if (!quoteData.time) return ''; // Need a time to create a timed event
  if (quoteData.eventColor === 'skip') return ''; // User opted out via the calendar prompt

  var settings = getSettings();
  var calId = (settings['Calendar ID'] || '').toString().trim();
  var cal = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
  if (!cal) return '';

  var storeName = quoteData.locationName || settings['Store Name (Active)'] || '';
  var customerName = quoteData.customerName || 'Unknown';

  // Parse order time (HH:MM from input[type=time]) and create event N min before order time
  var leadMin = getCalendarLeadMinutes_(settings);
  var orderTime = new Date(quoteData.date + 'T' + quoteData.time + ':00');
  var startTime = new Date(orderTime.getTime() - leadMin * 60 * 1000);

  var dateStr = orderTime.toLocaleDateString('en-US');
  var title = buildCalendarTitle(customerName, dateStr, quoteData.poNumber);

  // Build description with full quote details
  var items = quoteData.lineItems || [];
  if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
  var itemDiscTotal = 0;
  var itemLines = items.map(function(i) {
    var da = parseFloat(i.discountAmount) || 0;
    itemDiscTotal += da;
    var net = (parseFloat(i.amount) || 0) - da;
    var line = '  ' + i.quantity + 'x ' + i.description + ' — $' + net.toFixed(2);
    if (da > 0) line += ' (−$' + da.toFixed(2) + ' discount)';
    return line;
  }).join('\n');

  var orderDiscAmt = parseFloat(quoteData.orderDiscountAmount) || 0;
  if (!orderDiscAmt) {
    var afterItem = (parseFloat(quoteData.subtotal) || 0) - itemDiscTotal;
    var odv = parseFloat(quoteData.orderDiscountValue) || 0;
    var odt = quoteData.orderDiscountType || 'percent';
    orderDiscAmt = odt === 'percent' ? afterItem * (odv/100) : Math.min(odv, afterItem);
    if (orderDiscAmt < 0) orderDiscAmt = 0;
  }

  var desc = 'Quote: ' + quoteId + '\n'
    + 'Customer: ' + customerName + '\n'
    + 'Contact: ' + (quoteData.contactName || '') + '\n'
    + (quoteData.customerEmail ? 'Email: ' + quoteData.customerEmail + '\n' : '')
    + 'Order Type: ' + (quoteData.orderType || 'Pickup') + '\n'
    + (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress ? 'Delivery Address: ' + quoteData.deliveryAddress + '\nDirections: ' + mapsDirectionsUrl_(quoteData.deliveryAddress) + '\n' : '')
    + 'Location: ' + storeName + '\n'
    + 'PO: ' + (quoteData.poNumber === 'NO_PO_NEEDED' ? 'Not Required' : (quoteData.poNumber || 'PENDING')) + '\n'
    + '\n--- Order Details ---\n' + (itemLines || '(no items)') + '\n'
    + '\nSubtotal: $' + (parseFloat(quoteData.subtotal) || 0).toFixed(2)
    + (itemDiscTotal > 0 ? '\nItem Discounts: −$' + itemDiscTotal.toFixed(2) : '')
    + (orderDiscAmt > 0 ? '\nOrder Discount: −$' + orderDiscAmt.toFixed(2) : '')
    + '\nTax: $' + (parseFloat(quoteData.taxAmount) || 0).toFixed(2)
    + (quoteData.taxExempt ? ' (Tax Exempt)' : ' (' + (quoteData.taxRate || 0) + '%)')
    + '\nTotal: $' + (parseFloat(quoteData.total) || 0).toFixed(2)
    + (quoteData.quoteNotes ? '\n\n--- Notes ---\n' + quoteData.quoteNotes : '')
    + '\n\nQuote submitted: ' + new Date().toLocaleString();

  var event = cal.createEvent(title, startTime, orderTime, { description: desc });
  if (quoteData.eventColor && CalendarApp.EventColor[quoteData.eventColor]) {
    try { event.setColor(CalendarApp.EventColor[quoteData.eventColor]); } catch(e) {}
  }
  return event.getId();
}

// Full update of calendar event (date, time, title, description) on quote edit
function updateCalendarEvent(quoteData, quoteId, calendarEventId) {
  if (!calendarEventId) return '';

  var settings = getSettings();
  var calId = (settings['Calendar ID'] || '').toString().trim();
  var cal = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
  if (!cal) return calendarEventId;

  try {
    var event = cal.getEventById(calendarEventId);
    if (!event) {
      // Event was deleted from calendar — create a fresh one
      return createCalendarEvent(quoteData, quoteId) || '';
    }

    var storeName = quoteData.locationName || settings['Store Name (Active)'] || '';
    var customerName = quoteData.customerName || 'Unknown';

    // Update time if date and time are present
    if (quoteData.date && quoteData.time) {
      var leadMin = getCalendarLeadMinutes_(settings);
      var orderTime = new Date(quoteData.date + 'T' + quoteData.time + ':00');
      var startTime = new Date(orderTime.getTime() - leadMin * 60 * 1000);
      event.setTime(startTime, orderTime);
      var dateStr = orderTime.toLocaleDateString('en-US');
      event.setTitle(buildCalendarTitle(customerName, dateStr, quoteData.poNumber));
    } else if (quoteData.date) {
      var dateStr = new Date(quoteData.date + 'T00:00:00').toLocaleDateString('en-US');
      event.setTitle(buildCalendarTitle(customerName, dateStr, quoteData.poNumber));
    }

    // Rebuild description with updated quote details
    var items = quoteData.lineItems || [];
    if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
    var itemDiscTotal2 = 0;
    var itemLines = items.map(function(i) {
      var da = parseFloat(i.discountAmount) || 0;
      itemDiscTotal2 += da;
      var net = (parseFloat(i.amount) || 0) - da;
      var line = '  ' + i.quantity + 'x ' + i.description + ' — $' + net.toFixed(2);
      if (da > 0) line += ' (−$' + da.toFixed(2) + ' discount)';
      return line;
    }).join('\n');

    var orderDiscAmt2 = parseFloat(quoteData.orderDiscountAmount) || 0;
    if (!orderDiscAmt2) {
      var afterItem2 = (parseFloat(quoteData.subtotal) || 0) - itemDiscTotal2;
      var odv2 = parseFloat(quoteData.orderDiscountValue) || 0;
      var odt2 = quoteData.orderDiscountType || 'percent';
      orderDiscAmt2 = odt2 === 'percent' ? afterItem2 * (odv2/100) : Math.min(odv2, afterItem2);
      if (orderDiscAmt2 < 0) orderDiscAmt2 = 0;
    }

    var desc = 'Quote: ' + quoteId + '\n'
      + 'Customer: ' + customerName + '\n'
      + 'Contact: ' + (quoteData.contactName || '') + '\n'
      + (quoteData.customerEmail ? 'Email: ' + quoteData.customerEmail + '\n' : '')
      + 'Order Type: ' + (quoteData.orderType || 'Pickup') + '\n'
      + (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress ? 'Delivery Address: ' + quoteData.deliveryAddress + '\nDirections: ' + mapsDirectionsUrl_(quoteData.deliveryAddress) + '\n' : '')
      + 'Location: ' + storeName + '\n'
      + 'PO: ' + (quoteData.poNumber === 'NO_PO_NEEDED' ? 'Not Required' : (quoteData.poNumber || 'PENDING')) + '\n'
      + '\n--- Order Details ---\n' + (itemLines || '(no items)') + '\n'
      + '\nSubtotal: $' + (parseFloat(quoteData.subtotal) || 0).toFixed(2)
      + (itemDiscTotal2 > 0 ? '\nItem Discounts: −$' + itemDiscTotal2.toFixed(2) : '')
      + (orderDiscAmt2 > 0 ? '\nOrder Discount: −$' + orderDiscAmt2.toFixed(2) : '')
      + '\nTax: $' + (parseFloat(quoteData.taxAmount) || 0).toFixed(2)
      + (quoteData.taxExempt ? ' (Tax Exempt)' : ' (' + (quoteData.taxRate || 0) + '%)')
      + '\nTotal: $' + (parseFloat(quoteData.total) || 0).toFixed(2)
      + (quoteData.quoteNotes ? '\n\n--- Notes ---\n' + quoteData.quoteNotes : '')
      + '\n\nLast updated: ' + new Date().toLocaleString();

    event.setDescription(desc);
    return calendarEventId;
  } catch(e) {
    return calendarEventId;
  }
}

// Update calendar event title when PO changes
// Returns true on success, or a string reason on failure
function updateCalendarEventPO(sheetRow, poNumber, calendarEventId) {
  if (!calendarEventId) return 'Empty event ID';

  var settings = getSettings();
  var calId = (settings['Calendar ID'] || '').toString().trim();

  // Try the configured calendar first, then fall back to the default \u2014
  // covers the case where Settings was changed after the event was created.
  var event = null;
  var triedCalendars = [];
  function tryCal(c, label) {
    if (!c) return;
    triedCalendars.push(label);
    try { var e = c.getEventById(calendarEventId); if (e) event = e; } catch(_) {}
  }
  if (calId) tryCal(CalendarApp.getCalendarById(calId), 'configured (' + calId + ')');
  if (!event) tryCal(CalendarApp.getDefaultCalendar(), 'default');
  if (!event) return 'Event ID not found in: ' + triedCalendars.join(', ');

  try {
    var oldTitle = event.getTitle();
    var parts = oldTitle.split(' \u2014 ');
    var customerName = parts.length >= 2 ? parts[1] : '';
    var dateStr = parts.length >= 3 ? parts[2] : '';
    event.setTitle(buildCalendarTitle(customerName, dateStr, poNumber));
    return true;
  } catch(e) {
    return 'setTitle failed: ' + (e.message || String(e));
  }
}


// ── REMINDERS ────────────────────────────────────────────────

const TAB_REMINDERS = 'Reminders_Sent';

function initRemindersSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_REMINDERS);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_REMINDERS);
    sheet.getRange(1, 1, 1, 2).setValues([['Quote ID', 'Reminder Sent Date']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRemindedQuoteIds() {
  var sheet = getSpreadsheet().getSheetByName(TAB_REMINDERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return r[0].toString().trim(); })
    .filter(function(id) { return id !== ''; });
}

function markReminderSent(quoteId) {
  initRemindersSheet().appendRow([quoteId, new Date()]);
}

function processReminders() {
  try {
    return _processReminders();
  } catch (e) {
    var alertEmail = getSettings()['Reminder Internal Email'] || Session.getEffectiveUser().getEmail();
    if (alertEmail) {
      MailApp.sendEmail(alertEmail, 'Catering Reminder Trigger Error', 'processReminders failed:\n\n' + e.message + '\n\n' + e.stack);
    }
    throw e;
  }
}

function _processReminders() {
  var settings = getSettings();
  if (settings['Reminder Enabled'] !== 'TRUE') return 'Reminders are disabled.';

  var daysAfter      = parseInt(settings['Reminder After Days']) || 3;
  var toCustomer     = settings['Reminder Send To Customer'] === 'TRUE';
  var toInternal     = settings['Reminder Send To Internal'] === 'TRUE';
  var internalEmail  = (settings['Reminder Internal Email'] || '').toString().trim();
  var subject        = settings['Reminder Subject'] || 'Following up on your Catering Quote \u2014 {{quoteId}}';
  var body           = settings['Reminder Body']    || 'Hi {{customer}},\n\nWe wanted to follow up on the catering quote we sent you {{daysSince}} days ago (Quote {{quoteId}}, total: {{total}}).\n\nPlease let us know if you have any questions!\n\n{{contact}}\nChick-fil-A {{location}}';

  var quotes     = getQuotes();
  var remindedIds = getRemindedQuoteIds();
  var now        = new Date();
  var cutoffMs   = daysAfter * 24 * 60 * 60 * 1000;
  var count      = 0;

  quotes.forEach(function(q) {
    if (!q.createdDate) return;
    if (remindedIds.indexOf(q.quoteId.toString()) >= 0) return;

    var ageMs = now - new Date(q.createdDate);
    if (ageMs < cutoffMs) return;

    var daysSince = Math.floor(ageMs / (24 * 60 * 60 * 1000));
    var storeName = q.locationName || settings['Store Name (Active)'] || '';
    var reps = {
      '{{customer}}':  q.customerName || '',
      '{{contact}}':   settings['Quote Contact Name'] || '',
      '{{location}}':  storeName,
      '{{quoteId}}':   q.quoteId || '',
      '{{total}}':     '$' + (parseFloat(q.total) || 0).toFixed(2),
      '{{date}}':      q.createdDate ? new Date(q.createdDate).toLocaleDateString() : '',
      '{{daysSince}}': daysSince.toString()
    };
    var subj = subject, bdy = body;
    for (var k in reps) { subj = subj.split(k).join(reps[k]); bdy = bdy.split(k).join(reps[k]); }

    var sent = false;
    if (toCustomer && q.customerEmail) {
      MailApp.sendEmail(q.customerEmail, subj, bdy, { name: 'Chick-fil-A ' + storeName });
      sent = true;
    }
    if (toInternal && internalEmail) {
      var intBody = 'Customer: ' + (q.customerName || '') + '\nEmail: ' + (q.customerEmail || 'N/A') + '\n\n' + bdy;
      MailApp.sendEmail(internalEmail, '[Follow-Up] ' + subj, intBody, { name: 'Chick-fil-A ' + storeName });
      sent = true;
    }
    if (sent) {
      markReminderSent(q.quoteId);
      count++;
    }
  });

  return 'Sent ' + count + ' reminder(s).';
}

function setupReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processReminders').timeBased().everyDays(1).atHour(9).create();
  return 'Daily reminder trigger enabled (runs at 9am).';
}

function removeReminderTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processReminders') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed > 0 ? 'Reminder trigger removed.' : 'No active reminder trigger found.';
}

function getReminderTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'processReminders';
  });
}


// ── DAILY AUTOMATION: day-before confirmations + missing-PO alerts ──
// One trigger (3pm daily) runs both sweeps; each is gated by its own Settings toggle.

const TAB_CONFIRMATIONS = 'Confirmations_Sent';
const TAB_PO_ALERTS     = 'PO_Alerts_Sent';

// Get-or-create a hidden log tab. Same pattern as initRemindersSheet.
function initLogSheet_(tabName, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

// "Sat, Jul 18 at 11:30 AM" from stored yyyy-MM-dd + HH:mm.
function formatWhen_(eventDate, eventTime, tz) {
  if (!eventDate) return '';
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(eventDate);
  if (!m) return eventDate;
  var d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  var out = Utilities.formatDate(d, tz, 'EEE, MMM d');
  var t = /^(\d{1,2}):(\d{2})/.exec(eventTime || '');
  if (t) {
    var h = parseInt(t[1], 10), mn = parseInt(t[2], 10);
    var ampm = h >= 12 ? 'PM' : 'AM';
    var h12 = h % 12; if (h12 === 0) h12 = 12;
    out += ' at ' + h12 + ':' + (mn < 10 ? '0' + mn : mn) + ' ' + ampm;
  }
  return out;
}

function poAlertRecipients_(settings) {
  var raw = (settings['PO Alert Email'] || '').toString().trim();
  if (raw) return raw;
  try { return Session.getEffectiveUser().getEmail(); } catch(e) { return ''; }
}

function dailyCateringAutomation() {
  var alertTo = '';
  try { alertTo = poAlertRecipients_(getSettings()); } catch(e) {}
  try { _sendDayBeforeConfirmations(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Day-before confirmations failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) {}
  }
  try { _sendPoAlerts(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Missing-PO sweep failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) {}
  }
  try { _sendYearEndTaxReview(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Year-end tax review failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) {}
  }
}

function _sendDayBeforeConfirmations() {
  var settings = getSettings();
  if (settings['Confirmation Enabled'] !== 'TRUE') return 'Confirmations are disabled.';
  var tz = Session.getScriptTimeZone();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowKey = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');

  var logSheet = initLogSheet_(TAB_CONFIRMATIONS, ['Quote ID', 'Event Date', 'Sent At']);
  var sentKeys = {};
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
      sentKeys[r[0] + '|' + normalizeEventDate_(r[1], tz)] = true;
    });
  }

  var storeName = settings['Store Name (Active)'] || '';
  var storePhone = '';
  if (storeName === (settings['Location 2 Name'] || '')) storePhone = settings['Location 2 Phone'] || '';
  else storePhone = settings['Location 1 Phone'] || '';
  var subjectTpl = settings['Confirmation Subject'] || 'See you tomorrow — your Chick-fil-A catering order {{quoteId}}';
  var bodyTpl = settings['Confirmation Body'] || 'Hi {{contactPerson}},\n\nJust confirming your catering order for {{customer}} — we\'ll see you {{when}}.\n\n{{contact}}\nChick-fil-A {{location}}';
  var bcc = (settings['BCC Email'] || '').toString().trim();

  var count = 0;
  getQuotes().forEach(function(q) {
    if (q.eventDate !== tomorrowKey) return;
    var email = (q.customerEmail || '').toString().trim();
    if (!email) return;
    if (sentKeys[q.quoteId + '|' + q.eventDate]) return;
    var reps = {
      '{{contactPerson}}': q.contactName || '',
      '{{customer}}': q.customerName || '',
      '{{contact}}': settings['Quote Contact Name'] || '',
      '{{location}}': storeName,
      '{{phone}}': storePhone,
      '{{quoteId}}': q.quoteId || '',
      '{{total}}': '$' + (parseFloat(q.total) || 0).toFixed(2),
      '{{when}}': formatWhen_(q.eventDate, q.eventTime, tz),
      '{{date}}': formatWhen_(q.eventDate, '', tz),
      '{{time}}': formatWhen_(q.eventDate, q.eventTime, tz).split(' at ')[1] || ''
    };
    var subj = subjectTpl, body = bodyTpl;
    for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
    var opts = { name: 'Chick-fil-A ' + storeName };
    if (bcc) opts.bcc = bcc;
    MailApp.sendEmail(email, subj, body, opts);
    logSheet.appendRow([q.quoteId, q.eventDate, new Date()]);
    count++;
  });
  SpreadsheetApp.flush();
  return 'Sent ' + count + ' confirmation(s).';
}

function _sendPoAlerts() {
  var settings = getSettings();
  if (settings['PO Alert Enabled'] !== 'TRUE') return 'PO alerts are disabled.';
  var days = parseInt(settings['PO Alert Days Before'], 10);
  if (isNaN(days) || days < 1) days = 7;
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var max = new Date();
  max.setDate(max.getDate() + days);
  var maxKey = Utilities.formatDate(max, tz, 'yyyy-MM-dd');

  var logSheet = initLogSheet_(TAB_PO_ALERTS, ['Quote ID', 'Alerted At']);
  var alerted = {};
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
      if (r[0]) alerted[r[0].toString().trim()] = true;
    });
  }

  var missing = getQuotes().filter(function(q) {
    if (!q.eventDate || q.eventDate < todayKey || q.eventDate > maxKey) return false;
    if ((q.poNumber || '').toString().trim() !== '') return false; // HAVE PO and NO_PO_NEEDED both count as handled
    return !alerted[q.quoteId];
  });
  if (!missing.length) return 'No missing-PO quotes in the next ' + days + ' days.';

  var to = poAlertRecipients_(settings);
  if (!to) return 'No PO alert recipient configured.';

  var lines = missing.map(function(q) {
    var daysOut = Math.round((new Date(q.eventDate + 'T12:00:00') - new Date(todayKey + 'T12:00:00')) / 86400000);
    return '• ' + q.quoteId + ' — ' + q.customerName
      + (q.contactName ? ' (' + q.contactName + (q.customerPhone ? ', ' + q.customerPhone : '') + ')' : '')
      + ' — ' + formatWhen_(q.eventDate, q.eventTime, tz)
      + ' — $' + (parseFloat(q.total) || 0).toFixed(2)
      + ' — ' + (daysOut === 0 ? 'TODAY' : daysOut + ' day' + (daysOut === 1 ? '' : 's') + ' out');
  });
  var subj = '🔴 ' + missing.length + ' catering quote' + (missing.length === 1 ? '' : 's') + ' missing a PO (next ' + days + ' days)';
  var body = 'These upcoming orders have no PO on file yet:\n\n' + lines.join('\n')
    + '\n\nAdd the PO (or mark "No PO Needed") from the quote\'s popup in the catering tool. Each quote is alerted once.';
  MailApp.sendEmail(to, subj, body, { name: 'Catering Quote Tool' });
  missing.forEach(function(q) { logSheet.appendRow([q.quoteId, new Date()]); });
  SpreadsheetApp.flush();
  return 'Alerted on ' + missing.length + ' quote(s).';
}

function setupAutomationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyCateringAutomation') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyCateringAutomation').timeBased().everyDays(1).atHour(15).create();
  return 'Daily automation enabled (runs at 3pm).';
}

function removeAutomationTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyCateringAutomation') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed > 0 ? 'Automation trigger removed.' : 'No active automation trigger found.';
}

function getAutomationTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'dailyCateringAutomation';
  });
}


// ── TAX-EXEMPT FORM TRACKING ─────────────────────────────────
// Status lives in the hidden Tax_Forms tab (Quote ID | Status | Updated At),
// separate from the Quotes columns so nothing else shifts. Statuses:
// '' (never asked) / 'REQUESTED' (email sent) / 'ON_FILE' (form in the Drive folder).

const TAB_TAX_FORMS = 'Tax_Forms';               // per-quote answer (quote popup color)
const TAB_TAX_REGISTRY = 'Tax_Exempt_Registry';  // org-level master list (visible tab)
const TAB_TAX_PENDING = 'Tax_Form_Uploads';      // raw guest uploads awaiting review
const TAX_FORM_FOLDER_NAME = 'Catering Tax Exempt Forms';
const TAX_MAX_UPLOAD_BYTES = 10 * 1024 * 1024;   // 10 MB cap on guest uploads

// Find-or-create the Drive folder; remember its ID so renames don't orphan it.
function ensureTaxFormFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('TAX_FORM_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {} // deleted — fall through and recreate
  }
  var it = DriveApp.getFoldersByName(TAX_FORM_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(TAX_FORM_FOLDER_NAME);
  props.setProperty('TAX_FORM_FOLDER_ID', folder.getId());
  return folder;
}

// Decode a base64 PDF and drop it in the folder. Files stay private to the
// Drive owner (these are sensitive tax docs) — the team opens them via the folder.
function saveTaxPdf_(baseName, mimeType, dataBase64) {
  var bytes = Utilities.base64Decode(dataBase64);
  if (bytes.length > TAX_MAX_UPLOAD_BYTES) throw new Error('File is too large (10 MB max).');
  var safe = (baseName || 'Tax Exempt Form').replace(/[\\/:*?"<>|]/g, '-').slice(0, 120);
  var blob = Utilities.newBlob(bytes, mimeType || 'application/pdf', safe + '.pdf');
  return ensureTaxFormFolder_().createFile(blob);
}

// One call for the client: every quote's status plus the folder link.
function getTaxFormStatuses() {
  var out = { folderUrl: '', statuses: {} };
  try { out.folderUrl = ensureTaxFormFolder_().getUrl(); } catch(e) {} // Drive not authorized yet — statuses still work
  var sheet = getSpreadsheet().getSheetByName(TAB_TAX_FORMS);
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
    if (r[0]) out.statuses[r[0].toString().trim()] = (r[1] || '').toString().trim();
  });
  return out;
}

function setTaxFormStatus(quoteId, status) {
  quoteId = (quoteId || '').toString().trim();
  if (!quoteId) return false;
  var sheet = initLogSheet_(TAB_TAX_FORMS, ['Quote ID', 'Status', 'Updated At']);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0].toString().trim() === quoteId) {
        sheet.getRange(i + 2, 2, 1, 2).setValues([[status, new Date()]]);
        return true;
      }
    }
  }
  sheet.appendRow([quoteId, status, new Date()]);
  return true;
}

// Emails the guest asking them to reply with their form PDF, then marks REQUESTED.
function sendTaxFormRequest(quoteId, recipientEmail) {
  recipientEmail = (recipientEmail || '').toString().trim();
  if (!recipientEmail) return { success: false, message: 'No customer email on this quote.' };
  var quote = null;
  getQuotes().some(function(q) { if (q.quoteId === quoteId) { quote = q; return true; } return false; });
  if (!quote) return { success: false, message: 'Quote ' + quoteId + ' not found.' };
  var settings = getSettings();
  var storeName = quote.locationName || settings['Store Name (Active)'] || '';
  var storePhone = '';
  if (storeName === (settings['Location 2 Name'] || '')) storePhone = settings['Location 2 Phone'] || '';
  else storePhone = settings['Location 1 Phone'] || '';
  var uploadLink = '';
  try { uploadLink = ScriptApp.getService().getUrl() + '?view=taxform&quote=' + encodeURIComponent(quote.quoteId || ''); } catch(e) {}
  var reps = {
    '{{contactPerson}}': quote.contactName || '',
    '{{customer}}': quote.customerName || '',
    '{{contact}}': settings['Quote Contact Name'] || '',
    '{{location}}': storeName,
    '{{phone}}': storePhone,
    '{{quoteId}}': quote.quoteId || '',
    '{{uploadLink}}': uploadLink
  };
  var subj = settings['Tax Form Request Subject'] || 'Tax-exempt form needed for your catering quote {{quoteId}}';
  var body = settings['Tax Form Request Body'] || 'Hi {{contactPerson}},\n\nPlease reply with a PDF of your tax-exempt form, including your name and quote number {{quoteId}}.\n\nThank you!\n{{contact}}';
  for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
  var opts = { name: 'Chick-fil-A ' + storeName };
  var bcc = (settings['BCC Email'] || '').toString().trim();
  if (bcc) opts.bcc = bcc;
  MailApp.sendEmail(recipientEmail, subj, body, opts);
  setTaxFormStatus(quoteId, 'REQUESTED');
  return { success: true, message: 'Form request sent to ' + recipientEmail };
}

// Year-end review: on the last business day of December, one email to the
// PO-alert recipients listing the year's tax-exempt quotes and their form status.
function isLastBusinessDayOfYear_(d) {
  if (d.getMonth() !== 11) return false;
  var last = new Date(d.getFullYear(), 11, 31);
  while (last.getDay() === 0 || last.getDay() === 6) last.setDate(last.getDate() - 1);
  return d.getDate() === last.getDate();
}

function _sendYearEndTaxReview() {
  var today = new Date();
  if (!isLastBusinessDayOfYear_(today)) return 'Not the last business day of the year.';
  var props = PropertiesService.getScriptProperties();
  var guard = 'TAX_REVIEW_SENT_' + today.getFullYear();
  if (props.getProperty(guard)) return 'Already sent this year.';
  var settings = getSettings();
  var to = poAlertRecipients_(settings);
  if (!to) return 'No recipient configured.';
  var year = today.getFullYear();
  var reg = getTaxRegistry();
  var regLines = reg.entries.map(function(e) {
    var dot = e.status === 'ON_FILE' ? '🟢' : (e.status === 'REQUESTED' ? '🟠' : '🔴');
    return '• ' + dot + ' ' + e.organization + (e.dateOnFile ? ' — on file ' + e.dateOnFile : '') + (e.pdfUrl ? ' — ' + e.pdfUrl : '');
  });
  var exemptCount = getQuotes().filter(function(q) {
    var te = q.taxExempt === 'TRUE' || q.taxExempt === true;
    return te && q.createdDate && new Date(q.createdDate).getFullYear() === year;
  }).length;
  var body = 'It\'s the last business day of ' + year + ' — time to review the tax-exempt forms.\n\n'
    + 'You created ' + exemptCount + ' tax-exempt quote(s) this year.\n\n'
    + (reg.entries.length ? 'Tax-Exempt Registry (' + reg.entries.length + ' organizations):\n\n' + regLines.join('\n') : 'The registry is empty — nothing recorded yet.')
    + (reg.pending.length ? '\n\n⚠ ' + reg.pending.length + ' guest upload(s) still awaiting review in the app.' : '')
    + (reg.folderUrl ? '\n\nForms folder: ' + reg.folderUrl : '')
    + '\n\nOpen each 🟢 to confirm the form is current; chase anything 🟠 or 🔴 before the next order.';
  MailApp.sendEmail(to, '📋 Year-end review: catering tax-exempt forms (' + year + ')', body, { name: 'Catering Quote Tool' });
  props.setProperty(guard, 'sent');
  return 'Year-end tax review sent.';
}


// ── TAX-EXEMPT REGISTRY + GUEST UPLOADS ──────────────────────
// Registry (Tax_Exempt_Registry): one row per organization —
//   Organization | Status | PDF Link | PDF File ID | Date On File | Notes | Updated At
// Guest uploads land in Tax_Form_Uploads as PENDING; a human confirms each into
// the registry (they pick the matching org — the system never auto-matches names).

// Visible, team-reviewable tab (unlike the hidden log tabs).
function ensureTaxRegistrySheet_() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_TAX_REGISTRY);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_TAX_REGISTRY);
    sheet.getRange(1, 1, 1, 7).setValues([['Organization', 'Status', 'PDF Link', 'PDF File ID', 'Date On File', 'Notes', 'Updated At']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 260);
    sheet.setColumnWidth(3, 260);
  }
  return sheet;
}

function getTaxRegistry() {
  var out = { folderUrl: '', entries: [], pending: [] };
  try { out.folderUrl = ensureTaxFormFolder_().getUrl(); } catch(e) {}
  var reg = getSpreadsheet().getSheetByName(TAB_TAX_REGISTRY);
  if (reg && reg.getLastRow() > 1) {
    reg.getRange(2, 1, reg.getLastRow() - 1, 7).getValues().forEach(function(r, i) {
      if (!r[0] || r[0].toString().trim() === '') return;
      out.entries.push({
        organization: r[0].toString().trim(),
        status: (r[1] || '').toString().trim() || 'ON_FILE',
        pdfUrl: (r[2] || '').toString().trim(),
        dateOnFile: normalizeEventDate_(r[4], Session.getScriptTimeZone()),
        notes: (r[5] || '').toString(),
        row: i + 2
      });
    });
    out.entries.sort(function(a, b) { return a.organization.toLowerCase() < b.organization.toLowerCase() ? -1 : 1; });
  }
  var pend = getSpreadsheet().getSheetByName(TAB_TAX_PENDING);
  if (pend && pend.getLastRow() > 1) {
    pend.getRange(2, 1, pend.getLastRow() - 1, 7).getValues().forEach(function(r, i) {
      if ((r[6] || '').toString().trim() !== 'PENDING') return;
      out.pending.push({
        submittedAt: r[0] ? new Date(r[0]).toISOString() : '',
        organization: (r[1] || '').toString().trim(),
        contactName: (r[2] || '').toString().trim(),
        quoteId: (r[3] || '').toString().trim(),
        pdfUrl: (r[4] || '').toString().trim(),
        row: i + 2
      });
    });
  }
  return out;
}

// Add or update a registry row. row>0 updates that row; otherwise appends.
function saveTaxRegistryEntry(entry) {
  var org = (entry.organization || '').toString().trim();
  if (!org) return { success: false, message: 'Organization name is required.' };
  var sheet = ensureTaxRegistrySheet_();
  var vals = [org, entry.status || 'ON_FILE', entry.pdfUrl || '', entry.pdfFileId || '', entry.dateOnFile || '', entry.notes || '', new Date()];
  var row = parseInt(entry.row, 10);
  if (row && row > 1) {
    // Preserve the existing PDF link/id if this update didn't supply new ones
    var cur = sheet.getRange(row, 1, 1, 7).getValues()[0];
    if (!vals[2]) vals[2] = cur[2];
    if (!vals[3]) vals[3] = cur[3];
    sheet.getRange(row, 1, 1, 7).setValues([vals]);
  } else {
    sheet.appendRow(vals);
  }
  return { success: true };
}

function deleteTaxRegistryEntry(row) {
  var sheet = getSpreadsheet().getSheetByName(TAB_TAX_REGISTRY);
  if (sheet && row > 1) sheet.deleteRow(row);
  return true;
}

// PUBLIC — called by the guest upload page (anonymous users on the public deployment).
// Stores the PDF and logs a PENDING row for internal review. Never touches the registry directly.
function submitGuestTaxForm(payload) {
  payload = payload || {};
  var org = (payload.organization || '').toString().trim();
  var name = (payload.contactName || '').toString().trim();
  var quoteId = (payload.quoteId || '').toString().trim();
  var data = (payload.dataBase64 || '').toString();
  if (!org || !name) return { success: false, message: 'Please enter your name and organization.' };
  if (!data) return { success: false, message: 'Please attach a PDF of your form.' };
  var mime = (payload.mimeType || '').toString();
  if (mime && mime.indexOf('pdf') < 0) return { success: false, message: 'Please upload a PDF file.' };
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var file;
  try {
    file = saveTaxPdf_(org + ' - ' + name + (quoteId ? ' - ' + quoteId : '') + ' - ' + stamp, 'application/pdf', data);
  } catch (err) {
    return { success: false, message: err.message || 'Upload failed — please try again.' };
  }
  var sheet = initLogSheet_(TAB_TAX_PENDING, ['Submitted At', 'Organization', 'Contact Name', 'Quote ID', 'PDF Link', 'PDF File ID', 'Status']);
  sheet.appendRow([new Date(), org, name, quoteId, file.getUrl(), file.getId(), 'PENDING']);
  return { success: true, message: 'Thank you! We\'ve received your tax-exempt form.' };
}

// Internal: confirm a pending guest upload into the registry under a chosen org name.
function confirmPendingTaxUpload(pendingRow, organization) {
  var org = (organization || '').toString().trim();
  if (!org) return { success: false, message: 'Pick an organization name.' };
  var pend = getSpreadsheet().getSheetByName(TAB_TAX_PENDING);
  if (!pend || pendingRow < 2) return { success: false, message: 'Upload not found.' };
  var r = pend.getRange(pendingRow, 1, 1, 7).getValues()[0];
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  saveTaxRegistryEntry({ organization: org, status: 'ON_FILE', pdfUrl: r[4], pdfFileId: r[5], dateOnFile: stamp, notes: 'Uploaded by ' + (r[2] || 'guest') });
  pend.getRange(pendingRow, 7).setValue('CONFIRMED');
  return { success: true };
}

function dismissPendingTaxUpload(pendingRow) {
  var pend = getSpreadsheet().getSheetByName(TAB_TAX_PENDING);
  if (pend && pendingRow > 1) pend.getRange(pendingRow, 7).setValue('DISMISSED');
  return { success: true };
}

// Internal backfill: upload a PDF we already have on file straight into the registry.
function backfillTaxForm(payload) {
  payload = payload || {};
  var org = (payload.organization || '').toString().trim();
  if (!org) return { success: false, message: 'Organization name is required.' };
  var pdfUrl = '', pdfFileId = '';
  if (payload.dataBase64) {
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var file;
    try { file = saveTaxPdf_(org + ' - ' + stamp, payload.mimeType || 'application/pdf', payload.dataBase64); }
    catch (err) { return { success: false, message: err.message || 'Upload failed.' }; }
    pdfUrl = file.getUrl(); pdfFileId = file.getId();
  }
  saveTaxRegistryEntry({
    organization: org, status: 'ON_FILE', pdfUrl: pdfUrl, pdfFileId: pdfFileId,
    dateOnFile: payload.dateOnFile || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    notes: payload.notes || 'Backfilled'
  });
  return { success: true, message: 'Added ' + org + ' to the registry.' };
}


// ── DELIVERY RUNSHEET EMAIL ──────────────────────────────────

function emailRunsheet(dateStr, recipient) {
  recipient = (recipient || '').toString().trim();
  if (!recipient) return { success: false, message: 'No recipient given.' };
  var tz = Session.getScriptTimeZone();
  var quotes = getQuotes().filter(function(q) { return q.eventDate === dateStr; });
  if (!quotes.length) return { success: false, message: 'No orders on that day.' };
  quotes.sort(function(a, b) { return (a.eventTime || '').localeCompare(b.eventTime || ''); });

  var rows = quotes.map(function(q) {
    var items = [];
    try { items = JSON.parse(q.lineItems) || []; } catch(e) {}
    var itemsHtml = items.map(function(i) { return esc(i.quantity + '× ' + i.description); }).join('<br>');
    var when = formatWhen_(q.eventDate, q.eventTime, tz);
    var time = when.indexOf(' at ') >= 0 ? when.split(' at ')[1] : 'No time';
    var po = q.poNumber === 'NO_PO_NEEDED' ? '🟢 No PO needed' : (q.poNumber ? '🟢 PO: ' + esc(q.poNumber) : '🔴 NEEDS PO');
    var addr = q.orderType === 'Delivery' && q.deliveryAddress
      ? esc(q.deliveryAddress) + '<br><a href="' + mapsDirectionsUrl_(q.deliveryAddress) + '">Directions</a>'
      : esc(q.orderType || '');
    return '<tr>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;white-space:nowrap;"><b>' + esc(time) + '</b></td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;"><b>' + esc(q.customerName) + '</b><br>' + esc(q.contactName || '') + (q.customerPhone ? '<br>' + esc(q.customerPhone) : '') + '</td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;">' + esc(q.orderType || '') + '</td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;">' + addr + '</td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;font-size:12px;">' + itemsHtml + '</td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;white-space:nowrap;">' + po + '</td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;text-align:right;">$' + (parseFloat(q.total) || 0).toFixed(2) + '</td>'
      + '</tr>';
  });
  var dayLabel = formatWhen_(dateStr, '', tz);
  var html = '<h2 style="font-family:Arial,sans-serif;">Catering Runsheet — ' + esc(dayLabel) + '</h2>'
    + '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;width:100%;">'
    + '<tr style="background:#f3f4f6;"><th style="padding:8px;text-align:left;">Time</th><th style="padding:8px;text-align:left;">Customer</th><th style="padding:8px;text-align:left;">Type</th><th style="padding:8px;text-align:left;">Address</th><th style="padding:8px;text-align:left;">Items</th><th style="padding:8px;text-align:left;">PO</th><th style="padding:8px;text-align:right;">Total</th></tr>'
    + rows.join('') + '</table>';
  MailApp.sendEmail(recipient, 'Catering Runsheet — ' + dayLabel + ' (' + quotes.length + ' order' + (quotes.length === 1 ? '' : 's') + ')',
    'Runsheet for ' + dayLabel + ' — open in an HTML mail client.', { htmlBody: html, name: 'Catering Quote Tool' });
  return { success: true, message: 'Runsheet sent to ' + recipient };
}


