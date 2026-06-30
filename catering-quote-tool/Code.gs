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


// ── WEB APP ENTRY POINT ──────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CFA Catering Quotes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── SETTINGS ─────────────────────────────────────────────────

function getSettings() {
  var sheet = getSpreadsheet().getSheetByName(TAB_SETTINGS);
  var data  = sheet.getRange('A1:B50').getValues();
  var settings = {};
  data.forEach(function(row) {
    if (row[0] && row[0].toString().trim() !== '') {
      settings[row[0].toString().trim()] = row[1];
    }
  });
  return settings;
}

function updateSetting(label, value) {
  var sheet = getSpreadsheet().getSheetByName(TAB_SETTINGS);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var data = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === label) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  // Label not found — append a new row instead of silently dropping the save
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
  var width = Math.max(19, Math.min(22, lastCol));
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

function getNextQuoteId() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTE_SEQUENCE);
  var current = parseInt(sheet.getRange('B1').getValue()) || 0;
  var next = current + 1;
  sheet.getRange('B1').setValue(next);
  return 'Q-' + new Date().getFullYear() + '-' + ('0000' + next).slice(-4);
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
    var existing = sheet.getRange(existingRow, 1, 1, 22).getValues()[0];
    var eventId = existing[16] ? existing[16].toString().trim() : '';

    // Update existing calendar event, or create one if it doesn't exist yet
    if (eventId) {
      try { eventId = updateCalendarEvent(quoteData, existingQuoteId, eventId) || eventId; } catch(e) {}
    } else if (quoteData.date && quoteData.time) {
      try { eventId = createCalendarEvent(quoteData, existingQuoteId) || ''; } catch(e) { eventId = ''; }
    }

    sheet.getRange(existingRow, 1, 1, 22).setValues([[
      existingQuoteId, existing[1], quoteData.customerName, quoteData.contactName,
      quoteData.orderType, quoteData.deliveryAddress || '',
      JSON.stringify(quoteData.lineItems),
      parseFloat(quoteData.subtotal) || 0, parseFloat(quoteData.taxRate) || 0,
      parseFloat(quoteData.taxAmount) || 0, parseFloat(quoteData.total) || 0,
      quoteData.taxExempt ? 'TRUE' : 'FALSE', quoteData.locationName || '',
      quoteData.customerEmail || '', quoteData.poNumber || '', quoteData.date || '',
      eventId, new Date(), quoteData.time || '',
      parseFloat(quoteData.orderDiscountValue) || 0, quoteData.orderDiscountType || 'percent', quoteData.quoteNotes || ''
    ]]);
    return existingQuoteId;
  }

  // New quote
  var quoteId = getNextQuoteId();
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
    parseFloat(quoteData.orderDiscountValue) || 0, quoteData.orderDiscountType || 'percent', quoteData.quoteNotes || ''
  ]);
  return quoteId;
}

function deleteQuote(sheetRow) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var quoteId = sheet.getRange(sheetRow, 1).getValue().toString().trim();
  sheet.deleteRow(sheetRow);
  if (quoteId) deletePipelineEntry(quoteId);
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
    ['Logo (Base64)',                ''],
    ['Email Subject',               'Your Catering Quote from Chick-fil-A {{location}}'],
    ['Email Body',                  'Hi {{customer}},\n\nThank you for considering Chick-fil-A {{location}} for your catering needs! We appreciate you reaching out and would love to help make your event something special.\n\nPlease find your catering quote attached. If you have any questions or would like to make any changes, don\'t hesitate to reach out — we\'re happy to help.\n\nWe look forward to serving you!\n\nWarm regards,\n{{contact}}\nChick-fil-A {{location}}\n{{phone}}'],
    ['BCC Email',                   '']
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
    var h = ['Quote ID','Created Date','Customer Name','Contact Name','Order Type','Delivery Address','Line Items (JSON)','Subtotal','Tax Rate Used','Tax Amount','Total','Tax Exempt','Location Name','Customer Email','PO Number','Event Date','Calendar Event ID','Last Modified','Event Time','Order Discount Value','Order Discount Type','Quote Notes'];
    qSheet.getRange(1, 1, 1, h.length).setValues([h]);
    qSheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    qSheet.setFrozenRows(1);
  } else {
    // Migrate: add new column headers if missing
    var lastCol = qSheet.getLastColumn();
    var newHeaders = {'Calendar Event ID': 17, 'Last Modified': 18, 'Event Time': 19, 'Order Discount Value': 20, 'Order Discount Type': 21, 'Quote Notes': 22};
    for (var label in newHeaders) {
      if (lastCol < newHeaders[label]) {
        qSheet.getRange(1, newHeaders[label]).setValue(label).setFontWeight('bold');
      }
    }
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
    orderType: quoteData.orderType || 'Pickup', deliveryAddress: quoteData.deliveryAddress || '',
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
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;"><div>'+logo+'<div style="font-size:16px;font-weight:700;">'+esc(pd.storeName)+'</div><div style="font-size:13px;color:#6B7280;margin-top:2px;">'+esc(pd.storeAddress)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.storePhone)+'</div></div><div style="text-align:right;"><div style="font-size:36px;font-weight:300;color:#D1D5DB;letter-spacing:2px;">QUOTE</div><div style="font-size:13px;color:#6B7280;margin-top:4px;font-family:Courier New,monospace;">'+esc(pd.quoteId)+'</div>'+(pd.poNumber?'<div style="font-size:13px;color:#6B7280;margin-top:2px;">PO: '+esc(pd.poNumber)+'</div>':'')+'<div style="font-size:13px;color:#6B7280;margin-top:8px;">Date: '+esc(pd.date)+'</div>'+(pd.time?'<div style="font-size:13px;color:#6B7280;">Time: '+esc(pd.time)+'</div>':'')+'<div style="font-size:14px;font-weight:700;margin-top:8px;">For: '+esc(pd.customerName)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.orderType)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(addr)+'</div><div style="font-size:13px;font-weight:600;margin-top:4px;">'+esc(pd.contactPerson)+'</div></div></div>' +
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
  addPipelineEntry({ ...quoteData, recipientEmail: recipientEmail });
  return { success: true, message: 'Email sent to ' + recipientEmail };
}

function sendQuoteEmailDirect(quoteData, recipientEmail, subject, body) {
  var settings = getSettings();
  var storeName = quoteData.locationName || settings['Store Name (Active)'] || '';
  var bcc = (settings['BCC Email']||'').toString().trim();
  var opts = { attachments: [generatePdfBlob(quoteData)], name: 'Chick-fil-A ' + storeName };
  if (bcc) opts.bcc = bcc;
  MailApp.sendEmail(recipientEmail, subject, body, opts);
  addPipelineEntry({ ...quoteData, recipientEmail: recipientEmail });
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
    + (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress ? 'Delivery Address: ' + quoteData.deliveryAddress + '\n' : '')
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
      + (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress ? 'Delivery Address: ' + quoteData.deliveryAddress + '\n' : '')
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


// ── Pipeline Data Wrapper ─────────────────────────────────────
// Called by the UI to load the pipeline view.
// Returns both entries and stats in one round-trip.

function getPipelineData() {
  return {
    entries: getPipelineEntries(),
    stats:   getPipelineStats()
  };
}
