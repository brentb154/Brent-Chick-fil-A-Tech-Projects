// ============================================================
// CHICK-FIL-A CATERING QUOTE GENERATOR — Server-Side (Code.gs)
// ============================================================
// Handles all communication between the web app and the Google
// Sheet: settings, menu items (with categories), quotes,
// server-side PDF generation, and email sending.
// ============================================================

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

const TAB_SETTINGS       = 'Settings';
const TAB_MENU           = 'Menu';
const TAB_QUOTES         = 'Quotes';
const TAB_QUOTE_SEQUENCE = 'Quote_Sequence';


// ── WEB APP ENTRY POINT ──────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
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
  var data  = sheet.getRange('A1:B30').getValues();
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
  var data  = sheet.getRange('A1:A30').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString().trim() === label) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  return false;
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
      items.push({
        category: (row[0] || '').toString().trim(),
        name: row[1].toString().trim(),
        pickupPrice: parseFloat(row[2]) || 0,
        deliveryPrice: parseFloat(row[3]) || 0,
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
  var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  var quotes = [];
  data.forEach(function(row, index) {
    if (row[0] && row[0].toString().trim() !== '') {
      quotes.push({
        quoteId: row[0], createdDate: row[1] ? new Date(row[1]).toISOString() : '',
        customerName: row[2], contactName: row[3], orderType: row[4],
        deliveryAddress: row[5], lineItems: row[6], subtotal: row[7],
        taxRateUsed: row[8], taxAmount: row[9], total: row[10],
        taxExempt: row[11], locationName: row[12], customerEmail: row[13] || '',
        sheetRow: index + 2
      });
    }
  });
  quotes.sort(function(a, b) { return new Date(b.createdDate) - new Date(a.createdDate); });
  return quotes;
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
  var quoteId = getNextQuoteId();
  sheet.appendRow([
    quoteId, new Date(), quoteData.customerName, quoteData.contactName,
    quoteData.orderType, quoteData.deliveryAddress || '',
    JSON.stringify(quoteData.lineItems),
    parseFloat(quoteData.subtotal) || 0, parseFloat(quoteData.taxRate) || 0,
    parseFloat(quoteData.taxAmount) || 0, parseFloat(quoteData.total) || 0,
    quoteData.taxExempt ? 'TRUE' : 'FALSE', quoteData.locationName || '',
    quoteData.customerEmail || ''
  ]);
  return quoteId;
}

function deleteQuote(sheetRow) {
  getSpreadsheet().getSheetByName(TAB_QUOTES).deleteRow(sheetRow);
  return true;
}

function cleanOldQuotes() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 30);
  for (var i = data.length - 1; i >= 0; i--) {
    if (new Date(data[i][1]) < cutoff) sheet.deleteRow(i + 2);
  }
}


// ── SHEET INITIALIZATION ─────────────────────────────────────

function initializeSheet() {
  var ss = getSpreadsheet();

  // Settings
  var sSheet = ss.getSheetByName(TAB_SETTINGS);
  if (!sSheet) sSheet = ss.insertSheet(TAB_SETTINGS);
  if (!sSheet.getRange('A1').getValue()) {
    sSheet.getRange(1, 1, 13, 2).setValues([
      ['Store Name (Active)',    'Cockrell Hill DTO'],
      ['Location 1 Name',       'Cockrell Hill DTO'],
      ['Location 1 Address',    '1535 N Cockrell Hill Rd., Dallas, Texas 75211'],
      ['Location 1 Phone',      '214-331-2400'],
      ['Location 2 Name',       'Dallas Baptist University OCV'],
      ['Location 2 Address',    ''],
      ['Location 2 Phone',      ''],
      ['Quote Contact Name',    ''],
      ['Default Tax Rate (%)',   8.25],
      ['Logo (Base64)',          ''],
      ['Email Subject',         'Your Catering Quote from Chick-fil-A {{location}}'],
      ['Email Body',            'Hi {{customer}},\n\nThank you for considering Chick-fil-A {{location}} for your catering needs! We appreciate you reaching out and would love to help make your event something special.\n\nPlease find your catering quote attached. If you have any questions or would like to make any changes, don\'t hesitate to reach out — we\'re happy to help.\n\nWe look forward to serving you!\n\nWarm regards,\n{{contact}}\nChick-fil-A {{location}}\n{{phone}}'],
      ['BCC Email',             '']
    ]);
    sSheet.setColumnWidth(1, 220);
    sSheet.setColumnWidth(2, 500);
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
    var h = ['Quote ID','Created Date','Customer Name','Contact Name','Order Type','Delivery Address','Line Items (JSON)','Subtotal','Tax Rate Used','Tax Amount','Total','Tax Exempt','Location Name','Customer Email'];
    qSheet.getRange(1, 1, 1, h.length).setValues([h]);
    qSheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    qSheet.setFrozenRows(1);
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
    quoteId: quoteData.quoteId || ''
  };
}


// ── SERVER-SIDE PDF ──────────────────────────────────────────

function buildPdfHtml(pd) {
  var items = pd.lineItems || [];
  if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
  var liHtml = '';
  items.forEach(function(it) {
    liHtml += '<tr><td style="text-align:center;padding:10px 12px;border-bottom:1px solid #E5E7EB;font-size:14px;">' + it.quantity + '</td><td style="padding:10px 12px;border-bottom:1px solid #E5E7EB;font-size:14px;">' + esc(it.description) + '</td><td style="text-align:right;padding:10px 12px;border-bottom:1px solid #E5E7EB;font-family:Courier New,monospace;font-size:14px;">$' + (it.price||0).toFixed(2) + '</td><td style="text-align:right;padding:10px 12px;border-bottom:1px solid #E5E7EB;font-family:Courier New,monospace;font-size:14px;font-weight:600;">$' + (it.amount||0).toFixed(2) + '</td></tr>';
  });
  var taxEx = pd.taxExempt === true || pd.taxExempt === 'TRUE';
  var taxLine = taxEx ? '<tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">TAX EXEMPT</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">$0.00</td></tr>' : '<tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">SALES TAX (' + (pd.taxRate||0) + '%)</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">$' + (pd.taxAmount||0).toFixed(2) + '</td></tr>';
  var addr = pd.orderType === 'Delivery' ? (pd.deliveryAddress||'') : (pd.storeAddress||'');
  var logo = pd.logo ? '<img src="'+pd.logo+'" style="max-width:140px;max-height:80px;margin-bottom:8px;">' : '<div style="font-size:24px;font-weight:800;color:#E51636;margin-bottom:8px;">Chick-fil-A</div>';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>@page{size:letter;margin:0.6in;}body{font-family:Helvetica,Arial,sans-serif;color:#1F2937;margin:0;padding:40px;}</style></head><body>' +
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;"><div>'+logo+'<div style="font-size:16px;font-weight:700;">'+esc(pd.storeName)+'</div><div style="font-size:13px;color:#6B7280;margin-top:2px;">'+esc(pd.storeAddress)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.storePhone)+'</div></div><div style="text-align:right;"><div style="font-size:36px;font-weight:300;color:#D1D5DB;letter-spacing:2px;">QUOTE</div><div style="font-size:13px;color:#6B7280;margin-top:4px;font-family:Courier New,monospace;">'+esc(pd.quoteId)+'</div><div style="font-size:13px;color:#6B7280;margin-top:8px;">Date: '+esc(pd.date)+'</div>'+(pd.time?'<div style="font-size:13px;color:#6B7280;">Time: '+esc(pd.time)+'</div>':'')+'<div style="font-size:14px;font-weight:700;margin-top:8px;">For: '+esc(pd.customerName)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.orderType)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(addr)+'</div><div style="font-size:13px;font-weight:600;margin-top:4px;">'+esc(pd.contactPerson)+'</div></div></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;"><thead><tr style="background:#F3F4F6;"><th style="padding:10px 12px;text-align:center;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:10%;">QTY</th><th style="padding:10px 12px;text-align:left;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:50%;">DESCRIPTION</th><th style="padding:10px 12px;text-align:right;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:18%;">PRICE/ITEM</th><th style="padding:10px 12px;text-align:right;font-size:12px;font-weight:700;color:#4B5563;text-transform:uppercase;border-bottom:2px solid #E5E7EB;width:22%;">AMOUNT</th></tr></thead><tbody>'+liHtml+'</tbody></table>' +
    '<div style="display:flex;justify-content:space-between;align-items:flex-end;"><div style="font-size:13px;color:#6B7280;max-width:320px;">If you have any questions concerning this Quote, contact:<br><strong style="color:#1F2937;">'+esc(pd.contactName)+' at '+esc(pd.storePhone)+'</strong></div><table style="width:280px;background:#F9FAFB;border-radius:8px;overflow:hidden;"><tr><td style="padding:8px 16px;font-size:14px;color:#6B7280;">SUBTOTAL</td><td style="padding:8px 16px;text-align:right;font-family:Courier New,monospace;font-size:14px;">$'+(pd.subtotal||0).toFixed(2)+'</td></tr>'+taxLine+'<tr style="background:#E5E7EB;"><td style="padding:10px 16px;font-size:15px;font-weight:700;">TOTAL</td><td style="padding:10px 16px;text-align:right;font-family:Courier New,monospace;font-size:16px;font-weight:700;">$'+(pd.total||0).toFixed(2)+'</td></tr></table></div></body></html>';
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

function getEmailQuota() { return MailApp.getRemainingDailyQuota(); }

function esc(s) { return s ? s.toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;') : ''; }
