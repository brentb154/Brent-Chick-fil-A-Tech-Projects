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
const TAB_QUOTES         = 'Quotes';
const TAB_QUOTES_ARCHIVE = 'Quotes_Archive';
const TAB_QUOTE_SEQUENCE = 'Quote_Sequence';
const TAB_QUOTE_REVISIONS = 'Quote_Revisions';
const TAB_RECIPE_BOM     = 'Recipe_BOM';
const TAB_SUPPLIES_REF   = 'Supplies_Reference';
const TAB_CASE_CONV      = 'Case_Conversions';
const TAB_ITEM_ALIASES   = 'Item_Aliases';

// The menu + cheat-sheet tabs are labeled by store so nobody edits the wrong
// restaurant's prices. The label is a short prefix the OPERATOR picks in Settings
// ("Store 1 Tab Prefix" / "Store 2 Tab Prefix"). The Catering Setup wizard sets them;
// these defaults only apply if Settings is blank — a different restaurant
// just changes those two settings. Tab names are "<prefix> Menu" and
// "<prefix> Off-Menu". A store arg of '2' = store 2; anything else = store 1.
const DEFAULT_STORE_PREFIX = { '1': 'S1', '2': 'S2' };

var _tabPrefixCache = null; // memoized per execution so Settings is read once
function storeTabPrefix_(store) {
  if (!_tabPrefixCache) {
    var s = {};
    try { s = getSettings(); } catch (e) { /* Settings tab not built yet — use defaults */ }
    _tabPrefixCache = {
      '1': (s['Store 1 Tab Prefix'] || DEFAULT_STORE_PREFIX['1']).toString().trim() || DEFAULT_STORE_PREFIX['1'],
      '2': (s['Store 2 Tab Prefix'] || DEFAULT_STORE_PREFIX['2']).toString().trim() || DEFAULT_STORE_PREFIX['2']
    };
  }
  return _tabPrefixCache[String(store) === '2' ? '2' : '1'];
}
function menuTabForStore_(store)    { return storeTabPrefix_(store) + ' Menu'; }
function offMenuTabForStore_(store) { return storeTabPrefix_(store) + ' Off-Menu'; }

// Rename a tab in place (keeps all rows/prices). Safe: only renames when the old
// tab exists and the new name is free.
function renameSheetIfExists_(ss, oldName, newName) {
  if (!oldName || oldName === newName) return;
  var oldSheet = ss.getSheetByName(oldName);
  if (oldSheet && !ss.getSheetByName(newName)) oldSheet.setName(newName);
}
// Bring each store's menu + cheat-sheet tab to the currently-configured name —
// from the original generic name (first upgrade) or from the previously-applied
// name (operator changed the prefix). Idempotent, no data loss, safe to re-run.
function migrateTabNames_(ss) {
  var props = PropertiesService.getScriptProperties();
  applyStoreTabNames_(ss, props, '1', 'Menu',   'Off_Menu');
  applyStoreTabNames_(ss, props, '2', 'Menu_2', 'Off_Menu_2');
}
function applyStoreTabNames_(ss, props, store, legacyMenu, legacyOff) {
  var wantMenu = menuTabForStore_(store), wantOff = offMenuTabForStore_(store);
  renameSheetIfExists_(ss, props.getProperty('TAB_MENU_' + store)    || legacyMenu, wantMenu);
  renameSheetIfExists_(ss, legacyMenu, wantMenu);
  renameSheetIfExists_(ss, props.getProperty('TAB_OFFMENU_' + store) || legacyOff,  wantOff);
  renameSheetIfExists_(ss, legacyOff, wantOff);
  props.setProperty('TAB_MENU_' + store, wantMenu);
  props.setProperty('TAB_OFFMENU_' + store, wantOff);
}


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
      .setTitle('Upload Your Texas Sales and Use Tax Exemption Certificate')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  if (view === 'app') {
    logAdoptionPing_('catering-quote-tool');
    return HtmlService.createTemplateFromFile('App')
      .evaluate()
      .setTitle('Catering Tool')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutput(
      '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:460px;margin:80px auto;padding:0 24px;text-align:center;color:#374151;">'
    + '<div style="font-size:40px;">🍽️</div>'
    + '<h2 style="color:#E51636;">Chick-fil-A Catering</h2>'
    + '<p style="line-height:1.6;">This link isn\'t valid on its own. If a team member sent you here to upload a Texas Sales and Use Tax Exemption Certificate, please use the exact link from your email.</p>'
    + '</div>')
    .setTitle('Chick-fil-A Catering')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── SETTINGS ─────────────────────────────────────────────────

// Memoized for the life of ONE execution — a single save path reads settings 4-5 times
// (saveQuote → createCalendarEvent → getSettings, then sendQuoteEmail → getSettings).
// updateSetting() clears it so a write is never read back stale within the same run.
var _settingsMemo = null;

function getSettings() {
  if (_settingsMemo) return _settingsMemo;
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
  _settingsMemo = settings;
  return settings;
}

// Sheets converts the literal strings TRUE/FALSE into real BOOLEANS, so an on/off setting
// comes back as either boolean true or the string 'TRUE' depending on how it was written.
// Comparing against 'TRUE' alone silently reads every enabled flag as OFF — that's why
// reminder emails stopped. Always test on/off settings through this.
function isSettingOn_(v) {
  if (v === true) return true;
  return (v || '').toString().trim().toUpperCase() === 'TRUE';
}

function updateSetting(label, value, asText) {
  _settingsMemo = null; // this run must not keep serving the pre-write copy
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

function getMenuItems(store) {
  var sheet = getSpreadsheet().getSheetByName(menuTabForStore_(store));
  if (!sheet) return []; // Store 2 tab may not exist on an un-migrated sheet
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
function getPriceCheckItems(store) {
  var pool = getMenuItems(store);
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
function submitPriceVerification(entries, store) {
  var rowsByName = {};
  getMenuItems(store).forEach(function(it) {
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
  markPricesVerified_(store);
  return { ok: true };
}

// Per-store verification date so each store's menu is checked on its own clock.
function markPricesVerified_(store) {
  var key = String(store) === '2' ? 'Last Price Verification 2' : 'Last Price Verification';
  // asText keeps Sheets from converting the yyyy-MM-dd string into a Date cell
  updateSetting(key, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), true);
}

// ── OFF-MENU CHEAT SHEET ─────────────────────────────────────
// Off_Menu tab: A=Item Name | B=Base Price. Delivery price is computed
// in the app as base × (1 + "Off-Menu Markup (%)" / 100), delivery only.

function getOffMenuItems(store) {
  var sheet = getSpreadsheet().getSheetByName(offMenuTabForStore_(store));
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

function addOffMenuItem(name, basePrice, store) {
  var sheet = getSpreadsheet().getSheetByName(offMenuTabForStore_(store));
  if (!sheet) sheet = createOffMenuSheet_(store);
  var r = sheet.getLastRow() + 1;
  sheet.getRange(r, 1, 1, 2).setValues([[name, parseFloat(basePrice) || 0]]);
  sheet.getRange(r, 2).setNumberFormat('$#,##0.00');
  return true;
}

// Fallback creator if someone deleted the tab — same shape as initializeSheet builds.
function createOffMenuSheet_(store) {
  var sheet = getSpreadsheet().insertSheet(offMenuTabForStore_(store));
  sheet.getRange(1, 1, 1, 2).setValues([['Item Name', 'Base Price']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function updateOffMenuItem(row, name, basePrice, store) {
  var sheet = getSpreadsheet().getSheetByName(offMenuTabForStore_(store));
  sheet.getRange(row, 1, 1, 2).setValues([[name, parseFloat(basePrice) || 0]]);
  return true;
}

function deleteOffMenuItem(row, store) {
  getSpreadsheet().getSheetByName(offMenuTabForStore_(store)).deleteRow(row);
  return true;
}

function addMenuItem(category, name, pickupPrice, deliveryPrice, store) {
  var sheet = getSpreadsheet().getSheetByName(menuTabForStore_(store));
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([
    [category || '', name, parseFloat(pickupPrice) || 0, parseFloat(deliveryPrice) || 0]
  ]);
  return true;
}

function updateMenuItem(row, category, name, pickupPrice, deliveryPrice, store) {
  var sheet = getSpreadsheet().getSheetByName(menuTabForStore_(store));
  sheet.getRange(row, 1, 1, 4).setValues([
    [category || '', name, parseFloat(pickupPrice) || 0, parseFloat(deliveryPrice) || 0]
  ]);
  return true;
}

function deleteMenuItem(row, store) {
  var sheet = getSpreadsheet().getSheetByName(menuTabForStore_(store));
  sheet.deleteRow(row);
  return true;
}


// ── QUOTES ───────────────────────────────────────────────────

// The Quotes tab in header order. Reads locate each field by HEADER NAME, not by a
// hardcoded position — this sheet has gained columns three times (Status, Handoff Store,
// Driver) and anyone inserting or reordering a column by hand used to silently shift every
// field after it. A header that's missing or renamed falls back to its historical position,
// so an untouched sheet behaves exactly as before.
var QUOTE_COLUMNS = ['Quote ID', 'Created Date', 'Customer Name', 'Contact Name', 'Order Type',
  'Delivery Address', 'Line Items (JSON)', 'Subtotal', 'Tax Rate Used', 'Tax Amount', 'Total',
  'Tax Exempt', 'Location Name', 'Customer Email', 'PO Number', 'Event Date', 'Calendar Event ID',
  'Last Modified', 'Event Time', 'Order Discount Value', 'Order Discount Type', 'Quote Notes',
  'Customer Phone', 'Status', 'Handoff Store', 'Driver'];

// headers (row 1 values) -> { 'Quote ID': 0, … }. Case/space tolerant; first match wins.
function detectQuoteColumns_(headers) {
  var seen = {}, map = {};
  (headers || []).forEach(function(h, i) {
    var k = (h == null ? '' : h.toString()).trim().toLowerCase();
    if (k && seen[k] === undefined) seen[k] = i;
  });
  QUOTE_COLUMNS.forEach(function(name, fallbackIdx) {
    var found = seen[name.toLowerCase()];
    map[name] = (found === undefined) ? fallbackIdx : found;
  });
  return map;
}

function getQuotes() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  // Cap the window generously so inserted columns are still seen; header lookup makes any
  // extra columns harmless. Row 1 comes along in the SAME read — no second round-trip.
  var width = Math.max(19, Math.min(40, lastCol));
  var all = sheet.getRange(1, 1, lastRow, width).getValues();
  var C = detectQuoteColumns_(all[0]);
  var data = all.slice(1);
  var tz = Session.getScriptTimeZone();
  var quotes = [];
  data.forEach(function(row, index) {
    if (row[C['Quote ID']] && row[C['Quote ID']].toString().trim() !== '') {
      quotes.push({
        quoteId: row[C['Quote ID']], createdDate: row[C['Created Date']] ? new Date(row[C['Created Date']]).toISOString() : '',
        customerName: row[C['Customer Name']], contactName: row[C['Contact Name']], orderType: row[C['Order Type']],
        deliveryAddress: row[C['Delivery Address']], lineItems: row[C['Line Items (JSON)']], subtotal: row[C['Subtotal']],
        taxRateUsed: row[C['Tax Rate Used']], taxAmount: row[C['Tax Amount']], total: row[C['Total']],
        taxExempt: row[C['Tax Exempt']], locationName: row[C['Location Name']], customerEmail: row[C['Customer Email']] || '',
        poNumber: row[C['PO Number']] ? row[C['PO Number']].toString().trim() : '',
        eventDate: normalizeEventDate_(row[C['Event Date']], tz),
        calendarEventId: row[C['Calendar Event ID']] ? row[C['Calendar Event ID']].toString().trim() : '',
        lastModified: row[C['Last Modified']] ? new Date(row[C['Last Modified']]).toISOString() : '',
        eventTime: normalizeEventTime_(row[C['Event Time']], tz),
        orderDiscountValue: row[C['Order Discount Value']] != null ? row[C['Order Discount Value']] : 0,
        orderDiscountType: row[C['Order Discount Type']] ? row[C['Order Discount Type']].toString().trim() : 'percent',
        quoteNotes: row[C['Quote Notes']] ? row[C['Quote Notes']].toString() : '',
        customerPhone: row[C['Customer Phone']] ? row[C['Customer Phone']].toString().trim() : '',
        status: (row[C['Status']] && row[C['Status']].toString().trim()) ? row[C['Status']].toString().trim() : 'Active', // blank = Active (legacy rows)
        handoffStore: row[C['Handoff Store']] ? row[C['Handoff Store']].toString().trim() : '', // blank = not handed off
        driver: row[C['Driver']] ? row[C['Driver']].toString().trim() : '',
        sheetRow: index + 2
      });
    }
  });
  // TMS bag count for the runsheet — deliveries only. BOM/aliases are read once, and only if needed.
  var tms = null;
  quotes.forEach(function(q) {
    if (q.orderType !== 'Delivery') return;
    if (!tms) tms = tmsContext_();
    q.tmsBags = tmsBagsForItems_(q.lineItems, tms).label;
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
  // Read-increment-write MUST be atomic. Two stores saving at the same moment used to
  // read the same counter and mint the same ID — and saveQuote() resolves edits by the
  // FIRST matching ID, so editing one twin would silently overwrite the other.
  var next;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getSpreadsheet().getSheetByName(TAB_QUOTE_SEQUENCE);
    var current = parseInt(sheet.getRange('B1').getValue()) || 0;
    next = current + 1;
    sheet.getRange('B1').setValue(next);
    SpreadsheetApp.flush(); // release the lock only after the counter is durably written
  } finally {
    lock.releaseLock();
  }
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
    try { appendQuoteRevision_(existing); } catch(e) { /* revision archiving is best-effort — never block a save */ }

    // Update existing calendar event, or create one if it doesn't exist yet
    if (eventId) {
      try { eventId = updateCalendarEvent(quoteData, existingQuoteId, eventId) || eventId; } catch(e) { /* calendar sync best-effort — keep the existing eventId */ }
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

// Row numbers come from the client's cached getQuotes() snapshot. If rows shifted after
// that snapshot (someone deleted a quote, or edited the sheet by hand), the index now
// points at a DIFFERENT order — so every mutation below verifies the ID before writing.
// Returns the corrected row, or 0 if that quote is gone. quoteId blank = legacy caller,
// nothing to verify against, so trust the index as before.
function resolveQuoteRow_(sheetRow, quoteId) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var id = (quoteId || '').toString().trim();
  var row = parseInt(sheetRow, 10) || 0;
  if (!id) return row;
  var last = sheet.getLastRow();
  if (row >= 2 && row <= last && sheet.getRange(row, 1).getValue().toString().trim() === id) return row;
  if (last >= 2) {
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0].toString().trim() === id) return i + 2; // it moved — use where it actually is
    }
  }
  return 0;
}

function deleteQuote(sheetRow, quoteId) {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var row = resolveQuoteRow_(sheetRow, quoteId);
  if (!row) throw new Error('That quote is no longer in the sheet — refresh and try again.');
  sheet.deleteRow(row);
  return true;
}

function updateQuotePO(sheetRow, poNumber, calendarEventId, quoteId) {
  var qSheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  sheetRow = resolveQuoteRow_(sheetRow, quoteId);
  if (!sheetRow) throw new Error('That quote is no longer in the sheet — refresh and try again.');
  // Locate PO Number by HEADER, not position. getQuotes() reads it by header, so a
  // hardcoded column 15 here meant that the moment anyone inserted or moved a column, the
  // save landed in the wrong cell and the PO read back blank — the dot never went green
  // and nothing said why.
  if (qSheet) {
    var C = detectQuoteColumns_(qSheet.getRange(1, 1, 1, Math.max(qSheet.getLastColumn(), 1)).getValues()[0]);
    qSheet.getRange(sheetRow, C['PO Number'] + 1).setValue(poNumber || '');
  }
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

// Every OTHER active quote already carrying this PO number, so the UI can warn before a
// PO is attached twice. Case- and space-insensitive: "4417-og " and "4417-OG" are the same
// PO to a human, so they're the same here too.
// NO_PO_NEEDED and blanks never match — those are statuses, not PO numbers, and every
// no-PO order in the book would otherwise look like a duplicate of every other one.
// Deactivated quotes are skipped: a cancelled order isn't a live conflict.
function findQuotesWithPo(poNumber, excludeQuoteId) {
  var want = (poNumber || '').toString().trim().toLowerCase();
  if (!want || want === 'no_po_needed') return [];
  var skip = (excludeQuoteId || '').toString().trim();
  var out = [];
  getQuotes().forEach(function(q) {
    if (q.status === 'Inactive') return;
    if (skip && q.quoteId === skip) return; // a quote is never a duplicate of itself
    var p = (q.poNumber || '').toString().trim().toLowerCase();
    if (!p || p === 'no_po_needed' || p !== want) return;
    out.push({
      quoteId: q.quoteId, customerName: q.customerName,
      eventDate: q.eventDate, eventTime: q.eventTime, total: q.total
    });
  });
  return out;
}

// Guarantee the Quotes sheet has the 24th "Status" column + header before we read/write it.
function ensureStatusColumn_() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  if (sheet.getMaxColumns() < 24) sheet.insertColumnsAfter(sheet.getMaxColumns(), 24 - sheet.getMaxColumns());
  if (sheet.getRange(1, 24).getValue() !== 'Status') sheet.getRange(1, 24).setValue('Status').setFontWeight('bold');
  return sheet;
}

// Guarantee the Quotes sheet has the 25th "Handoff Store" + 26th "Driver" columns + headers.
function ensureHandoffColumns_() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  if (sheet.getMaxColumns() < 26) sheet.insertColumnsAfter(sheet.getMaxColumns(), 26 - sheet.getMaxColumns());
  if (sheet.getRange(1, 25).getValue() !== 'Handoff Store') sheet.getRange(1, 25).setValue('Handoff Store').setFontWeight('bold');
  if (sheet.getRange(1, 26).getValue() !== 'Driver') sheet.getRange(1, 26).setValue('Driver').setFontWeight('bold');
  return sheet;
}

// Team-editable driver list from the Drivers tab (single column, header in row 1).
function getDrivers() {
  var sheet = getSpreadsheet().getSheetByName('Drivers');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return (r[0] || '').toString().trim(); })
    .filter(function(v) { return v; });
}

// Rewrite the Drivers tab from the edited list (from the Settings page). De-dups + trims.
function saveDrivers(list) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Drivers') || ss.insertSheet('Drivers');
  var seen = {}, names = [];
  (list || []).forEach(function(n) { n = (n || '').toString().trim(); if (n && !seen[n.toLowerCase()]) { seen[n.toLowerCase()] = true; names.push(n); } });
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 1).clearContent();
  sheet.getRange(1, 1).setValue('Driver Name').setFontWeight('bold');
  if (names.length) sheet.getRange(2, 1, names.length, 1).setValues(names.map(function(n) { return [n]; }));
  SpreadsheetApp.flush();
  return true;
}

// Assign (or clear) a driver on a quote: write col 26, then fold the driver name
// into the Google Calendar event title as <customer>-"<driver>". Returns a UI status.
function assignDriver(sheetRow, driver, expectedQuoteId) {
  var drv = (driver || '').toString().trim();
  var sheet = ensureHandoffColumns_();
  sheetRow = resolveQuoteRow_(sheetRow, expectedQuoteId);
  if (!sheetRow) throw new Error('That quote is no longer in the sheet — refresh and try again.');
  var row = sheet.getRange(sheetRow, 1, 1, 26).getValues()[0];
  var quoteId = (row[0] || '').toString().trim();
  var eventId = row[16] ? row[16].toString().trim() : '';
  var out = { driver: drv, calendarChanged: false, calendarReason: '' };

  sheet.getRange(sheetRow, 26).setValue(drv);

  if (eventId) {
    var tz = Session.getScriptTimeZone();
    var quoteData = {
      date: normalizeEventDate_(row[15], tz), time: normalizeEventTime_(row[18], tz),
      customerName: row[2], contactName: row[3], orderType: row[4], deliveryAddress: row[5] || '',
      lineItems: row[6], subtotal: row[7], taxRate: row[8], taxAmount: row[9], total: row[10],
      taxExempt: (row[11] === 'TRUE' || row[11] === true), locationName: row[12] || '',
      customerEmail: row[13] || '', poNumber: row[14] || '',
      orderDiscountValue: row[19], orderDiscountType: row[20], quoteNotes: row[21] || '',
      customerPhone: row[22] || '', driver: drv
    };
    try {
      var res = updateCalendarEvent(quoteData, quoteId, eventId);
      out.calendarChanged = !!res;
    } catch(e) { out.calendarReason = e.message || String(e); }
  }
  SpreadsheetApp.flush();
  return out;
}

// ── HAND-OFF TO PARTNER STORE ────────────────────────────────

// Mark (or clear) which store an order was handed off to. Blank = not handed off.
function setHandoffStore(sheetRow, storeName, expectedQuoteId) {
  var sheet = ensureHandoffColumns_();
  sheetRow = resolveQuoteRow_(sheetRow, expectedQuoteId);
  if (!sheetRow) throw new Error('That quote is no longer in the sheet — refresh and try again.');
  sheet.getRange(sheetRow, 25).setValue((storeName || '').toString().trim());
  SpreadsheetApp.flush();
  return { handoffStore: (storeName || '').toString().trim() };
}

// Partner store contacts, one object per row of the Partner Store tab.
function getPartnerStores() {
  var sheet = getSpreadsheet().getSheetByName('Partner Store');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var out = [];
  data.forEach(function(r) {
    var loc = (r[0] || '').toString().trim();
    if (!loc) return;
    out.push({ storeLocation: loc, operatorName: (r[1]||'').toString().trim(), storePhone: (r[2]||'').toString().trim(), storeEmail: (r[3]||'').toString().trim(), cateringPOC: (r[4]||'').toString().trim() });
  });
  return out;
}

// Add or update a partner store, keyed on Store Location. Returns the saved list.
function savePartnerStore(store) {
  var sheet = getSpreadsheet().getSheetByName('Partner Store');
  if (!sheet) throw new Error('Partner Store tab missing — run initializeSheet.');
  var loc = (store.storeLocation || '').toString().trim();
  if (!loc) throw new Error('Store Location is required.');
  var row = [loc, (store.operatorName||'').toString().trim(), (store.storePhone||'').toString().trim(), (store.storeEmail||'').toString().trim(), (store.cateringPOC||'').toString().trim()];
  var found = 0;
  if (sheet.getLastRow() > 1) {
    var locs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < locs.length; i++) {
      if ((locs[i][0] || '').toString().trim().toLowerCase() === loc.toLowerCase()) { found = i + 2; break; }
    }
  }
  if (found) sheet.getRange(found, 1, 1, 5).setValues([row]);
  else sheet.appendRow(row);
  SpreadsheetApp.flush();
  return getPartnerStores();
}

function deletePartnerStore(storeLocation) {
  var sheet = getSpreadsheet().getSheetByName('Partner Store');
  if (!sheet || sheet.getLastRow() < 2) return getPartnerStores();
  var loc = (storeLocation || '').toString().trim().toLowerCase();
  var locs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < locs.length; i++) {
    if ((locs[i][0] || '').toString().trim().toLowerCase() === loc) { sheet.deleteRow(i + 2); break; }
  }
  SpreadsheetApp.flush();
  return getPartnerStores();
}

// Build an EDITABLE hand-off email draft (never auto-sent). Fills the Settings
// template with the quote + chosen partner-store values. Returns {to, subject, body}.
function buildHandoffEmailDraft(quoteId) {
  var quotes = getQuotes();
  var q = null;
  for (var i = 0; i < quotes.length; i++) { if (quotes[i].quoteId === quoteId) { q = quotes[i]; break; } }
  if (!q) throw new Error('Quote not found: ' + quoteId);
  var storeName = (q.handoffStore || '').toString().trim();
  if (!storeName) throw new Error('This order is not marked as handed off to a store yet.');

  var partner = null;
  getPartnerStores().forEach(function(p) { if (p.storeLocation.toLowerCase() === storeName.toLowerCase()) partner = p; });

  var settings = getSettings();
  var fromStore = (q.locationName || settings['Store Name (Active)'] || '').toString();
  var items = [];
  try { items = JSON.parse(q.lineItems) || []; } catch(e) { items = []; }
  var itemLines = items.map(function(it) { return '  ' + (it.quantity||'') + 'x ' + (it.description||''); }).join('\n');

  var reps = {
    '{{customer}}': q.customerName || '', '{{contact}}': q.contactName || '',
    '{{poc}}': (partner && partner.cateringPOC) || (partner && partner.operatorName) || 'there',
    '{{partnerStore}}': storeName, '{{date}}': q.eventDate || '', '{{time}}': q.eventTime || '',
    '{{orderType}}': q.orderType || '', '{{address}}': q.deliveryAddress || '',
    '{{total}}': '$' + (parseFloat(q.total) || 0).toFixed(2), '{{items}}': itemLines || '(no items)',
    '{{fromContact}}': settings['Quote Contact Name'] || '', '{{fromStore}}': fromStore
  };
  var subj = settings['Handoff Email Subject'] || 'Catering hand-off: {{customer}} on {{date}}';
  var body = settings['Handoff Email Body'] || 'Hi {{poc}},\n\nWe\'d like to hand off {{customer}} on {{date}} to {{partnerStore}}.\n\n{{items}}\n\nThanks!\n{{fromContact}}';
  for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
  return { to: (partner && partner.storeEmail) || '', subject: subj, body: body };
}

// Send the (human-reviewed) hand-off email — plain email, no quote PDF attached.
function sendHandoffEmail(to, subject, body) {
  var recipient = (to || '').toString().trim();
  if (!recipient) throw new Error('No recipient email.');
  var settings = getSettings();
  var fromName = 'Chick-fil-A ' + (settings['Store Name (Active)'] || '');
  MailApp.sendEmail(recipient, subject, body, { name: fromName });
  return { success: true };
}

// ── "WE GOT YOUR PO" CONFIRMATION ────────────────────────────
// Offered when a PO is first attached to an order: tells the guest (a teacher, usually)
// that the PO landed and we're executing at their requested time. Always human-reviewed
// in the preview modal before it sends — never fired automatically by a trigger.
const TAB_PO_CONFIRMATIONS = 'PO_Confirmations_Sent';
const PO_CONFIRMATION_HEADERS = ['Quote ID', 'PO Number', 'Sent At', 'Sent To', 'Sent By'];

// Store phone for whichever location the quote belongs to (falls back to Location 1).
function storePhoneForName_(settings, storeName) {
  if (storeName && storeName === (settings['Location 2 Name'] || '')) return settings['Location 2 Phone'] || '';
  return settings['Location 1 Phone'] || '';
}

// Most recent send for this quote, or '' — drives the "already sent on…" warning.
function lastPoConfirmationSent_(quoteId) {
  var id = (quoteId || '').toString().trim();
  if (!id) return '';
  var sheet = getSpreadsheet().getSheetByName(TAB_PO_CONFIRMATIONS);
  if (!sheet || sheet.getLastRow() < 2) return '';
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PO_CONFIRMATION_HEADERS.length).getValues();
  var tz = Session.getScriptTimeZone(), out = '';
  rows.forEach(function(r) {
    if ((r[0] || '').toString().trim() !== id || !r[2]) return;
    out = Utilities.formatDate(new Date(r[2]), tz, 'MMM d, yyyy') + (r[3] ? ' to ' + r[3] : '');
  });
  return out; // last matching row wins — the log is append-only, so that's the newest
}

// Draft only — nothing is sent here. Returns { to, subject, body, alreadySentAt, quoteId }.
function buildPoConfirmationDraft(quoteId) {
  var quotes = getQuotes(), q = null;
  for (var i = 0; i < quotes.length; i++) { if (quotes[i].quoteId === quoteId) { q = quotes[i]; break; } }
  if (!q) throw new Error('Quote not found: ' + quoteId);
  var po = (q.poNumber || '').toString().trim();
  if (!po || po === 'NO_PO_NEEDED') throw new Error('This order has no PO number to confirm.');

  var settings = getSettings();
  var tz = Session.getScriptTimeZone();
  var storeName = (q.locationName || settings['Store Name (Active)'] || '').toString();
  var items = [];
  try { items = JSON.parse(q.lineItems) || []; } catch(e) { items = []; }
  var itemLines = items.map(function(it) { return '  ' + (it.quantity || '') + 'x ' + (it.description || ''); }).join('\n');
  var when = formatWhen_(q.eventDate, q.eventTime, tz);
  // One line that reads right for either order type, so the template doesn't need an if.
  var orderLine = q.orderType === 'Delivery'
    ? ('We\'ll deliver to ' + (q.deliveryAddress || 'your address') + '.')
    : 'Your order will be ready for pickup at our restaurant.';

  var reps = {
    '{{contactPerson}}': q.contactName || '', '{{contact}}': settings['Quote Contact Name'] || '',
    '{{customer}}': q.customerName || '', '{{poNumber}}': po, '{{quoteId}}': q.quoteId || '',
    '{{when}}': when || 'your requested time', '{{date}}': formatWhen_(q.eventDate, '', tz),
    '{{time}}': (when.split(' at ')[1] || ''), '{{orderType}}': q.orderType || '',
    '{{address}}': q.deliveryAddress || '', '{{orderLine}}': orderLine,
    '{{total}}': '$' + (parseFloat(q.total) || 0).toFixed(2), '{{items}}': itemLines || '(no items listed)',
    '{{location}}': storeName, '{{phone}}': storePhoneForName_(settings, storeName)
  };
  var subj = settings['PO Confirmation Subject'] || 'We received your PO — your catering order {{quoteId}} is confirmed';
  var body = settings['PO Confirmation Body'] || 'Hi {{contactPerson}},\n\nWe\'ve received PO {{poNumber}} and your catering order is confirmed for {{when}}.\n\n{{contact}}\nChick-fil-A {{location}}';
  for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
  return {
    to: (q.customerEmail || '').toString().trim(), subject: subj, body: body,
    quoteId: q.quoteId, poNumber: po, alreadySentAt: lastPoConfirmationSent_(q.quoteId)
  };
}

// Send the (human-reviewed) PO confirmation and log it. Plain email — no quote PDF:
// they already have the quote, this is just "we got the PO".
function sendPoConfirmationEmail(quoteId, to, subject, body) {
  var recipient = (to || '').toString().trim();
  if (!recipient) throw new Error('No recipient email.');
  var settings = getSettings();
  var quotes = getQuotes(), q = null;
  for (var i = 0; i < quotes.length; i++) { if (quotes[i].quoteId === quoteId) { q = quotes[i]; break; } }
  var storeName = ((q && q.locationName) || settings['Store Name (Active)'] || '').toString();
  var opts = { name: 'Chick-fil-A ' + storeName };
  var bcc = (settings['BCC Email'] || '').toString().trim();
  if (bcc) opts.bcc = bcc;
  MailApp.sendEmail(recipient, subject, body, opts);
  var who = ''; try { who = Session.getEffectiveUser().getEmail(); } catch(e) { /* not resolvable — leave blank */ }
  initLogSheet_(TAB_PO_CONFIRMATIONS, PO_CONFIRMATION_HEADERS)
    .appendRow([quoteId || '', (q && q.poNumber) || '', new Date(), recipient, who]);
  SpreadsheetApp.flush();
  return { success: true };
}

// ── "WE NEED YOUR PO" REQUESTS ───────────────────────────────
// Customer-facing chase for orders with no PO. Manual only — a person reviews the
// list and clicks Send. One email per customer address, listing all their orders,
// so a contact with 3 orders gets one email (never a group blast — addresses stay private).
const TAB_PO_REQUESTS = 'PO_Requests_Sent';
const PO_REQUEST_HEADERS = ['Quote ID', 'Event Date', 'Sent At', 'Sent To', 'Sent By'];

// quoteId → newest send Date from the log
function poRequestLastSent_() {
  var out = {};
  var sheet = getSpreadsheet().getSheetByName(TAB_PO_REQUESTS);
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function(r) {
    var id = (r[0] || '').toString().trim();
    if (id && r[2]) out[id] = new Date(r[2]); // append-only log, so the last row wins
  });
  return out;
}

// Orders in [startKey, endKey] (never before today) with no PO, grouped by customer email.
// Returns { start, end, groups: [...], noEmail: [...] }. Each group carries its rendered email.
function getPoRequestList(startKey, endKey) {
  var settings = getSettings();
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var start = (startKey || todayKey) < todayKey ? todayKey : (startKey || todayKey);
  var end = endKey || start;
  var lastSent = poRequestLastSent_();

  var missing = getQuotes().filter(function(q) {
    if (q.status === 'Inactive') return false;
    if (!q.eventDate || q.eventDate < start || q.eventDate > end) return false;
    return (q.poNumber || '').toString().trim() === ''; // HAVE PO and NO_PO_NEEDED both count as handled
  });
  missing.sort(function(a, b) { return (a.eventDate + (a.eventTime || '')).localeCompare(b.eventDate + (b.eventTime || '')); });

  var groups = [], byEmail = {}, noEmail = [];
  missing.forEach(function(q) {
    var order = {
      quoteId: q.quoteId, eventDate: q.eventDate, customerName: q.customerName || '', contactName: q.contactName || '',
      phone: q.customerPhone || '', when: formatWhen_(q.eventDate, q.eventTime, tz) || q.eventDate,
      total: '$' + (parseFloat(q.total) || 0).toFixed(2),
      lastAsked: lastSent[q.quoteId] ? Utilities.formatDate(lastSent[q.quoteId], tz, 'MMM d') : ''
    };
    var email = (q.customerEmail || '').toString().trim();
    if (!email) { noEmail.push(order); return; }
    var k = email.toLowerCase();
    if (!byEmail[k]) {
      byEmail[k] = { email: email, contactName: q.contactName || '', customerName: q.customerName || '',
                     storeName: (q.locationName || settings['Store Name (Active)'] || '').toString(), orders: [] };
      groups.push(byEmail[k]);
    }
    byEmail[k].orders.push(order);
  });

  groups.forEach(function(g) {
    var mail = buildPoRequestEmail_(g, settings);
    g.subject = mail.subject;
    g.body = mail.body;
    g.lastAsked = g.orders.map(function(o) { return o.lastAsked; }).filter(function(x) { return x; }).pop() || '';
  });
  return { start: start, end: end, groups: groups, noEmail: noEmail };
}

function buildPoRequestEmail_(g, settings) {
  var orderLines = g.orders.map(function(o) {
    return '  • ' + o.when + ' — ' + o.customerName + ' (quote ' + o.quoteId + ', ' + o.total + ')';
  }).join('\n');
  var reps = {
    '{{contactPerson}}': g.contactName || 'there', '{{customer}}': g.customerName,
    '{{orders}}': orderLines, '{{orderCount}}': String(g.orders.length),
    '{{contact}}': settings['Quote Contact Name'] || '', '{{location}}': g.storeName,
    '{{phone}}': storePhoneForName_(settings, g.storeName)
  };
  var subj = settings['PO Request Subject'] || 'Reminder: we need a PO for your upcoming catering order';
  var body = settings['PO Request Body'] || 'Hi {{contactPerson}},\n\nJust a friendly reminder — we don\'t have a PO on file yet for your upcoming catering order:\n\n{{orders}}\n\nWhen you get a chance, please reply to this email with the PO number so we can lock everything in. If you have any questions, call us at {{phone}}.\n\nThank you!\n{{contact}}\nChick-fil-A {{location}}';
  for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
  return { subject: subj, body: body };
}

// Sends to the selected addresses only. Rebuilds the list server-side so the emails
// always reflect the sheet at send time (a PO added a minute ago drops out).
function sendPoRequests(startKey, endKey, emails) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another send is in progress — try again in a moment.');
  try {
    var pick = {};
    (emails || []).forEach(function(e) { pick[(e || '').toString().trim().toLowerCase()] = true; });
    var list = getPoRequestList(startKey, endKey);
    var groups = list.groups.filter(function(g) { return pick[g.email.toLowerCase()]; });
    if (!groups.length) return { success: false, message: 'Nothing to send — those orders may already have a PO.' };
    var quota = MailApp.getRemainingDailyQuota();
    if (groups.length > quota) return { success: false, message: 'Only ' + quota + ' emails left in today\'s Google limit — uncheck some and try again.' };

    var settings = getSettings();
    var bcc = (settings['BCC Email'] || '').toString().trim();
    var who = ''; try { who = Session.getEffectiveUser().getEmail(); } catch(e) { /* not resolvable — leave blank */ }
    var log = initLogSheet_(TAB_PO_REQUESTS, PO_REQUEST_HEADERS);
    var now = new Date(), rows = [], sent = 0, failed = [];
    groups.forEach(function(g) {
      try {
        var opts = { name: 'Chick-fil-A ' + g.storeName };
        if (bcc) opts.bcc = bcc;
        MailApp.sendEmail(g.email, g.subject, g.body, opts);
        sent++;
        g.orders.forEach(function(o) { rows.push([o.quoteId, o.eventDate, now, g.email, who]); });
      } catch(e) { failed.push(g.email); }
    });
    if (rows.length) log.getRange(log.getLastRow() + 1, 1, rows.length, PO_REQUEST_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    var msg = 'Sent ' + sent + ' PO request' + (sent === 1 ? '' : 's') + '.';
    if (failed.length) msg += ' Failed: ' + failed.join(', ');
    return { success: failed.length === 0, message: msg };
  } finally {
    lock.releaseLock();
  }
}

// Non-blocking distance check for the hand-off suggestion. Returns
// { miles, threshold, suggest } — suggest true when the delivery is far from the start point.
function handoffDistanceCheck(deliveryAddress) {
  var addr = (deliveryAddress || '').toString().trim();
  if (!addr) return { miles: null, suggest: false };
  var settings = getSettings();
  var threshold = parseFloat(settings['Handoff Distance Threshold (Miles)']); if (isNaN(threshold) || threshold <= 0) threshold = 20;
  var geoCache = loadGeocodeCache_();
  var start = geocodeAddress_(settings['Start Address'] || '', geoCache);
  var dest = geocodeAddress_(addr, geoCache);
  if (!start || !dest) return { miles: null, threshold: threshold, suggest: false };
  var miles = Math.round(haversineMiles_(start, dest) * 10) / 10;
  return { miles: miles, threshold: threshold, suggest: miles > threshold };
}

// Resolve an event by ID across the configured then default calendar, and delete it.
// Returns true on success, or a string reason on failure. Mirrors updateCalendarEventPO's lookup.
function deleteCalendarEventById_(calendarEventId) {
  if (!calendarEventId) return 'Empty event ID';
  var settings = getSettings();
  var calId = (settings['Calendar ID'] || '').toString().trim();
  var event = null;
  function tryCal(c) { if (!c) return; try { var e = c.getEventById(calendarEventId); if (e) event = e; } catch(_) { /* may live elsewhere — keep looking */ } }
  if (calId) tryCal(CalendarApp.getCalendarById(calId));
  if (!event) tryCal(CalendarApp.getDefaultCalendar());
  if (!event) return 'Event not found on calendar';
  try { event.deleteEvent(); return true; } catch(e) { return e.message || String(e); }
}

// Deactivate = hide from the active calendar view AND pull its Google Calendar event.
// Reactivate = re-create the calendar event. The quote row itself is always kept in history.
// status is 'Active' or 'Inactive'. Returns a UI status object like updateQuotePO.
function setQuoteStatus(sheetRow, status, expectedQuoteId) {
  var newStatus = (status === 'Inactive') ? 'Inactive' : 'Active';
  var sheet = ensureStatusColumn_();
  sheetRow = resolveQuoteRow_(sheetRow, expectedQuoteId);
  if (!sheetRow) throw new Error('That quote is no longer in the sheet — refresh and try again.');
  var row = sheet.getRange(sheetRow, 1, 1, 24).getValues()[0];
  var quoteId = (row[0] || '').toString().trim();
  var eventId = row[16] ? row[16].toString().trim() : '';
  var out = { status: newStatus, calendarChanged: false, calendarReason: '' };

  sheet.getRange(sheetRow, 24).setValue(newStatus);

  if (newStatus === 'Inactive') {
    if (eventId) {
      var del = deleteCalendarEventById_(eventId);
      if (del === true) { sheet.getRange(sheetRow, 17).setValue(''); out.calendarChanged = true; }
      else { out.calendarReason = (typeof del === 'string') ? del : 'Calendar event not removed'; }
    }
  } else {
    // Reactivate: rebuild a quoteData object from the row and re-create the event
    var tz = Session.getScriptTimeZone();
    if (!eventId) {
      var quoteData = {
        date: normalizeEventDate_(row[15], tz), time: normalizeEventTime_(row[18], tz),
        customerName: row[2], contactName: row[3], orderType: row[4], deliveryAddress: row[5] || '',
        lineItems: row[6], subtotal: row[7], taxRate: row[8], taxAmount: row[9], total: row[10],
        taxExempt: (row[11] === 'TRUE' || row[11] === true), locationName: row[12] || '',
        customerEmail: row[13] || '', poNumber: row[14] || '',
        orderDiscountValue: row[19], orderDiscountType: row[20], quoteNotes: row[21] || '',
        customerPhone: row[22] || ''
      };
      if (quoteData.date && quoteData.time) {
        try {
          var newId = createCalendarEvent(quoteData, quoteId) || '';
          if (newId) { sheet.getRange(sheetRow, 17).setValue(newId); out.calendarChanged = true; }
        } catch(e) { out.calendarReason = e.message || String(e); }
      }
    }
  }
  SpreadsheetApp.flush();
  return out;
}

// Total every menu item needed across all of one day's active orders.
// Returns { orderCount, items: [{item, qty}] } sorted by quantity desc, for the kitchen prep view.
// allQuotes is an optional pre-read snapshot — the range version passes one in so a
// 31-day span doesn't re-read the whole Quotes tab 31 times.
function getDayItemTotals(dateStr, allQuotes) {
  var quotes = (allQuotes || getQuotes()).filter(function(q) { return q.eventDate === dateStr && q.status !== 'Inactive'; });
  var totals = {}; // normalized name -> {item (display), qty}
  quotes.forEach(function(q) {
    var items = [];
    try { items = JSON.parse(q.lineItems) || []; } catch(e) { items = []; }
    items.forEach(function(it) {
      var name = (it.description || '').toString().trim();
      if (!name) return;
      var qty = parseFloat(it.quantity) || 0;
      if (qty <= 0) return;
      var key = name.toLowerCase();
      if (!totals[key]) totals[key] = { item: name, qty: 0 };
      totals[key].qty += qty;
    });
  });
  var list = Object.keys(totals).map(function(k) { return totals[k]; });
  list.sort(function(a, b) { return b.qty - a.qty || a.item.localeCompare(b.item); });
  return { orderCount: quotes.length, items: list };
}

// Inclusive list of yyyy-MM-dd date strings from start to end (capped, safe).
function eachDate_(startStr, endStr) {
  var tz = Session.getScriptTimeZone();
  var s = new Date((startStr || '') + 'T12:00:00'), e = new Date((endStr || startStr || '') + 'T12:00:00');
  if (isNaN(s.getTime())) return [];
  if (isNaN(e.getTime()) || e < s) e = s;
  var out = [], cur = new Date(s), guard = 0;
  while (cur <= e && guard < 400) { out.push(Utilities.formatDate(cur, tz, 'yyyy-MM-dd')); cur.setDate(cur.getDate() + 1); guard++; }
  return out;
}

// Item totals over a date range: per-day sections + a grand total.
function getDayItemTotalsRange(startStr, endStr) {
  var dates = eachDate_(startStr, endStr);
  var allQuotes = getQuotes(); // read once, filter per day in memory
  var days = [], grand = {}, grandOrders = 0;
  dates.forEach(function(d) {
    var r = getDayItemTotals(d, allQuotes);
    days.push({ date: d, orderCount: r.orderCount, items: r.items });
    grandOrders += r.orderCount;
    r.items.forEach(function(it) { var k = it.item.toLowerCase(); if (!grand[k]) grand[k] = { item: it.item, qty: 0 }; grand[k].qty += it.qty; });
  });
  var total = Object.keys(grand).map(function(k) { return grand[k]; }).sort(function(a, b) { return b.qty - a.qty || a.item.localeCompare(b.item); });
  return { start: dates[0] || startStr, end: dates[dates.length - 1] || startStr, days: days, total: { orderCount: grandOrders, items: total } };
}

// Supplies over a date range. getSuppliesForDate is multi-store, so this keeps the
// per-store shape: days = [full getSuppliesForDate results]; total.stores aggregates
// each store's ingredients across the range (case conversions re-applied to the grand).
function getSuppliesForRange(startStr, endStr) {
  var dates = eachDate_(startStr, endStr);
  var ctx = { quotes: getQuotes(), bom: loadRecipeBom_(), conv: loadCaseConversions_(), aliasMap: getItemAliasMap_() };
  var days = [], byStore = {}, byStoreTag = {}, byStoreNo = {}, storeOrder = [], gOrders = 0;
  dates.forEach(function(d) {
    var r = getSuppliesForDate(d, ctx);
    days.push(r);
    gOrders += r.orderCount;
    (r.stores || []).forEach(function(st) {
      if (!byStore[st.store]) { byStore[st.store] = {}; byStoreTag[st.store] = {}; byStoreNo[st.store] = {}; storeOrder.push(st.store); }
      st.ingredients.forEach(function(ig) { byStore[st.store][ig.component] = (byStore[st.store][ig.component] || 0) + ig.qty; if (ig.tag) byStoreTag[st.store][ig.component] = ig.tag; });
      st.noRecipe.forEach(function(nr) { var k = nr.item.toLowerCase(); if (!byStoreNo[st.store][k]) byStoreNo[st.store][k] = { item: nr.item, qty: 0 }; byStoreNo[st.store][k].qty += nr.qty; });
    });
  });
  var conv = ctx.conv;
  var totalStores = storeOrder.map(function(store) {
    var ings = Object.keys(byStore[store]).map(function(comp) {
      var qty = byStore[store][comp], c = conv[comp.toLowerCase()] || { perCase: 0, perBag: 0 };
      return { component: comp, qty: qty, tag: byStoreTag[store][comp] || '', perCase: c.perCase, perBag: c.perBag,
        cases: c.perCase > 0 ? Math.round((qty / c.perCase) * 100) / 100 : null,
        bags:  c.perBag  > 0 ? Math.round((qty / c.perBag)  * 10)  / 10  : null };
    }).sort(function(a, b) { return b.qty - a.qty || a.component.localeCompare(b.component); });
    var no = Object.keys(byStoreNo[store]).map(function(k) { return byStoreNo[store][k]; }).sort(function(a, b) { return b.qty - a.qty; });
    return { store: store, ingredients: ings, noRecipe: no };
  });
  totalStores.sort(function(a, b) { return a.store.localeCompare(b.store); });
  return { start: dates[0] || startStr, end: dates[dates.length - 1] || startStr, days: days, total: { orderCount: gOrders, stores: totalStores } };
}

// ── DELIVERY ROUTE SUGGESTIONS ───────────────────────────────
// For one day, suggest pairs of deliveries that are close together AND due
// around the same time, so a human can decide to run them on one drive.
// Straight-line distance (crow-flies) — no traffic, no roads. A hint, not a route.
const TAB_GEOCODE = 'Geocode_Cache';

// Geocode an address, caching the result in a hidden tab so repeat addresses
// don't burn the Maps quota. Returns {lat, lng} or null if it can't be resolved.
// Read the whole cache tab once into { lowercased address -> {lat, lng} }.
function loadGeocodeCache_() {
  var sheet = initLogSheet_(TAB_GEOCODE, ['Address', 'Lat', 'Lng', 'Cached At']);
  var map = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function(r) {
      var k = (r[0] || '').toString().trim().toLowerCase();
      var lat = parseFloat(r[1]), lng = parseFloat(r[2]);
      if (k && !isNaN(lat) && !isNaN(lng)) map[k] = { lat: lat, lng: lng };
    });
  }
  return { sheet: sheet, map: map };
}

// cache is an optional loadGeocodeCache_() result — callers that geocode in a loop pass
// one in so the cache tab is read once instead of once per address.
function geocodeAddress_(address, cache) {
  var addr = (address || '').toString().trim();
  if (!addr) return null;
  cache = cache || loadGeocodeCache_();
  var key = addr.toLowerCase();
  if (cache.map[key]) return cache.map[key];
  try {
    var res = Maps.newGeocoder().geocode(addr);
    if (res && res.status === 'OK' && res.results && res.results.length) {
      var loc = res.results[0].geometry.location;
      cache.sheet.appendRow([addr, loc.lat, loc.lng, new Date()]); // only cache successes — a failed lookup retries next run
      cache.map[key] = { lat: loc.lat, lng: loc.lng }; // and don't re-geocode it within this run
      return cache.map[key];
    }
  } catch(e) { /* quota or transient error — leave uncached so it retries later */ }
  return null;
}

// Straight-line miles between two lat/lng points (haversine).
function haversineMiles_(a, b) {
  var R = 3958.8; // earth radius, miles
  var toRad = function(d) { return d * Math.PI / 180; };
  var dLat = toRad(b.lat - a.lat), dLng = toRad(b.lng - a.lng);
  var lat1 = toRad(a.lat), lat2 = toRad(b.lat);
  var h = Math.sin(dLat/2)*Math.sin(dLat/2) + Math.sin(dLng/2)*Math.sin(dLng/2)*Math.cos(lat1)*Math.cos(lat2);
  return R * 2 * Math.atan2(Math.sqrt(h), Math.sqrt(1-h));
}

function eventTimeToMinutes_(t) {
  var m = /^(\d{1,2}):(\d{2})/.exec((t || '').toString());
  return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : null;
}

// Suggest deliveries on dateStr that could share a drive.
// Returns { radius, windowMin, pairs:[{a,b,miles,minutesApart}], deliveryCount, ungeocoded:[address...] }.
function suggestDeliveryRoutes(dateStr) {
  var settings = getSettings();
  var radius = parseFloat(settings['Route Suggest Radius (Miles)']); if (isNaN(radius) || radius <= 0) radius = 5;
  var windowMin = parseInt(settings['Route Suggest Window (Minutes)'], 10); if (isNaN(windowMin) || windowMin < 0) windowMin = 90;

  var deliveries = getQuotes().filter(function(q) {
    return q.eventDate === dateStr && q.status !== 'Inactive'
      && q.orderType === 'Delivery' && (q.deliveryAddress || '').toString().trim();
  });

  var pts = [], ungeocoded = [], geoCache = loadGeocodeCache_();
  deliveries.forEach(function(q) {
    var geo = geocodeAddress_(q.deliveryAddress, geoCache);
    if (!geo) { ungeocoded.push(q.customerName + ' — ' + q.deliveryAddress); return; }
    pts.push({
      quoteId: q.quoteId, customerName: q.customerName, address: q.deliveryAddress,
      eventTime: q.eventTime, mins: eventTimeToMinutes_(q.eventTime), geo: geo
    });
  });

  var pairs = [];
  for (var i = 0; i < pts.length; i++) {
    for (var j = i + 1; j < pts.length; j++) {
      var miles = haversineMiles_(pts[i].geo, pts[j].geo);
      if (miles > radius) continue;
      var minutesApart = (pts[i].mins != null && pts[j].mins != null) ? Math.abs(pts[i].mins - pts[j].mins) : null;
      if (minutesApart != null && minutesApart > windowMin) continue; // too far apart in time to combine
      pairs.push({
        a: { quoteId: pts[i].quoteId, customerName: pts[i].customerName, address: pts[i].address, eventTime: pts[i].eventTime },
        b: { quoteId: pts[j].quoteId, customerName: pts[j].customerName, address: pts[j].address, eventTime: pts[j].eventTime },
        miles: Math.round(miles * 10) / 10,
        minutesApart: minutesApart
      });
    }
  }
  pairs.sort(function(x, y) { return x.miles - y.miles; });
  return { radius: radius, windowMin: windowMin, pairs: pairs, deliveryCount: deliveries.length, ungeocoded: ungeocoded };
}

// ── SUPPLIES CALC ────────────────────────────────────────────
// Explode a day's quoted menu items into raw ingredient/supply totals using
// the team-editable Recipe_BOM sheet. Anything without a BOM row is surfaced,
// never guessed. Reference/contacts live in the Supplies_Reference sheet.

// Read Case_Conversions into { lowercased component -> {perCase, perBag} }.
function loadCaseConversions_() {
  var sheet = getSpreadsheet().getSheetByName(TAB_CASE_CONV);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  data.forEach(function(r) {
    var comp = (r[0] || '').toString().trim();
    if (!comp) return;
    map[comp.toLowerCase()] = { perCase: parseFloat(r[1]) || 0, perBag: parseFloat(r[2]) || 0 };
  });
  return map;
}

// Read Recipe_BOM into { lowercased menu item -> [{component, qtyPer}] }.
function loadRecipeBom_() {
  var sheet = getSpreadsheet().getSheetByName(TAB_RECIPE_BOM);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  data.forEach(function(r) {
    var name = (r[0] || '').toString().trim();
    var comp = (r[1] || '').toString().trim();
    var qty = parseFloat(r[2]) || 0;
    if (!name || !comp || qty <= 0) return;
    var key = name.toLowerCase();
    if (!map[key]) map[key] = [];
    map[key].push({ component: comp, qtyPer: qty });
  });
  return map;
}

// Component -> prep tag, for row highlighting on the supplies view.
// protein = raw protein to pull/thaw (red); thaw = mac bags to thaw day-before (yellow).
// Keys must match the Case_Conversions / BOM component names (lowercased).
function supplyTagMap_() {
  var map = {};
  ['filet', 'spicy filet', 'breakfast filet', 'spicy breakfast filet', 'nuggets', 'strips', 'sausage patty']
    .forEach(function(p) { map[p] = 'protein'; });
  map['mac & cheese bag'] = 'thaw';
  return map;
}

// Explode a list of {item, qty} against the BOM into sorted ingredient rows +
// a noRecipe list. Shared by the per-store buckets in getSuppliesForDate.
function explodeItemsToIngredients_(items, bom, conv, tags, aliasMap) {
  aliasMap = aliasMap || {};
  var totals = {};   // component -> qty
  var noRecipe = [];
  items.forEach(function(di) {
    // Translate the item name through the alias map before matching the BOM
    // (retroactive — an alias added later fixes all past orders). Report unmatched under the ORIGINAL name.
    var canon = aliasMap[di.item.toLowerCase().trim()] || di.item;
    var comps = bom[canon.toLowerCase().trim()];
    if (!comps || !comps.length) { noRecipe.push({ item: di.item, qty: di.qty }); return; }
    comps.forEach(function(c) {
      if (!totals[c.component]) totals[c.component] = 0;
      totals[c.component] += c.qtyPer * di.qty;
    });
  });
  var ingredients = Object.keys(totals).map(function(k) {
    var qty = totals[k];
    var c = conv[k.toLowerCase()] || { perCase: 0, perBag: 0 };
    return {
      component: k, qty: qty,
      perCase: c.perCase, perBag: c.perBag,
      cases: c.perCase > 0 ? Math.round((qty / c.perCase) * 100) / 100 : null,
      bags:  c.perBag  > 0 ? Math.round((qty / c.perBag)  * 10)  / 10  : null,
      tag: tags[k.toLowerCase()] || ''
    };
  });
  ingredients.sort(function(a, b) { return b.qty - a.qty || a.component.localeCompare(b.component); });
  return { ingredients: ingredients, noRecipe: noRecipe };
}

// Day's supplies, split per store (by locationName). Orders handed off to another
// store are pulled OUT of the totals and listed separately (visibility only).
// Returns { date, orderCount, stores:[{store, ingredients, noRecipe}], handedOff:[{guest, store, from, items}] }.
// ctx is an optional pre-read bundle {quotes, bom, conv, aliasMap} — the range version
// builds it once so a 31-day span doesn't re-read four tabs per day (124 reads → 4).
function getSuppliesForDate(dateStr, ctx) {
  ctx = ctx || {};
  var quotes = (ctx.quotes || getQuotes()).filter(function(q) { return q.eventDate === dateStr && q.status !== 'Inactive'; });
  var bom = ctx.bom || loadRecipeBom_();
  var conv = ctx.conv || loadCaseConversions_();
  var tags = supplyTagMap_();
  var aliasMap = ctx.aliasMap || getItemAliasMap_();

  var storeItems = {};   // store -> { lowerName -> {item, qty} }
  var storeOrder = [];   // stores in first-seen order
  var handedOff = [];    // orders passed to another store (excluded from totals)

  quotes.forEach(function(q) {
    var lineItems = [];
    try { lineItems = JSON.parse(q.lineItems) || []; } catch(e) { lineItems = []; }
    var pickList = lineItems.map(function(it) {
      return { item: (it.description || '').toString().trim(), qty: parseFloat(it.quantity) || 0 };
    }).filter(function(x) { return x.item && x.qty > 0; });

    // handoffStore populated by §8; blank/undefined until then (harmless).
    var handoff = (q.handoffStore || '').toString().trim();
    if (handoff) {
      handedOff.push({
        guest: q.customerName || q.contactName || '(no name)',
        store: handoff, from: (q.locationName || '').toString().trim(), items: pickList
      });
      return;
    }

    var store = (q.locationName || '').toString().trim() || '(no store set)';
    if (!storeItems[store]) { storeItems[store] = {}; storeOrder.push(store); }
    pickList.forEach(function(x) {
      var key = x.item.toLowerCase();
      if (!storeItems[store][key]) storeItems[store][key] = { item: x.item, qty: 0 };
      storeItems[store][key].qty += x.qty;
    });
  });

  var stores = storeOrder.map(function(store) {
    var items = Object.keys(storeItems[store]).map(function(k) { return storeItems[store][k]; });
    var ex = explodeItemsToIngredients_(items, bom, conv, tags, aliasMap);
    return { store: store, ingredients: ex.ingredients, noRecipe: ex.noRecipe };
  });
  stores.sort(function(a, b) { return a.store.localeCompare(b.store); });

  return { date: dateStr, orderCount: quotes.length, stores: stores, handedOff: handedOff };
}

// Grouped reference boxes from Supplies_Reference: [{section, rows:[{item, detail}]}] in sheet order.
function getSuppliesReference() {
  var sheet = getSpreadsheet().getSheetByName(TAB_SUPPLIES_REF);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var sections = [], bySection = {};
  data.forEach(function(r) {
    var section = (r[0] || '').toString().trim();
    var item = (r[1] || '').toString().trim();
    var detail = (r[2] || '').toString().trim();
    if (!section && !item && !detail) return;
    if (!bySection[section]) { bySection[section] = { section: section, rows: [] }; sections.push(bySection[section]); }
    bySection[section].rows.push({ item: item, detail: detail });
  });
  return sections;
}

// Rewrite the whole Supplies_Reference sheet from the edited sections (small sheet, full rewrite is safe).
function saveSuppliesReference(sections) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_SUPPLIES_REF);
  if (!sheet) sheet = ss.insertSheet(TAB_SUPPLIES_REF);
  var rows = [];
  (sections || []).forEach(function(sec) {
    var section = (sec.section || '').toString();
    (sec.rows || []).forEach(function(r) {
      var item = (r.item || '').toString();
      var detail = (r.detail || '').toString();
      if (!section && !item && !detail) return; // skip fully-blank rows
      rows.push([section, item, detail]);
    });
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  sheet.getRange('A:C').setNumberFormat('@'); // keep phone numbers / codes as text
  sheet.getRange(1, 1, 1, 3).setValues([['Section', 'Item', 'Detail']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  SpreadsheetApp.flush();
  return true;
}

// ── STORE CONTACTS ───────────────────────────────────────────
// Its own nav page, but NOT its own data: it's the "Important Contacts" section of
// Supplies_Reference. One source of truth, two pages rendering different slices —
// the Supplies Calculator view filters this section out so it isn't shown twice.
const STORE_CONTACTS_SECTION = 'Important Contacts';

function getStoreContacts() {
  var found = null;
  getSuppliesReference().forEach(function(sec) {
    if (sec.section.toLowerCase() === STORE_CONTACTS_SECTION.toLowerCase()) found = sec;
  });
  return found ? found.rows : [];
}

// Replace just this section, leaving every other Supplies_Reference section untouched.
// Appends the section if the sheet doesn't have it yet (e.g. an older install).
function saveStoreContacts(rows) {
  var sections = getSuppliesReference(), hit = false;
  sections.forEach(function(sec) {
    if (sec.section.toLowerCase() === STORE_CONTACTS_SECTION.toLowerCase()) { sec.rows = rows || []; hit = true; }
  });
  if (!hit) sections.push({ section: STORE_CONTACTS_SECTION, rows: rows || [] });
  saveSuppliesReference(sections);
  return true;
}


// ── FAQ ──────────────────────────────────────────────────────
// Content lives in the sheet so it's editable without touching code.
const TAB_FAQ = 'FAQ';

function ensureFaqSheet_() {
  var ss = getSpreadsheet();
  var s = ss.getSheetByName(TAB_FAQ);
  if (!s) {
    s = ss.insertSheet(TAB_FAQ);
    s.getRange(1, 1, 1, 2).setValues([['Question', 'Answer']]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 340); s.setColumnWidth(2, 620);
    var seed = seedFaqRows_();
    s.getRange(2, 1, seed.length, 2).setValues(seed);
    s.getRange('A:B').setWrap(true);
  }
  return s;
}

function seedFaqRows_() {
  return [
    ['How do I build a quote?', 'Catering Quotes → New Quote. Pick Pickup or Delivery first (it sets which prices are used), then the customer, then the items. Save, Save & Email, or Preview/Print when you\'re done.'],
    ['A customer is tax exempt — what do I do?', 'Flip the Tax Exempt toggle on the quote. It asks whether their Texas Sales and Use Tax Exemption Certificate is on file and gives you a "look it up" button. If it isn\'t on file, use "Request from Guest" to email them the upload link.'],
    ['How do I find an old order?', 'Catering Quotes → History. Search by customer name or filter by event date. "Edit & Reuse" copies it into a new quote; "Reorder" is the one-click version for repeat customers.'],
    ['What\'s the difference between Calendar and Upcoming Orders?', 'Calendar is the full month/week/day view of everything. Upcoming Orders is the tight version — just the next 7 days, so you can see what\'s actually coming without scrolling a month.'],
    ['Why did the Supplies Calculator warn me before showing anything?', 'Because its numbers depend on the recipe table being right, and it only covers the highest-volume items. Always cross-check before you place an order.'],
    ['An item on a quote isn\'t showing up in the Supplies Calculator.', 'It needs to be mapped to a menu item. Settings → Item Name Matching lists everything unmatched. Mapping is retroactive, so fixing it there also fixes past orders.'],
    ['How do I hand an order off to another store?', 'Open the order and use "Hand off". Partner store details are on the Partner Stores page. Handed-off orders drop out of your store\'s supplies totals and are listed separately.'],
    ['Who do I call about bread, produce, or a credit?', 'Store Contacts has the numbers, plus the calling script — greet with your name and store number, then state the reason for the call.'],
    ['Can I change prices, tax rate, or email templates myself?', 'Yes — all of it is in Settings, or in each store\'s Menu sheet tab. No code changes needed.'],
    ['The tool locked me out asking me to verify prices.', 'That\'s the quarterly price check. Type the current POS prices for the three items it lists and it unlocks. It exists so quotes never go out at stale prices.']
  ];
}

function getFaq() {
  var s = ensureFaqSheet_();
  if (s.getLastRow() < 2) return [];
  var out = [];
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function(r) {
    var q = (r[0] || '').toString().trim(), a = (r[1] || '').toString().trim();
    if (q || a) out.push({ question: q, answer: a });
  });
  return out;
}

// Idempotent: create + seed the two Supplies sheets only when missing/empty.
function initSuppliesSheets_() {
  var ss = getSpreadsheet();

  var bom = ss.getSheetByName(TAB_RECIPE_BOM);
  if (!bom) bom = ss.insertSheet(TAB_RECIPE_BOM);
  if (!bom.getRange('A1').getValue()) {
    var bomRows = seedRecipeBomRows_();
    bom.getRange(1, 1, 1, 3).setValues([['Menu Item', 'Component', 'Qty per Unit']]).setFontWeight('bold');
    bom.getRange(2, 1, bomRows.length, 3).setValues(bomRows);
    bom.setFrozenRows(1);
    bom.setColumnWidth(1, 330); bom.setColumnWidth(2, 210); bom.setColumnWidth(3, 110);
  }

  var ref = ss.getSheetByName(TAB_SUPPLIES_REF);
  if (!ref) ref = ss.insertSheet(TAB_SUPPLIES_REF);
  if (!ref.getRange('A1').getValue()) {
    var refRows = seedSuppliesReferenceRows_();
    ref.getRange('A:C').setNumberFormat('@'); // keep phone numbers / codes as text
    ref.getRange(1, 1, 1, 3).setValues([['Section', 'Item', 'Detail']]).setFontWeight('bold');
    ref.getRange(2, 1, refRows.length, 3).setValues(refRows);
    ref.setFrozenRows(1);
    ref.setColumnWidth(1, 230); ref.setColumnWidth(2, 220); ref.setColumnWidth(3, 430);
  }

  var cc = ss.getSheetByName(TAB_CASE_CONV);
  if (!cc) cc = ss.insertSheet(TAB_CASE_CONV);
  if (!cc.getRange('A1').getValue()) {
    var ccRows = seedCaseConversionRows_();
    cc.getRange(1, 1, 1, 3).setValues([['Component', 'Per Case', 'Per Bag']]).setFontWeight('bold');
    cc.getRange(2, 1, ccRows.length, 3).setValues(ccRows);
    cc.setFrozenRows(1);
    cc.setColumnWidth(1, 220); cc.setColumnWidth(2, 110); cc.setColumnWidth(3, 110);
  }
}

// Case_Conversions seed: Component | Per Case | Per Bag. Component names must
// match the BOM's component names exactly (case-insensitive). Per Bag optional —
// blank = don't show a bag figure. Team-editable; fill in the ones left blank.
function seedCaseConversionRows_() {
  return [
    ['Filet', 150, 25],
    ['Spicy filet', 150, 25],
    ['Breakfast filet', 240, ''],
    ['Spicy breakfast filet', 240, ''],
    ['Nuggets', 1080, ''],
    ['Strips', 420, ''],
    ['Sausage patty', 420, ''],
    ['Bun', 60, ''],          // bread rack of 60
    ['Cookie', 196, ''],
    ['Chips', 60, ''],
    ['Dipping sauce', 492, ''],
    ['Mac & cheese bag', 4, '']   // 4 bags per case
  ];
}

// Run ONCE from the editor to force-refresh Case_Conversions, OVERWRITING it.
// Same idea as resetRecipeBom — the normal seed won't touch a non-empty sheet.
function resetCaseConversions() {
  var ss = getSpreadsheet();
  var cc = ss.getSheetByName(TAB_CASE_CONV) || ss.insertSheet(TAB_CASE_CONV);
  cc.clear();
  var ccRows = seedCaseConversionRows_();
  cc.getRange(1, 1, 1, 3).setValues([['Component', 'Per Case', 'Per Bag']]).setFontWeight('bold');
  cc.getRange(2, 1, ccRows.length, 3).setValues(ccRows);
  cc.setFrozenRows(1);
  cc.setColumnWidth(1, 220); cc.setColumnWidth(2, 110); cc.setColumnWidth(3, 110);
  SpreadsheetApp.flush();
  return 'Case_Conversions reseeded with ' + ccRows.length + ' rows.';
}

// Run ONCE from the editor to force-refresh Recipe_BOM with the current recipe
// list, OVERWRITING whatever is there (the normal seed won't touch a non-empty
// sheet). Use after changing seedRecipeBomRows_, or if the tab was created stale/empty.
// This wipes manual BOM edits — it's a deliberate reset.
function resetRecipeBom() {
  var ss = getSpreadsheet();
  var bom = ss.getSheetByName(TAB_RECIPE_BOM) || ss.insertSheet(TAB_RECIPE_BOM);
  bom.clear();
  var bomRows = seedRecipeBomRows_();
  bom.getRange(1, 1, 1, 3).setValues([['Menu Item', 'Component', 'Qty per Unit']]).setFontWeight('bold');
  bom.getRange(2, 1, bomRows.length, 3).setValues(bomRows);
  bom.setFrozenRows(1);
  bom.setColumnWidth(1, 330); bom.setColumnWidth(2, 210); bom.setColumnWidth(3, 110);
  SpreadsheetApp.flush();
  return 'Recipe_BOM reseeded with ' + bomRows.length + ' rows.';
}

// Recipe_BOM seed (v1 scope), keyed on EXACT menu item names from the CFA sheet.
// Package meals flattened; a 4pc's Fruit OR Kale Crunch side isn't tracked in v1
// so its tracked components equal the 3pc. Trays include their "Chilled" twins
// (same protein) and use the real "<Item> Tray, <Size>" comma naming.
function seedRecipeBomRows_() {
  var rows = [];
  function bom(item, comps) { comps.forEach(function(c) { rows.push([item, c[0], c[1]]); }); }

  // Breakfast biscuits
  bom('Biscuit w/ Chick-fil-A, Egg, Cheese', [['Breakfast filet', 1], ['Folded egg', 1], ['American cheese slice', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Chick-fil-A Chicken',      [['Breakfast filet', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Spicy CFA,Egg,Cheese',     [['Spicy breakfast filet', 1], ['Folded egg', 1], ['American cheese slice', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Spicy Chicken',            [['Spicy breakfast filet', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Sausage, Egg, Cheese',     [['Sausage patty', 1], ['Folded egg', 1], ['American cheese slice', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Sausage',                  [['Sausage patty', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Bacon, Egg, Cheese',       [['Bacon', 1], ['Folded egg', 1], ['American cheese slice', 1], ['Foil wrap', 1]]);
  bom('Biscuit w/ Bacon',                    [['Bacon', 1], ['Foil wrap', 1]]);

  // Sandwiches (a la carte, regular & spicy)
  bom('Chick-fil-A Sandwich',  [['Bun', 1], ['Filet', 1], ['Pickles', 2], ['Sandwich foil bag', 1]]);
  bom('Spicy Chicken Sandwich', [['Bun', 1], ['Spicy filet', 1], ['Pickles', 2], ['Sandwich foil bag', 1]]);

  // Package meals (flattened, exact names). 3pc and 4pc share tracked components.
  var cfaMeal   = [['Bun', 1], ['Filet', 1], ['Pickles', 2], ['Sandwich foil bag', 1], ['Chips', 1], ['Cookie', 1]];
  var spicyMeal = [['Bun', 1], ['Spicy filet', 1], ['Pickles', 2], ['Sandwich foil bag', 1], ['Chips', 1], ['Cookie', 1]];
  var nugMeal   = [['Nuggets', 8], ['Dipping sauce', 1], ['Chips', 1], ['Cookie', 1]];
  bom('Packaged Meal-CFA Sandwich 3pc', cfaMeal);
  bom('Packaged Meal-CFA Sandwich w/ Fruit OR Kale Crunch 4pc', cfaMeal);
  bom('Packaged Meal-Spicy Sandwich 3pc', spicyMeal);
  bom('Packaged Meal-Spicy Sandwich w/ Fruit OR Kale Crunch 4pc', spicyMeal);
  bom('Packaged Meal-Nuggets (8) 3pc', nugMeal);
  bom('Packaged Meal-Nuggets (8) w/ Fruit OR Kale Crunch 4pc', nugMeal);

  // Nugget trays (+ Chilled twins). Small 64 / Medium 120 / Large 200.
  [['Small', 64, 1], ['Medium', 120, 1], ['Large', 200, 2]].forEach(function(t) {
    var comps = [['Nuggets', t[1]], ['Sauce bottle (8oz)', t[2]], ['Serving spoon', 1]];
    bom('Nugget Tray, ' + t[0], comps);
    bom('Nugget Tray Chilled, ' + t[0], comps);
  });
  // Strip trays (+ Chilled twins). Small 24 / Medium 45 / Large 75.
  [['Small', 24, 1], ['Medium', 45, 1], ['Large', 75, 2]].forEach(function(t) {
    var comps = [['Strips', t[1]], ['Sauce bottle (8oz)', t[2]], ['Serving tong', 1]];
    bom('Chick-n-Strips Tray, ' + t[0], comps);
    bom('Chick-n-Strips Tray Chilled, ' + t[0], comps);
  });
  // Chilled grilled sub sandwich trays (regular + spicy). Halves 6 / 12 / 16.
  [['Small', 6], ['Medium', 12], ['Large', 16]].forEach(function(t) {
    bom('Chilled Grilled Sandwich Tray, ' + t[0], [['Sandwich halves', t[1]]]);
    bom('Spicy Chilled Grilled Sand Tray, ' + t[0], [['Sandwich halves', t[1]]]);
  });
  // Grilled chicken cool wrap trays. Halves 6 / 10 / 14.
  [['Small', 6], ['Medium', 10], ['Large', 14]].forEach(function(t) {
    bom('Grilled Chicken Wrap Tray, ' + t[0], [['Wrap halves', t[1]]]);
  });
  // Mac & Cheese trays. Small 1 bag / Large 2 bags (thaw the day before).
  bom('Mac & Cheese Tray, Small', [['Mac & cheese bag', 1]]);
  bom('Mac & Cheese Tray, Large', [['Mac & cheese bag', 2]]);

  // A la carte entree counts — protein drivers, name → count.
  [1, 5, 8, 12, 30].forEach(function(n) { bom('Nuggets (' + n + ')', [['Nuggets', n]]); });
  bom('Nuggets (30) w/ Football Tray', [['Nuggets', 30]]);
  bom('Nuggets (30) w/ Heart Tray', [['Nuggets', 30]]);
  [1, 5, 8, 12, 30].forEach(function(n) { bom('Grilled Nuggets (' + n + ')', [['Grilled nuggets', n]]); });
  [2, 3, 4, 10].forEach(function(n) { bom('Chick-n-Strips (' + n + ')', [['Strips', n]]); });

  return rows;
}

// Supplies_Reference seed. Section | Item | Detail. Blank Item = full-width note.
function seedSuppliesReferenceRows_() {
  var S = 'Case/Truck Conversions';
  var rows = [
    [S, 'Filet case', '150 filets'],
    [S, 'Breakfast Filet case', '240 breakfast filets'],
    [S, 'Nugget case', '1080 nuggets'],
    [S, 'Sausage case', '420 sausages'],
    [S, 'Tender case', '420 strips'],
    [S, 'Coater', '562 filets'],
    [S, 'Peanut oil', '500 filets'],
    [S, 'Pickles bucket', '1200 sandwiches'],
    [S, 'Bread rack', '60 buns'],
    [S, 'Bun oil case', '4270 buns'],
    [S, 'Chips case', '60 chips'],
    [S, 'Cookies case', '192 cookies'],
    [S, 'Sauce case', '492 sauces'],
    [S, 'Ketchup case', '500'],
    [S, 'Corn/bean case', '150 veggie wraps'],
    [S, 'Yogurt case', '30 parfaits'],
    [S, 'Blueberries case', '87 parfaits, or 83 medium fruit cups'],
    [S, 'Apples case', '125 medium fruit cups, or 3 lbs apples per Large fruit tray'],
    [S, 'Mandarin oranges case', '120 medium fruit cups, or 26oz per Large fruit tray'],
    [S, 'Biscuit mix box', '320 biscuits']
  ];
  var O = 'Order Checklist';
  rows.push([O, '', '(Add your standard order checklist items here — edit this sheet.)']);
  var L = 'Catering Order Lead-Time Rules';
  [
    'Non-protein supplies: order to be received the DAY BEFORE the catering order date.',
    'Proteins (chicken, sausage): order 3 DAYS before the catering order date (allows thaw + filet prep time).',
    'Large catering orders: split ordering — proteins on the 3-day timeline, everything else on the day-before timeline.',
    'Orders involving DISD: require #8 bags.',
    'Box meals required if guest count < 50; #8 bags required if guest count ≥ 50.',
    'Large bun orders: divide total sandwich count across all orders by 60 to get additional bun racks needed; call the bread contact with that number.'
  ].forEach(function(t) { rows.push([L, '', t]); });
  var T = 'Truck Ordering Schedule';
  [['Monday', 'Wednesday'], ['Tuesday', 'Thursday'], ['Wednesday', 'Friday'], ['Thursday', 'Saturday'], ['Friday', 'Monday'], ['Saturday', 'Tuesday']]
    .forEach(function(p) { rows.push([T, 'Order ' + p[0], 'Receive ' + p[1]]); });
  rows.push(['Suggested Ordering Formula', '', '(Forecast Usage) + (Safety Stock) − (On Hand) − (On Order) = Suggested Quantity.  Forecast = Store Usage + Catering Usage.  Reference only — not automated in v1.']);
  var B = 'Boil-Outs';
  rows.push([B, 'Fries', '3 canola']);
  rows.push([B, 'Chicken', '2 peanut']);
  var C = 'Daily Chicken Usage Baselines';
  [['Filet', '5'], ['Spicy', '8'], ['Nugget', '12'], ['Strip', '3'], ['Grill Filet', '2'], ['Grill Nugget', '2'], ['Breakfast Filet', '1'], ['Spicy Breakfast Filet', 'Build to 2 cases in freezer']]
    .forEach(function(p) { rows.push([C, p[0], p[1]]); });
  rows.push(['Beverages (reference)', '', 'Gallon jugs include 12 foam 14oz cups per gallon (lids/straws not included).']);
  var K = 'Important Contacts';
  // Blank on purpose — each store types its own contacts into this tab
  rows.push([K, 'Bread contact', '']);
  rows.push([K, 'Distributor customer service', '']);
  rows.push([K, 'Produce vendor', '']);
  rows.push([K, 'Store email', '']);
  rows.push([K, 'Store phone', '']);
  rows.push([K, '', 'Calling script: greet with your name + store number, then state the reason for the call. For credit requests, describe the item/issue and ask for email confirmation of the credit.']);
  return rows;
}

// Move quotes older than the "Archive After Days" setting (default 120) to a
// hidden archive sheet instead of deleting.
// Nothing is lost — the archive keeps the full row so old quotes can always be reviewed.
// Trigger entry point — install with setupArchiveTrigger(). This is the ONLY automation
// that deletes rows, so a silent mid-run failure is the worst kind: it alerts.
function cleanOldQuotes() {
  try {
    return _cleanOldQuotes();
  } catch (e) {
    var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e2) { /* no recipient */ }
    if (alertTo) {
      try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'cleanOldQuotes failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e3) { /* give up quietly */ }
    }
    return 'Archive sweep failed: ' + (e.message || e);
  }
}

function _cleanOldQuotes() {
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'Nothing to archive.';

  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var archiveAfterDays = getArchiveAfterDays_(getSettings());
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - archiveAfterDays);
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var toArchive = [];   // full rows to copy into the archive
  var rowsToDelete = []; // sheet row numbers, deleted bottom-up
  for (var i = data.length - 1; i >= 0; i--) {
    if (!(new Date(data[i][1]) < cutoff)) continue; // created date not old enough yet
    // NEVER archive an order whose event hasn't happened. Catering books months ahead, and
    // nothing reads the archive back — archiving a future order would drop it off the
    // calendar, runsheets, supplies calc, PO warnings and reminders BEFORE the event,
    // while its Google Calendar entry stayed behind. No event date → created-date rule only.
    var eventKey = normalizeEventDate_(data[i][15], tz);
    if (eventKey && eventKey >= todayKey) continue;
    toArchive.push(data[i]);
    rowsToDelete.push(i + 2);
  }
  if (!toArchive.length) return 'Nothing old enough to archive.';

  var archive = ensureArchiveSheet_(sheet, lastCol);
  archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, lastCol).setValues(toArchive);
  SpreadsheetApp.flush();

  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]); // already sorted descending, safe to delete in order
  }
  return 'Archived ' + toArchive.length + ' quote(s).';
}

// Hidden archive that mirrors the Quotes header. Widens itself and re-syncs the header
// every run — Quotes has gained columns before (Status/Handoff/Driver), and a narrower
// archive would make the setValues range throw.
function ensureArchiveSheet_(quotesSheet, needCols) {
  var ss = getSpreadsheet();
  var archive = ss.getSheetByName(TAB_QUOTES_ARCHIVE);
  if (!archive) {
    archive = ss.insertSheet(TAB_QUOTES_ARCHIVE);
    archive.setFrozenRows(1);
    archive.hideSheet();
  }
  if (archive.getMaxColumns() < needCols) {
    archive.insertColumnsAfter(archive.getMaxColumns(), needCols - archive.getMaxColumns());
  }
  archive.getRange(1, 1, 1, needCols)
    .setValues(quotesSheet.getRange(1, 1, 1, needCols).getValues())
    .setFontWeight('bold');
  return archive;
}

// The nightly archive was the one automation with no in-app switch — the setup guide told
// operators to add the trigger by hand, so it was usually never installed at all.
function setupArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'cleanOldQuotes') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('cleanOldQuotes').timeBased().everyDays(1).atHour(0).create();
  return 'Nightly archive enabled (runs daily around midnight).';
}
function removeArchiveTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'cleanOldQuotes') { ScriptApp.deleteTrigger(t); removed++; } });
  return removed > 0 ? 'Nightly archive trigger removed.' : 'No active archive trigger found.';
}
function getArchiveTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'cleanOldQuotes'; });
}


// ── SHEET INITIALIZATION ─────────────────────────────────────

// initializeSheet is what a non-technical maintainer opens first, so it reads as a short
// list of "make sure each tab exists" steps. Each helper below is the verbatim block that
// used to be inlined — same order, same idempotent "only touch it if it's missing" rules.

// The Settings seed, grouped the way the tab reads. Same pattern as seedRecipeBomRows_ /
// seedFaqRows_ elsewhere in this file: data lives in its own function so the code that
// writes it stays short. Order here is the order rows appear on a fresh sheet.
function seedSettingsRows_() {
  return seedStoreSettings_().concat(seedAutomationSettings_(), seedTemplateSettings_());
}

// Store identity, both locations, tax, and the operational numbers.
function seedStoreSettings_() {
  return [
    // Store details start blank — the Catering Setup wizard asks for them once.
    // Seeding only adds missing rows, so an existing sheet keeps its values.
    ['Store Name (Active)',         ''],
    // Short tab labels — change these for a different restaurant, then re-run
    // initializeSheet (it renames the tabs in place, keeping your prices).
    ['Store 1 Tab Prefix',          'S1'],  // Setup asks for the real code
    ['Store 2 Tab Prefix',          'S2'],
    ['Location 1 Name',             ''],
    ['Location 1 Address',          ''],
    ['Location 1 Phone',            ''],
    ['Location 2 Name',             ''],
    ['Location 2 Address',          ''],
    ['Location 2 Phone',            ''],
    ['Quote Contact Name',          ''],
    ['Default Tax Rate (%)',         8.25],
    ['Calendar Lead Time (Minutes)', 30],
    ['Archive After Days',           120],
    ['Price Check Interval (Days)',  90],
    ['Price Check Categories',      'Catering Items, Catering - Bulk Drinks'],
    ['Quote Valid For (Days)',       30],
    ['Delivery Warning Count',       4],
    ['Delivery Warning Window (Minutes)', 30],
    ['Route Suggest Radius (Miles)',  5],
    ['Route Suggest Window (Minutes)', 90],
    // TMS bag capacities for delivery orders (runsheet + calendar). Hot trays only;
    // mac & cheese is the one tray that fits a small.
    ['TMS Large: Sandwiches',        80],
    ['TMS Small: Sandwiches',        40],
    ['TMS Large: Biscuits',          100],
    ['TMS Small: Biscuits',          50],
    ['TMS Large: Hot Trays',         2],
    ['TMS Small: Mac & Cheese Trays', 2],
    ['Confirmation Enabled',        'FALSE'],
    ['Confirmation Subject',        'See you tomorrow — your Chick-fil-A catering order {{quoteId}}'],
    ['Confirmation Body',           'Hi {{contactPerson}},\n\nJust confirming your catering order for {{customer}} — we\'ll see you {{when}}.\n\nOrder total: {{total}}\n\nIf anything has changed, reply to this email or call us at {{phone}} and we\'ll take care of it.\n\nSee you soon!\n{{contact}}\nChick-fil-A {{location}}'],
    ['PO Alert Enabled',            'TRUE'],
    ['PO Alert Days Before',         7],
    ['PO Alert Email',              ''],
    ['POS Reminder Enabled',        'FALSE'],
    ['POS Reminder Email',          ''],
    ['POS Reminder Message',        'Reminder: please log all catering orders for the upcoming week into our POS system.'],
    ['Incoming Order Sender',       'one@email.chick-fil-a.com'],
    ['Incoming Order Subject',      'Incoming Catering Order'],
    ['Auto Import Enabled',         'FALSE'],
    // Calendar colour for orders the tool creates from an incoming order email, so they
    // stand out from hand-built quotes. Use a colour name from the Calendar Color Names
    // card in Settings (Banana, Tomato, Basil…), or "None" for no colour at all.
    ['Email Import Calendar Color', 'Banana']
  ];
}

// The automation switches and their schedules/thresholds.
function seedAutomationSettings_() {
  return [
    // Weekly DISD delivered-orders reminder (Saturday 12pm)
    ['DISD Reminder Enabled',       'FALSE'],
    ['DISD Reminder Recipients',    ''],  // blank = same as PO Alert Email
    // Comma-separated. Matched against customer name + contact name + email, with spaces
    // and punctuation stripped from both sides — so "Dallas ISD" also matches dallasisd.org.
    ['DISD Reminder Match Terms',   ''],
    ['DISD Reminder Subject',       'Weekly DISD catering recap — {{count}} order(s)'],
    ['DISD Reminder Message',       'DISD orders delivered in the last 7 days:'],
    ['Off-Menu Markup (%)',          30],
    ['Tax Form Request Subject',    'Texas Sales and Use Tax Exemption Certificate needed for your catering quote {{quoteId}}'],
    ['Tax Form Request Body',       'Hi {{contactPerson}},\n\nThanks for your catering order with Chick-fil-A {{location}}! To honor the tax exemption on quote {{quoteId}}, we need a copy of your organization\'s Texas Sales and Use Tax Exemption Certificate on file.\n\nPlease upload a PDF of the Certificate using this secure link:\n{{uploadLink}}\n\nIt takes about a minute and asks for your name and quote number ({{quoteId}}) so we can match it up.\n\nThank you!\n{{contact}}\nChick-fil-A {{location}}\n{{phone}}'],
    ['Logo (Base64)',                ''],
    ['Email Subject',               'Your Catering Quote from Chick-fil-A {{location}}'],
    ['Email Body',                  'Hi {{customer}},\n\nThank you for considering Chick-fil-A {{location}} for your catering needs! We appreciate you reaching out and would love to help make your event something special.\n\nPlease find your catering quote attached. If you have any questions or would like to make any changes, don\'t hesitate to reach out — we\'re happy to help.\n\nWe look forward to serving you!\n\nWarm regards,\n{{contact}}\nChick-fil-A {{location}}\n{{phone}}'],
    ['BCC Email',                   ''],
    // Quote follow-up reminders (9am trigger, installed separately from Settings)
    ['Reminder Enabled',            'FALSE'],
    ['Reminder After Days',          3],
    // Blast guards. Without these, turning reminders on sweeps the ENTIRE quote history
    // in one run — every quote older than "After Days" that isn't already logged.
    ['Reminder Max Age Days',        21],  // older than this = stale, never follow up
    ['Reminder Max Per Run',         10],  // hard ceiling on emails from one run
    ['Reminder Send To Customer',   'FALSE'],
    ['Reminder Send To Internal',   'TRUE'],
    ['Reminder Internal Email',     ''],
    ['Reminder Subject',            'Following up on your Catering Quote — {{quoteId}}'],
    ['Reminder Body',               'Hi {{customer}},\n\nWe wanted to follow up on the catering quote we sent you {{daysSince}} days ago (Quote {{quoteId}}, total: {{total}}).\n\nPlease let us know if you have any questions!\n\n{{contact}}\nChick-fil-A {{location}}']
  ];
}

// Operator-editable email templates, hand-off defaults, and calendar colour labels.
function seedTemplateSettings_() {
  return [
    // Hand-off to a partner store
    ['Start Address',               ''],  // Setup fills this from the store address
    ['Handoff Distance Threshold (Miles)', 20],
    ['Handoff Email Subject',       'Catering hand-off: {{customer}} on {{date}}'],
    ['Handoff Email Body',          'Hi {{poc}},\n\nWe\'d like to hand off a catering order to {{partnerStore}} — it\'s closer to your store for this one.\n\nCustomer: {{customer}}\nContact: {{contact}}\nDate / Time: {{date}} {{time}}\nOrder type: {{orderType}}\nDelivery address: {{address}}\nOrder total: {{total}}\n\n--- Items ---\n{{items}}\n\nLet us know if you can take it and we\'ll send over the full details.\n\nThanks!\n{{fromContact}}\nChick-fil-A {{fromStore}}'],
    // "We got your PO" — offered when a PO is first attached to an order.
    ['PO Confirmation Enabled',     'TRUE'],
    ['PO Confirmation Subject',     'We received your PO — your catering order {{quoteId}} is confirmed'],
    ['PO Confirmation Body',        'Hi {{contactPerson}},\n\nGood news — we\'ve received PO {{poNumber}} for {{customer}}, and your catering order is confirmed.\n\nWe\'ll have everything ready for {{when}}.\n{{orderLine}}\nOrder total: {{total}}\n\n--- Your order ---\n{{items}}\n\nIf anything changes, just reply to this email or call us at {{phone}} and we\'ll take care of it.\n\nThank you for letting us serve you!\n{{contact}}\nChick-fil-A {{location}}'],
    // "We need your PO" — sent by hand from Calendar week view or the Email tab.
    ['PO Request Subject',          'Reminder: we need a PO for your upcoming catering order'],
    ['PO Request Body',             'Hi {{contactPerson}},\n\nJust a friendly reminder — we don\'t have a PO on file yet for your upcoming catering order:\n\n{{orders}}\n\nWhen you get a chance, please reply to this email with the PO number so we can lock everything in. If you have any questions, call us at {{phone}}.\n\nThank you!\n{{contact}}\nChick-fil-A {{location}}'],
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
}

function initSettingsSheet_(ss) {

  // Settings
  var sSheet = ss.getSheetByName(TAB_SETTINGS);
  if (!sSheet) sSheet = ss.insertSheet(TAB_SETTINGS);
  var seedSettings = seedSettingsRows_();
  var existingKeys = {}; // declared out here so the migrations below can read it in both branches
  if (!sSheet.getRange('A1').getValue()) {
    sSheet.getRange(1, 1, seedSettings.length, 2).setValues(seedSettings);
    sSheet.setColumnWidth(1, 220);
    sSheet.setColumnWidth(2, 500);
  } else {
    // Migrate: append any seed keys that don't already exist (preserves user edits)
    var existing = sSheet.getRange('A1:A' + Math.max(sSheet.getLastRow(), 1)).getValues();
    existing.forEach(function(r) { if (r[0]) existingKeys[r[0].toString().trim()] = true; });
    seedSettings.forEach(function(pair) {
      if (!existingKeys[pair[0]]) {
        sSheet.getRange(sSheet.getLastRow() + 1, 1, 1, 2).setValues([pair]);
      }
    });
  }

  // The capacity warning changed meaning: it used to count DELIVERIES in a 60-min window,
  // it now counts ALL orders in a 30-min window. Old values (3/60) are stale by definition
  // under the new rule and would nag constantly, so re-point them once at the new defaults.
  // Keyed on a marker setting so a value you tune afterwards is never overwritten again.
  if (!existingKeys['Capacity Rule Version']) {
    updateSetting('Delivery Warning Count', 4);
    updateSetting('Delivery Warning Window (Minutes)', 30);
    updateSetting('Capacity Rule Version', 2);
  }
}

function initMenuSheets_(ss) {
  // Menu — 4 columns: Category, Item Name, Pickup Price, Delivery Price
  var mSheet = ss.getSheetByName(menuTabForStore_('1'));
  if (!mSheet) mSheet = ss.insertSheet(menuTabForStore_('1'));
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

  // Menu (Store 2) — same shape, starts empty. Team fills it from Store 2's POS.
  var m2Sheet = ss.getSheetByName(menuTabForStore_('2'));
  if (!m2Sheet) m2Sheet = ss.insertSheet(menuTabForStore_('2'));
  if (!m2Sheet.getRange('A1').getValue()) {
    m2Sheet.getRange(1, 1, 1, 4).setValues([['Category', 'Item Name', 'Pickup Price', 'Delivery Price']]);
    m2Sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    m2Sheet.setFrozenRows(1);
    m2Sheet.setColumnWidth(1, 160);
    m2Sheet.setColumnWidth(2, 300);
    m2Sheet.setColumnWidth(3, 120);
    m2Sheet.setColumnWidth(4, 120);
    m2Sheet.getRange(2, 3, 30, 2).setNumberFormat('$#,##0.00');
  }
}

function initQuotesSheet_(ss) {
  // Quotes
  var qSheet = ss.getSheetByName(TAB_QUOTES);
  if (!qSheet) qSheet = ss.insertSheet(TAB_QUOTES);
  if (!qSheet.getRange('A1').getValue()) {
    var h = ['Quote ID','Created Date','Customer Name','Contact Name','Order Type','Delivery Address','Line Items (JSON)','Subtotal','Tax Rate Used','Tax Amount','Total','Tax Exempt','Location Name','Customer Email','PO Number','Event Date','Calendar Event ID','Last Modified','Event Time','Order Discount Value','Order Discount Type','Quote Notes','Customer Phone','Status','Handoff Store','Driver'];
    qSheet.getRange(1, 1, 1, h.length).setValues([h]);
    qSheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    qSheet.setFrozenRows(1);
  } else {
    // Migrate: add new column headers if missing
    if (qSheet.getMaxColumns() < 26) qSheet.insertColumnsAfter(qSheet.getMaxColumns(), 26 - qSheet.getMaxColumns()); // a trimmed grid would make the new-col writes throw
    var lastCol = qSheet.getLastColumn();
    // Columns 17-26 in order. Write the whole missing tail as ONE range instead of a
    // setValue() per column — up to 10 separate writes on an older sheet.
    var tailHeaders = ['Calendar Event ID', 'Last Modified', 'Event Time', 'Order Discount Value',
                       'Order Discount Type', 'Quote Notes', 'Customer Phone', 'Status', 'Handoff Store', 'Driver'];
    var firstMissing = Math.max(lastCol + 1, 17);
    if (firstMissing <= 26) {
      var slice = tailHeaders.slice(firstMissing - 17);
      qSheet.getRange(1, firstMissing, 1, slice.length).setValues([slice]).setFontWeight('bold');
    }
  }
}

function initTeamSheets_(ss) {
  // Drivers list (visible, team-editable, single column)
  var drv = ss.getSheetByName('Drivers');
  if (!drv) drv = ss.insertSheet('Drivers');
  if (!drv.getRange('A1').getValue()) {
    drv.getRange(1, 1).setValue('Driver Name').setFontWeight('bold');
    drv.setFrozenRows(1);
    drv.setColumnWidth(1, 260);
  }

  // Partner Store contacts for hand-offs (visible, one row per store, team-editable)
  var ps = ss.getSheetByName('Partner Store');
  if (!ps) ps = ss.insertSheet('Partner Store');
  if (!ps.getRange('A1').getValue()) {
    ps.getRange(1, 1, 1, 5).setValues([['Store Location', 'Operator Name', 'Store Phone', 'Store Email', 'Catering POC']]).setFontWeight('bold');
    ps.setFrozenRows(1);
    ps.setColumnWidth(1, 220); ps.setColumnWidth(2, 180); ps.setColumnWidth(3, 130); ps.setColumnWidth(4, 220); ps.setColumnWidth(5, 160);
  }
}

function initOffMenuSheets_(ss) {
  // Off-menu cheat sheet — visible tab, team-editable. Base prices start blank
  // on purpose: fill them from the POS; the app computes delivery price (+markup).
  var oSheet = ss.getSheetByName(offMenuTabForStore_('1'));
  if (!oSheet) oSheet = ss.insertSheet(offMenuTabForStore_('1'));
  if (!oSheet.getRange('A1').getValue()) {
    oSheet.getRange(1, 1, 1, 2).setValues([['Item Name', 'Base Price']]).setFontWeight('bold');
    oSheet.setFrozenRows(1);
    oSheet.setColumnWidth(1, 300);
    oSheet.setColumnWidth(2, 120);
    var offMenuSeed = ['Sliced Tomatoes', 'Lettuce', 'Sliced Cheese', 'Folded Yellow Egg', 'Folded White Egg', 'English Muffins', 'Shredded Cheese', 'Blue Cheese Crumbles', 'Bacon Crumbles', 'Sliced Hardboiled Egg', '12ct Regular Nugget Package Meal (chips & cookie)', 'Strawberries', 'Blueberries', 'Apples'];
    oSheet.getRange(2, 1, offMenuSeed.length, 2).setValues(offMenuSeed.map(function(n) { return [n, '']; }));
    oSheet.getRange(2, 2, offMenuSeed.length, 1).setNumberFormat('$#,##0.00');
  }

  // Off-menu cheat sheet (Store 2) — headers only; team fills its own list.
  var o2Sheet = ss.getSheetByName(offMenuTabForStore_('2'));
  if (!o2Sheet) o2Sheet = ss.insertSheet(offMenuTabForStore_('2'));
  if (!o2Sheet.getRange('A1').getValue()) {
    o2Sheet.getRange(1, 1, 1, 2).setValues([['Item Name', 'Base Price']]).setFontWeight('bold');
    o2Sheet.setFrozenRows(1);
    o2Sheet.setColumnWidth(1, 300);
    o2Sheet.setColumnWidth(2, 120);
    o2Sheet.getRange(2, 2, 30, 1).setNumberFormat('$#,##0.00');
  }
}

function initQuoteSequenceSheet_(ss) {
  // Sequence
  var seqSheet = ss.getSheetByName(TAB_QUOTE_SEQUENCE);
  if (!seqSheet) seqSheet = ss.insertSheet(TAB_QUOTE_SEQUENCE);
  if (!seqSheet.getRange('A1').getValue()) {
    seqSheet.getRange('A1').setValue('Last Quote Number Used');
    seqSheet.getRange('A1').setFontWeight('bold');
    seqSheet.getRange('B1').setValue(0);
    seqSheet.setColumnWidth(1, 220);
  }
}

function initializeSheet() {
  var ss = getSpreadsheet();
  _settingsMemo = null; // this run seeds Settings directly, so drop any memoized copy

  initSettingsSheet_(ss);

  // Now that the tab-prefix settings exist, rename the menu/cheat-sheet tabs to
  // the configured names IN PLACE (keeps all prices). Safe to re-run; handles
  // both first upgrade (Menu -> "<prefix> Menu") and a later prefix change.
  migrateTabNames_(ss);

  initMenuSheets_(ss);
  initQuotesSheet_(ss);

  // Revisions — prior versions of edited quotes, hidden. Created by the shared helper.
  getRevisionsSheet_();

  // Follow-up reminder log (visible; created here so a fresh install has it)
  initRemindersSheet();

  // Hidden automation logs (day-before confirmations, missing-PO alerts)
  initLogSheet_(TAB_CONFIRMATIONS, ['Quote ID', 'Event Date', 'Sent At']);
  initLogSheet_(TAB_PO_ALERTS, ['Quote ID', 'Alerted At']);
  initLogSheet_(TAB_PO_WARN_DAYBEFORE, ['Quote ID', 'Event Date', 'Sent At']);
  initLogSheet_(TAB_PO_WARN_DAYOF, ['Quote ID', 'Event Date', 'Sent At']);
  initLogSheet_(TAB_TAX_FORMS, ['Quote ID', 'Status', 'Updated At']);
  initLogSheet_(TAB_EMAIL_REVIEW, EMAIL_REVIEW_HEADERS); // auto-import review queue
  var pendSheet = initLogSheet_(TAB_TAX_PENDING, ['Submitted At', 'Organization', 'Contact Name', 'Quote ID', 'PDF Link', 'PDF File ID', 'Status', 'Tax-Exempt Number']);
  // Migrate an existing 7-column pending tab to include the tax-exempt number column
  if (pendSheet.getRange(1, 8).getValue() !== 'Tax-Exempt Number') pendSheet.getRange(1, 8).setValue('Tax-Exempt Number').setFontWeight('bold');
  ensureTaxRegistrySheet_();

  // Supplies Calc reference sheets (BOM + ordering reference)
  initSuppliesSheets_();
  ensureCustomerSheets_(); // customer database + alias-merge tabs
  ensureItemAliasSheet_(); // item-name → menu translation for imports/supplies
  ensureFaqSheet_();       // FAQ page content, seeded once then team-editable

  initTeamSheets_(ss);

  // Create the tax-form Drive folder now, so this setup run triggers the Drive
  // authorization prompt — click ALLOW. Without it, the deployed web app runs
  // without Drive access and the public guest upload page can't save files.
  try { ensureTaxFormFolder_(); } catch (e) { /* Drive not granted — the guest upload won't work until you re-run this and Allow, then redeploy */ }

  initOffMenuSheets_(ss);
  initQuoteSequenceSheet_(ss);

  try { var d = ss.getSheetByName('Sheet1'); if (d && d.getLastRow() <= 1 && d.getLastColumn() <= 1) ss.deleteSheet(d); } catch(e) { /* Sheet1 already gone or protected — non-fatal cleanup */ }
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
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;"><div>'+logo+'<div style="font-size:16px;font-weight:700;">'+esc(pd.storeName)+'</div><div style="font-size:13px;color:#6B7280;margin-top:2px;">'+esc(pd.storeAddress)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.storePhone)+'</div></div><div style="text-align:right;"><div style="font-size:36px;font-weight:300;color:#D1D5DB;letter-spacing:2px;">QUOTE</div><div style="font-size:13px;color:#6B7280;margin-top:4px;font-family:Courier New,monospace;">'+esc(pd.quoteId)+'</div>'+(pd.poNumber?'<div style="font-size:13px;color:#6B7280;margin-top:2px;">PO: '+esc(pd.poNumber)+'</div>':'')+'<div style="font-size:13px;color:#6B7280;margin-top:8px;">Date: '+esc(pd.date)+'</div>'+(pd.time?'<div style="font-size:13px;color:#6B7280;">Time: '+esc(pd.time)+'</div>':'')+(pd.quotedOn?'<div style="font-size:12px;color:#9CA3AF;margin-top:2px;">Quoted on '+esc(pd.quotedOn)+'</div>':'')+(pd.validThrough?'<div style="font-size:12px;color:#9CA3AF;">Quote valid through '+esc(pd.validThrough)+'</div>':'')+'<div style="font-size:14px;font-weight:700;margin-top:8px;">For: '+esc(pd.customerName)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(pd.orderType)+'</div><div style="font-size:13px;color:#6B7280;">'+esc(addr)+'</div>'+'<div style="font-size:13px;font-weight:600;margin-top:4px;">'+esc(pd.contactPerson)+'</div>'+(pd.customerEmail?'<div style="font-size:13px;color:#6B7280;">'+esc(pd.customerEmail)+'</div>':'')+(pd.customerPhone?'<div style="font-size:13px;color:#6B7280;">'+esc(pd.customerPhone)+'</div>':'')+'</div></div>' +
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


// ── TMS BAGS ─────────────────────────────────────────────────

// Settings + BOM + item aliases in one bundle, so a caller pricing many quotes reads each once.
function tmsContext_() {
  return { settings: getSettings(), bom: loadRecipeBom_(), aliasMap: getItemAliasMap_() };
}

// Settings capacity, falling back to the default if blank/zero (e.g. before initializeSheet runs).
function tmsCap_(settings, key, fallback) {
  var n = parseFloat(settings[key]);
  return n > 0 ? n : fallback;
}

// Fill large bags, then one small if the leftover fits it, otherwise one more large.
function packTms_(count, largeCap, smallCap) {
  if (count <= 0) return { large: 0, small: 0 };
  var large = Math.floor(count / largeCap);
  var left = count - large * largeCap;
  if (left > 0 && left <= smallCap) return { large: large, small: 1 };
  if (left > 0) large++;
  return { large: large, small: 0 };
}

// TMS bags for one order's line items. Each kind packs on its own:
//   sandwiches = Recipe_BOM "Sandwich foil bag" (so boxed-meal sandwiches count; sides don't)
//   biscuits   = Recipe_BOM "Foil wrap"
//   hot trays  = nugget/strip/mac & cheese trays, not Chilled. Larges only, except mac &
//                cheese fills any empty large slot first, then goes in smalls.
// Returns { large, small, label } — label '' means nothing rides in a TMS bag.
function tmsBagsForItems_(lineItems, ctx) {
  var items = lineItems || [];
  if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
  var s = ctx.settings;
  var sandwiches = 0, biscuits = 0, hotTrays = 0, macTrays = 0;
  items.forEach(function(i) {
    var qty = parseFloat(i.quantity) || 0;
    var name = (i.description || '').toString().trim();
    if (qty <= 0 || !name) return;
    var canon = ctx.aliasMap[name.toLowerCase()] || name;
    if (/tray/i.test(canon) && !/chilled/i.test(canon)) {
      if (/mac/i.test(canon)) { macTrays += qty; return; }
      if (/nugget|strip/i.test(canon)) { hotTrays += qty; return; }
    }
    (ctx.bom[canon.toLowerCase()] || []).forEach(function(c) {
      var comp = c.component.toLowerCase();
      if (comp === 'sandwich foil bag') sandwiches += c.qtyPer * qty;
      else if (comp === 'foil wrap') biscuits += c.qtyPer * qty;
    });
  });

  var sw = packTms_(sandwiches, tmsCap_(s, 'TMS Large: Sandwiches', 80), tmsCap_(s, 'TMS Small: Sandwiches', 40));
  var bis = packTms_(biscuits, tmsCap_(s, 'TMS Large: Biscuits', 100), tmsCap_(s, 'TMS Small: Biscuits', 50));
  var trayCap = tmsCap_(s, 'TMS Large: Hot Trays', 2);
  var trayLarge = Math.ceil(hotTrays / trayCap);
  var macLeft = Math.max(0, macTrays - (trayLarge * trayCap - hotTrays)); // empty large slots take mac first
  var macSmall = Math.ceil(macLeft / tmsCap_(s, 'TMS Small: Mac & Cheese Trays', 2));

  var large = sw.large + bis.large + trayLarge;
  var small = sw.small + bis.small + macSmall;
  var parts = [];
  if (large) parts.push(large + ' Large');
  if (small) parts.push(small + ' Small');
  return { large: large, small: small, label: parts.join(' + ') };
}

// ── CALENDAR EVENTS ──────────────────────────────────────────

// Title = <prefix> \u2014 <customer>[-"driver"] \u2014 <date>. The driver rides INSIDE the
// customer segment (not a new delimiter) so updateCalendarEventPO's split on ' \u2014 '
// still recovers customer+date and round-trips the driver untouched.
function buildCalendarTitle(customerName, dateStr, poNumber, driver, viaEmail) {
  // "VERIFY PO" not "NEEDS PO" \u2014 the tool has no idea whether this order actually requires
  // a PO, only that nobody has confirmed one either way. It's an instruction to a human,
  // not a claim about the order.
  var prefix = '\uD83D\uDD34 VERIFY PO';
  if (poNumber === 'NO_PO_NEEDED') prefix = '\uD83D\uDFE2 NO PO NEEDED';
  else if (poNumber && poNumber.toString().trim() !== '') prefix = '\uD83D\uDFE2 HAVE PO';
  if (viaEmail) prefix = '\uD83D\uDCE7 ' + prefix; // \uD83D\uDCE7 = auto-created from an email; drops off once a human edits/PO's it
  var name = (customerName || '').toString();
  var drv = (driver || '').toString().trim();
  if (drv) name += '-"' + drv + '"';
  return prefix + ' \u2014 ' + name + ' \u2014 ' + dateStr;
}

// Google's friendly colour name -> the CalendarApp.EventColor key it actually maps to.
// Same list, same order, as CALENDAR_COLORS in App.html — keep the two in step.
var CAL_COLOR_KEYS = {
  'lavender': 'PALE_BLUE', 'sage': 'PALE_GREEN', 'grape': 'MAUVE', 'flamingo': 'PALE_RED',
  'banana': 'YELLOW', 'tangerine': 'ORANGE', 'peacock': 'CYAN', 'graphite': 'GRAY',
  'blueberry': 'BLUE', 'basil': 'GREEN', 'tomato': 'RED'
};

// Turn a Settings colour value into an EventColor key. Accepts all three things an
// operator or the client might supply: the friendly name ('Banana'), the raw enum key
// the colour picker sends ('YELLOW'), or the operator's own RENAMED label (they can
// rewrite "Calendar Color: Banana" to "DISD" in Settings). Blank/'None' = leave the
// event its default colour. Never returns 'skip' — that value suppresses the event entirely.
function resolveCalendarColor_(value, settings) {
  var v = (value || '').toString().trim();
  if (!v || v.toLowerCase() === 'none') return '';
  var byName = CAL_COLOR_KEYS[v.toLowerCase()];
  if (byName) return byName;
  var asKey = v.toUpperCase().replace(/[\s-]+/g, '_');
  for (var n in CAL_COLOR_KEYS) {
    if (CAL_COLOR_KEYS[n] === asKey) return asKey;
    var label = n.charAt(0).toUpperCase() + n.slice(1);
    var custom = ((settings || {})['Calendar Color: ' + label] || '').toString().trim();
    if (custom && custom.toLowerCase() === v.toLowerCase()) return CAL_COLOR_KEYS[n];
  }
  return ''; // unrecognised — default colour beats guessing wrong
}

// Pull a previously-embedded driver name out of a title's customer segment (parts[1]),
// so a normal quote edit re-sync doesn't silently drop the assigned driver.
function driverFromEventTitle_(title) {
  try {
    var seg = ((title || '').toString().split(' \u2014 ')[1]) || '';
    var m = seg.match(/-"([^"]*)"\s*$/);
    return m ? m[1] : '';
  } catch(_) { return ''; }
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
  var title = buildCalendarTitle(customerName, dateStr, quoteData.poNumber, quoteData.driver, quoteData.viaEmail);

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

  var bags = quoteData.orderType === 'Delivery' ? tmsBagsForItems_(items, tmsContext_()).label : '';

  var desc = 'Quote: ' + quoteId + '\n'
    + 'Customer: ' + customerName + '\n'
    + 'Contact: ' + (quoteData.contactName || '') + '\n'
    + (quoteData.customerPhone ? 'Phone: ' + quoteData.customerPhone + '\n' : '')
    + (quoteData.customerEmail ? 'Email: ' + quoteData.customerEmail + '\n' : '')
    + 'Order Type: ' + (quoteData.orderType || 'Pickup') + '\n'
    + (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress ? 'Delivery Address: ' + quoteData.deliveryAddress + '\nDirections: ' + mapsDirectionsUrl_(quoteData.deliveryAddress) + '\n' : '')
    + (bags ? 'TMS Bags: ' + bags + '\n' : '')
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
    try { event.setColor(CalendarApp.EventColor[quoteData.eventColor]); } catch(e) { /* color is cosmetic — ignore an unsupported value */ }
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
    // Preserve an assigned driver across normal edits (saveQuote doesn't carry it).
    var driver = (quoteData.driver != null && quoteData.driver !== '') ? quoteData.driver : driverFromEventTitle_(event.getTitle());

    // Update time if date and time are present
    if (quoteData.date && quoteData.time) {
      var leadMin = getCalendarLeadMinutes_(settings);
      var orderTime = new Date(quoteData.date + 'T' + quoteData.time + ':00');
      var startTime = new Date(orderTime.getTime() - leadMin * 60 * 1000);
      event.setTime(startTime, orderTime);
      var dateStr = orderTime.toLocaleDateString('en-US');
      event.setTitle(buildCalendarTitle(customerName, dateStr, quoteData.poNumber, driver));
    } else if (quoteData.date) {
      var dateStr = new Date(quoteData.date + 'T00:00:00').toLocaleDateString('en-US');
      event.setTitle(buildCalendarTitle(customerName, dateStr, quoteData.poNumber, driver));
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

    var bags = quoteData.orderType === 'Delivery' ? tmsBagsForItems_(items, tmsContext_()).label : '';

    var desc = 'Quote: ' + quoteId + '\n'
      + 'Customer: ' + customerName + '\n'
      + 'Contact: ' + (quoteData.contactName || '') + '\n'
      + (quoteData.customerPhone ? 'Phone: ' + quoteData.customerPhone + '\n' : '')
      + (quoteData.customerEmail ? 'Email: ' + quoteData.customerEmail + '\n' : '')
      + 'Order Type: ' + (quoteData.orderType || 'Pickup') + '\n'
      + (quoteData.orderType === 'Delivery' && quoteData.deliveryAddress ? 'Delivery Address: ' + quoteData.deliveryAddress + '\nDirections: ' + mapsDirectionsUrl_(quoteData.deliveryAddress) + '\n' : '')
      + (bags ? 'TMS Bags: ' + bags + '\n' : '')
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
    try { var e = c.getEventById(calendarEventId); if (e) event = e; } catch(_) { /* event may live on another calendar — keep looking */ }
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
    // The description's "PO:" line was written once at event creation and never touched
    // again, so it kept saying PENDING forever even after a PO was saved — the title went
    // green while the body still said the opposite. Rewrite just that one line.
    try { event.setDescription(applyPoToDescription_(event.getDescription(), poNumber)); }
    catch(e2) { /* title is the signal that matters; a description write failure isn't fatal */ }
    return true;
  } catch(e) {
    return 'setTitle failed: ' + (e.message || String(e));
  }
}

// "PO: 4417-OG" / "PO: Not Required" / "PO: PENDING" — matches what createCalendarEvent writes.
function poLineForDescription_(poNumber) {
  var p = (poNumber || '').toString().trim();
  return 'PO: ' + (p === 'NO_PO_NEEDED' ? 'Not Required' : (p || 'PENDING'));
}

// Surgical: replace the FIRST line starting with "PO: " and leave the rest of the
// description byte-for-byte alone. Rebuilding the whole body isn't possible here — this
// function only receives the PO, not the quote — and rebuilding from partial data would
// wipe the item list. An older event with no PO line gets one appended.
function applyPoToDescription_(desc, poNumber) {
  var line = poLineForDescription_(poNumber);
  var d = (desc == null) ? '' : desc.toString();
  if (/^PO: .*$/m.test(d)) return d.replace(/^PO: .*$/m, line);
  return d ? (d + '\n' + line) : line;
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
  if (!isSettingOn_(settings['Reminder Enabled'])) return 'Reminders are disabled.';

  var daysAfter      = parseInt(settings['Reminder After Days']) || 3;
  var toCustomer     = isSettingOn_(settings['Reminder Send To Customer']);
  var toInternal     = isSettingOn_(settings['Reminder Send To Internal']);
  // "Reminder Internal Email" ships blank, and with "send to customer" off by default that
  // meant an enabled reminder sent NOTHING and reported success. Fall back to the script
  // owner, the same way poAlertRecipients_ does, so enabling it is enough on its own.
  var internalEmail  = (settings['Reminder Internal Email'] || '').toString().trim();
  if (!internalEmail) { try { internalEmail = Session.getEffectiveUser().getEmail(); } catch(e) { internalEmail = ''; } }
  if (!toCustomer && !(toInternal && internalEmail)) {
    return 'Reminders are on, but there is no one to send to \u2014 turn on "send to customer", or set a Reminder Email.';
  }
  var subject        = settings['Reminder Subject'] || 'Following up on your Catering Quote \u2014 {{quoteId}}';
  var body           = settings['Reminder Body']    || 'Hi {{customer}},\n\nWe wanted to follow up on the catering quote we sent you {{daysSince}} days ago (Quote {{quoteId}}, total: {{total}}).\n\nPlease let us know if you have any questions!\n\n{{contact}}\nChick-fil-A {{location}}';

  // Reminders_Sent only stops a REPEAT — it can't stop the first run. Turning this on
  // (or a fresh trigger install) used to sweep every quote ever written and send one
  // email each: ~100 in a morning. A follow-up is only useful in a narrow window, so
  // the sweep is now bounded on BOTH ends, skips events that already happened, and
  // can never send more than "Reminder Max Per Run" emails in a single run.
  var maxAgeDays = parseInt(settings['Reminder Max Age Days'], 10);
  if (isNaN(maxAgeDays) || maxAgeDays < 1) maxAgeDays = 21;
  var maxPerRun = parseInt(settings['Reminder Max Per Run'], 10);
  if (isNaN(maxPerRun) || maxPerRun < 1) maxPerRun = 10;
  var tz         = Session.getScriptTimeZone();
  var todayKey   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var quotes     = getQuotes().filter(function(q) { return q.status !== 'Inactive'; }); // deactivated quotes don't get follow-ups
  var remindedIds = getRemindedQuoteIds();
  var now        = new Date();
  var cutoffMs   = daysAfter * 24 * 60 * 60 * 1000;
  var maxAgeMs   = maxAgeDays * 24 * 60 * 60 * 1000;
  var count      = 0;
  var skippedStale = 0;

  quotes.forEach(function(q) {
    if (count >= maxPerRun) return;
    if (!q.createdDate) return;
    if (remindedIds.indexOf(q.quoteId.toString()) >= 0) return;

    var ageMs = now - new Date(q.createdDate);
    if (ageMs < cutoffMs) return;
    if (ageMs > maxAgeMs) { skippedStale++; return; } // too old to be worth chasing
    // The event already happened — following up on it now is noise, not a nudge.
    if (q.eventDate && q.eventDate < todayKey) { skippedStale++; return; }

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

  return 'Sent ' + count + ' reminder(s).'
    + (count >= maxPerRun ? ' Hit the per-run cap of ' + maxPerRun + ' — the rest go out on the next run.' : '')
    + (skippedStale ? ' Skipped ' + skippedStale + ' stale quote(s) (older than ' + maxAgeDays + ' days, or the event already passed).' : '');
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
const TAB_PO_WARN_DAYBEFORE = 'PO_Warn_DayBefore_Sent';
const TAB_PO_WARN_DAYOF     = 'PO_Warn_DayOf_Sent';

// ── MANUAL "RUN NOW" ─────────────────────────────────────────
// Each button in Settings calls one of these. They run the SAME core the trigger runs, so a
// manual run behaves identically. The per-ORDER sweeps (confirmations, PO alerts + warnings,
// DISD recap, auto-import) each keep a dedup log, so running early never re-sends the same
// order. The POS logging reminder has NO log — it's a standing weekly nudge, so a manual run
// really does send a second copy. Errors come back as text instead of throwing at the UI.
function runNow_(label, fn) {
  try {
    var msg = fn();
    return (msg == null || msg === '') ? (label + ' finished.') : msg.toString();
  } catch (e) {
    return label + ' failed: ' + (e.message || e);
  }
}

function runPoBriefNow()    { return runNow_('Daily PO brief',      _sendPoBrief); }
function runRemindersNow()  { return runNow_('Follow-up reminders', _processReminders); }
function runPosReminderNow(){ return runNow_('POS reminder',        _sendPosLoggingReminder); }
function runAutoImportNow() { return runNow_('Auto-import',         _autoImportOrders); }
function runDisdReminderNow(){ return runNow_('DISD recap',         sendDisdWeeklyReminder_); }

// The 3pm sweep is three independent jobs — report each so a partial result is visible.
function runDailyAutomationNow() {
  return [
    'Confirmations: '  + runNow_('Confirmations',  _sendDayBeforeConfirmations),
    'Missing-PO: '     + runNow_('Missing-PO',     _sendPoAlerts),
    'Year-end review: '+ runNow_('Year-end review', _sendYearEndTaxReview)
  ].join('\n');
}

// Both urgent phases are one switch in the UI, so run both and report both.
function runPoWarningsNow() {
  return [
    'Day-before: ' + runNow_('Day-before warning', _sendPoDayBefore),
    'Day-of: '     + runNow_('Day-of warning',     _sendPoDayOf)
  ].join('\n');
}


// ── INTERNAL DIGEST EMAILS ───────────────────────────────────
// One look for every internal automation digest: CFA header bar, intro line, a table,
// a footer note. Email clients strip <style>, so everything is inline and table-based.
// Customer-facing mail is NOT built here — quotes, confirmations and follow-ups stay
// plain text driven by the operator-editable Settings templates.
const DIGEST_ACCENT = '#E51636';

// Wrap a pre-built HTML string so buildDigestHtml_ won't escape it (e.g. a link).
function digestRaw_(html) { return { __raw: html }; }
// esc() returns '' for ANY falsy input, so esc(0) and esc(false) come back blank. Stringify
// first, otherwise a numeric column (a count, a zero total) silently renders as an empty cell.
function digestCell_(v) {
  if (v && v.__raw !== undefined) return v.__raw;
  return (v == null) ? '' : esc(v.toString());
}
function digestPlain_(v) { return (v && v.__raw !== undefined) ? v.__raw.replace(/<[^>]*>/g, '') : (v == null ? '' : v.toString()); }

// o = { title, intro, columns:[], rows:[[cell,…]], align:['left'|'right',…],
//       nowrap:[bool,…], footer, accentText }
// nowrap keeps short columns (quote ids, dates, money) on one line so only the free-text
// columns wrap — otherwise a quote id like Q-2026-0273-GM breaks across two lines.
function buildDigestHtml_(o) {
  var align = o.align || [], nowrap = o.nowrap || [];
  function ta(i) { return align[i] === 'right' ? 'right' : 'left'; }
  function nw(i) { return (nowrap[i] || align[i] === 'right') ? 'white-space:nowrap;' : ''; }

  var head = (o.columns || []).map(function(c, i) {
    return '<th style="padding:9px 12px;text-align:' + ta(i) + ';font-family:Arial,Helvetica,sans-serif;'
      + 'font-size:11px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:#6B7280;'
      + 'border-bottom:2px solid #E5E7EB;white-space:nowrap;">' + esc(c) + '</th>';
  }).join('');

  var body = (o.rows || []).map(function(r, ri) {
    return '<tr style="background:' + (ri % 2 ? '#FAFAFA' : '#FFFFFF') + ';">' + r.map(function(cell, i) {
      return '<td style="padding:10px 12px;text-align:' + ta(i) + ';' + nw(i) + 'font-family:Arial,Helvetica,sans-serif;'
        + 'font-size:14px;color:#1F2937;border-bottom:1px solid #EEF0F2;">' + digestCell_(cell) + '</td>';
    }).join('') + '</tr>';
  }).join('');

  return '<div style="background:#F3F4F6;padding:20px 12px;font-family:Arial,Helvetica,sans-serif;">'
    + '<div style="max-width:720px;margin:0 auto;background:#FFFFFF;border-radius:10px;overflow:hidden;border:1px solid #E5E7EB;">'
    + '<div style="background:' + DIGEST_ACCENT + ';padding:16px 20px;">'
    +   '<div style="color:#FFFFFF;font-size:17px;font-weight:700;">' + esc(o.title || '') + '</div>'
    +   (o.accentText ? '<div style="color:#FFFFFF;opacity:.9;font-size:13px;margin-top:3px;">' + esc(o.accentText) + '</div>' : '')
    + '</div>'
    + '<div style="padding:18px 20px;">'
    +   (o.intro ? '<p style="margin:0 0 14px;font-size:14px;line-height:1.55;color:#374151;">' + esc(o.intro) + '</p>' : '')
    +   '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">'
    +     '<thead><tr>' + head + '</tr></thead><tbody>' + body + '</tbody></table>'
    +   (o.footer ? '<p style="margin:16px 0 0;font-size:12.5px;line-height:1.55;color:#6B7280;">' + esc(o.footer) + '</p>' : '')
    + '</div></div>'
    + '<div style="max-width:720px;margin:10px auto 0;text-align:center;font-size:11px;color:#9CA3AF;">Sent by the Catering Tool</div>'
    + '</div>';
}

// Plain-text twin of the same data, for clients that won't render HTML.
function buildDigestText_(o) {
  var lines = [];
  if (o.intro) lines.push(o.intro, '');
  (o.rows || []).forEach(function(r) {
    lines.push('• ' + r.map(digestPlain_).filter(function(v) { return v !== ''; }).join(' — '));
  });
  if (o.footer) { lines.push('', o.footer); }
  return lines.join('\n');
}

// Send an internal digest as HTML with a plain-text fallback.
function sendDigest_(to, subject, o) {
  MailApp.sendEmail(to, subject, buildDigestText_(o), {
    htmlBody: buildDigestHtml_(o),
    name: 'Catering Tool'
  });
}

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
  try { alertTo = poAlertRecipients_(getSettings()); } catch(e) { /* no recipient resolved — alerts below simply won't send */ }
  try { _sendDayBeforeConfirmations(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Day-before confirmations failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* can't email about a failed email — give up quietly */ }
  }
  try { _sendPoAlerts(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Missing-PO sweep failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* can't email about a failed email — give up quietly */ }
  }
  try { _sendYearEndTaxReview(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Year-end tax review failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* can't email about a failed email — give up quietly */ }
  }
}

function _sendDayBeforeConfirmations() {
  var settings = getSettings();
  if (!isSettingOn_(settings['Confirmation Enabled'])) return 'Confirmations are disabled.';
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
    if (q.status === 'Inactive') return; // deactivated orders don't get confirmations
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
  if (!isSettingOn_(settings['PO Alert Enabled'])) return 'PO alerts are disabled.';
  // The 6am brief covers this same ground with a NEXT-N-DAYS section, so running both
  // would report the same orders twice a day — exactly the noise the brief exists to kill.
  // Gated on the trigger rather than deleted outright: if the brief was never enabled,
  // this still runs and PO alerts don't silently stop.
  try { if (getPoBriefTriggerStatus()) return 'Superseded by the daily PO brief.'; } catch(e) { /* fall through and run */ }
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
    if (q.status === 'Inactive') return false; // deactivated orders don't need a PO chase
    if (!q.eventDate || q.eventDate < todayKey || q.eventDate > maxKey) return false;
    if ((q.poNumber || '').toString().trim() !== '') return false; // HAVE PO and NO_PO_NEEDED both count as handled
    return !alerted[q.quoteId];
  });
  if (!missing.length) return 'No missing-PO quotes in the next ' + days + ' days.';

  var to = poAlertRecipients_(settings);
  if (!to) return 'No PO alert recipient configured.';

  var rows = missing.map(function(q) {
    var daysOut = Math.round((new Date(q.eventDate + 'T12:00:00') - new Date(todayKey + 'T12:00:00')) / 86400000);
    var due = daysOut === 0 ? 'TODAY' : daysOut + ' day' + (daysOut === 1 ? '' : 's');
    return [
      q.quoteId,
      q.customerName,
      (q.contactName || '') + (q.customerPhone ? ' · ' + q.customerPhone : ''),
      formatWhen_(q.eventDate, q.eventTime, tz),
      '$' + (parseFloat(q.total) || 0).toFixed(2),
      daysOut === 0 ? digestRaw_('<b style="color:' + DIGEST_ACCENT + ';">TODAY</b>') : due
    ];
  });
  var subj = '🔴 ' + missing.length + ' catering quote' + (missing.length === 1 ? '' : 's') + ' missing a PO (next ' + days + ' days)';
  sendDigest_(to, subj, {
    title: missing.length + ' order' + (missing.length === 1 ? '' : 's') + ' missing a PO',
    accentText: 'Next ' + days + ' days',
    intro: 'These upcoming orders have no PO on file yet.',
    columns: ['Quote', 'Customer', 'Contact', 'When', 'Total', 'Due in'],
    align: ['left', 'left', 'left', 'left', 'right', 'right'],
    nowrap: [true, false, false, true, true, true],
    rows: rows,
    footer: 'Add the PO (or mark "No PO Needed") from the quote\'s popup in the catering tool. Each quote is alerted once.'
  });
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


// ── URGENT PO WARNINGS: day-before (2pm) + day-of (6am) ──────
// Separate from the 7-day _sendPoAlerts sweep and its log, so an order that
// already got the early heads-up still gets these urgent nudges. Each phase
// dedups on quoteId|eventDate in its own log tab, so a re-quote for a new date
// can warn again but the same order is never warned twice in one phase.

function sendPoWarning_(targetKey, phaseLabel, logTab) {
  var settings = getSettings();
  if (!isSettingOn_(settings['PO Alert Enabled'])) return 'PO alerts are disabled.';
  // Same reason as _sendPoAlerts: the brief's TODAY/TOMORROW sections are these two warnings.
  // setupPoBriefTrigger deletes these triggers, but a stale one left behind stays quiet too.
  try { if (getPoBriefTriggerStatus()) return 'Superseded by the daily PO brief.'; } catch(e) { /* fall through and run */ }
  var tz = Session.getScriptTimeZone();

  var logSheet = initLogSheet_(logTab, ['Quote ID', 'Event Date', 'Sent At']);
  var sentKeys = {};
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
      sentKeys[r[0] + '|' + normalizeEventDate_(r[1], tz)] = true;
    });
  }

  var missing = getQuotes().filter(function(q) {
    if (q.status === 'Inactive') return false;
    if (q.eventDate !== targetKey) return false;
    if ((q.poNumber || '').toString().trim() !== '') return false; // HAVE PO / NO_PO_NEEDED both count as handled
    return !sentKeys[q.quoteId + '|' + q.eventDate];
  });
  if (!missing.length) return 'No missing-PO orders ' + phaseLabel + '.';

  var to = poAlertRecipients_(settings);
  if (!to) return 'No PO alert recipient configured.';

  var rows = missing.map(function(q) {
    return [
      q.quoteId,
      q.customerName,
      (q.contactName || '') + (q.customerPhone ? ' · ' + q.customerPhone : ''),
      formatWhen_(q.eventDate, q.eventTime, tz),
      '$' + (parseFloat(q.total) || 0).toFixed(2)
    ];
  });
  var subj = '🔴 PO needed ' + phaseLabel + ' — ' + missing.length + ' order' + (missing.length === 1 ? '' : 's');
  sendDigest_(to, subj, {
    title: 'PO needed ' + phaseLabel,
    accentText: missing.length + ' order' + (missing.length === 1 ? '' : 's') + ' still unresolved',
    intro: 'These orders ' + phaseLabel + ' still have NO PO on file.',
    columns: ['Quote', 'Customer', 'Contact', 'When', 'Total'],
    align: ['left', 'left', 'left', 'left', 'right'],
    nowrap: [true, false, false, true, true],
    rows: rows,
    footer: 'Add the PO (or mark "No PO Needed") from the quote\'s popup in the catering tool.'
  });
  missing.forEach(function(q) { logSheet.appendRow([q.quoteId, q.eventDate, new Date()]); });
  SpreadsheetApp.flush();
  return 'Warned on ' + missing.length + ' order(s) ' + phaseLabel + '.';
}

// ── DAILY PO BRIEF (6am) — replaces the 6am/2pm/3pm PO emails ────────────────
// Three separate emails a day, all about POs, all to the same person, differing only in
// how far out the order was. By the third look-alike 🔴 subject line you stop reading
// them — and the one that's genuinely unrecoverable if missed (due TODAY) looks exactly
// like the other two. One message, sections ordered by urgency, sent at the hour the most
// urgent section still matters.
//
// No dedup log, deliberately: this is a live snapshot of what still needs a PO, not an
// event notification. An order that stays unresolved SHOULD keep appearing — that's a
// to-do list. The volume control isn't dedup, it's the send rule below.
function poBriefSection_(label, list, tz) {
  if (!list.length) return [];
  var rows = [[digestRaw_('<b style="color:' + DIGEST_ACCENT + ';letter-spacing:.05em;">' + esc(label) + ' — ' + list.length + '</b>'), '', '', '', '']];
  list.forEach(function(q) {
    rows.push([
      q.quoteId,
      q.customerName,
      (q.contactName || '') + (q.customerPhone ? ' · ' + q.customerPhone : ''),
      formatWhen_(q.eventDate, q.eventTime, tz),
      '$' + (parseFloat(q.total) || 0).toFixed(2)
    ]);
  });
  return rows;
}

function _sendPoBrief() {
  var settings = getSettings();
  if (!isSettingOn_(settings['PO Alert Enabled'])) return 'PO alerts are disabled.';
  var days = parseInt(settings['PO Alert Days Before'], 10);
  if (isNaN(days) || days < 1) days = 7;
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var d1 = new Date(); d1.setDate(d1.getDate() + 1);
  var tomorrowKey = Utilities.formatDate(d1, tz, 'yyyy-MM-dd');
  var dMax = new Date(); dMax.setDate(dMax.getDate() + days);
  var maxKey = Utilities.formatDate(dMax, tz, 'yyyy-MM-dd');

  var todayList = [], tomorrowList = [], laterList = [];
  getQuotes().forEach(function(q) {
    if (q.status === 'Inactive') return;
    if (!q.eventDate || q.eventDate < todayKey || q.eventDate > maxKey) return;
    if ((q.poNumber || '').toString().trim() !== '') return; // HAVE PO / NO PO NEEDED both handled
    if (q.eventDate === todayKey) todayList.push(q);
    else if (q.eventDate === tomorrowKey) tomorrowList.push(q);
    else laterList.push(q);
  });
  function byTime(a, b) { return (a.eventTime || '99:99').localeCompare(b.eventTime || '99:99'); }
  todayList.sort(byTime); tomorrowList.sort(byTime);
  laterList.sort(function(a, b) { return (a.eventDate + (a.eventTime || '')).localeCompare(b.eventDate + (b.eventTime || '')); });

  // THE noise control: nothing time-critical means no email at all. Planning-horizon items
  // aren't lost — they ride along the next time something IS due, and the weekly POS
  // reminder already lists the coming week.
  if (!todayList.length && !tomorrowList.length) {
    return laterList.length
      ? 'Nothing urgent. ' + laterList.length + ' order(s) further out still need a PO — held until something is due.'
      : 'Nothing needs a PO.';
  }

  var to = poAlertRecipients_(settings);
  if (!to) return 'No PO alert recipient configured.';

  var rows = poBriefSection_('DUE TODAY', todayList, tz)
    .concat(poBriefSection_('DUE TOMORROW', tomorrowList, tz))
    .concat(poBriefSection_('NEXT ' + days + ' DAYS', laterList, tz));

  var subj = todayList.length
    ? '🔴 ' + todayList.length + ' order' + (todayList.length === 1 ? '' : 's') + ' TODAY still need' + (todayList.length === 1 ? 's' : '') + ' a PO'
    : '🟠 ' + tomorrowList.length + ' order' + (tomorrowList.length === 1 ? '' : 's') + ' tomorrow still need' + (tomorrowList.length === 1 ? 's' : '') + ' a PO';

  sendDigest_(to, subj, {
    title: 'Catering PO brief',
    accentText: todayList.length + ' today · ' + tomorrowList.length + ' tomorrow · ' + laterList.length + ' further out',
    intro: 'Everything below is still missing a PO. Top section first — those are running today.',
    columns: ['Quote', 'Customer', 'Contact', 'When', 'Total'],
    align: ['left', 'left', 'left', 'left', 'right'],
    nowrap: [true, false, false, true, true],
    rows: rows,
    footer: 'Add the PO (or mark "No PO Needed") from the quote\'s popup. This is the only PO email — you get it at 6am on days something is due today or tomorrow, and not at all otherwise.'
  });
  return 'Brief sent: ' + todayList.length + ' today, ' + tomorrowList.length + ' tomorrow, ' + laterList.length + ' further out.';
}

// Trigger wrapper — alerts on failure, matching the other scheduled jobs.
function dailyPoBrief() {
  var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e) { /* no recipient */ }
  try { _sendPoBrief(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Daily PO brief failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* give up quietly */ }
  }
}

// Installing the brief retires the two urgent-warning triggers it replaces. The 3pm
// 7-day sweep isn't a trigger of its own (it rides dailyCateringAutomation), so it's
// skipped in code instead — see the guard in _sendPoAlerts.
function setupPoBriefTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var f = t.getHandlerFunction();
    if (f === 'dailyPoBrief' || f === 'poWarningDayBefore' || f === 'poWarningDayOf') { ScriptApp.deleteTrigger(t); if (f !== 'dailyPoBrief') removed++; }
  });
  ScriptApp.newTrigger('dailyPoBrief').timeBased().everyDays(1).atHour(6).create();
  return 'Daily PO brief enabled (6am).' + (removed ? ' Retired ' + removed + ' old PO warning trigger(s) it replaces.' : '');
}
function removePoBriefTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'dailyPoBrief') { ScriptApp.deleteTrigger(t); removed++; } });
  return removed > 0 ? 'Daily PO brief trigger removed.' : 'No active PO brief trigger found.';
}
function getPoBriefTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'dailyPoBrief'; });
}

function _sendPoDayBefore() {
  var tz = Session.getScriptTimeZone();
  var d = new Date(); d.setDate(d.getDate() + 1);
  return sendPoWarning_(Utilities.formatDate(d, tz, 'yyyy-MM-dd'), 'due tomorrow', TAB_PO_WARN_DAYBEFORE);
}
function _sendPoDayOf() {
  var tz = Session.getScriptTimeZone();
  return sendPoWarning_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), 'due TODAY', TAB_PO_WARN_DAYOF);
}

// Trigger wrappers — alert on failure, matching dailyCateringAutomation's pattern.
function poWarningDayBefore() {
  var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e) { /* no recipient */ }
  try { _sendPoDayBefore(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Day-before PO warning failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* give up quietly */ }
  }
}
function poWarningDayOf() {
  var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e) { /* no recipient */ }
  try { _sendPoDayOf(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'Day-of PO warning failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* give up quietly */ }
  }
}

// Both urgent-warning triggers are managed as one unit (one Settings switch).
function setupPoWarningTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var f = t.getHandlerFunction();
    if (f === 'poWarningDayBefore' || f === 'poWarningDayOf') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('poWarningDayBefore').timeBased().everyDays(1).atHour(14).create();
  ScriptApp.newTrigger('poWarningDayOf').timeBased().everyDays(1).atHour(6).create();
  return 'Urgent PO warnings enabled (day-before 2pm + day-of 6am).';
}
function removePoWarningTriggers() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var f = t.getHandlerFunction();
    if (f === 'poWarningDayBefore' || f === 'poWarningDayOf') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed > 0 ? 'Urgent PO warning triggers removed.' : 'No active PO warning triggers found.';
}
function getPoWarningTriggerStatus() {
  var has = {};
  ScriptApp.getProjectTriggers().forEach(function(t) { has[t.getHandlerFunction()] = true; });
  return !!(has['poWarningDayBefore'] && has['poWarningDayOf']);
}


// ── WEEKLY POS-LOGGING REMINDER (Fridays at 12pm) ────────────
// Reminds the team to log the upcoming week's catering orders into the POS.
// Editable message + recipients in Settings; lists the next 7 days of active orders.
function _sendPosLoggingReminder() {
  var settings = getSettings();
  if (!isSettingOn_(settings['POS Reminder Enabled'])) return 'POS reminder is disabled.';
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var end = new Date(); end.setDate(end.getDate() + 7);
  var endKey = Utilities.formatDate(end, tz, 'yyyy-MM-dd');
  var to = (settings['POS Reminder Email'] || '').toString().trim() || poAlertRecipients_(settings);
  if (!to) return 'No POS reminder recipient configured.';

  var upcoming = getQuotes().filter(function(q) {
    return q.status !== 'Inactive' && q.eventDate && q.eventDate >= todayKey && q.eventDate <= endKey;
  });
  upcoming.sort(function(a, b) { return (a.eventDate + (a.eventTime || '')).localeCompare(b.eventDate + (b.eventTime || '')); });

  var msg = (settings['POS Reminder Message'] || 'Reminder: please log all catering orders for the upcoming week into our POS system.').toString();
  var rows = upcoming.map(function(q) {
    return [
      storeCodeForQuote_(q.locationName, settings),
      formatWhen_(q.eventDate, q.eventTime, tz),
      q.customerName,
      '$' + (parseFloat(q.total) || 0).toFixed(2)
    ];
  });
  // Nothing next week means there is nothing to log — an email saying "None scheduled" is
  // pure noise, and noise is what teaches you to stop reading these.
  if (!rows.length) return 'No orders in the next 7 days — POS reminder skipped.';
  sendDigest_(to, '🧾 Log catering orders into POS — upcoming week', {
    title: 'Log catering orders into the POS',
    accentText: upcoming.length + ' order' + (upcoming.length === 1 ? '' : 's') + ' in the next 7 days',
    intro: msg,
    columns: ['Store', 'When', 'Customer', 'Total'],
    align: ['left', 'left', 'left', 'right'],
    nowrap: [true, true, false, true],
    rows: rows
  });
  return 'POS reminder sent to ' + to + '.';
}

function posLoggingReminder() {
  var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e) { /* no recipient */ }
  try { _sendPosLoggingReminder(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'POS reminder failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* give up quietly */ }
  }
}

function setupPosReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'posLoggingReminder') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('posLoggingReminder').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(12).create();
  return 'POS reminder enabled (Fridays at 12pm).';
}
function removePosReminderTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'posLoggingReminder') { ScriptApp.deleteTrigger(t); removed++; } });
  return removed > 0 ? 'POS reminder trigger removed.' : 'No active POS reminder trigger found.';
}
function getPosReminderTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'posLoggingReminder'; });
}


// ── INCOMING ORDER EMAIL IMPORT (manual, human-in-the-loop) ──
// Scans Gmail for the CFA "Incoming Catering Order" emails, parses each, and
// hands the parsed order to the client to PRE-FILL the New Quote form (never
// auto-creates). Imported message IDs are logged so orders don't re-appear.
const TAB_IMPORTED = 'Imported_Orders';

// "Friday 7/24/2026 at 7:45am" -> { date:'2026-07-24', time:'07:45' }
function parseWhen_(str) {
  var out = { date: '', time: '' };
  var d = (str || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (d) { out.date = d[3] + '-' + ('0' + d[1]).slice(-2) + '-' + ('0' + d[2]).slice(-2); }
  var t = (str || '').match(/(\d{1,2}):(\d{2})\s*([ap]m)/i);
  if (t) {
    var h = parseInt(t[1], 10), mn = t[2], ap = t[3].toLowerCase();
    if (ap === 'pm' && h < 12) h += 12;
    if (ap === 'am' && h === 12) h = 0;
    out.time = ('0' + h).slice(-2) + ':' + mn;
  }
  return out;
}

// Best-effort parse of one order email into a draft order. Scope is deliberately
// light (readable info to pre-fill, not a perfectly-priced quote). Uses whole-body
// regex scans for the key fields so it tolerates however Gmail flattens the HTML.
function parseIncomingOrderEmail_(message) {
  var body = (message.getPlainBody() || '').replace(/\r/g, '');
  var subject = message.getSubject() || '';
  var lines = body.split('\n');
  var t = lines.map(function(l) { return l.trim(); });
  function idxOf(re) { for (var i = 0; i < t.length; i++) { if (re.test(t[i])) return i; } return -1; }

  var isDelivery = /delivery/i.test(subject) || /Catering Delivery Order|Delivery Time|Delivery Address/i.test(body);
  var orderType = isDelivery ? 'Delivery' : 'Pickup';

  // Date + time: scan the ~150 chars right after the "Pickup/Delivery Time" label; fall back to whole body.
  var whenChunk = body;
  var wm = body.match(/(?:Pickup|Delivery)\s*Time[\s\S]{0,150}/i);
  if (wm) whenChunk = wm[0];
  var dt = parseWhen_(whenChunk);
  if (!dt.date || !dt.time) { var dt2 = parseWhen_(body); dt.date = dt.date || dt2.date; dt.time = dt.time || dt2.time; }

  // Email: first address that isn't the CFA sender / store domain.
  var email = '';
  var emails = body.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || [];
  for (var e = 0; e < emails.length; e++) { if (!/chick-fil-a\.com$/i.test(emails[e])) { email = emails[e]; break; } }

  // Phone: first phone-like run of digits.
  var phone = '';
  var pm = body.match(/\+?\d[\d().\-\s]{8,}\d/);
  if (pm) phone = pm[0].replace(/\s{2,}/g, ' ').trim();

  // Name: first sensible line after "Customer Information".
  var name = '';
  var custIdx = idxOf(/Customer Information/i);
  if (custIdx >= 0) {
    for (var i = custIdx + 1; i < Math.min(custIdx + 9, t.length); i++) {
      var l = t[i];
      if (!l) continue;
      if (/@/.test(l) || /^\+?\d[\d().\-\s]{6,}$/.test(l) || /^(Guest Count|Paper Goods|Item Name)/i.test(l)) continue;
      name = l; break;
    }
  }

  // Delivery address: everything between the "Delivery Address" label and
  // "Customer Information" (which the email jams onto the last address line).
  var address = '';
  if (isDelivery) {
    var addrM = body.match(/Delivery Address\*?\s*([\s\S]*?)\*?\s*Customer Information/i);
    if (addrM) {
      address = addrM[1].replace(/\*/g, ' ').replace(/[\r\n]+/g, ', ')
        .replace(/\s+,/g, ',').replace(/,\s*,/g, ',').replace(/\s{2,}/g, ' ')
        .replace(/^[,\s]+|[,\s]+$/g, '').trim();
    }
  }

  // Labels are wrapped in asterisks in the plain body: "*Guest Count:*  15".
  var guest = '', paper = '';
  var gm = body.match(/Guest Count:?\*?\s*([0-9]+)/i); if (gm) guest = gm[1];
  var pm2 = body.match(/Paper Goods:?\*?\s*([A-Za-z]+)/i); if (pm2) paper = pm2[1];

  // Items: between "Item Name" header and "Subtotal". Handle BOTH layouts —
  // "Name  10  $3.89" on one line, OR a name line followed by "10  $3.89".
  // A line with qty but NO price is a condiment modifier (folded into notes).
  var items = [], modifiers = [];
  var itemHdr = idxOf(/Item Name/i), subIdx = idxOf(/^\s*Subtotal/i);
  if (itemHdr >= 0) {
    var stop = subIdx > itemHdr ? subIdx : t.length;
    for (var k = itemHdr + 1; k < stop; k++) {
      var raw = lines[k], tl = t[k]; if (!tl) continue;
      var indented = /^(\s{2,}|\t)/.test(raw);
      var one = tl.match(/^(.+?)[\s\t]+(\d+(?:\.\d+)?)[\s\t]+\$([\d,]+\.\d{2})\s*$/); // Name qty $price
      if (one && !indented) { items.push({ description: one[1].trim(), quantity: parseFloat(one[2]), price: parseFloat(one[3].replace(/,/g, '')) }); continue; }
      var next = t[k + 1] || '';
      var two = next.match(/^(\d+(?:\.\d+)?)[\s\t]*\$([\d,]+\.\d{2})\s*$/); // name line, then "qty $price"
      if (two && !indented && !/\$/.test(tl) && !/^\d/.test(tl)) { items.push({ description: tl, quantity: parseFloat(two[1]), price: parseFloat(two[2].replace(/,/g, '')) }); k++; continue; }
      var modM = tl.match(/^(.+?)[\s\t]+\d+\s*$/); // "Grape Jam   5" — modifier
      if (indented || (modM && !/\$/.test(tl))) { modifiers.push((modM ? modM[1] : tl).replace(/\t.*$/, '').replace(/[\s\t]+\d+(\s*\$[\d.]+)?\s*$/, '').trim()); }
    }
  }

  function money_(re) { var m = body.match(re); return m ? m[1].replace(/,/g, '') : ''; }
  var total = money_(/\bTotal[\s\t]*\$?([\d,]+\.\d{2})/i);
  var subtotal = money_(/\bSubtotal[\s\t]*\$?([\d,]+\.\d{2})/i);
  var tax = money_(/\bTax[\s\t]*\$?([\d,]+\.\d{2})/i);

  return {
    orderType: orderType, date: dt.date, time: dt.time,
    customerName: name, phone: phone, email: email, address: address,
    guestCount: guest, paperGoods: paper,
    items: items, modifiers: modifiers,
    subtotal: subtotal, tax: tax, total: total,
    receivedAt: message.getDate() ? message.getDate().toISOString() : ''
  };
}

// DIAGNOSTIC — run from the editor and read the log to see exactly what the parser
// receives, so the field patterns can be tuned to the real email format.
function debugIncomingRaw() {
  var settings = getSettings();
  var sender = (settings['Incoming Order Sender'] || 'one@email.chick-fil-a.com').toString().trim();
  var subj = (settings['Incoming Order Subject'] || 'Incoming Catering Order').toString().trim();
  var th = GmailApp.search('from:(' + sender + ') subject:("' + subj + '") newer_than:60d', 0, 1);
  if (!th.length) { Logger.log('No matching emails found.'); return 'No matching emails found.'; }
  var m = th[0].getMessages()[0];
  var body = m.getPlainBody() || '';
  Logger.log('SUBJECT: ' + m.getSubject());
  Logger.log('--- PLAIN BODY START ---\n' + body + '\n--- PLAIN BODY END ---');
  Logger.log('--- PARSED ---\n' + JSON.stringify(parseIncomingOrderEmail_(m), null, 2));
  return { subject: m.getSubject(), body: body, parsed: parseIncomingOrderEmail_(m) };
}

// Scan Gmail for not-yet-imported order emails. Returns [{messageId, subject, ...parsed}].
function scanIncomingOrders() {
  var settings = getSettings();
  var sender = (settings['Incoming Order Sender'] || 'one@email.chick-fil-a.com').toString().trim();
  var subjectPhrase = (settings['Incoming Order Subject'] || 'Incoming Catering Order').toString().trim();
  var log = initLogSheet_(TAB_IMPORTED, ['Message ID', 'Imported At', 'Quote ID']);
  var done = {};
  if (log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 1).getValues().forEach(function(r) { if (r[0]) done[r[0].toString().trim()] = true; });

  var query = 'from:(' + sender + ') subject:("' + subjectPhrase + '") newer_than:60d';
  var threads = [];
  try { threads = GmailApp.search(query, 0, 40); } catch(e) { return { error: e.message || String(e), orders: [] }; }
  var out = [];
  threads.forEach(function(th) {
    th.getMessages().forEach(function(m) {
      if ((m.getSubject() || '').toLowerCase().indexOf(subjectPhrase.toLowerCase()) < 0) return;
      var id = m.getId();
      if (done[id]) return;
      try {
        var parsed = parseIncomingOrderEmail_(m);
        parsed.messageId = id; parsed.subject = m.getSubject() || '';
        out.push(parsed);
      } catch(e) { /* skip an unparseable email rather than fail the whole scan */ }
    });
  });
  out.sort(function(a, b) { return (a.date + a.time).localeCompare(b.date + b.time); });
  return { error: '', orders: out };
}

// Mark one email as imported so it won't show again (called after the human saves the quote).
function markIncomingImported(messageId, quoteId) {
  if (!messageId) return false;
  initLogSheet_(TAB_IMPORTED, ['Message ID', 'Imported At', 'Quote ID']).appendRow([messageId, new Date(), quoteId || '']);
  return true;
}


// ── AUTO-IMPORT SWEEP (daily) ────────────────────────────────
// Parses new order emails and auto-creates the quote + calendar entry, flagged 📧.
// VALIDATES first (date, time, name, ≥1 item); a thin parse is skipped + alerted,
// never dropped on the calendar as garbage. Dedups on Gmail message-id.

// Turn a parsed order into a saveQuote() payload (viaEmail flag → 📧 in the title).
// An order email only ever names a PERSON — the parser reads "Customer Information",
// which is a human, not an organization. That left the Customer/Organization field holding
// a person ("Sylvia Perez" instead of "Dallas ISD"), which quietly broke three things:
// duplicate detection couldn't match the order against hand-built quotes, the Customer
// Database split the person off from their org, and the DISD recap stopped seeing them.
//
// Recover the organization the way a human would: find the most recent past quote where
// this same person was the CONTACT, and reuse that quote's organization. Returns '' when
// there's no usable match, and never returns the person's own name back — otherwise a
// previous bad auto-import would teach the next one the same wrong answer.
function orgNameForContact_(personName, existingQuotes) {
  var person = (personName || '').toString().toLowerCase().trim();
  if (!person) return '';
  var best = null;
  (existingQuotes || []).forEach(function(q) {
    if ((q.contactName || '').toString().toLowerCase().trim() !== person) return;
    var org = (q.customerName || '').toString().trim();
    if (!org || org.toLowerCase() === person) return; // that quote names a person too — no help
    if (!best || new Date(q.createdDate) > new Date(best.createdDate)) best = q;
  });
  return best ? (best.customerName || '').toString().trim() : '';
}

// Were this organization's past orders tax exempt? Used ONLY to add a "check this" note.
// It deliberately does NOT set the flag: wrongly marking an order exempt is silent and
// costs real money, while wrongly taxing it is visible and gets fixed. Same asymmetry
// that keeps PO status off auto-inference.
function pastOrdersWereTaxExempt_(orgName, existingQuotes) {
  var org = (orgName || '').toString().toLowerCase().trim();
  if (!org) return false;
  return (existingQuotes || []).some(function(q) {
    return (q.customerName || '').toString().toLowerCase().trim() === org
      && (q.taxExempt === 'TRUE' || q.taxExempt === true);
  });
}

// existingQuotes is the sweep's pre-read getQuotes() snapshot — passed in so the org
// lookup costs no extra sheet reads.
function buildQuoteFromParsed_(o, existingQuotes) {
  var settings = getSettings();
  var storeName = settings['Store Name (Active)'] || settings['Location 1 Name'] || '';
  var person = (o.customerName || '').toString().trim();
  var org = orgNameForContact_(person, existingQuotes) || person;
  var notes = ['Added automatically via email — please double-check.'];
  if (org && org.toLowerCase() !== person.toLowerCase()) notes.push('Organization "' + org + '" filled in from ' + person + '\'s past orders.');
  if (pastOrdersWereTaxExempt_(org, existingQuotes)) notes.push('⚠ Past orders for ' + org + ' were TAX EXEMPT — verify the tax on this one.');
  if (o.guestCount) notes.push('Guest Count: ' + o.guestCount);
  if (o.paperGoods) notes.push('Paper Goods: ' + o.paperGoods);
  if (o.modifiers && o.modifiers.length) notes.push('Add-ons: ' + o.modifiers.join(', '));
  var lineItems = (o.items || []).map(function(it) {
    var price = parseFloat(it.price) || 0, qty = parseFloat(it.quantity) || 0;
    return { description: it.description, quantity: qty, price: price, amount: price * qty, discountValue: 0, discountType: 'percent', discountAmount: 0 };
  });
  var subtotal = parseFloat(o.subtotal) || lineItems.reduce(function(s, it) { return s + it.amount; }, 0);
  var tax = parseFloat(o.tax) || 0, total = parseFloat(o.total) || (subtotal + tax);
  var taxRate = subtotal > 0 ? Math.round((tax / subtotal) * 10000) / 100 : 0;
  return {
    quoteId: '', customerName: org, contactName: person,
    orderType: o.orderType || 'Pickup', deliveryAddress: o.orderType === 'Delivery' ? (o.address || '') : '',
    lineItems: lineItems, subtotal: subtotal, taxRate: taxRate, taxAmount: tax, total: total,
    taxExempt: false, locationName: storeName, customerEmail: o.email || '', customerPhone: o.phone || '',
    poNumber: '', date: o.date || '', time: o.time || '',
    orderDiscountValue: 0, orderDiscountType: 'percent', quoteNotes: notes.join('\n'),
    viaEmail: true,
    // Colour-code email-created orders on the calendar (banana by default) so they read
    // as "the tool made this, check it" at a glance — the 📧 in the title says the same
    // thing but you have to open the event to see it.
    eventColor: resolveCalendarColor_(settings['Email Import Calendar Color'] || 'Banana', settings)
  };
}

// ── DUPLICATE ORDER DETECTION ────────────────────────────────
// The Gmail message-id log stops the same EMAIL being read twice. It does nothing about
// the same ORDER arriving twice (an order plus its confirmation/revision mail), or an
// order somebody already typed in by hand. That's what this catches.
//
// It matches against the QUOTES SHEET, not against Google Calendar, on purpose: the sheet
// is structured and exact, while an event title is a moving target — it starts
// "🔴 VERIFY PO", becomes "🟢 HAVE PO" the moment a PO lands, and grows a -"driver" segment
// when one is assigned. Matching on that would miss real duplicates and invent fake ones.
// The sheet also answers "is it on the calendar" anyway: every quote row carries its
// Calendar Event ID.

// Same fingerprint the History ⚠️ badge uses (quoteSignature_ in App.html) — keep the two
// in step so the automation and the UI never disagree about what a duplicate is.
function orderSignature_(dateStr, timeStr, customerName, total, items) {
  var sig = (items || []).map(function(it) {
    return (it.description || '').toString().trim().toLowerCase() + ':' + (parseFloat(it.quantity) || 0);
  }).sort().join('|');
  return [(dateStr || ''), (timeStr || ''), (customerName || '').toString().toLowerCase().trim(),
          (parseFloat(total) || 0).toFixed(2), sig].join('#');
}

// Does this parsed order already exist? Returns:
//   { match: 'exact', quoteId }  same slot AND same contents — a straight duplicate
//   { match: 'slot',  quoteId }  same customer + date + time, contents differ — likely a revision
//   { match: '' }                nothing like it on the books
// existing is a pre-read getQuotes() snapshot so a 40-email sweep reads the tab once.
function findDuplicateOrder_(parsed, existing) {
  var lineItems = (parsed.items || []).map(function(it) {
    return { description: it.description, quantity: it.quantity };
  });
  var subtotal = parseFloat(parsed.subtotal) || 0;
  var total = parseFloat(parsed.total) || (subtotal + (parseFloat(parsed.tax) || 0));
  var wantExact = orderSignature_(parsed.date, parsed.time, parsed.customerName, total, lineItems);
  var wantSlot = [(parsed.date || ''), (parsed.time || ''), (parsed.customerName || '').toLowerCase().trim()].join('#');
  var slotHit = '';
  for (var i = 0; i < existing.length; i++) {
    var q = existing[i];
    if (q.status === 'Inactive') continue; // a deactivated order isn't occupying that slot
    if (q.eventDate !== parsed.date || q.eventTime !== parsed.time) continue;
    if ((q.customerName || '').toString().toLowerCase().trim() !== (parsed.customerName || '').toLowerCase().trim()) continue;
    var qItems = [];
    try { qItems = JSON.parse(q.lineItems) || []; } catch(e) { qItems = []; }
    if (orderSignature_(q.eventDate, q.eventTime, q.customerName, q.total, qItems) === wantExact) {
      return { match: 'exact', quoteId: q.quoteId };
    }
    if (!slotHit) slotHit = q.quoteId; // same slot, different contents — remember but keep looking for an exact hit
  }
  return slotHit ? { match: 'slot', quoteId: slotHit } : { match: '' };
}

// ── DUPLICATE PAIRS ALREADY ON THE BOOKS ─────────────────────
// findDuplicateOrder_ stops NEW duplicates at the door. It can't help with pairs already
// written before that existed. These power the side-by-side compare + merge in History.

// Every group of 2+ active quotes sharing a signature, newest event first.
// Full field set per quote so the UI can diff them without a second round-trip.
function getDuplicateQuoteGroups() {
  var bySig = {};
  getQuotes().forEach(function(q) {
    if (q.status === 'Inactive') return; // a deactivated copy isn't competing for the slot
    var items = [];
    try { items = JSON.parse(q.lineItems) || []; } catch(e) { items = []; }
    var sig = orderSignature_(q.eventDate, q.eventTime, q.customerName, q.total, items);
    if (!bySig[sig]) bySig[sig] = [];
    bySig[sig].push({
      quoteId: q.quoteId, sheetRow: q.sheetRow, createdDate: q.createdDate,
      customerName: q.customerName, contactName: q.contactName,
      customerEmail: q.customerEmail, customerPhone: q.customerPhone,
      orderType: q.orderType, deliveryAddress: q.deliveryAddress,
      eventDate: q.eventDate, eventTime: q.eventTime, total: q.total,
      poNumber: q.poNumber, locationName: q.locationName, driver: q.driver,
      handoffStore: q.handoffStore, quoteNotes: q.quoteNotes,
      lineItems: q.lineItems, calendarEventId: q.calendarEventId
    });
  });
  var groups = [];
  Object.keys(bySig).forEach(function(sig) {
    if (bySig[sig].length < 2) return;
    // Oldest first — the original is almost always the one to keep.
    bySig[sig].sort(function(a, b) { return new Date(a.createdDate) - new Date(b.createdDate); });
    groups.push({ signature: sig, quotes: bySig[sig] });
  });
  groups.sort(function(a, b) {
    return (b.quotes[0].eventDate || '').localeCompare(a.quotes[0].eventDate || '');
  });
  return { groups: groups };
}

// Fields a merge can rescue from the copy being removed. If the keeper's value is blank
// and the other one has something, the keeper inherits it — so a PO typed on the "wrong"
// copy isn't thrown away just because you kept the other one.
var MERGE_FILL_FIELDS = ['PO Number', 'Customer Email', 'Customer Phone', 'Delivery Address', 'Quote Notes'];

// Keep one quote, remove the other. NOT a content merge — two duplicates are identical by
// definition; the only thing combined is blank fields (above). Deletes the dropped quote's
// CALENDAR EVENT first: deleteQuote only removes the sheet row, so skipping this would
// leave the duplicate event sitting on the calendar — the exact thing being cleaned up.
function mergeDuplicateQuotes(keepQuoteId, dropQuoteId) {
  var keepId = (keepQuoteId || '').toString().trim();
  var dropId = (dropQuoteId || '').toString().trim();
  if (!keepId || !dropId) return { success: false, message: 'Need both quote ids.' };
  if (keepId === dropId) return { success: false, message: 'Those are the same quote.' };

  var keep = null, drop = null;
  getQuotes().forEach(function(q) {
    if (q.quoteId === keepId) keep = q;
    if (q.quoteId === dropId) drop = q;
  });
  if (!keep || !drop) return { success: false, message: 'One of those quotes is no longer in the sheet — refresh and try again.' };

  // Guard against a mis-click merging two unrelated orders. They must at least occupy the
  // same slot (customer + date + time); identical contents are not required, because a
  // near-duplicate with a corrected item list is exactly what you'd want to reconcile.
  var sameSlot = (keep.eventDate === drop.eventDate)
    && (keep.eventTime === drop.eventTime)
    && ((keep.customerName || '').toString().toLowerCase().trim() === (drop.customerName || '').toString().toLowerCase().trim());
  if (!sameSlot) return { success: false, message: 'Those two are not the same customer, date and time — merge refused.' };

  var out = { success: true, keptQuoteId: keepId, removedQuoteId: dropId, filled: [], calendarRemoved: false, calendarReason: '' };

  // 1. Rescue anything the keeper is missing.
  var sheet = getSpreadsheet().getSheetByName(TAB_QUOTES);
  var C = detectQuoteColumns_(sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]);
  var keepRow = resolveQuoteRow_(keep.sheetRow, keepId);
  if (!keepRow) return { success: false, message: 'The quote you kept is no longer in the sheet — refresh and try again.' };
  var byField = {
    'PO Number': ['poNumber'], 'Customer Email': ['customerEmail'], 'Customer Phone': ['customerPhone'],
    'Delivery Address': ['deliveryAddress'], 'Quote Notes': ['quoteNotes']
  };
  MERGE_FILL_FIELDS.forEach(function(field) {
    var prop = byField[field][0];
    var keepVal = (keep[prop] || '').toString().trim();
    var dropVal = (drop[prop] || '').toString().trim();
    if (keepVal === '' && dropVal !== '') {
      sheet.getRange(keepRow, C[field] + 1).setValue(dropVal);
      out.filled.push(field);
    }
  });

  // 2. Pull the dropped copy's calendar event BEFORE its row goes away — once the row is
  //    gone the event id is gone with it and the event would be orphaned forever.
  if (drop.calendarEventId) {
    var del = deleteCalendarEventById_(drop.calendarEventId);
    if (del === true) out.calendarRemoved = true;
    else out.calendarReason = (typeof del === 'string') ? del : 'Calendar event not removed';
  }

  // 3. Remove the row (re-resolve: step 1 didn't shift rows, but the sheet may have moved
  //    under us between the read and now).
  var dropRow = resolveQuoteRow_(drop.sheetRow, dropId);
  if (!dropRow) return { success: false, message: 'The duplicate is already gone — nothing removed.' };
  sheet.deleteRow(dropRow);

  // 4. Retire its review-queue row so a merged-away order stops asking to be reviewed.
  try { closeEmailReviewFor_(dropId, 'MERGED'); } catch(e) { /* queue hygiene is best-effort */ }

  SpreadsheetApp.flush();
  return out;
}

// Mark any UNREVIEWED review row for a quote as resolved (used when it's merged away).
function closeEmailReviewFor_(quoteId, status) {
  var sheet = getSpreadsheet().getSheetByName(TAB_EMAIL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return;
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < ids.length; i++) {
    if ((ids[i][0] || '').toString().trim() === quoteId && (ids[i][2] || '').toString().trim() === 'UNREVIEWED') {
      sheet.getRange(i + 2, 3, 1, 3).setValues([[status || 'RESOLVED', new Date(), '']]);
    }
  }
}

// Trigger entry point — the wrapper is what's installed, so a throw anywhere in the
// core (getSettings, the Gmail scan) alerts instead of dying silently at 5am.
function autoImportOrders() {
  try {
    return _autoImportOrders();
  } catch (e) {
    var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e2) { /* no recipient */ }
    if (alertTo) {
      try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'autoImportOrders failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e3) { /* give up quietly */ }
    }
    return 'Auto-import failed: ' + (e.message || e);
  }
}

function _autoImportOrders() {
  var settings = getSettings();
  if (!isSettingOn_(settings['Auto Import Enabled'])) return 'Auto-import is disabled.';
  var res = scanIncomingOrders();
  if (res.error) {
    // The scan already failed; if the heads-up email ALSO fails there's nothing left to
    // try, and the caller still gets 'Scan failed.' back (shown by the Run Now button).
    try { MailApp.sendEmail(poAlertRecipients_(settings), 'Catering Auto-Import Error', 'Gmail scan failed:\n\n' + res.error); } catch(e) { /* can't email about a failed email — give up quietly */ }
    return 'Scan failed: ' + res.error;
  }
  var created = 0, skipped = [], dupes = [];
  // Read the Quotes tab ONCE for the whole sweep, then keep the snapshot current as we
  // create rows — otherwise two copies of the same order inside a single run would both
  // pass the check and both get written.
  var existing = getQuotes();
  (res.orders || []).forEach(function(o) {
    if (!o.date || !o.time || !(o.customerName || '').trim() || !(o.items || []).length) {
      skipped.push((o.customerName || '(no name)') + ' — ' + (o.subject || '') + ' [missing: ' + [(!o.date ? 'date' : ''), (!o.time ? 'time' : ''), (!(o.customerName || '').trim() ? 'name' : ''), (!(o.items || []).length ? 'items' : '')].filter(String).join(', ') + ']');
      markIncomingImported(o.messageId, 'SKIPPED'); // don't re-alert daily; human handles via manual import
      return;
    }
    // Never put a second copy of an order on the calendar. An exact repeat is dropped;
    // a same-slot order with different contents is a probable revision, so it's held back
    // for a human rather than silently creating a competing event.
    var dup = findDuplicateOrder_(o, existing);
    if (dup.match) {
      dupes.push((o.customerName || '(no name)') + ' ' + o.date + ' ' + o.time
        + (dup.match === 'exact' ? ' — already on the books as ' + dup.quoteId
                                 : ' — same slot as ' + dup.quoteId + ' but different items/total (possible revision — reconcile by hand)'));
      markIncomingImported(o.messageId, dup.match === 'exact' ? 'DUPLICATE of ' + dup.quoteId : 'REVISION? of ' + dup.quoteId);
      return;
    }
    try {
      var quoteId = saveQuote(buildQuoteFromParsed_(o, existing));
      markIncomingImported(o.messageId, quoteId);
      markEmailImportForReview_(quoteId); // lands live, but flagged for a human to confirm
      // Fold the new order into the snapshot so a twin later in THIS sweep is caught too.
      // Total is derived exactly as findDuplicateOrder_ derives it, so a parsed twin
      // compares equal instead of sliding from an 'exact' match down to a 'slot' one.
      existing.push({ quoteId: quoteId, status: 'Active', eventDate: o.date, eventTime: o.time,
        customerName: o.customerName,
        total: parseFloat(o.total) || ((parseFloat(o.subtotal) || 0) + (parseFloat(o.tax) || 0)),
        lineItems: JSON.stringify((o.items || []).map(function(it) { return { description: it.description, quantity: it.quantity }; })) });
      created++;
    } catch(e) { skipped.push((o.customerName || '(no name)') + ' — save failed: ' + (e.message || e)); }
  });
  // One email covers both piles — a thin parse and a suppressed duplicate are the same
  // kind of news ("the sweep didn't create this; look at it yourself").
  if (skipped.length || dupes.length) {
    var body = '';
    if (skipped.length) {
      body += 'COULD NOT BE READ — create them from "📥 Import from Email" if still needed:\n\n• ' + skipped.join('\n• ') + '\n\n';
    }
    if (dupes.length) {
      body += 'ALREADY ON THE BOOKS — skipped so they didn\'t land on the calendar twice:\n\n• ' + dupes.join('\n• ') + '\n\n'
            + 'Anything marked "possible revision" means an order for that same customer, date and time already exists with different items — check which version is right and edit the existing quote.\n';
    }
    var n = skipped.length + dupes.length;
    try { MailApp.sendEmail(poAlertRecipients_(settings), '⚠️ Catering Auto-Import: ' + n + ' order(s) need manual review', body); }
    catch(e) { /* review list is a courtesy — the counts are in the return value either way */ }
  }
  return 'Auto-imported ' + created + ', skipped ' + skipped.length + ', duplicates blocked ' + dupes.length + '.';
}

function setupAutoImportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'autoImportOrders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoImportOrders').timeBased().everyDays(1).atHour(5).create();
  return 'Auto-import enabled (runs daily at 5am).';
}
function removeAutoImportTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'autoImportOrders') { ScriptApp.deleteTrigger(t); removed++; } });
  return removed > 0 ? 'Auto-import trigger removed.' : 'No active auto-import trigger found.';
}
function getAutoImportTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'autoImportOrders'; });
}


// ── EMAIL IMPORT REVIEW QUEUE ────────────────────────────────
// Auto-import still creates the quote AND its calendar event immediately — an order the
// robot found is a real order, and burying it behind an approval step would hide it from
// the kitchen, runsheets and supplies calc until somebody happened to look. Instead each
// one is flagged UNREVIEWED here until a human eyeballs it, so nothing is ever missing —
// only ever unconfirmed. Only the automatic sweep writes rows: an order a human built
// through "📥 Import from Email" was reviewed by definition, at the form.
const TAB_EMAIL_REVIEW = 'Email_Import_Review';
const EMAIL_REVIEW_HEADERS = ['Quote ID', 'Added At', 'Status', 'Reviewed At', 'Reviewed By'];

function markEmailImportForReview_(quoteId) {
  initLogSheet_(TAB_EMAIL_REVIEW, EMAIL_REVIEW_HEADERS).appendRow([quoteId, new Date(), 'UNREVIEWED', '', '']);
}

// Pending reviews joined to their live quote row. A quote that's since been deleted drops
// out silently — deleting it WAS the review decision, there's nothing left to confirm.
// Orders whose event date has already passed drop off too, and get closed out in the
// sheet as EXPIRED so they never come back: there is nothing left to catch on an order
// that already happened, and a queue full of history hides the ones that still matter.
// Only the review row is touched — the quote itself is never deleted.
function getEmailImportReviews() {
  var sheet = getSpreadsheet().getSheetByName(TAB_EMAIL_REVIEW);
  if (!sheet || sheet.getLastRow() < 2) return { pending: [] };
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, EMAIL_REVIEW_HEADERS.length);
  var rows = range.getValues();
  var addedAt = {};
  rows.forEach(function(r) {
    var id = (r[0] || '').toString().trim();
    if (id && (r[2] || '').toString().trim() === 'UNREVIEWED') addedAt[id] = r[1] ? new Date(r[1]).toISOString() : '';
  });
  var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var expired = {}, out = [];
  getQuotes().forEach(function(q) {
    if (!(q.quoteId in addedAt)) return;
    // Blank event date never expires — it's unscheduled, not past.
    if (q.eventDate && q.eventDate < todayKey) { expired[q.quoteId] = true; return; }
    out.push({
      quoteId: q.quoteId, sheetRow: q.sheetRow, customerName: q.customerName, contactName: q.contactName,
      customerEmail: q.customerEmail, customerPhone: q.customerPhone, orderType: q.orderType,
      deliveryAddress: q.deliveryAddress, eventDate: q.eventDate, eventTime: q.eventTime,
      total: q.total, lineItems: q.lineItems, quoteNotes: q.quoteNotes, status: q.status,
      poNumber: q.poNumber, calendarEventId: q.calendarEventId, addedAt: addedAt[q.quoteId]
    });
  });
  // Close the expired rows out in one write. Best-effort: this runs on a read path, so a
  // failure here must never stop the list from rendering.
  var expiredIds = Object.keys(expired);
  if (expiredIds.length) {
    try {
      var now = new Date(), changed = false;
      rows.forEach(function(r) {
        var id = (r[0] || '').toString().trim();
        if (!expired[id] || (r[2] || '').toString().trim() !== 'UNREVIEWED') return;
        r[2] = 'EXPIRED'; r[3] = now; r[4] = 'auto (event date passed)';
        changed = true;
      });
      if (changed) { range.setValues(rows); SpreadsheetApp.flush(); }
    } catch (e) { /* leave them pending; the next load will try again */ }
  }
  // Soonest event first — the one happening tomorrow is the one worth checking now.
  out.sort(function(a, b) { return (a.eventDate + (a.eventTime || '')).localeCompare(b.eventDate + (b.eventTime || '')); });
  return { pending: out };
}

// Cheap count for the nav badge.
function getUnreviewedEmailImportCount() { return getEmailImportReviews().pending.length; }

// Orders auto-imported BEFORE this queue existed have no review row, so the list shows
// empty even though the calendar is full of 📧 entries. This catches them up.
// Identified by the note stamp buildQuoteFromParsed_ writes — the manual "Import from
// Email" path writes a DIFFERENT stamp ("Imported from CFA …"), so quotes a human built
// through the form stay out of the queue, same as they do going forward.
// Idempotent: a quote already listed (reviewed or not) is never added twice.
const EMAIL_IMPORT_NOTE_MARK = 'Added automatically via email';

function backfillEmailImportReviews() {
  var sheet = initLogSheet_(TAB_EMAIL_REVIEW, EMAIL_REVIEW_HEADERS);
  var known = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var id = (r[0] || '').toString().trim(); if (id) known[id] = true;
    });
  }
  var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var rows = [], past = 0;
  getQuotes().forEach(function(q) {
    if (known[q.quoteId]) return;
    if ((q.quoteNotes || '').toString().trim().indexOf(EMAIL_IMPORT_NOTE_MARK) !== 0) return;
    // Its event already happened — there's nothing left to catch, so don't make someone
    // clear a backlog of history to get to the orders that still matter.
    if (q.eventDate && q.eventDate < todayKey) { past++; return; }
    rows.push([q.quoteId, q.createdDate ? new Date(q.createdDate) : new Date(), 'UNREVIEWED', '', '']);
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, EMAIL_REVIEW_HEADERS.length).setValues(rows);
  SpreadsheetApp.flush();
  return {
    added: rows.length, skippedPast: past,
    message: rows.length
      ? 'Found ' + rows.length + ' email-added order(s) to review.' + (past ? ' Ignored ' + past + ' whose event already passed.' : '')
      : (past ? 'No upcoming ones to review — ' + past + ' email-added order(s) found, but their events have all passed.'
              : 'No email-added orders found that aren\'t already listed.')
  };
}

// Mark one or more reviewed. Also drops the 📧 from each calendar title — that flag means
// "a machine wrote this, nobody has checked it", which stops being true the moment this
// runs. buildCalendarTitle is fed the whole customer segment back, so an assigned driver
// and the PO colour survive untouched.
//
// Batched deliberately: approving 15 orders one at a time would re-read the ENTIRE Quotes
// tab 15 times and rewrite the review sheet 15 times. This reads each once and writes the
// status column back in a single setValues.
function approveEmailImports(quoteIds) {
  var ids = (quoteIds || []).map(function(v) { return (v || '').toString().trim(); }).filter(function(v) { return v; });
  if (!ids.length) return { success: false, message: 'Nothing selected.' };
  var want = {};
  ids.forEach(function(id) { want[id] = true; });

  var sheet = initLogSheet_(TAB_EMAIL_REVIEW, EMAIL_REVIEW_HEADERS);
  var last = sheet.getLastRow();
  if (last < 2) return { success: false, message: 'Nothing is awaiting review.' };
  var range = sheet.getRange(2, 1, last - 1, EMAIL_REVIEW_HEADERS.length);
  var rows = range.getValues();
  var who = ''; try { who = Session.getEffectiveUser().getEmail(); } catch(e) { /* not resolvable — leave blank */ }
  var now = new Date(), hit = {};
  rows.forEach(function(r) {
    var id = (r[0] || '').toString().trim();
    if (!want[id]) return;
    if ((r[2] || '').toString().trim() !== 'UNREVIEWED') return; // already handled by someone else
    r[2] = 'REVIEWED'; r[3] = now; r[4] = who;
    hit[id] = true;
  });
  var approvedIds = Object.keys(hit);
  if (!approvedIds.length) {
    return { success: false, message: 'None of those are still awaiting review — the list is out of date. Refresh and try again.' };
  }
  range.setValues(rows);

  // Clear the 📧 per order. Best-effort and INDIVIDUALLY caught: one unreachable event
  // must not cost the other fourteen their review, which is already written above.
  var calendarFailed = [];
  getQuotes().forEach(function(q) {
    if (!hit[q.quoteId] || !q.calendarEventId) return;
    try {
      var res = updateCalendarEventPO(q.sheetRow, q.poNumber, q.calendarEventId);
      if (res !== true) calendarFailed.push(q.quoteId);
    } catch(e) { calendarFailed.push(q.quoteId); }
  });
  SpreadsheetApp.flush();
  return {
    success: true, approved: approvedIds.length,
    stale: ids.length - approvedIds.length, // asked for but no longer pending
    calendarFailed: calendarFailed
  };
}

// Single-order wrapper, kept so an older deployed App.html that calls this still works.
function approveEmailImport(quoteId) {
  var r = approveEmailImports([quoteId]);
  if (!r.success) return r;
  return {
    success: true,
    calendarChanged: !(r.calendarFailed || []).length,
    calendarReason: (r.calendarFailed || []).length ? 'Calendar title not updated' : ''
  };
}


// ── ITEM NAME ALIASES + MATCHING QUEUE ───────────────────────
// Maps an item name (as it appears on quotes / from email imports) to a canonical
// MENU item, applied at supplies read-time so mapping is retroactive. A row with a
// blank "Menu Item" = "ignore" (known non-menu item; stop flagging it). The master
// menu is never touched — this is a translation layer only.
function ensureItemAliasSheet_() {
  var ss = getSpreadsheet();
  var s = ss.getSheetByName(TAB_ITEM_ALIASES);
  if (!s) { s = ss.insertSheet(TAB_ITEM_ALIASES); s.getRange(1, 1, 1, 2).setValues([['Item Name', 'Menu Item (blank = ignore)']]).setFontWeight('bold'); s.setFrozenRows(1); s.setColumnWidth(1, 320); s.setColumnWidth(2, 320); }
  return s;
}
// { lowercased item name -> canonical menu item } — only rows with a non-blank menu item.
function getItemAliasMap_() {
  var s = ensureItemAliasSheet_(), map = {};
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function(r) { var a = (r[0] || '').toString().trim(), m = (r[1] || '').toString().trim(); if (a && m) map[a.toLowerCase()] = m; });
  return map;
}
// Set of every name already in the alias table (mapped OR ignored) — excluded from the queue.
function getItemAliasedNames_() {
  var s = ensureItemAliasSheet_(), set = {};
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().forEach(function(r) { var a = (r[0] || '').toString().trim(); if (a) set[a.toLowerCase()] = true; });
  return set;
}
// Union of both stores' menu item names (for the mapping dropdown + match check).
function getAllMenuItemNames_() {
  var names = {}, out = [];
  ['1', '2'].forEach(function(st) { getMenuItems(st).forEach(function(it) { var n = it.name; if (n && !names[n.toLowerCase()]) { names[n.toLowerCase()] = true; out.push(n); } }); });
  out.sort(function(a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : 1; });
  return out;
}

// Distinct item names across all quotes that DON'T resolve to a menu item and aren't
// aliased/ignored yet — the matching queue. Returns { pending:[{name, count}], menuNames:[...] }.
function getUnmatchedItemNames() {
  var menuSet = {}, menuNames = getAllMenuItemNames_();
  menuNames.forEach(function(n) { menuSet[n.toLowerCase()] = true; });
  var aliased = getItemAliasedNames_();
  var counts = {}, display = {};
  getQuotes().forEach(function(q) {
    var items = []; try { items = JSON.parse(q.lineItems) || []; } catch(e) { /* skip */ }
    items.forEach(function(it) {
      var name = (it.description || '').toString().trim(); if (!name) return;
      var key = name.toLowerCase();
      if (menuSet[key] || aliased[key]) return; // already resolvable
      counts[key] = (counts[key] || 0) + 1;
      if (!display[key]) display[key] = name;
    });
  });
  var pending = Object.keys(counts).map(function(k) { return { name: display[k], count: counts[k] }; });
  pending.sort(function(a, b) { return b.count - a.count || a.name.toLowerCase().localeCompare(b.name.toLowerCase()); });
  return { pending: pending, menuNames: menuNames };
}
// Lightweight count for the tab badge.
function getUnmatchedItemCount() { return getUnmatchedItemNames().pending.length; }

// Upsert an alias. menuItem blank/empty = "ignore" (known non-menu item).
function setItemAlias(aliasName, menuItem) {
  var a = (aliasName || '').toString().trim(); if (!a) return { success: false, message: 'No item name.' };
  var m = (menuItem || '').toString().trim();
  var s = ensureItemAliasSheet_(), last = s.getLastRow(), found = 0;
  if (last > 1) { var rows = s.getRange(2, 1, last - 1, 1).getValues(); for (var i = 0; i < rows.length; i++) { if ((rows[i][0] || '').toString().trim().toLowerCase() === a.toLowerCase()) { found = i + 2; break; } } }
  if (found) s.getRange(found, 1, 1, 2).setValues([[a, m]]); else s.appendRow([a, m]);
  SpreadsheetApp.flush();
  return { success: true };
}


// ── CUSTOMER DATABASE (non-destructive alias merge + editable overrides) ──
// Merging is an ALIAS map (variant -> canonical); quote history is never rewritten.
// A soft password (Script Property) gates reads/writes — deterrent only, not real ACL.
const TAB_CUSTOMERS = 'Customers';
const TAB_CUST_ALIASES = 'Customer_Aliases';

function ensureCustomerSheets_() {
  var ss = getSpreadsheet();
  var c = ss.getSheetByName(TAB_CUSTOMERS);
  if (!c) { c = ss.insertSheet(TAB_CUSTOMERS); c.getRange(1, 1, 1, 5).setValues([['Canonical Name', 'Email', 'Phone', 'Contact', 'Notes']]).setFontWeight('bold'); c.setFrozenRows(1); c.setColumnWidth(1, 220); c.setColumnWidth(5, 280); }
  var a = ss.getSheetByName(TAB_CUST_ALIASES);
  if (!a) { a = ss.insertSheet(TAB_CUST_ALIASES); a.getRange(1, 1, 1, 2).setValues([['Variant Name', 'Canonical Name']]).setFontWeight('bold'); a.setFrozenRows(1); a.setColumnWidth(1, 220); a.setColumnWidth(2, 220); }
  return { customers: c, aliases: a };
}

function getCustomerAliasMap_() {
  var a = ensureCustomerSheets_().aliases, map = {};
  if (a.getLastRow() > 1) a.getRange(2, 1, a.getLastRow() - 1, 2).getValues().forEach(function(r) { var v = (r[0] || '').toString().trim(), c = (r[1] || '').toString().trim(); if (v && c) map[v.toLowerCase()] = c; });
  return map;
}
function canonicalName_(name, aliasMap) { var n = (name || '').toString().trim(); return aliasMap[n.toLowerCase()] || n; }

// Soft password gate. No password set in Script Properties → open.
function customerDbHasPassword() { return !!(PropertiesService.getScriptProperties().getProperty('CUSTOMER_DB_PASSWORD') || ''); }
function checkCustomerDbPassword_(pw) {
  var stored = PropertiesService.getScriptProperties().getProperty('CUSTOMER_DB_PASSWORD') || '';
  return !stored || (pw || '').toString() === stored;
}
function setCustomerDbPassword(pw) { PropertiesService.getScriptProperties().setProperty('CUSTOMER_DB_PASSWORD', (pw || '').toString()); return true; }

// Grouped customer list: canonical name + aggregated quote stats + editable overrides + merged variants.
function getCustomerDatabase(password) {
  if (!checkCustomerDbPassword_(password)) return { authorized: false, customers: [] };
  var aliasMap = getCustomerAliasMap_();
  var sheets = ensureCustomerSheets_();
  var overrides = {}, cs = sheets.customers;
  if (cs.getLastRow() > 1) cs.getRange(2, 1, cs.getLastRow() - 1, 5).getValues().forEach(function(r) {
    var n = (r[0] || '').toString().trim(); if (!n) return;
    overrides[n.toLowerCase()] = { name: n, email: (r[1] || '').toString().trim(), phone: (r[2] || '').toString().trim(), contact: (r[3] || '').toString().trim(), notes: (r[4] || '').toString() };
  });
  var byCanon = {};
  getQuotes().forEach(function(q) {
    var canon = canonicalName_(q.customerName, aliasMap), key = canon.toLowerCase();
    if (!byCanon[key]) byCanon[key] = { name: canon, orderCount: 0, totalSpent: 0, lastOrder: '', email: '', phone: '', contact: '', variants: {} };
    var c = byCanon[key];
    c.orderCount++; c.totalSpent += parseFloat(q.total) || 0;
    if (q.eventDate && q.eventDate > c.lastOrder) c.lastOrder = q.eventDate;
    if (q.customerEmail && !c.email) c.email = q.customerEmail;
    if (q.customerPhone && !c.phone) c.phone = q.customerPhone;
    if (q.contactName && !c.contact) c.contact = q.contactName;
    var orig = (q.customerName || '').toString().trim();
    if (orig && orig.toLowerCase() !== key) c.variants[orig] = true;
  });
  Object.keys(overrides).forEach(function(k) { if (!byCanon[k]) byCanon[k] = { name: overrides[k].name, orderCount: 0, totalSpent: 0, lastOrder: '', email: '', phone: '', contact: '', variants: {} }; });
  var list = Object.keys(byCanon).map(function(k) {
    var c = byCanon[k], ov = overrides[k] || {};
    return { name: c.name, orderCount: c.orderCount, totalSpent: Math.round(c.totalSpent * 100) / 100, lastOrder: c.lastOrder,
      email: ov.email || c.email, phone: ov.phone || c.phone, contact: ov.contact || c.contact, notes: ov.notes || '', variants: Object.keys(c.variants) };
  });
  list.sort(function(a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
  return { authorized: true, customers: list };
}

// Merge: alias variant -> canonical (upsert). Non-destructive; history untouched.
function mergeCustomer(variantName, canonicalName) {
  var v = (variantName || '').toString().trim(), c = (canonicalName || '').toString().trim();
  if (!v || !c || v.toLowerCase() === c.toLowerCase()) return { success: false, message: 'Pick two different names.' };
  var a = ensureCustomerSheets_().aliases, last = a.getLastRow(), found = 0;
  if (last > 1) { var rows = a.getRange(2, 1, last - 1, 1).getValues(); for (var i = 0; i < rows.length; i++) { if ((rows[i][0] || '').toString().trim().toLowerCase() === v.toLowerCase()) { found = i + 2; break; } } }
  if (found) a.getRange(found, 2).setValue(c); else a.appendRow([v, c]);
  SpreadsheetApp.flush();
  return { success: true };
}
function unmergeCustomer(variantName) {
  var v = (variantName || '').toString().trim(); if (!v) return false;
  var a = ensureCustomerSheets_().aliases, last = a.getLastRow();
  if (last > 1) { var rows = a.getRange(2, 1, last - 1, 1).getValues(); for (var i = last - 2; i >= 0; i--) { if ((rows[i][0] || '').toString().trim().toLowerCase() === v.toLowerCase()) a.deleteRow(i + 2); } }
  SpreadsheetApp.flush();
  return true;
}

// Upsert editable override info for a canonical customer.
function saveCustomerInfo(entry) {
  var name = (entry.name || '').toString().trim(); if (!name) return { success: false, message: 'Name required.' };
  var cs = ensureCustomerSheets_().customers, last = cs.getLastRow(), found = 0;
  if (last > 1) { var rows = cs.getRange(2, 1, last - 1, 1).getValues(); for (var i = 0; i < rows.length; i++) { if ((rows[i][0] || '').toString().trim().toLowerCase() === name.toLowerCase()) { found = i + 2; break; } } }
  var vals = [name, entry.email || '', entry.phone || '', entry.contact || '', entry.notes || ''];
  if (found) cs.getRange(found, 1, 1, 5).setValues([vals]); else cs.appendRow(vals);
  SpreadsheetApp.flush();
  return { success: true };
}


// ── WEEKLY DISD RECAP (Saturday 12pm) ────────────────────────
// Emails PO recipients a summary of DISD orders delivered in the last 7 days
// (guest + total only, no line items). Dedups on quoteId so a manual re-run
// the same week won't double-report an order. Text half intentionally omitted.
const TAB_DISD_WEEKLY = 'DISD_Weekly_Sent';

// The old rule was a plain indexOf('disd') over customer + contact name, which matched
// NEITHER of the two ways this district actually appears: "Dallas ISD" has a space, and
// "dallasisd.org" never appeared because the email wasn't searched at all. Strip everything
// to letters+digits on both sides and compare against an operator-editable term list, so
// "Dallas ISD", "DISD", and either email domain all land.
function disdNormalize_(s) { return (s || '').toString().toLowerCase().replace(/[^a-z0-9]/g, ''); }

function disdMatchTerms_(settings) {
  var raw = ((settings || {})['DISD Reminder Match Terms'] || 'DISD, Dallas ISD, disd.org, dallasisd.org').toString();
  return raw.split(',').map(disdNormalize_).filter(function(t) { return t; });
}

// Customer, contact AND email are searched — an auto-imported order often has a person in
// both name fields, leaving the @disd.org address as the only thing identifying the district.
//
// Two match modes, because one alone gets it wrong:
//   long terms  → substring of the stripped text, so "Dallas ISD" matches "dallasisd"
//   short terms → must stand alone in the raw text, or "disd" hits paradisdiner.com
function matchesDisd_(q, terms) {
  var raw = ((q.customerName || '') + ' ' + (q.contactName || '') + ' ' + (q.customerEmail || '')).toLowerCase();
  if (!raw.trim()) return false;
  var norm = disdNormalize_(raw);
  return (terms || []).some(function(t) {
    if (!t) return false;
    if (t.length >= 7) return norm.indexOf(t) >= 0;
    return new RegExp('(^|[^a-z0-9])' + t + '([^a-z0-9]|$)').test(raw);
  });
}

function sendDisdWeeklyReminder_() {
  var settings = getSettings();
  if (!isSettingOn_(settings['DISD Reminder Enabled'])) return 'DISD reminder is disabled.';
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var weekAgo = new Date(); weekAgo.setDate(weekAgo.getDate() - 7);
  var weekAgoKey = Utilities.formatDate(weekAgo, tz, 'yyyy-MM-dd');

  var logSheet = initLogSheet_(TAB_DISD_WEEKLY, ['Quote ID', 'Event Date', 'Sent At']);
  var sent = {};
  if (logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
      if (r[0]) sent[r[0].toString().trim()] = true;
    });
  }

  // "Delivered" = a delivery order whose event date fell in the last 7 days (up to today).
  var terms = disdMatchTerms_(settings); // hoisted — same for every row
  var disd = getQuotes().filter(function(q) {
    if (q.status === 'Inactive') return false;
    if (!q.eventDate || q.eventDate < weekAgoKey || q.eventDate > todayKey) return false;
    if (!matchesDisd_(q, terms)) return false;
    return !sent[q.quoteId];
  });
  if (!disd.length) return 'No DISD orders delivered in the last 7 days.';

  var to = (settings['DISD Reminder Recipients'] || '').toString().trim() || poAlertRecipients_(settings);
  if (!to) return 'No DISD reminder recipient configured.';

  // Quote ID + PO so a line can be tracked down without opening the tool to find it.
  // A delivered order with NO PO is a billing problem, not a blank — flag it in red rather
  // than leaving an empty cell that reads like "nothing to see here".
  var rows = disd.map(function(q) {
    var p = (q.poNumber || '').toString().trim();
    var poCell = p === 'NO_PO_NEEDED' ? 'No PO needed'
      : (p ? p : digestRaw_('<b style="color:' + DIGEST_ACCENT + ';">none on file</b>'));
    return [q.quoteId, poCell, q.customerName, formatWhen_(q.eventDate, q.eventTime, tz), '$' + (parseFloat(q.total) || 0).toFixed(2)];
  });
  var grand = disd.reduce(function(s, q) { return s + (parseFloat(q.total) || 0); }, 0);
  // Footer row must carry one cell per column or the table goes ragged.
  rows.push([digestRaw_('<b>Total</b>'), '', '', '', digestRaw_('<b>$' + grand.toFixed(2) + '</b>')]);
  var msg = (settings['DISD Reminder Message'] || 'DISD orders delivered in the last 7 days:');
  var subj = (settings['DISD Reminder Subject'] || 'Weekly DISD catering recap — {{count}} order(s)')
    .split('{{count}}').join(disd.length);
  sendDigest_(to, subj, {
    title: 'Weekly DISD recap',
    accentText: disd.length + ' order' + (disd.length === 1 ? '' : 's') + ' delivered in the last 7 days',
    intro: msg,
    columns: ['Quote', 'PO', 'Customer', 'Delivered', 'Total'],
    align: ['left', 'left', 'left', 'left', 'right'],
    nowrap: [true, true, false, true, true],
    rows: rows
  });
  disd.forEach(function(q) { logSheet.appendRow([q.quoteId, q.eventDate, new Date()]); });
  SpreadsheetApp.flush();
  return 'Reported ' + disd.length + ' DISD order(s).';
}

// Trigger wrapper — alert on failure, matching poWarningDayBefore's pattern.
function disdWeeklyReminder() {
  var alertTo = ''; try { alertTo = poAlertRecipients_(getSettings()); } catch(e) { /* no recipient */ }
  try { sendDisdWeeklyReminder_(); } catch(e) {
    try { MailApp.sendEmail(alertTo, 'Catering Automation Error', 'DISD weekly reminder failed:\n\n' + e.message + '\n\n' + e.stack); } catch(e2) { /* give up quietly */ }
  }
}

function setupDisdReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'disdWeeklyReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('disdWeeklyReminder').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(12).create();
  return 'DISD weekly reminder enabled (Saturday 12pm).';
}
function removeDisdReminderTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'disdWeeklyReminder') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed > 0 ? 'DISD reminder trigger removed.' : 'No active DISD reminder trigger found.';
}
function getDisdReminderTriggerStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'disdWeeklyReminder';
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
    try { return DriveApp.getFolderById(id); } catch(e) { /* folder deleted — fall through and recreate */ }
  }
  var it = DriveApp.getFoldersByName(TAX_FORM_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(TAX_FORM_FOLDER_NAME);
  props.setProperty('TAX_FORM_FOLDER_ID', folder.getId());
  return folder;
}

// Run this FROM THE EDITOR (Run button) to diagnose + fix the Drive permission.
// Because it references DriveApp, running it forces the OAuth consent screen — accept
// it to grant Drive, then push a NEW deployment version so the web app picks it up.
function checkTaxFormAuth() {
  var out = { reauthRequired: false, authUrl: '', driveWorks: false, driveError: '' };
  var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  out.reauthRequired = info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED;
  if (out.reauthRequired) out.authUrl = info.getAuthorizationUrl();
  // The real test — a pinned manifest can hide a missing scope from getAuthorizationInfo.
  try { DriveApp.getRootFolder().getName(); out.driveWorks = true; }
  catch (e) { out.driveError = e.message; }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// Decode a base64 PDF and drop it in the folder. Files stay private to the
// Drive owner (these are sensitive tax docs) — the team opens them via the folder.
function saveTaxPdf_(baseName, mimeType, dataBase64) {
  var bytes = Utilities.base64Decode(dataBase64);
  if (bytes.length > TAX_MAX_UPLOAD_BYTES) throw new Error('File is too large (10 MB max).');
  var safe = (baseName || 'Tax Exemption Certificate').replace(/[\\/:*?"<>|]/g, '-').slice(0, 120);
  var mime = mimeType || 'application/pdf';
  // Photos keep their own extension so Drive previews them correctly
  var ext = mime.indexOf('pdf') >= 0 ? '.pdf' : (mime === 'image/jpeg' ? '.jpg' : '.' + (mime.split('/')[1] || 'jpg'));
  var blob = Utilities.newBlob(bytes, mime, safe + ext);
  return ensureTaxFormFolder_().createFile(blob);
}

// One call for the client: every quote's status plus the folder link.
function getTaxFormStatuses() {
  var out = { folderUrl: '', statuses: {} };
  try { out.folderUrl = ensureTaxFormFolder_().getUrl(); } catch(e) { /* Drive not authorized yet — statuses still work */ }
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
  try { uploadLink = ScriptApp.getService().getUrl() + '?view=taxform&quote=' + encodeURIComponent(quote.quoteId || ''); } catch(e) { /* no deployment URL available — link stays blank */ }
  var reps = {
    '{{contactPerson}}': quote.contactName || '',
    '{{customer}}': quote.customerName || '',
    '{{contact}}': settings['Quote Contact Name'] || '',
    '{{location}}': storeName,
    '{{phone}}': storePhone,
    '{{quoteId}}': quote.quoteId || '',
    '{{uploadLink}}': uploadLink
  };
  var subj = settings['Tax Form Request Subject'] || 'Texas Sales and Use Tax Exemption Certificate needed for your catering quote {{quoteId}}';
  var body = settings['Tax Form Request Body'] || 'Hi {{contactPerson}},\n\nPlease reply with a PDF of your Texas Sales and Use Tax Exemption Certificate, including your name and quote number {{quoteId}}.\n\nThank you!\n{{contact}}';
  for (var k in reps) { subj = subj.split(k).join(reps[k]); body = body.split(k).join(reps[k]); }
  var opts = { name: 'Chick-fil-A ' + storeName };
  var bcc = (settings['BCC Email'] || '').toString().trim();
  if (bcc) opts.bcc = bcc;
  MailApp.sendEmail(recipientEmail, subj, body, opts);
  // A request can now be sent even when a certificate is already on file (they expire).
  // Asking for a fresh copy must NOT downgrade 🟢 to 🟠 — the form we have is still on
  // file until a new one actually arrives and someone confirms it.
  var current = '';
  try { current = (getTaxFormStatuses().statuses[quoteId] || '').toString().trim(); } catch(e) { /* status read is advisory */ }
  if (current !== 'ON_FILE') setTaxFormStatus(quoteId, 'REQUESTED');
  return { success: true, message: 'Form request sent to ' + recipientEmail, statusKept: current === 'ON_FILE' };
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
  var rows = reg.entries.map(function(e) {
    var dot = e.status === 'ON_FILE' ? '🟢' : (e.status === 'REQUESTED' ? '🟠' : '🔴');
    var label = e.status === 'ON_FILE' ? 'On file' : (e.status === 'REQUESTED' ? 'Requested' : 'Missing');
    return [
      e.organization,
      dot + ' ' + label,
      e.dateOnFile || '—',
      e.pdfUrl ? digestRaw_('<a href="' + esc(e.pdfUrl) + '" style="color:' + DIGEST_ACCENT + ';">View PDF</a>') : '—'
    ];
  });
  if (!rows.length) rows = [[digestRaw_('<i style="color:#9CA3AF;">The registry is empty — nothing recorded yet.</i>'), '', '', '']];
  var exemptCount = getQuotes().filter(function(q) {
    var te = q.taxExempt === 'TRUE' || q.taxExempt === true;
    return te && q.createdDate && new Date(q.createdDate).getFullYear() === year;
  }).length;
  sendDigest_(to, '📋 Year-end review: catering Tax Exemption Certificates (' + year + ')', {
    title: 'Year-end tax certificate review — ' + year,
    accentText: reg.entries.length + ' organization' + (reg.entries.length === 1 ? '' : 's') + ' on the registry',
    intro: 'It\'s the last business day of ' + year + '. You created ' + exemptCount + ' tax-exempt quote(s) this year.'
      + (reg.pending.length ? '  ⚠ ' + reg.pending.length + ' guest upload(s) still await review in the app.' : ''),
    columns: ['Organization', 'Status', 'On file', 'Certificate'],
    align: ['left', 'left', 'left', 'left'],
    nowrap: [false, true, true, true],
    rows: rows,
    footer: 'Open each 🟢 to confirm the form is current; chase anything 🟠 or 🔴 before the next order.'
      + (reg.folderUrl ? '  Forms folder: ' + reg.folderUrl : '')
  });
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
  var out = { folderUrl: '', uploadUrl: '', entries: [], pending: [] };
  try { out.folderUrl = ensureTaxFormFolder_().getUrl(); } catch(e) { /* Drive not authorized yet — registry still loads */ }
  // Public guest upload page (no quote param — guests type their quote number on it)
  try { out.uploadUrl = ScriptApp.getService().getUrl() + '?view=taxform'; } catch(e) { /* no deployment URL available — link stays blank */ }
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
    pend.getRange(2, 1, pend.getLastRow() - 1, pend.getLastColumn()).getValues().forEach(function(r, i) {
      if ((r[6] || '').toString().trim() !== 'PENDING') return;
      out.pending.push({
        submittedAt: r[0] ? new Date(r[0]).toISOString() : '',
        organization: (r[1] || '').toString().trim(),
        contactName: (r[2] || '').toString().trim(),
        quoteId: (r[3] || '').toString().trim(),
        pdfUrl: (r[4] || '').toString().trim(),
        taxNumber: (r[7] || '').toString().trim(),
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
  var taxNumber = (payload.taxNumber || '').toString().trim();
  var data = (payload.dataBase64 || '').toString();
  if (!org || !name) return { success: false, message: 'Please enter your name and organization.' };
  if (!taxNumber) return { success: false, message: 'Please enter your tax-exempt number.' };
  if (!data) return { success: false, message: 'Please attach a PDF or photo of your form.' };
  var mime = (payload.mimeType || '').toString();
  if (mime && mime.indexOf('pdf') < 0 && mime.indexOf('image/') !== 0) return { success: false, message: 'Please upload a PDF or a photo.' };
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var file;
  try {
    file = saveTaxPdf_(org + ' - ' + name + (quoteId ? ' - ' + quoteId : '') + ' - ' + stamp, mime || 'application/pdf', data);
  } catch (err) {
    return { success: false, message: 'We couldn\'t save your form right now — please try again in a moment. If it keeps happening, reply to our email and we\'ll take it directly.' };
  }
  var sheet = initLogSheet_(TAB_TAX_PENDING, ['Submitted At', 'Organization', 'Contact Name', 'Quote ID', 'PDF Link', 'PDF File ID', 'Status', 'Tax-Exempt Number']);
  sheet.appendRow([new Date(), org, name, quoteId, file.getUrl(), file.getId(), 'PENDING', taxNumber]);
  return { success: true, message: 'Thank you! We\'ve received your Texas Sales and Use Tax Exemption Certificate.' };
}

// Internal: confirm a pending guest upload into the registry under a chosen org name.
function confirmPendingTaxUpload(pendingRow, organization) {
  var org = (organization || '').toString().trim();
  if (!org) return { success: false, message: 'Pick an organization name.' };
  var pend = getSpreadsheet().getSheetByName(TAB_TAX_PENDING);
  if (!pend || pendingRow < 2) return { success: false, message: 'Upload not found.' };
  var r = pend.getRange(pendingRow, 1, 1, pend.getLastColumn()).getValues()[0];
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var taxNumber = (r[7] || '').toString().trim();
  var note = (taxNumber ? 'Tax-exempt #: ' + taxNumber + ' · ' : '') + 'Uploaded by ' + (r[2] || 'guest');
  saveTaxRegistryEntry({ organization: org, status: 'ON_FILE', pdfUrl: r[4], pdfFileId: r[5], dateOnFile: stamp, notes: note });
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

// Short store code (CH / DBU) for a quote's location, from the store-prefix settings.
function storeCodeForQuote_(locationName, settings) {
  var loc = (locationName || '').toString().trim();
  var l2 = (settings['Location 2 Name'] || '').toString().trim();
  var p1 = (settings['Store 1 Tab Prefix'] || 'CH').toString().trim() || 'CH';
  var p2 = (settings['Store 2 Tab Prefix'] || 'DBU').toString().trim() || 'DBU';
  return (l2 && loc === l2) ? p2 : p1;
}

function emailRunsheet(dateStr, recipient) {
  recipient = (recipient || '').toString().trim();
  if (!recipient) return { success: false, message: 'No recipient given.' };
  var tz = Session.getScriptTimeZone();
  var settings = getSettings();
  var quotes = getQuotes().filter(function(q) { return q.eventDate === dateStr && q.status !== 'Inactive'; });
  if (!quotes.length) return { success: false, message: 'No orders on that day.' };
  quotes.sort(function(a, b) { return (a.eventTime || '').localeCompare(b.eventTime || ''); });

  var rows = quotes.map(function(q) {
    var items = [];
    try { items = JSON.parse(q.lineItems) || []; } catch(e) { /* malformed line-item JSON → no items */ }
    var itemsHtml = items.map(function(i) { return esc(i.quantity + '× ' + i.description); }).join('<br>');
    var when = formatWhen_(q.eventDate, q.eventTime, tz);
    var time = when.indexOf(' at ') >= 0 ? when.split(' at ')[1] : 'No time';
    var po = q.poNumber === 'NO_PO_NEEDED' ? '🟢 No PO needed' : (q.poNumber ? '🟢 PO: ' + esc(q.poNumber) : '🔴 VERIFY PO');
    var addr = q.orderType === 'Delivery' && q.deliveryAddress
      ? esc(q.deliveryAddress) + '<br><a href="' + mapsDirectionsUrl_(q.deliveryAddress) + '">Directions</a>'
      : esc(q.orderType || '');
    return '<tr>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;white-space:nowrap;"><b>' + esc(time) + '</b></td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;"><b>' + esc(q.customerName) + '</b><br>' + esc(q.contactName || '') + (q.customerPhone ? '<br>' + esc(q.customerPhone) : '') + '</td>'
      + '<td style="padding:8px;border-bottom:1px solid #ddd;font-weight:600;">' + esc(storeCodeForQuote_(q.locationName, settings)) + '</td>'
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
    + '<tr style="background:#f3f4f6;"><th style="padding:8px;text-align:left;">Time</th><th style="padding:8px;text-align:left;">Customer</th><th style="padding:8px;text-align:left;">Store</th><th style="padding:8px;text-align:left;">Type</th><th style="padding:8px;text-align:left;">Address</th><th style="padding:8px;text-align:left;">Items</th><th style="padding:8px;text-align:left;">PO</th><th style="padding:8px;text-align:right;">Total</th></tr>'
    + rows.join('') + '</table>';
  MailApp.sendEmail(recipient, 'Catering Runsheet — ' + dayLabel + ' (' + quotes.length + ' order' + (quotes.length === 1 ? '' : 's') + ')',
    'Runsheet for ' + dayLabel + ' — open in an HTML mail client.', { htmlBody: html, name: 'Catering Tool' });
  return { success: true, message: 'Runsheet sent to ' + recipient };
}


