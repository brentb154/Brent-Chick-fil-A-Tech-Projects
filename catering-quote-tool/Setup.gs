// ── FIRST-TIME SETUP ─────────────────────────────────────────
// A new store opens its copy of the sheet → "🍽️ Catering Setup" menu → Start Setup.
// The wizard asks for the store details once, saves them to the Settings tab, and
// builds every tab. Safe to re-run: it shows the current answers so they can be edited.

// The questions the wizard asks. key = the Settings row it fills.
var SETUP_FIELDS = [
  { key: 'Location 1 Name',      label: 'Store name',             hint: 'How guests know you, e.g. "Main Street FSU". Shows on quotes and emails.' },
  { key: 'Store 1 Tab Prefix',   label: 'Short store code',       hint: '2–4 letters for tab names. "MS" makes a tab called "MS Menu".' },
  { key: 'Location 1 Address',   label: 'Store address',          hint: 'Street, city, state, ZIP. Used for pickup and delivery directions.' },
  { key: 'Location 1 Phone',     label: 'Store phone',            hint: 'The number guests should call about catering.' },
  { key: 'Quote Contact Name',   label: 'Catering contact name',  hint: 'Printed on quotes as the person to call with questions.' },
  { key: 'Default Tax Rate (%)', label: 'Sales tax rate (%)',     hint: 'Your local rate, e.g. 8.25. It can still be changed on any quote.' },
  { key: 'PO Alert Email',       label: 'Email for order alerts', hint: 'Who gets missing-PO and order reminders. Separate several with commas.' }
];

// Asked only when the store runs catering from 2 locations.
var SETUP_STORE2_FIELDS = [
  { key: 'Location 2 Name',      label: 'Second store name',       hint: 'How guests know your second location.' },
  { key: 'Store 2 Tab Prefix',   label: 'Second store short code', hint: '2–4 letters, different from the first store\'s code.' },
  { key: 'Location 2 Address',   label: 'Second store address',    hint: 'Street, city, state, ZIP.' },
  { key: 'Location 2 Phone',     label: 'Second store phone',      hint: 'The number guests call for this location.' }
];

// "Number of Stores" is written only by the wizard, never seeded — a sheet without it
// (like one set up before this existed) keeps working as 2 stores.
function setupStoreCount_(s) {
  var n = (s['Number of Stores'] || '').toString().trim();
  if (n === '1' || n === '2') return Number(n);
  return (s['Location 2 Name'] || '').toString().trim() ? 2 : 1;
}

// Hide the second store's tabs for a 1-store setup (hidden, never deleted).
function applyStoreCountToTabs_(ss, count) {
  [menuTabForStore_('2'), offMenuTabForStore_('2')].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    if (count === 1) sh.hideSheet(); else sh.showSheet();
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍽️ Catering Setup')
    .addItem('Start Setup', 'showSetupWizard')
    .addToUi();
}

function showSetupWizard() {
  SpreadsheetApp.getActive().toast('Building your tabs — this takes about 30 seconds the first time.', 'Catering Setup', 30);
  initializeSheet(); // creates missing tabs only; never overwrites existing data
  var html = HtmlService.createTemplateFromFile('SetupWizard').evaluate().setWidth(520).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Catering Setup');
}

// Current answers + what's still missing. Also used by the web app's "finish setup" banner.
function getSetupStatus() {
  var s = {};
  try { s = getSettings(); } catch (e) { /* Settings tab not built yet — everything is missing */ }
  function withValue(f) {
    var v = s[f.key];
    return { key: f.key, label: f.label, hint: f.hint, value: (v === undefined || v === null) ? '' : v.toString() };
  }
  var fields = SETUP_FIELDS.map(withValue), store2Fields = SETUP_STORE2_FIELDS.map(withValue);
  var storeCount = setupStoreCount_(s);
  var asked = storeCount === 2 ? fields.concat(store2Fields) : fields;
  var missing = asked.filter(function(f) { return f.value.trim() === ''; }).map(function(f) { return f.label; });
  return { fields: fields, store2Fields: store2Fields, storeCount: storeCount, missing: missing, complete: missing.length === 0 };
}

function saveSetupAnswers(answers) {
  answers = answers || {};
  var storeCount = String(answers['Number of Stores']) === '2' ? 2 : 1;
  var asked = storeCount === 2 ? SETUP_FIELDS.concat(SETUP_STORE2_FIELDS) : SETUP_FIELDS;
  var clean = {}, missing = [];
  asked.forEach(function(f) {
    var v = (answers[f.key] || '').toString().trim();
    if (/Tab Prefix$/.test(f.key)) v = v.replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(0, 4);
    if (f.key === 'Default Tax Rate (%)') { v = parseFloat(v.replace('%', '')); if (isNaN(v)) v = ''; }
    if (v === '') missing.push(f.label); else clean[f.key] = v;
  });
  if (missing.length) return { success: false, message: 'Please fill in: ' + missing.join(', ') + '.' };
  if (storeCount === 2 && clean['Store 1 Tab Prefix'] === clean['Store 2 Tab Prefix']) {
    return { success: false, message: 'The two stores need different short codes.' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSpreadsheet();
    // Remember the current tab names so a new store code renames them in place
    var oldTabs = {};
    ['1', '2'].forEach(function(n) { oldTabs[n] = [menuTabForStore_(n), offMenuTabForStore_(n)]; });
    var s = getSettings();
    Object.keys(clean).forEach(function(k) { updateSetting(k, clean[k], typeof clean[k] === 'string'); });
    updateSetting('Number of Stores', String(storeCount), true);
    updateSetting('Store Name (Active)', clean['Location 1 Name'], true);
    if (!(s['Start Address'] || '').toString().trim()) updateSetting('Start Address', clean['Location 1 Address'], true);
    _tabPrefixCache = null;
    ['1', '2'].forEach(function(n) {
      renameSheetIfExists_(ss, oldTabs[n][0], menuTabForStore_(n));
      renameSheetIfExists_(ss, oldTabs[n][1], offMenuTabForStore_(n));
    });
    initializeSheet();
    applyStoreCountToTabs_(ss, storeCount);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { success: true, status: getSetupStatus() };
}
