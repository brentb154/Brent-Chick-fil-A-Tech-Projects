// ============================================================
// FOH Links Menu — Menu Builder & Setup
// ============================================================

// --- Global link opener pool (supports up to 200 links) ---
for (var i = 1; i <= 200; i++) {
  this['openLink_' + i] = (function(idx) {
    return function() {
      openLinkByIndex_(idx);
    };
  })(i);
}

/**
 * Builds the FOH Links menu and sidebar trigger on spreadsheet open.
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Links');
  var ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.createMenu('FOH Links')
      .addItem('Run Initial Setup', 'runInitialSetup')
      .addToUi();
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.createMenu('FOH Links')
      .addItem('Open Links Panel', 'showSidebar')
      .addToUi();
    return;
  }

  // Group links by category, build URL map
  var categories = {};
  var urlMap = {};
  var idx = 0;

  for (var r = 1; r < data.length; r++) {
    var category = String(data[r][0]).trim();
    var name = String(data[r][1]).trim();
    var url = String(data[r][2]).trim();
    var sortOrder = data[r][3] || 0;
    if (!category || !name || !url) continue;
    idx++;
    if (!categories[category]) categories[category] = [];
    categories[category].push({ name: name, idx: idx, sort: sortOrder });
    urlMap[idx] = url;
  }

  // Store URL map for opener functions
  PropertiesService.getScriptProperties().setProperty('linkUrlMap', JSON.stringify(urlMap));

  // Sort links within each category
  Object.keys(categories).forEach(function(cat) {
    categories[cat].sort(function(a, b) {
      return (a.sort || 0) - (b.sort || 0) || a.name.localeCompare(b.name);
    });
  });

  // Build menu — standard category order, then any extras
  var menu = ui.createMenu('FOH Links');
  var catOrder = ['Training', 'Communication', 'Operations', 'HR', 'Finance',
                  'Scheduling', 'Safety', 'Marketing', 'General', 'Other'];

  var added = {};
  catOrder.forEach(function(cat) {
    if (categories[cat] && categories[cat].length > 0) {
      var sub = ui.createMenu(cat);
      categories[cat].forEach(function(link) {
        sub.addItem(link.name, 'openLink_' + link.idx);
      });
      menu.addSubMenu(sub);
      added[cat] = true;
    }
  });

  // Any categories not in the standard list
  Object.keys(categories).forEach(function(cat) {
    if (!added[cat] && categories[cat].length > 0) {
      var sub = ui.createMenu(cat);
      categories[cat].forEach(function(link) {
        sub.addItem(link.name, 'openLink_' + link.idx);
      });
      menu.addSubMenu(sub);
    }
  });

  menu.addSeparator();
  menu.addItem('Open Links Panel', 'showSidebar');
  menu.addItem('Send Donation Request', 'showDonationDialog');
  menu.addSeparator();
  var resetMenu = ui.createMenu('Sheet Reset');
  resetMenu.addItem('Test Reset Monday Only', 'testResetMonday');
  resetMenu.addItem('Reset All Days Now', 'manualReset');
  resetMenu.addSeparator();
  resetMenu.addItem('Setup Sunday Auto-Reset', 'setupResetTrigger');
  resetMenu.addItem('Remove Auto-Reset', 'deleteResetTriggers');
  menu.addSubMenu(resetMenu);
  menu.addToUi();
}

/**
 * Opens a link by its index (called from the generated pool functions).
 */
function openLinkByIndex_(idx) {
  var urlMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('linkUrlMap') || '{}');
  var url = urlMap[idx];
  if (!url) {
    SpreadsheetApp.getUi().alert('Link not found. Try refreshing the page (this rebuilds the menu).');
    return;
  }
  var safe = url.replace(/"/g, '&quot;');
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + safe + '", "_blank");google.script.host.close();</script>'
  ).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

/**
 * Opens the Donation Request email dialog.
 */
function showDonationDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DonationDialog')
    .setWidth(480)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Donation Request');
}

/**
 * Sends the donation request email. Called from the dialog.
 */
function sendDonationEmail(formData) {
  var to = formData.email;
  var name = formData.name;
  var subject = formData.subject;
  var body = formData.message;

  if (!to || !name || !subject || !body) {
    throw new Error('All fields are required.');
  }

  var fromName = getSetting('Donation From Name');
  GmailApp.sendEmail(to, subject, body, {
    name: fromName,
    htmlBody: body.replace(/\n/g, '<br>')
  });

  // Log the send to a DonationLog sheet (create if needed)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('DonationLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('DonationLog');
    logSheet.getRange(1, 1, 1, 5).setValues([['Date', 'Name', 'Email', 'Subject', 'Sent By']]);
    logSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    logSheet.setFrozenRows(1);
  }
  logSheet.appendRow([
    new Date(),
    name,
    to,
    subject,
    Session.getActiveUser().getEmail() || 'Unknown'
  ]);
  SpreadsheetApp.flush();

  return { success: true };
}

/**
 * Returns donation dialog defaults: form URL from Links sheet + settings.
 */
function getDonationDefaults() {
  // Find form URL from Links sheet
  var formUrl = '';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var name = String(data[r][1]).toLowerCase();
      if (name.indexOf('donation') > -1 || name.indexOf('giving') > -1) {
        formUrl = String(data[r][2]).trim();
        break;
      }
    }
  }

  return {
    formUrl: formUrl,
    subject: getSetting('Donation Subject'),
    fromName: getSetting('Donation From Name'),
    message: getSetting('Donation Message')
  };
}

/**
 * Opens the Links sidebar panel.
 */
function showSidebar() {
  logAdoptionPing_('foh-links-menu');
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('FOH Links');
  SpreadsheetApp.getUi().showSidebar(html);
}

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

/**
 * Creates the Links sheet with headers and example data. Safe to re-run.
 */
function runInitialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Links');

  if (!sheet) {
    sheet = ss.insertSheet('Links');
  }

  // Check if headers already exist
  var existing = sheet.getRange('A1').getValue();
  if (existing === 'Category') {
    SpreadsheetApp.getUi().alert('Links sheet already set up. Refresh the page to see the menu.');
    return;
  }

  // Write headers
  var headers = ['Category', 'Name', 'URL', 'Sort Order', 'Pinned'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');

  // Set column widths
  sheet.setColumnWidth(1, 130); // Category
  sheet.setColumnWidth(2, 200); // Name
  sheet.setColumnWidth(3, 350); // URL
  sheet.setColumnWidth(4, 90);  // Sort Order
  sheet.setColumnWidth(5, 70);  // Pinned

  // Add example links
  var examples = [
    ['Training', 'Pathway', 'https://pathway.example.com', 1, true],
    ['Training', 'LMS Portal', 'https://lms.example.com', 2, false],
    ['Communication', 'Team GroupMe', 'https://groupme.example.com', 1, true],
    ['Operations', 'Food Safety Log', 'https://foodsafety.example.com', 1, false],
    ['Scheduling', 'HotSchedules', 'https://hotschedules.example.com', 1, true],
    ['General', 'CFA Operator Portal', 'https://operator.example.com', 1, false],
  ];
  sheet.getRange(2, 1, examples.length, examples[0].length).setValues(examples);

  // Add data validation for Pinned column (checkboxes)
  var pinnedRange = sheet.getRange(2, 5, 100);
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  pinnedRange.setDataValidation(rule);

  // Add data validation for Category column (dropdown)
  var cats = ['Training', 'Communication', 'Operations', 'HR', 'Finance',
              'Scheduling', 'Safety', 'Marketing', 'General', 'Other'];
  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cats, true)
    .setAllowInvalid(true) // allow custom categories too
    .build();
  sheet.getRange(2, 1, 100).setDataValidation(catRule);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Create Settings sheet
  setupSettingsSheet();

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Setup complete! Refresh the page to see the FOH Links menu.');
}
