// ============================================================
// FOH Links Menu — CRUD & Sidebar Helpers
// ============================================================

var LINK_HEADERS = ['Category', 'Name', 'URL', 'Sort Order', 'Pinned'];
var DEFAULT_CATEGORIES = ['Training', 'Communication', 'Operations', 'HR', 'Finance',
                          'Scheduling', 'Safety', 'Marketing', 'General', 'Other'];

/**
 * Returns all links with row numbers for the sidebar.
 */
function getLinksData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (!sheet) return { links: [], categories: DEFAULT_CATEGORIES };

  var data = sheet.getDataRange().getValues();
  var links = [];

  for (var r = 1; r < data.length; r++) {
    var category = String(data[r][0]).trim();
    var name = String(data[r][1]).trim();
    var url = String(data[r][2]).trim();
    if (!category || !name || !url) continue;

    links.push({
      row: r + 1, // sheet row (1-indexed, skip header)
      category: category,
      name: name,
      url: url,
      sortOrder: data[r][3] || 0,
      pinned: data[r][4] === true || data[r][4] === 'TRUE'
    });
  }

  // Sort: pinned first, then by sort order, then alphabetical
  links.sort(function(a, b) {
    if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
    return (a.sortOrder || 0) - (b.sortOrder || 0) || a.name.localeCompare(b.name);
  });

  // Collect unique categories from data + defaults
  var catSet = {};
  DEFAULT_CATEGORIES.forEach(function(c) { catSet[c] = true; });
  links.forEach(function(l) { catSet[l.category] = true; });
  var categories = Object.keys(catSet);

  return { links: links, categories: categories };
}

/**
 * Adds a new link to the Links sheet.
 */
function addLink(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (!sheet) throw new Error('Links sheet not found. Run Initial Setup first.');

  var row = [
    data.category || 'General',
    data.name,
    data.url,
    data.sortOrder || 0,
    data.pinned || false
  ];
  sheet.appendRow(row);

  // Add checkbox validation to the new pinned cell
  var lastRow = sheet.getLastRow();
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(lastRow, 5).setDataValidation(rule);

  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * Updates an existing link by sheet row number.
 */
function updateLink(rowNum, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (!sheet) throw new Error('Links sheet not found.');

  var row = [
    data.category || 'General',
    data.name,
    data.url,
    data.sortOrder || 0,
    data.pinned || false
  ];
  sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * Deletes a link by sheet row number.
 */
function deleteLink(rowNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (!sheet) throw new Error('Links sheet not found.');
  sheet.deleteRow(rowNum);
  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * Toggles the pinned status of a link.
 */
function togglePin(rowNum) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  if (!sheet) throw new Error('Links sheet not found.');

  var cell = sheet.getRange(rowNum, 5);
  var current = cell.getValue();
  var newVal = !(current === true || current === 'TRUE');
  cell.setValue(newVal);
  SpreadsheetApp.flush();
  return { success: true, pinned: newVal };
}
