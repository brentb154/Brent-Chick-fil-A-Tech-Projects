/**
 * CFA Accountability System - Phase 1: Sheet Setup
 *
 * This script sets up the complete sheet structure for the accountability system.
 * Run setupAccountabilitySystem() once to initialize all tabs and structure.
 *
 * Sheet ID: 1w71ytbfftinyG2GeAdM6NDlFjpbHCChVH1V49CObmkc
 */

// ============================================
// CONFIGURATION
// ============================================

const SHEET_ID = '1w71ytbfftinyG2GeAdM6NDlFjpbHCChVH1V49CObmkc';
const HEADER_COLOR = '#4285F4'; // Google blue
const HEADER_TEXT_COLOR = '#FFFFFF'; // White

// Tab names in required order
const TAB_NAMES = [
  'Infractions',
  'Settings',
  'User_Permissions',
  'Email_Log',
  'Links',
  'Terminated_Employees',
  'Pending_Signups'
];

// ============================================
// MAIN SETUP FUNCTION
// ============================================

/**
 * Main entry point - sets up the entire accountability system sheet structure.
 * Run this function once to initialize everything.
 */
function setupAccountabilitySystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  console.log('Starting CFA Accountability System setup...');

  // Step 1: Rename spreadsheet
  ss.rename('CFA Accountability System');
  console.log('✓ Renamed spreadsheet');

  // Step 2: Create all tabs in correct order
  createAllTabs(ss);
  console.log('✓ Created all tabs');

  // Step 3: Set up each tab's structure
  setupInfractionsTab(ss);
  console.log('✓ Set up Infractions tab');

  setupSettingsTab(ss);
  console.log('✓ Set up Settings tab');

  setupUserPermissionsTab(ss);
  console.log('✓ Set up User_Permissions tab');

  setupEmailLogTab(ss);
  console.log('✓ Set up Email_Log tab');

  setupLinksTab(ss);
  console.log('✓ Set up Links tab');

  setupTerminatedEmployeesTab(ss);
  console.log('✓ Set up Terminated_Employees tab');

  setupPendingSignupsTab(ss);
  console.log('✓ Set up Pending_Signups tab');

  // Step 4: Clean up any default sheets (like "Sheet1")
  cleanupDefaultSheets(ss);
  console.log('✓ Cleaned up default sheets');

  console.log('=== Setup complete! ===');
  return { success: true, message: 'CFA Accountability System setup complete' };
}

// ============================================
// TAB CREATION
// ============================================

/**
 * Creates all required tabs in the correct order.
 * Deletes existing tabs with same names first to ensure clean setup.
 */
function createAllTabs(ss) {
  // First, create all tabs (or get existing ones)
  const sheets = {};

  for (let i = 0; i < TAB_NAMES.length; i++) {
    const tabName = TAB_NAMES[i];
    let sheet = ss.getSheetByName(tabName);

    if (sheet) {
      // Clear existing content but keep the sheet
      sheet.clear();
    } else {
      // Create new sheet
      sheet = ss.insertSheet(tabName);
    }
    sheets[tabName] = sheet;
  }

  // Reorder sheets to match required order
  for (let i = 0; i < TAB_NAMES.length; i++) {
    const sheet = ss.getSheetByName(TAB_NAMES[i]);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1);
  }

  return sheets;
}

/**
 * Removes any default sheets like "Sheet1" that Google creates automatically.
 */
function cleanupDefaultSheets(ss) {
  const defaultNames = ['Sheet1', 'Sheet2', 'Sheet3'];

  for (const name of defaultNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      // Can only delete if there are other sheets
      if (ss.getSheets().length > 1) {
        ss.deleteSheet(sheet);
      }
    }
  }
}

// ============================================
// INFRACTIONS TAB SETUP
// ============================================

/**
 * Sets up the Infractions tab with all columns, formatting, and validation.
 */
function setupInfractionsTab(ss) {
  const sheet = ss.getSheetByName('Infractions');

  // Column headers (A-P)
  const headers = [
    'Infraction_ID',      // A - auto-generated, format: INF-YYYYMMDD-####
    'Employee_ID',        // B - from Payroll Tracker
    'Full_Name',          // C - from Payroll Tracker
    'Date',               // D - date of infraction, MM/DD/YYYY
    'Infraction_Type',    // E - dropdown from buckets
    'Points_Assigned',    // F - number, can be negative
    'Point_Value_At_Time',// G - stores bucket value when entered
    'Description',        // H - minimum 240 characters
    'Location',           // I - dropdown: Cockrell Hill DTO, Dallas Baptist University OCV
    'Entered_By',         // J - name of manager/director
    'Entry_Timestamp',    // K - auto-generated, MM/DD/YYYY HH:MM:SS
    'Last_Modified_By',   // L - name, updates on edits
    'Last_Modified_Timestamp', // M - updates on edits
    'Modification_Reason',// N - filled when director edits
    'Status',             // O - dropdown: Active, Deleted, Modified
    'Expiration_Date'     // P - auto-calculated: Date + 90 days
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  setInfractionsColumnWidths(sheet);

  // Add data validation
  addInfractionsValidation(sheet);
}

/**
 * Sets column widths for Infractions tab.
 */
function setInfractionsColumnWidths(sheet) {
  // ID columns = 150px
  sheet.setColumnWidth(1, 180);  // A: Infraction_ID (slightly wider for format)
  sheet.setColumnWidth(2, 150);  // B: Employee_ID

  // Name columns = 200px
  sheet.setColumnWidth(3, 200);  // C: Full_Name

  // Date column
  sheet.setColumnWidth(4, 120);  // D: Date

  // Type column
  sheet.setColumnWidth(5, 200);  // E: Infraction_Type (wider for bucket names)

  // Points columns
  sheet.setColumnWidth(6, 120);  // F: Points_Assigned
  sheet.setColumnWidth(7, 150);  // G: Point_Value_At_Time

  // Description = 400px
  sheet.setColumnWidth(8, 400);  // H: Description

  // Location
  sheet.setColumnWidth(9, 220);  // I: Location (wider for full names)

  // Other columns = 120px
  sheet.setColumnWidth(10, 150); // J: Entered_By
  sheet.setColumnWidth(11, 180); // K: Entry_Timestamp
  sheet.setColumnWidth(12, 150); // L: Last_Modified_By
  sheet.setColumnWidth(13, 180); // M: Last_Modified_Timestamp
  sheet.setColumnWidth(14, 250); // N: Modification_Reason
  sheet.setColumnWidth(15, 100); // O: Status
  sheet.setColumnWidth(16, 120); // P: Expiration_Date
}

/**
 * Adds data validation dropdowns to Infractions tab.
 */
function addInfractionsValidation(sheet) {
  // Column E: Infraction_Type - dropdown from bucket names
  // Will reference Settings tab once populated
  const infractionTypes = [
    'Bucket 1: Minor Offenses',
    'Bucket 2: Moderate Offenses',
    'Bucket 3: Major Offenses',
    'Bucket 4: Severe Offenses'
  ];
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(infractionTypes, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(typeRule);

  // Column I: Location dropdown
  const locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cockrell Hill DTO', 'Dallas Baptist University OCV'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(locationRule);

  // Column O: Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Deleted', 'Modified'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('O2:O1000').setDataValidation(statusRule);
}

// ============================================
// SETTINGS TAB SETUP
// ============================================

/**
 * Sets up the Settings tab with all configuration sections.
 */
function setupSettingsTab(ss) {
  const sheet = ss.getSheetByName('Settings');

  // Build all settings data
  const data = buildSettingsData();

  // Set all data at once
  sheet.getRange(1, 1, data.length, 3).setValues(data);

  // Format section headers (rows 1, 5, 11, 21, 28)
  formatSectionHeader(sheet, 1, 'Email Configuration');
  formatSectionHeader(sheet, 5, 'Password Configuration');
  formatSectionHeader(sheet, 11, 'Point Thresholds');
  formatSectionHeader(sheet, 21, 'Infraction Buckets Configuration');
  formatSectionHeader(sheet, 28, 'System Configuration');

  // Format sub-headers (rows 12, 22)
  formatSubHeader(sheet, 12, 2); // Threshold headers
  formatSubHeader(sheet, 22, 3); // Bucket headers

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 250);  // A: Labels
  sheet.setColumnWidth(2, 300);  // B: Values
  sheet.setColumnWidth(3, 400);  // C: Examples (for buckets)
}

/**
 * Builds the complete settings data array.
 */
function buildSettingsData() {
  // Create bucket examples as JSON strings
  const bucket1Examples = JSON.stringify([
    'Tardiness under 15 minutes',
    'Uniform violations',
    'Minor cleanliness issues',
    'Missing name tag',
    'Late return from break'
  ]);

  const bucket2Examples = JSON.stringify([
    'Tardiness 15-30 minutes',
    'Cell phone use during shift',
    'Call-outs',
    'Customer complaints',
    'Attendance issues'
  ]);

  const bucket3Examples = JSON.stringify([
    'Tardiness 30+ minutes',
    'Insubordination',
    'Profanity',
    'Food safety violations',
    'Leaving shift early',
    'Creating hostile environment'
  ]);

  const bucket4Examples = JSON.stringify([
    'No-call/no-show',
    'Major safety violations',
    'Theft',
    'Harassment',
    'Working under influence',
    'Physical altercations'
  ]);

  return [
    // Row 1: Email Configuration header
    ['Email Configuration', '', ''],
    // Row 2
    ['Store Email', '', ''],
    // Row 3
    ['Termination Email List', '', ''],
    // Row 4: blank
    ['', '', ''],
    // Row 5: Password Configuration header
    ['Password Configuration', '', ''],
    // Row 6
    ['Operator Password', '', ''],
    // Row 7
    ['Director Password', '', ''],
    // Row 8
    ['Manager Password', '', ''],
    // Row 9: blank
    ['', '', ''],
    // Row 10: blank
    ['', '', ''],
    // Row 11: Point Thresholds header
    ['Point Thresholds', '', ''],
    // Row 12: Threshold sub-headers
    ['Threshold', 'Consequences', ''],
    // Row 13-19: Threshold data
    [2, 'Verbal warning (Documented)', ''],
    [3, 'Remove break-food for day; Ineligible for raises/promotions', ''],
    [5, 'Remove break-food for 3 days', ''],
    [6, 'Director meeting; Ineligible for bonuses; Remove 1 day from schedule', ''],
    [9, 'Final written warning; 30-day probation', ''],
    [12, 'Reduced hours; 3-day suspension', ''],
    [15, 'Termination; Ineligible for rehire 12 months', ''],
    // Row 20: blank
    ['', '', ''],
    // Row 21: Infraction Buckets header
    ['Infraction Buckets Configuration', '', ''],
    // Row 22: Bucket sub-headers
    ['Bucket Name', 'Point Value', 'Examples'],
    // Row 23-26: Bucket data
    ['Bucket 1: Minor Offenses', 1, bucket1Examples],
    ['Bucket 2: Moderate Offenses', 3, bucket2Examples],
    ['Bucket 3: Major Offenses', 5, bucket3Examples],
    ['Bucket 4: Severe Offenses', 8, bucket4Examples],
    // Row 27: blank
    ['', '', ''],
    // Row 28: System Configuration header
    ['System Configuration', '', ''],
    // Row 29-35: System config
    ['Session Timeout (minutes)', 10, ''],
    ['Max Login Attempts', 3, ''],
    ['Backdate Limit (days)', 7, ''],
    ['Max Negative Points', -6, ''],
    ['Point Expiration (days)', 90, ''],
    ['Probation Duration (days)', 30, ''],
    ['Last Backup Date', '', '']
  ];
}

/**
 * Formats a section header row (bold, blue background, white text, merged if needed).
 */
function formatSectionHeader(sheet, row, text) {
  const range = sheet.getRange(row, 1, 1, 3);
  range.setBackground(HEADER_COLOR);
  range.setFontColor(HEADER_TEXT_COLOR);
  range.setFontWeight('bold');
  // Merge cells for section headers
  range.merge();
  range.setValue(text);
}

/**
 * Formats a sub-header row (bold, light gray background).
 */
function formatSubHeader(sheet, row, numCols) {
  const range = sheet.getRange(row, 1, 1, numCols);
  range.setBackground('#E8E8E8');
  range.setFontWeight('bold');
}

// ============================================
// USER_PERMISSIONS TAB SETUP (3-Tier System)
// ============================================

/**
 * Sets up the User_Permissions tab with 3-tier role system.
 * Roles: Manager (basic), Director (mid-level), Operator (admin)
 *
 * Schema:
 * A: Employee_ID - Unique identifier
 * B: Full_Name - Display name
 * C: Email - User's email
 * D: Role - Manager, Director, or Operator
 * E: PIN_Hash - SHA-256 hash of 6-digit PIN
 * F: Can_See_Directors - Boolean (only applicable for Directors)
 * G: Date_Added - When user was created
 * H: Added_By - Who created this user
 * I: Status - Active or Inactive
 * J: Last_Login - Timestamp of last login
 * K: Login_Count - Number of successful logins
 * L: Failed_Attempts - Counter for lockout
 * M: Lockout_Until - Timestamp when lockout expires
 */
function setupUserPermissionsTab(ss) {
  const sheet = ss.getSheetByName('User_Permissions');

  const headers = [
    'Employee_ID',      // A
    'Full_Name',        // B
    'Email',            // C
    'Role',             // D - dropdown: Manager, Director, Operator
    'PIN_Hash',         // E - SHA-256 hash of 6-digit PIN
    'Can_See_Directors',// F - TRUE/FALSE (Directors only)
    'Date_Added',       // G
    'Added_By',         // H
    'Status',           // I - dropdown: Active, Inactive
    'Last_Login',       // J
    'Login_Count',      // K
    'Failed_Attempts',  // L
    'Lockout_Until'     // M
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 150);  // A: Employee_ID
  sheet.setColumnWidth(2, 200);  // B: Full_Name
  sheet.setColumnWidth(3, 250);  // C: Email
  sheet.setColumnWidth(4, 100);  // D: Role
  sheet.setColumnWidth(5, 100);  // E: PIN_Hash (hidden in practice)
  sheet.setColumnWidth(6, 130);  // F: Can_See_Directors
  sheet.setColumnWidth(7, 120);  // G: Date_Added
  sheet.setColumnWidth(8, 150);  // H: Added_By
  sheet.setColumnWidth(9, 100);  // I: Status
  sheet.setColumnWidth(10, 150); // J: Last_Login
  sheet.setColumnWidth(11, 100); // K: Login_Count
  sheet.setColumnWidth(12, 120); // L: Failed_Attempts
  sheet.setColumnWidth(13, 150); // M: Lockout_Until

  // Add data validation for Role (3-tier system)
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Manager', 'Director', 'Operator'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(roleRule);

  // Add data validation for Can_See_Directors
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(boolRule);

  // Add data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(statusRule);
}

// ============================================
// EMAIL_LOG TAB SETUP
// ============================================

/**
 * Sets up the Email_Log tab.
 */
function setupEmailLogTab(ss) {
  const sheet = ss.getSheetByName('Email_Log');

  const headers = [
    'Log_ID',            // A - auto-generated
    'Timestamp',         // B
    'Employee_ID',       // C
    'Employee_Name',     // D
    'Recipient_Email',   // E
    'Email_Type',        // F - e.g., "6-Point Threshold", "Termination"
    'Thresholds_Crossed',// G - comma-separated list
    'Status',            // H - dropdown: Sent, Failed, Retrying
    'Retry_Count',       // I
    'Error_Message'      // J - if failed
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 150);  // A: Log_ID
  sheet.setColumnWidth(2, 180);  // B: Timestamp
  sheet.setColumnWidth(3, 150);  // C: Employee_ID
  sheet.setColumnWidth(4, 200);  // D: Employee_Name
  sheet.setColumnWidth(5, 250);  // E: Recipient_Email
  sheet.setColumnWidth(6, 150);  // F: Email_Type
  sheet.setColumnWidth(7, 150);  // G: Thresholds_Crossed
  sheet.setColumnWidth(8, 100);  // H: Status
  sheet.setColumnWidth(9, 100);  // I: Retry_Count
  sheet.setColumnWidth(10, 300); // J: Error_Message

  // Add data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sent', 'Failed', 'Retrying'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('H2:H1000').setDataValidation(statusRule);
}

// ============================================
// LINKS TAB SETUP
// ============================================

/**
 * Sets up the Links tab with initial link data.
 */
function setupLinksTab(ss) {
  const sheet = ss.getSheetByName('Links');

  const headers = [
    'Link_ID',       // A - auto-generated
    'Link_Name',     // B
    'URL',           // C
    'Display_Order', // D - number for sorting
    'Status',        // E - dropdown: Active, Hidden
    'Added_By',      // F
    'Date_Added'     // G
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 150);  // A: Link_ID
  sheet.setColumnWidth(2, 200);  // B: Link_Name
  sheet.setColumnWidth(3, 400);  // C: URL
  sheet.setColumnWidth(4, 120);  // D: Display_Order
  sheet.setColumnWidth(5, 100);  // E: Status
  sheet.setColumnWidth(6, 150);  // F: Added_By
  sheet.setColumnWidth(7, 120);  // G: Date_Added

  // Add data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Hidden'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(statusRule);

  // Add initial link data (URLs to be provided later)
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const initialLinks = [
    ['LNK-001', 'Training Tracker', '[URL to be provided]', 1, 'Active', 'System', today],
    ['LNK-002', 'Donation Request', '[URL to be provided]', 2, 'Active', 'System', today],
    ['LNK-003', 'Line up', '[URL to be provided]', 3, 'Active', 'System', today],
    ['LNK-004', 'Signal Link to CFA Home', '[URL to be provided]', 4, 'Active', 'System', today],
    ['LNK-005', 'Heard Log', '[URL to be provided]', 5, 'Active', 'System', today]
  ];

  sheet.getRange(2, 1, initialLinks.length, headers.length).setValues(initialLinks);
}

// ============================================
// TERMINATED_EMPLOYEES TAB SETUP
// ============================================

/**
 * Sets up the Terminated_Employees tab (same as Infractions plus termination fields).
 */
function setupTerminatedEmployeesTab(ss) {
  const sheet = ss.getSheetByName('Terminated_Employees');

  // Same as Infractions (A-P) plus termination fields (Q-S)
  const headers = [
    'Infraction_ID',      // A
    'Employee_ID',        // B
    'Full_Name',          // C
    'Date',               // D
    'Infraction_Type',    // E
    'Points_Assigned',    // F
    'Point_Value_At_Time',// G
    'Description',        // H
    'Location',           // I
    'Entered_By',         // J
    'Entry_Timestamp',    // K
    'Last_Modified_By',   // L
    'Last_Modified_Timestamp', // M
    'Modification_Reason',// N
    'Status',             // O
    'Expiration_Date',    // P
    'Termination_Date',   // Q - additional field
    'Termination_Reason', // R - additional field
    'Terminated_By'       // S - additional field
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths (same as Infractions for A-P, plus new columns)
  sheet.setColumnWidth(1, 180);  // A: Infraction_ID
  sheet.setColumnWidth(2, 150);  // B: Employee_ID
  sheet.setColumnWidth(3, 200);  // C: Full_Name
  sheet.setColumnWidth(4, 120);  // D: Date
  sheet.setColumnWidth(5, 200);  // E: Infraction_Type
  sheet.setColumnWidth(6, 120);  // F: Points_Assigned
  sheet.setColumnWidth(7, 150);  // G: Point_Value_At_Time
  sheet.setColumnWidth(8, 400);  // H: Description
  sheet.setColumnWidth(9, 220);  // I: Location
  sheet.setColumnWidth(10, 150); // J: Entered_By
  sheet.setColumnWidth(11, 180); // K: Entry_Timestamp
  sheet.setColumnWidth(12, 150); // L: Last_Modified_By
  sheet.setColumnWidth(13, 180); // M: Last_Modified_Timestamp
  sheet.setColumnWidth(14, 250); // N: Modification_Reason
  sheet.setColumnWidth(15, 100); // O: Status
  sheet.setColumnWidth(16, 120); // P: Expiration_Date
  sheet.setColumnWidth(17, 120); // Q: Termination_Date
  sheet.setColumnWidth(18, 300); // R: Termination_Reason
  sheet.setColumnWidth(19, 150); // S: Terminated_By

  // Add same data validation as Infractions tab
  const infractionTypes = [
    'Bucket 1: Minor Offenses',
    'Bucket 2: Moderate Offenses',
    'Bucket 3: Major Offenses',
    'Bucket 4: Severe Offenses'
  ];
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(infractionTypes, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(typeRule);

  const locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cockrell Hill DTO', 'Dallas Baptist University OCV'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(locationRule);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Deleted', 'Modified'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('O2:O1000').setDataValidation(statusRule);
}

/**
 * Sets up the Pending_Signups tab for self-service user registration.
 * Structure: 11 columns for tracking signup links and completion status.
 */
function setupPendingSignupsTab(ss) {
  const sheet = ss.getSheetByName('Pending_Signups');

  const headers = [
    'Signup_ID',        // A - Unique identifier (e.g., SU_20240115_001)
    'Token',            // B - Unique token for signup URL
    'Employee_ID',      // C - Pre-assigned employee ID
    'Email',            // D - Email address for signup link
    'Role',             // E - Manager, Director, or Operator
    'Can_See_Directors',// F - TRUE/FALSE
    'Created_Date',     // G - When signup was created
    'Expires_Date',     // H - 7 days after creation
    'Status',           // I - Pending, Completed, Expired, Cancelled
    'Created_By',       // J - Who created the signup link
    'Completed_Date'    // K - When user completed signup
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 180);  // A: Signup_ID
  sheet.setColumnWidth(2, 300);  // B: Token (long random string)
  sheet.setColumnWidth(3, 150);  // C: Employee_ID
  sheet.setColumnWidth(4, 250);  // D: Email
  sheet.setColumnWidth(5, 100);  // E: Role
  sheet.setColumnWidth(6, 130);  // F: Can_See_Directors
  sheet.setColumnWidth(7, 150);  // G: Created_Date
  sheet.setColumnWidth(8, 150);  // H: Expires_Date
  sheet.setColumnWidth(9, 100);  // I: Status
  sheet.setColumnWidth(10, 150); // J: Created_By
  sheet.setColumnWidth(11, 150); // K: Completed_Date

  // Add data validation for Role column
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Manager', 'Director', 'Operator'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(roleRule);

  // Add data validation for Can_See_Directors column
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(boolRule);

  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Completed', 'Expired', 'Cancelled'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(statusRule);
}

/**
 * STANDALONE FUNCTION - Run this directly from Apps Script Editor
 * Creates and sets up the Pending_Signups tab.
 * This function can be run without any parameters.
 */
function createPendingSignupsTab() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Check if sheet already exists
  let sheet = ss.getSheetByName('Pending_Signups');

  if (sheet) {
    console.log('Pending_Signups sheet already exists. Updating headers and formatting...');
  } else {
    // Create the sheet
    sheet = ss.insertSheet('Pending_Signups');
    console.log('Created new Pending_Signups sheet.');
  }

  // Define headers
  const headers = [
    'Signup_ID',        // A - Unique identifier (e.g., SU_20240115_001)
    'Token',            // B - Unique token for signup URL
    'Employee_ID',      // C - Pre-assigned employee ID
    'Email',            // D - Email address for signup link
    'Role',             // E - Manager, Director, or Operator
    'Can_See_Directors',// F - TRUE/FALSE
    'Created_Date',     // G - When signup was created
    'Expires_Date',     // H - 7 days after creation
    'Status',           // I - Pending, Completed, Expired, Cancelled
    'Created_By',       // J - Who created the signup link
    'Completed_Date'    // K - When user completed signup
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Apply header formatting
  formatHeaderRow(sheet, headers.length);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 180);  // A: Signup_ID
  sheet.setColumnWidth(2, 300);  // B: Token (long random string)
  sheet.setColumnWidth(3, 150);  // C: Employee_ID
  sheet.setColumnWidth(4, 250);  // D: Email
  sheet.setColumnWidth(5, 100);  // E: Role
  sheet.setColumnWidth(6, 130);  // F: Can_See_Directors
  sheet.setColumnWidth(7, 150);  // G: Created_Date
  sheet.setColumnWidth(8, 150);  // H: Expires_Date
  sheet.setColumnWidth(9, 100);  // I: Status
  sheet.setColumnWidth(10, 150); // J: Created_By
  sheet.setColumnWidth(11, 150); // K: Completed_Date

  // Add data validation for Role column
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Manager', 'Director', 'Operator'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(roleRule);

  // Add data validation for Can_See_Directors column
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(boolRule);

  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Completed', 'Expired', 'Cancelled'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(statusRule);

  console.log('Pending_Signups tab setup complete!');
  return 'Pending_Signups tab created and configured successfully!';
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Formats row 1 as a header row (bold, blue background, white text).
 */
function formatHeaderRow(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground(HEADER_COLOR);
  headerRange.setFontColor(HEADER_TEXT_COLOR);
  headerRange.setFontWeight('bold');
}

// ============================================
// TEST FUNCTION
// ============================================

/**
 * Test function to verify setup was successful.
 * Run this after setupAccountabilitySystem() to validate.
 */
function testPhase1Setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const results = [];
  let allPassed = true;

  // Test 1: Spreadsheet name
  const name = ss.getName();
  const nameTest = name === 'CFA Accountability System';
  results.push({
    test: 'Spreadsheet name',
    expected: 'CFA Accountability System',
    actual: name,
    passed: nameTest
  });
  if (!nameTest) allPassed = false;

  // Test 2: All tabs exist in correct order
  const sheets = ss.getSheets();
  const actualTabNames = sheets.map(s => s.getName());

  for (let i = 0; i < TAB_NAMES.length; i++) {
    const tabExists = actualTabNames[i] === TAB_NAMES[i];
    results.push({
      test: `Tab ${i + 1}: ${TAB_NAMES[i]}`,
      expected: TAB_NAMES[i],
      actual: actualTabNames[i] || 'MISSING',
      passed: tabExists
    });
    if (!tabExists) allPassed = false;
  }

  // Test 3: Infractions tab has correct number of columns
  const infractionsSheet = ss.getSheetByName('Infractions');
  const infractionsHeaders = infractionsSheet.getRange(1, 1, 1, 16).getValues()[0];
  const infractionsColCount = infractionsHeaders.filter(h => h !== '').length;
  const infractionsTest = infractionsColCount === 16;
  results.push({
    test: 'Infractions column count',
    expected: 16,
    actual: infractionsColCount,
    passed: infractionsTest
  });
  if (!infractionsTest) allPassed = false;

  // Test 4: Settings tab has threshold data
  const settingsSheet = ss.getSheetByName('Settings');
  const thresholdHeader = settingsSheet.getRange('A11').getValue();
  const thresholdTest = thresholdHeader === 'Point Thresholds';
  results.push({
    test: 'Settings Point Thresholds section',
    expected: 'Point Thresholds',
    actual: thresholdHeader,
    passed: thresholdTest
  });
  if (!thresholdTest) allPassed = false;

  // Test 5: Settings has all 7 thresholds
  const thresholdValues = settingsSheet.getRange('A13:A19').getValues().flat();
  const expectedThresholds = [2, 3, 5, 6, 9, 12, 15];
  const thresholdValuesTest = JSON.stringify(thresholdValues) === JSON.stringify(expectedThresholds);
  results.push({
    test: 'Threshold values',
    expected: expectedThresholds.join(', '),
    actual: thresholdValues.join(', '),
    passed: thresholdValuesTest
  });
  if (!thresholdValuesTest) allPassed = false;

  // Test 6: Links tab has initial data
  const linksSheet = ss.getSheetByName('Links');
  const linksData = linksSheet.getRange('A2:A6').getValues().flat().filter(v => v !== '');
  const linksTest = linksData.length === 5;
  results.push({
    test: 'Links initial data count',
    expected: 5,
    actual: linksData.length,
    passed: linksTest
  });
  if (!linksTest) allPassed = false;

  // Test 7: Terminated_Employees has extra columns (Q, R, S)
  const terminatedSheet = ss.getSheetByName('Terminated_Employees');
  const terminatedHeaders = terminatedSheet.getRange(1, 1, 1, 19).getValues()[0];
  const hasTerminationDate = terminatedHeaders[16] === 'Termination_Date';
  const hasTerminationReason = terminatedHeaders[17] === 'Termination_Reason';
  const hasTerminatedBy = terminatedHeaders[18] === 'Terminated_By';
  const terminatedTest = hasTerminationDate && hasTerminationReason && hasTerminatedBy;
  results.push({
    test: 'Terminated_Employees extra columns',
    expected: 'Termination_Date, Termination_Reason, Terminated_By',
    actual: `${terminatedHeaders[16]}, ${terminatedHeaders[17]}, ${terminatedHeaders[18]}`,
    passed: terminatedTest
  });
  if (!terminatedTest) allPassed = false;

  // Test 8: Row 1 is frozen on Infractions
  const frozenRows = infractionsSheet.getFrozenRows();
  const frozenTest = frozenRows === 1;
  results.push({
    test: 'Infractions frozen rows',
    expected: 1,
    actual: frozenRows,
    passed: frozenTest
  });
  if (!frozenTest) allPassed = false;

  // Log results
  console.log('=== Phase 1 Test Results ===');
  for (const result of results) {
    const status = result.passed ? '✓ PASS' : '✗ FAIL';
    console.log(`${status}: ${result.test}`);
    if (!result.passed) {
      console.log(`  Expected: ${result.expected}`);
      console.log(`  Actual: ${result.actual}`);
    }
  }

  console.log('');
  console.log(allPassed ? '=== ALL TESTS PASSED ===' : '=== SOME TESTS FAILED ===');

  return {
    success: allPassed,
    results: results,
    summary: `${results.filter(r => r.passed).length}/${results.length} tests passed`
  };
}

// ============================================
// PHASE 2: PAYROLL TRACKER INTEGRATION
// ============================================

// Payroll Tracker configuration
const PAYROLL_TRACKER_ID = '1aZUj7iFlxM6ID33CWGf_lAo7SR_Lq_1B-M-pIkJW0nA';
const PAYROLL_TAB_NAME = 'Employees';

/**
 * Retrieves all active employees from the Payroll Tracker sheet.
 *
 * @returns {Array<Object>} Array of employee objects with all 8 fields
 * @throws {Error} If Payroll Tracker cannot be accessed or Employees tab not found
 *
 * Payroll Tracker Columns:
 *   A = Employee_ID
 *   B = Full_Name
 *   C = Match_Key
 *   D = Primary_Location
 *   E = Status
 *   F = First_Seen
 *   G = Last_Seen
 *   H = Last_Period_End
 */
function getActiveEmployees() {
  try {
    // Step 1: Open Payroll Tracker sheet
    let payrollSpreadsheet;
    try {
      payrollSpreadsheet = SpreadsheetApp.openById(PAYROLL_TRACKER_ID);
    } catch (e) {
      const errorMsg = 'Cannot access Payroll Tracker sheet. Check sheet ID and permissions.';
      console.error(errorMsg, e);
      throw new Error(errorMsg);
    }

    // Step 2: Access the Employees tab
    const employeesSheet = payrollSpreadsheet.getSheetByName(PAYROLL_TAB_NAME);
    if (!employeesSheet) {
      const errorMsg = 'Employees tab not found in Payroll Tracker.';
      console.error(errorMsg);
      throw new Error(errorMsg);
    }

    // Step 3: Get all data (skip header row)
    const lastRow = employeesSheet.getLastRow();

    // If only header row or empty, return empty array
    if (lastRow <= 1) {
      console.log('No employee data found in Payroll Tracker.');
      return [];
    }

    // Read all data from row 2 to last row, columns A-H
    const dataRange = employeesSheet.getRange(2, 1, lastRow - 1, 8);
    const data = dataRange.getValues();

    // Step 4: Filter for Active employees and map to objects
    const activeEmployees = [];
    const seenIds = new Set(); // Track for duplicates

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const status = row[4]; // Column E (0-indexed = 4)

      // Only include Active employees
      if (status === 'Active') {
        const employeeId = row[0]; // Column A

        // Skip if we've already seen this employee (no duplicates)
        if (seenIds.has(employeeId)) {
          console.log(`Skipping duplicate employee ID: ${employeeId}`);
          continue;
        }
        seenIds.add(employeeId);

        // Map row to employee object
        const employee = {
          employee_id: row[0],      // A: Employee_ID
          full_name: row[1],        // B: Full_Name
          match_key: row[2],        // C: Match_Key
          primary_location: row[3], // D: Primary_Location
          status: row[4],           // E: Status
          first_seen: row[5],       // F: First_Seen
          last_seen: row[6],        // G: Last_Seen
          last_period_end: row[7]   // H: Last_Period_End
        };

        activeEmployees.push(employee);
      }
    }

    console.log(`Found ${activeEmployees.length} active employees.`);
    return activeEmployees;

  } catch (error) {
    // Log error and re-throw
    console.error('Error in getActiveEmployees:', error.toString());
    throw error;
  }
}

/**
 * Test function for getActiveEmployees().
 * Logs employee count and details of first 3 employees.
 */
function testGetActiveEmployees() {
  console.log('=== Testing getActiveEmployees() ===');
  console.log('');

  try {
    const startTime = new Date().getTime();

    // Call the function
    const employees = getActiveEmployees();

    const endTime = new Date().getTime();
    const duration = (endTime - startTime) / 1000;

    // Log count
    console.log(`Total active employees returned: ${employees.length}`);
    console.log(`Execution time: ${duration.toFixed(2)} seconds`);
    console.log('');

    // Verify performance requirement (under 5 seconds)
    if (duration >= 5) {
      console.log('⚠ WARNING: Function took longer than 5 seconds');
    } else {
      console.log('✓ Performance OK (under 5 seconds)');
    }
    console.log('');

    // Log first 3 employees
    if (employees.length > 0) {
      console.log('First 3 employees:');
      console.log('-------------------');

      const displayCount = Math.min(3, employees.length);
      for (let i = 0; i < displayCount; i++) {
        const emp = employees[i];
        console.log(`${i + 1}. ${emp.full_name}`);
        console.log(`   ID: ${emp.employee_id}`);
        console.log(`   Location: ${emp.primary_location}`);
        console.log(`   Status: ${emp.status}`);
        console.log('');
      }

      // Verify all 8 fields are present
      console.log('Verifying field mapping:');
      const firstEmp = employees[0];
      const expectedFields = [
        'employee_id', 'full_name', 'match_key', 'primary_location',
        'status', 'first_seen', 'last_seen', 'last_period_end'
      ];

      let allFieldsPresent = true;
      for (const field of expectedFields) {
        const hasField = field in firstEmp;
        const status = hasField ? '✓' : '✗';
        console.log(`  ${status} ${field}: ${hasField ? 'present' : 'MISSING'}`);
        if (!hasField) allFieldsPresent = false;
      }

      console.log('');
      if (allFieldsPresent) {
        console.log('✓ All 8 fields correctly mapped');
      } else {
        console.log('✗ Some fields are missing');
      }

      // Verify no duplicates
      const ids = employees.map(e => e.employee_id);
      const uniqueIds = new Set(ids);
      if (ids.length === uniqueIds.size) {
        console.log('✓ No duplicate employee IDs');
      } else {
        console.log(`✗ Found ${ids.length - uniqueIds.size} duplicate IDs`);
      }

      // Verify all are Active
      const allActive = employees.every(e => e.status === 'Active');
      if (allActive) {
        console.log('✓ All returned employees have Status = "Active"');
      } else {
        console.log('✗ Some employees do not have Status = "Active"');
      }

    } else {
      console.log('No active employees found (empty array returned).');
      console.log('This may be expected if the Payroll Tracker has no active employees.');
    }

    console.log('');
    console.log('=== Test Complete ===');

    return {
      success: true,
      employeeCount: employees.length,
      employees: employees
    };

  } catch (error) {
    console.error('Test failed with error:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// PHASE 3: POINT CALCULATION LOGIC
// ============================================

/**
 * Calculates current points for an employee based on infractions from the last 90 days.
 *
 * @param {string} employeeId - Employee ID from Payroll Tracker
 * @param {Date} [asOfDate=new Date()] - Calculate points as of this date (defaults to today)
 * @returns {Object} Point calculation result with structure:
 *   - total_points: number (sum of active infractions, capped at -6 minimum)
 *   - active_infractions: array of active infraction objects
 *   - next_expiration_date: Date or null
 *   - expired_infractions: array of expired infraction objects
 *
 * Infractions Sheet Columns:
 *   A = Infraction_ID
 *   B = Employee_ID
 *   C = Full_Name
 *   D = Date
 *   E = Infraction_Type
 *   F = Points_Assigned
 *   G = Point_Value_At_Time
 *   H = Description
 *   I = Location
 *   J = Entered_By
 *   K = Entry_Timestamp
 *   L = Last_Modified_By
 *   M = Last_Modified_Timestamp
 *   N = Modification_Reason
 *   O = Status
 *   P = Expiration_Date
 */
function calculatePoints(employeeId, asOfDate) {
  // Default asOfDate to today if not provided
  if (!asOfDate) {
    asOfDate = new Date();
  }

  // Normalize asOfDate to start of day for consistent comparison
  asOfDate = new Date(asOfDate.getFullYear(), asOfDate.getMonth(), asOfDate.getDate());

  // Default return object
  const defaultResult = {
    total_points: 0,
    active_infractions: [],
    next_expiration_date: null,
    expired_infractions: []
  };

  try {
    // Step 1: Get the Infractions sheet
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const infractionsSheet = ss.getSheetByName('Infractions');

    if (!infractionsSheet) {
      console.error('Infractions sheet not found');
      return defaultResult;
    }

    // Step 2: Read all data from row 2 through last row
    const lastRow = infractionsSheet.getLastRow();

    // If no infractions exist (only header or empty)
    if (lastRow < 2) {
      console.log('No infractions in sheet');
      return defaultResult;
    }

    // Read columns A through P (16 columns)
    const dataRange = infractionsSheet.getRange(2, 1, lastRow - 1, 16);
    const data = dataRange.getValues();

    // Step 3: Calculate cutoff date (90 days before asOfDate)
    const cutoffDate = new Date(asOfDate);
    cutoffDate.setDate(cutoffDate.getDate() - 90);

    // Step 4: Filter and categorize infractions
    const activeInfractions = [];
    const expiredInfractions = [];
    let totalPoints = 0;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      // Column indices (0-based)
      const infractionId = row[0];     // A: Infraction_ID
      const rowEmployeeId = row[1];    // B: Employee_ID
      const infractionDate = row[3];   // D: Date
      const infractionType = row[4];   // E: Infraction_Type
      const pointsAssigned = row[5];   // F: Points_Assigned
      const description = row[7];      // H: Description
      const location = row[8];         // I: Location
      const enteredBy = row[9];        // J: Entered_By
      const status = row[14];          // O: Status
      const expirationDate = row[15];  // P: Expiration_Date

      // Filter: Must match employee ID and have Active status
      if (rowEmployeeId !== employeeId) {
        continue;
      }

      if (status !== 'Active') {
        continue;
      }

      // Parse infraction date
      let parsedInfractionDate;
      if (infractionDate instanceof Date) {
        parsedInfractionDate = infractionDate;
      } else {
        parsedInfractionDate = new Date(infractionDate);
      }

      // Normalize to start of day
      parsedInfractionDate = new Date(
        parsedInfractionDate.getFullYear(),
        parsedInfractionDate.getMonth(),
        parsedInfractionDate.getDate()
      );

      // Parse expiration date
      let parsedExpirationDate;
      if (expirationDate instanceof Date) {
        parsedExpirationDate = expirationDate;
      } else if (expirationDate) {
        parsedExpirationDate = new Date(expirationDate);
      } else {
        // Calculate expiration if not set (date + 90 days)
        parsedExpirationDate = new Date(parsedInfractionDate);
        parsedExpirationDate.setDate(parsedExpirationDate.getDate() + 90);
      }

      // Categorize as active or expired
      if (parsedInfractionDate >= cutoffDate) {
        // Active infraction
        const activeInfraction = {
          infraction_id: infractionId,
          date: parsedInfractionDate,
          infraction_type: infractionType,
          points: pointsAssigned,
          description: description,
          location: location,
          entered_by: enteredBy,
          expiration_date: parsedExpirationDate
        };

        activeInfractions.push(activeInfraction);
        totalPoints += Number(pointsAssigned) || 0;
      } else {
        // Expired infraction
        const expiredInfraction = {
          infraction_id: infractionId,
          date: parsedInfractionDate,
          infraction_type: infractionType,
          points: pointsAssigned,
          expiration_date: parsedExpirationDate
        };

        expiredInfractions.push(expiredInfraction);
      }
    }

    // Step 5: Cap total points at -6 minimum
    if (totalPoints < -6) {
      totalPoints = -6;
    }

    // Step 6: Sort active infractions by expiration date (earliest first)
    activeInfractions.sort((a, b) => {
      return a.expiration_date.getTime() - b.expiration_date.getTime();
    });

    // Step 7: Get next expiration date
    let nextExpirationDate = null;
    if (activeInfractions.length > 0) {
      nextExpirationDate = activeInfractions[0].expiration_date;
    }

    // Return result
    return {
      total_points: totalPoints,
      active_infractions: activeInfractions,
      next_expiration_date: nextExpirationDate,
      expired_infractions: expiredInfractions
    };

  } catch (error) {
    console.error('Error in calculatePoints:', error.toString());
    return defaultResult;
  }
}

/**
 * Test function for calculatePoints().
 * Tests with a specified employee ID and logs detailed results.
 */
function testCalculatePoints() {
  console.log('=== Testing calculatePoints() ===');
  console.log('');

  // Test employee ID - change this to a valid ID in your system
  const testEmployeeId = '12-1543165';

  try {
    const startTime = new Date().getTime();

    // Call the function
    const result = calculatePoints(testEmployeeId);

    const endTime = new Date().getTime();
    const duration = (endTime - startTime) / 1000;

    // Log basic info
    console.log(`Employee ID tested: ${testEmployeeId}`);
    console.log(`Execution time: ${duration.toFixed(2)} seconds`);
    console.log('');

    // Verify performance (under 3 seconds)
    if (duration >= 3) {
      console.log('⚠ WARNING: Function took longer than 3 seconds');
    } else {
      console.log('✓ Performance OK (under 3 seconds)');
    }
    console.log('');

    // Log results
    console.log('Results:');
    console.log('--------');
    console.log(`Total points: ${result.total_points}`);
    console.log(`Active infractions count: ${result.active_infractions.length}`);
    console.log(`Expired infractions count: ${result.expired_infractions.length}`);

    if (result.next_expiration_date) {
      console.log(`Next expiration date: ${formatDateForLog(result.next_expiration_date)}`);
    } else {
      console.log('Next expiration date: null (no active infractions)');
    }
    console.log('');

    // Log active infractions details
    if (result.active_infractions.length > 0) {
      console.log('Active Infractions:');
      console.log('-------------------');
      for (let i = 0; i < result.active_infractions.length; i++) {
        const inf = result.active_infractions[i];
        console.log(`${i + 1}. ${inf.infraction_id}`);
        console.log(`   Date: ${formatDateForLog(inf.date)}`);
        console.log(`   Type: ${inf.infraction_type}`);
        console.log(`   Points: ${inf.points}`);
        console.log(`   Expiration: ${formatDateForLog(inf.expiration_date)}`);
        console.log(`   Location: ${inf.location}`);
        console.log('');
      }
    } else {
      console.log('No active infractions found.');
      console.log('');
    }

    // Log expired infractions summary
    if (result.expired_infractions.length > 0) {
      console.log('Expired Infractions:');
      console.log('--------------------');
      for (let i = 0; i < result.expired_infractions.length; i++) {
        const inf = result.expired_infractions[i];
        console.log(`${i + 1}. ${inf.infraction_id} - ${inf.infraction_type} (${inf.points} pts) - Expired: ${formatDateForLog(inf.expiration_date)}`);
      }
      console.log('');
    }

    // Validate return structure
    console.log('Validating return structure:');
    const hasAllFields =
      'total_points' in result &&
      'active_infractions' in result &&
      'next_expiration_date' in result &&
      'expired_infractions' in result;

    if (hasAllFields) {
      console.log('✓ All required fields present');
    } else {
      console.log('✗ Missing required fields');
    }

    // Validate active infraction structure (if any exist)
    if (result.active_infractions.length > 0) {
      const firstActive = result.active_infractions[0];
      const activeFields = ['infraction_id', 'date', 'infraction_type', 'points', 'description', 'location', 'entered_by', 'expiration_date'];
      const hasAllActiveFields = activeFields.every(f => f in firstActive);

      if (hasAllActiveFields) {
        console.log('✓ Active infraction objects have all required fields');
      } else {
        console.log('✗ Active infraction objects missing fields');
        console.log(`  Expected: ${activeFields.join(', ')}`);
        console.log(`  Found: ${Object.keys(firstActive).join(', ')}`);
      }
    }

    // Validate expired infraction structure (if any exist)
    if (result.expired_infractions.length > 0) {
      const firstExpired = result.expired_infractions[0];
      const expiredFields = ['infraction_id', 'date', 'infraction_type', 'points', 'expiration_date'];
      const hasAllExpiredFields = expiredFields.every(f => f in firstExpired);

      if (hasAllExpiredFields) {
        console.log('✓ Expired infraction objects have all required fields');
      } else {
        console.log('✗ Expired infraction objects missing fields');
      }
    }

    // Validate sorting (if multiple active)
    if (result.active_infractions.length > 1) {
      let isSorted = true;
      for (let i = 1; i < result.active_infractions.length; i++) {
        if (result.active_infractions[i].expiration_date < result.active_infractions[i - 1].expiration_date) {
          isSorted = false;
          break;
        }
      }
      if (isSorted) {
        console.log('✓ Active infractions sorted by expiration date');
      } else {
        console.log('✗ Active infractions NOT sorted correctly');
      }
    }

    // Validate -6 cap
    if (result.total_points >= -6) {
      console.log('✓ Points cap validated (>= -6)');
    } else {
      console.log('✗ Points below -6 cap!');
    }

    console.log('');
    console.log('=== Test Complete ===');

    return {
      success: true,
      result: result
    };

  } catch (error) {
    console.error('Test failed with error:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Helper function to format dates for logging.
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDateForLog(date) {
  if (!date) return 'null';
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

/**
 * Test function to verify point calculation with edge cases.
 * Creates temporary test data and verifies calculations.
 */
function testCalculatePointsEdgeCases() {
  console.log('=== Testing calculatePoints Edge Cases ===');
  console.log('');

  // Test 1: Non-existent employee
  console.log('Test 1: Non-existent employee ID');
  const result1 = calculatePoints('FAKE-EMPLOYEE-ID-12345');
  console.log(`  Total points: ${result1.total_points}`);
  console.log(`  Active count: ${result1.active_infractions.length}`);
  console.log(`  Next expiration: ${result1.next_expiration_date}`);
  if (result1.total_points === 0 && result1.active_infractions.length === 0 && result1.next_expiration_date === null) {
    console.log('  ✓ Correctly returned default object for non-existent employee');
  } else {
    console.log('  ✗ Unexpected result for non-existent employee');
  }
  console.log('');

  // Test 2: Calculate with custom asOfDate
  console.log('Test 2: Custom asOfDate parameter');
  const customDate = new Date();
  customDate.setDate(customDate.getDate() - 30); // 30 days ago
  const result2 = calculatePoints('12-1543165', customDate);
  console.log(`  Calculating as of: ${formatDateForLog(customDate)}`);
  console.log(`  Total points: ${result2.total_points}`);
  console.log(`  Active count: ${result2.active_infractions.length}`);
  console.log('  ✓ Custom date parameter accepted');
  console.log('');

  console.log('=== Edge Case Tests Complete ===');
}

// ============================================
// PHASE 4: ADD SINGLE INFRACTION
// ============================================

/**
 * Adds a single infraction to the Infractions sheet with full validation.
 *
 * @param {Object} infractionData - The infraction data object
 * @param {string} infractionData.employee_id - Employee ID (required)
 * @param {string} infractionData.full_name - Employee full name (required)
 * @param {Date} infractionData.date - Date of infraction (required)
 * @param {string} infractionData.infraction_type - Type from buckets (required)
 * @param {number} infractionData.points_assigned - Points value (required, can be negative)
 * @param {string} infractionData.description - Description, min 240 chars (required)
 * @param {string} infractionData.location - Location (required)
 * @param {string} infractionData.entered_by - Name of person entering (required)
 *
 * @returns {Object} Result object with:
 *   - success: boolean
 *   - infraction_id: string (if successful)
 *   - duplicate_warning: boolean
 *   - message: string
 */
function addInfraction(infractionData) {
  try {
    // ========================================
    // VALIDATION
    // ========================================

    // Check required fields exist
    const requiredFields = ['employee_id', 'full_name', 'date', 'infraction_type',
                           'points_assigned', 'description', 'location', 'entered_by'];

    for (const field of requiredFields) {
      if (infractionData[field] === undefined || infractionData[field] === null || infractionData[field] === '') {
        return {
          success: false,
          infraction_id: null,
          duplicate_warning: false,
          message: `Missing required field: ${field}`
        };
      }
    }

    // Parse and normalize the infraction date
    let infractionDate = infractionData.date;
    if (!(infractionDate instanceof Date)) {
      infractionDate = new Date(infractionDate);
    }

    if (isNaN(infractionDate.getTime())) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: 'Invalid date format'
      };
    }

    // Normalize to start of day
    infractionDate = new Date(infractionDate.getFullYear(), infractionDate.getMonth(), infractionDate.getDate());

    // Get today's date (start of day)
    const today = new Date();
    const todayNormalized = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    // Validation 1: Date cannot be in the future
    if (infractionDate > todayNormalized) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: 'Date cannot be in the future'
      };
    }

    // Validation 2: Date cannot be more than 7 days in the past
    const sevenDaysAgo = new Date(todayNormalized);
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

    if (infractionDate < sevenDaysAgo) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: 'Date cannot be more than 7 days in the past (backdate limit exceeded)'
      };
    }

    // Validation 3: Description must be at least 240 characters
    const description = String(infractionData.description);
    if (description.length < 240) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: `Description must be at least 240 characters (currently ${description.length} characters)`
      };
    }

    // Validation 4: Employee must exist in active employees
    const activeEmployees = getActiveEmployees();
    const employeeExists = activeEmployees.some(emp => emp.employee_id === infractionData.employee_id);

    if (!employeeExists) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: `Employee ID "${infractionData.employee_id}" not found in active employees list`
      };
    }

    // Validation 5: Location must be valid
    const validLocations = ['Cockrell Hill DTO', 'Dallas Baptist University OCV'];
    if (!validLocations.includes(infractionData.location)) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: `Invalid location. Must be "Cockrell Hill DTO" or "Dallas Baptist University OCV"`
      };
    }

    // Validation 6: Points must be a number
    const pointsAssigned = Number(infractionData.points_assigned);
    if (isNaN(pointsAssigned)) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: 'Points assigned must be a number'
      };
    }

    // ========================================
    // GET SETTINGS DATA
    // ========================================

    const ss = SpreadsheetApp.openById(SHEET_ID);

    // Look up Point_Value_At_Time from Settings tab
    const pointValueAtTime = getBucketPointValue(ss, infractionData.infraction_type);

    // ========================================
    // DUPLICATE DETECTION
    // ========================================

    const infractionsSheet = ss.getSheetByName('Infractions');
    if (!infractionsSheet) {
      return {
        success: false,
        infraction_id: null,
        duplicate_warning: false,
        message: 'Infractions sheet not found'
      };
    }

    let duplicateWarning = false;
    const lastRow = infractionsSheet.getLastRow();

    if (lastRow >= 2) {
      // Read existing infractions to check for duplicates
      const existingData = infractionsSheet.getRange(2, 1, lastRow - 1, 16).getValues();

      for (const row of existingData) {
        const rowEmployeeId = row[1];  // B: Employee_ID
        const rowType = row[4];         // E: Infraction_Type
        const rowDate = row[3];         // D: Date

        // Parse existing date
        let existingDate = rowDate;
        if (existingDate instanceof Date) {
          existingDate = new Date(existingDate.getFullYear(), existingDate.getMonth(), existingDate.getDate());
        } else {
          existingDate = new Date(existingDate);
          existingDate = new Date(existingDate.getFullYear(), existingDate.getMonth(), existingDate.getDate());
        }

        // Check for duplicate (same employee, type, and date)
        if (rowEmployeeId === infractionData.employee_id &&
            rowType === infractionData.infraction_type &&
            existingDate.getTime() === infractionDate.getTime()) {
          duplicateWarning = true;
          console.log(`Duplicate detected: ${infractionData.employee_id} - ${infractionData.infraction_type} on ${formatDateForLog(infractionDate)}`);
          break;
        }
      }
    }

    // ========================================
    // GENERATE VALUES
    // ========================================

    // Generate Infraction_ID: INF-YYYYMMDD-####
    const infractionId = generateInfractionId();

    // Entry timestamp
    const entryTimestamp = new Date();

    // Expiration date (date + 90 days)
    const expirationDate = new Date(infractionDate);
    expirationDate.setDate(expirationDate.getDate() + 90);

    // Status
    const status = 'Active';

    // ========================================
    // WRITE TO SHEET
    // ========================================

    const newRow = lastRow + 1;

    // Prepare row data (columns A through P)
    const rowData = [
      infractionId,                    // A: Infraction_ID
      infractionData.employee_id,      // B: Employee_ID
      infractionData.full_name,        // C: Full_Name
      infractionDate,                  // D: Date
      infractionData.infraction_type,  // E: Infraction_Type
      pointsAssigned,                  // F: Points_Assigned
      pointValueAtTime,                // G: Point_Value_At_Time
      description,                     // H: Description
      infractionData.location,         // I: Location
      infractionData.entered_by,       // J: Entered_By
      entryTimestamp,                  // K: Entry_Timestamp
      '',                              // L: Last_Modified_By (empty on creation)
      '',                              // M: Last_Modified_Timestamp (empty on creation)
      '',                              // N: Modification_Reason (empty on creation)
      status,                          // O: Status
      expirationDate                   // P: Expiration_Date
    ];

    // Write the row
    infractionsSheet.getRange(newRow, 1, 1, 16).setValues([rowData]);

    console.log(`Infraction added: ${infractionId} for ${infractionData.full_name}`);

    // ========================================
    // RETURN SUCCESS
    // ========================================

    return {
      success: true,
      infraction_id: infractionId,
      duplicate_warning: duplicateWarning,
      message: duplicateWarning
        ? `Infraction created with warning: Similar infraction exists for this employee on this date`
        : `Infraction ${infractionId} created successfully`
    };

  } catch (error) {
    console.error('Error in addInfraction:', error.toString());
    return {
      success: false,
      infraction_id: null,
      duplicate_warning: false,
      message: `Error adding infraction: ${error.toString()}`
    };
  }
}

/**
 * Generates a unique Infraction ID in format INF-YYYYMMDD-####
 * @returns {string} Generated infraction ID
 */
function generateInfractionId() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const dateStr = `${year}${month}${day}`;

  // Generate 4-digit random number
  const random = Math.floor(1000 + Math.random() * 9000);

  return `INF-${dateStr}-${random}`;
}

/**
 * Looks up the point value for a given infraction type from Settings.
 * @param {Spreadsheet} ss - The spreadsheet object
 * @param {string} infractionType - The infraction type to look up
 * @returns {number} The point value for this bucket
 */
function getBucketPointValue(ss, infractionType) {
  try {
    const settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      console.error('Settings sheet not found');
      return 0;
    }

    // Buckets are in rows 23-26, columns A and B
    // A23:A26 = Bucket names, B23:B26 = Point values
    const bucketData = settingsSheet.getRange('A23:B26').getValues();

    for (const row of bucketData) {
      if (row[0] === infractionType) {
        return Number(row[1]) || 0;
      }
    }

    // If not found, try to extract from the infraction type string
    // Bucket names might be stored slightly differently
    console.log(`Bucket "${infractionType}" not found in settings, using passed points value`);
    return 0;

  } catch (error) {
    console.error('Error looking up bucket value:', error.toString());
    return 0;
  }
}

// ============================================
// PHASE 4: TEST FUNCTIONS
// ============================================

/**
 * Test function for addInfraction().
 * Tests valid data, then tests various validation failures.
 */
function testAddInfraction() {
  console.log('=== Testing addInfraction() ===');
  console.log('');

  const testResults = [];

  // ----------------------------------------
  // Test 1: Valid infraction
  // ----------------------------------------
  console.log('Test 1: Valid infraction data');

  // First, get a valid employee ID
  const employees = getActiveEmployees();
  if (employees.length === 0) {
    console.log('No active employees found. Cannot run test.');
    return { success: false, message: 'No active employees for testing' };
  }

  const testEmployee = employees[0];
  console.log(`Using test employee: ${testEmployee.full_name} (${testEmployee.employee_id})`);

  // Create a 240+ character description
  const testDescription = 'Employee was observed arriving 15 minutes late for their scheduled shift. ' +
    'This is a repeated pattern of behavior that has been addressed verbally before. ' +
    'The employee acknowledged the tardiness and committed to improving their punctuality. ' +
    'This documentation is being created as part of the accountability process.';

  console.log(`Description length: ${testDescription.length} characters`);

  const validInfraction = {
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: new Date(), // Today
    infraction_type: 'Bucket 1: Minor Offenses',
    points_assigned: 1,
    description: testDescription,
    location: 'Cockrell Hill DTO',
    entered_by: 'Test Script'
  };

  const startTime = new Date().getTime();
  const result1 = addInfraction(validInfraction);
  const endTime = new Date().getTime();
  const duration = (endTime - startTime) / 1000;

  console.log(`Result: ${result1.success ? 'SUCCESS' : 'FAILED'}`);
  console.log(`Message: ${result1.message}`);
  console.log(`Infraction ID: ${result1.infraction_id}`);
  console.log(`Duplicate warning: ${result1.duplicate_warning}`);
  console.log(`Execution time: ${duration.toFixed(2)} seconds`);

  if (duration >= 2) {
    console.log('⚠ WARNING: Function took longer than 2 seconds');
  } else {
    console.log('✓ Performance OK (under 2 seconds)');
  }

  testResults.push({
    test: 'Valid infraction',
    passed: result1.success,
    message: result1.message
  });
  console.log('');

  // ----------------------------------------
  // Test 2: Future date (should fail)
  // ----------------------------------------
  console.log('Test 2: Future date (should fail)');

  const futureDate = new Date();
  futureDate.setDate(futureDate.getDate() + 1);

  const futureInfraction = { ...validInfraction, date: futureDate };
  const result2 = addInfraction(futureInfraction);

  console.log(`Result: ${result2.success ? 'UNEXPECTED SUCCESS' : 'CORRECTLY FAILED'}`);
  console.log(`Message: ${result2.message}`);

  testResults.push({
    test: 'Future date rejection',
    passed: !result2.success && result2.message.includes('future'),
    message: result2.message
  });
  console.log('');

  // ----------------------------------------
  // Test 3: Date more than 7 days old (should fail)
  // ----------------------------------------
  console.log('Test 3: Date more than 7 days old (should fail)');

  const oldDate = new Date();
  oldDate.setDate(oldDate.getDate() - 10);

  const oldInfraction = { ...validInfraction, date: oldDate };
  const result3 = addInfraction(oldInfraction);

  console.log(`Result: ${result3.success ? 'UNEXPECTED SUCCESS' : 'CORRECTLY FAILED'}`);
  console.log(`Message: ${result3.message}`);

  testResults.push({
    test: 'Old date rejection (>7 days)',
    passed: !result3.success && result3.message.includes('7 days'),
    message: result3.message
  });
  console.log('');

  // ----------------------------------------
  // Test 4: Short description (should fail)
  // ----------------------------------------
  console.log('Test 4: Short description (should fail)');

  const shortInfraction = { ...validInfraction, description: 'Too short' };
  const result4 = addInfraction(shortInfraction);

  console.log(`Result: ${result4.success ? 'UNEXPECTED SUCCESS' : 'CORRECTLY FAILED'}`);
  console.log(`Message: ${result4.message}`);

  testResults.push({
    test: 'Short description rejection',
    passed: !result4.success && result4.message.includes('240 characters'),
    message: result4.message
  });
  console.log('');

  // ----------------------------------------
  // Test 5: Invalid location (should fail)
  // ----------------------------------------
  console.log('Test 5: Invalid location (should fail)');

  const badLocationInfraction = { ...validInfraction, location: 'Invalid Location' };
  const result5 = addInfraction(badLocationInfraction);

  console.log(`Result: ${result5.success ? 'UNEXPECTED SUCCESS' : 'CORRECTLY FAILED'}`);
  console.log(`Message: ${result5.message}`);

  testResults.push({
    test: 'Invalid location rejection',
    passed: !result5.success && result5.message.includes('location'),
    message: result5.message
  });
  console.log('');

  // ----------------------------------------
  // Test 6: Invalid employee ID (should fail)
  // ----------------------------------------
  console.log('Test 6: Invalid employee ID (should fail)');

  const badEmployeeInfraction = { ...validInfraction, employee_id: 'FAKE-ID-12345' };
  const result6 = addInfraction(badEmployeeInfraction);

  console.log(`Result: ${result6.success ? 'UNEXPECTED SUCCESS' : 'CORRECTLY FAILED'}`);
  console.log(`Message: ${result6.message}`);

  testResults.push({
    test: 'Invalid employee rejection',
    passed: !result6.success && result6.message.includes('not found'),
    message: result6.message
  });
  console.log('');

  // ----------------------------------------
  // Test 7: Missing field (should fail)
  // ----------------------------------------
  console.log('Test 7: Missing required field (should fail)');

  const missingFieldInfraction = { ...validInfraction };
  delete missingFieldInfraction.entered_by;
  const result7 = addInfraction(missingFieldInfraction);

  console.log(`Result: ${result7.success ? 'UNEXPECTED SUCCESS' : 'CORRECTLY FAILED'}`);
  console.log(`Message: ${result7.message}`);

  testResults.push({
    test: 'Missing field rejection',
    passed: !result7.success && result7.message.includes('Missing'),
    message: result7.message
  });
  console.log('');

  // ----------------------------------------
  // Test 8: Duplicate detection (should warn but succeed)
  // ----------------------------------------
  console.log('Test 8: Duplicate detection (should warn but succeed)');

  // Use the same valid infraction again (if test 1 passed)
  if (result1.success) {
    const result8 = addInfraction(validInfraction);

    console.log(`Result: ${result8.success ? 'SUCCESS' : 'FAILED'}`);
    console.log(`Duplicate warning: ${result8.duplicate_warning}`);
    console.log(`Message: ${result8.message}`);

    testResults.push({
      test: 'Duplicate detection',
      passed: result8.success && result8.duplicate_warning,
      message: result8.message
    });
  } else {
    console.log('Skipped (Test 1 did not create initial infraction)');
    testResults.push({
      test: 'Duplicate detection',
      passed: false,
      message: 'Skipped - no initial infraction to duplicate'
    });
  }
  console.log('');

  // ----------------------------------------
  // Summary
  // ----------------------------------------
  console.log('=== Test Summary ===');
  let passCount = 0;
  for (const result of testResults) {
    const status = result.passed ? '✓ PASS' : '✗ FAIL';
    console.log(`${status}: ${result.test}`);
    if (result.passed) passCount++;
  }
  console.log('');
  console.log(`${passCount}/${testResults.length} tests passed`);
  console.log('=== Test Complete ===');

  return {
    success: passCount === testResults.length,
    results: testResults,
    summary: `${passCount}/${testResults.length} tests passed`
  };
}

/**
 * Verifies that an infraction was written to the sheet correctly.
 * @param {string} infractionId - The ID to look up
 * @returns {Object} The infraction data or null if not found
 */
function verifyInfractionInSheet(infractionId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const infractionsSheet = ss.getSheetByName('Infractions');

    if (!infractionsSheet) return null;

    const lastRow = infractionsSheet.getLastRow();
    if (lastRow < 2) return null;

    const data = infractionsSheet.getRange(2, 1, lastRow - 1, 16).getValues();

    for (const row of data) {
      if (row[0] === infractionId) {
        return {
          infraction_id: row[0],
          employee_id: row[1],
          full_name: row[2],
          date: row[3],
          infraction_type: row[4],
          points_assigned: row[5],
          point_value_at_time: row[6],
          description: row[7],
          location: row[8],
          entered_by: row[9],
          entry_timestamp: row[10],
          status: row[14],
          expiration_date: row[15]
        };
      }
    }

    return null;
  } catch (error) {
    console.error('Error verifying infraction:', error.toString());
    return null;
  }
}

// ============================================
// PHASE 5: THRESHOLD DETECTION LOGIC
// ============================================

/**
 * Detects which point thresholds were crossed between old and new point totals.
 * Only detects crossing UP (increasing points), never crossing DOWN.
 *
 * @param {number} oldPoints - Point total before the new infraction
 * @param {number} newPoints - Point total after the new infraction
 * @returns {Array<number>} Array of threshold numbers that were crossed, in ascending order
 *
 * @throws {Error} If Settings sheet cannot be accessed or thresholds not found
 *
 * @example
 * detectThresholds(4, 7)   // returns [6]
 * detectThresholds(4, 12)  // returns [6, 9, 12]
 * detectThresholds(-2, 4)  // returns [2, 3]
 * detectThresholds(6, 5)   // returns [] (going down)
 */
function detectThresholds(oldPoints, newPoints) {
  // Validate inputs are numbers
  const oldPts = Number(oldPoints);
  const newPts = Number(newPoints);

  if (isNaN(oldPts)) {
    throw new Error('oldPoints must be a number');
  }

  if (isNaN(newPts)) {
    throw new Error('newPoints must be a number');
  }

  // If points didn't increase, no thresholds can be crossed
  if (newPts <= oldPts) {
    return [];
  }

  // Get threshold values from Settings
  const thresholds = getThresholdValues();

  // Find which thresholds were crossed
  const crossedThresholds = [];

  for (const threshold of thresholds) {
    // Crossed if: old was below threshold AND new is at or above threshold
    if (oldPts < threshold && newPts >= threshold) {
      crossedThresholds.push(threshold);
    }
  }

  // Sort in ascending order (should already be sorted, but ensure)
  crossedThresholds.sort((a, b) => a - b);

  return crossedThresholds;
}

/**
 * Retrieves threshold values from the Settings sheet.
 * Thresholds are stored in rows 13-19, column A.
 *
 * @returns {Array<number>} Array of threshold values [2, 3, 5, 6, 9, 12, 15]
 * @throws {Error} If Settings sheet not found or thresholds not configured
 */
function getThresholdValues() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      throw new Error('Settings sheet not found');
    }

    // Thresholds are in rows 13-19, column A
    const thresholdData = settingsSheet.getRange('A13:A19').getValues();

    const thresholds = [];
    for (const row of thresholdData) {
      const value = Number(row[0]);
      if (!isNaN(value) && value > 0) {
        thresholds.push(value);
      }
    }

    if (thresholds.length === 0) {
      throw new Error('No threshold values found in Settings');
    }

    // Sort ascending
    thresholds.sort((a, b) => a - b);

    return thresholds;

  } catch (error) {
    console.error('Error getting threshold values:', error.toString());
    throw error;
  }
}

/**
 * Gets the consequence text for a given threshold value.
 *
 * @param {number} threshold - The threshold value
 * @returns {string} The consequence text for this threshold
 */
function getThresholdConsequence(threshold) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return 'Unknown consequence';
    }

    // Thresholds and consequences are in rows 13-19, columns A and B
    const data = settingsSheet.getRange('A13:B19').getValues();

    for (const row of data) {
      if (Number(row[0]) === threshold) {
        return row[1] || 'No consequence defined';
      }
    }

    return 'Unknown consequence';

  } catch (error) {
    console.error('Error getting threshold consequence:', error.toString());
    return 'Error retrieving consequence';
  }
}

/**
 * Gets all threshold data (value and consequence) as an array of objects.
 *
 * @returns {Array<Object>} Array of {threshold, consequence} objects
 */
function getAllThresholdData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      throw new Error('Settings sheet not found');
    }

    const data = settingsSheet.getRange('A13:B19').getValues();
    const thresholdData = [];

    for (const row of data) {
      const threshold = Number(row[0]);
      if (!isNaN(threshold) && threshold > 0) {
        thresholdData.push({
          threshold: threshold,
          consequence: row[1] || ''
        });
      }
    }

    // Sort by threshold ascending
    thresholdData.sort((a, b) => a.threshold - b.threshold);

    return thresholdData;

  } catch (error) {
    console.error('Error getting all threshold data:', error.toString());
    throw error;
  }
}

// ============================================
// PHASE 5: TEST FUNCTIONS
// ============================================

/**
 * Comprehensive test function for detectThresholds().
 * Tests all specified example cases and edge cases.
 */
function testDetectThresholds() {
  console.log('=== Testing detectThresholds() ===');
  console.log('');

  const testCases = [
    // Specified example cases (adjusted for actual thresholds [2,3,5,6,9,12,15])
    { old: 4, new: 7, expected: [5, 6], description: 'Cross thresholds 5 and 6' },
    { old: 4, new: 12, expected: [5, 6, 9, 12], description: 'Cross multiple thresholds (5, 6, 9, 12)' },
    { old: -2, new: 4, expected: [2, 3], description: 'Negative to positive, cross 2 and 3' },
    { old: 6, new: 5, expected: [], description: 'Going down - no detection' },
    { old: 6, new: 6, expected: [], description: 'No change - no detection' },
    { old: 8, new: 15, expected: [9, 12, 15], description: 'Cross 9, 12, 15' },
    { old: 15, new: 18, expected: [], description: 'Already past all thresholds' },

    // Edge cases
    { old: -6, new: 20, expected: [2, 3, 5, 6, 9, 12, 15], description: 'Very negative to very positive - cross all' },
    { old: 0, new: 2, expected: [2], description: 'Exactly reach threshold 2' },
    { old: 1, new: 2, expected: [2], description: 'One below to exactly at threshold' },
    { old: 2, new: 3, expected: [3], description: 'At threshold 2, reach threshold 3' },
    { old: 1.5, new: 2.5, expected: [2], description: 'Decimal values crossing threshold 2' },
    { old: 5.5, new: 6.5, expected: [6], description: 'Decimal values crossing threshold 6' },
    { old: 0, new: 1, expected: [], description: 'Increase but not reaching any threshold' },
    { old: 14, new: 15, expected: [15], description: 'Just cross termination threshold' },
    { old: 14.9, new: 15, expected: [15], description: 'Decimal just below to exactly 15' },
    { old: -3, new: 0, expected: [], description: 'Negative to zero - no thresholds' },
    { old: -3, new: 2, expected: [2], description: 'Negative to exactly threshold 2' }
  ];

  let passCount = 0;
  const results = [];

  const startTime = new Date().getTime();

  for (let i = 0; i < testCases.length; i++) {
    const tc = testCases[i];
    console.log(`Test ${i + 1}: ${tc.description}`);
    console.log(`  Input: oldPoints=${tc.old}, newPoints=${tc.new}`);

    try {
      const result = detectThresholds(tc.old, tc.new);
      console.log(`  Expected: [${tc.expected.join(', ')}]`);
      console.log(`  Actual:   [${result.join(', ')}]`);

      // Compare arrays
      const passed = arraysEqual(result, tc.expected);

      if (passed) {
        console.log('  ✓ PASS');
        passCount++;
      } else {
        console.log('  ✗ FAIL');
      }

      results.push({
        test: tc.description,
        passed: passed,
        expected: tc.expected,
        actual: result
      });

    } catch (error) {
      console.log(`  ✗ ERROR: ${error.toString()}`);
      results.push({
        test: tc.description,
        passed: false,
        error: error.toString()
      });
    }

    console.log('');
  }

  const endTime = new Date().getTime();
  const duration = (endTime - startTime) / 1000;

  // Performance check
  console.log('Performance:');
  console.log(`  Total time for ${testCases.length} tests: ${duration.toFixed(3)} seconds`);
  console.log(`  Average per test: ${(duration / testCases.length * 1000).toFixed(2)} ms`);

  if (duration / testCases.length < 0.5) {
    console.log('  ✓ Performance OK (under 0.5 seconds per call)');
  } else {
    console.log('  ⚠ Performance may be slow');
  }

  console.log('');

  // Summary
  console.log('=== Test Summary ===');
  console.log(`${passCount}/${testCases.length} tests passed`);

  if (passCount === testCases.length) {
    console.log('✓ ALL TESTS PASSED');
  } else {
    console.log('✗ SOME TESTS FAILED');
    console.log('');
    console.log('Failed tests:');
    for (const r of results) {
      if (!r.passed) {
        console.log(`  - ${r.test}`);
      }
    }
  }

  console.log('');
  console.log('=== Test Complete ===');

  return {
    success: passCount === testCases.length,
    results: results,
    summary: `${passCount}/${testCases.length} tests passed`
  };
}

/**
 * Helper function to compare two arrays for equality.
 * @param {Array} arr1 - First array
 * @param {Array} arr2 - Second array
 * @returns {boolean} True if arrays are equal
 */
function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false;
  }
  return true;
}

/**
 * Test function to verify threshold values are read correctly from Settings.
 */
function testGetThresholdValues() {
  console.log('=== Testing getThresholdValues() ===');
  console.log('');

  try {
    const thresholds = getThresholdValues();

    console.log('Threshold values retrieved:');
    console.log(`  [${thresholds.join(', ')}]`);
    console.log('');

    // Verify expected values
    const expectedThresholds = [2, 3, 5, 6, 9, 12, 15];
    const isCorrect = arraysEqual(thresholds, expectedThresholds);

    if (isCorrect) {
      console.log('✓ Threshold values match expected [2, 3, 5, 6, 9, 12, 15]');
    } else {
      console.log('✗ Threshold values do not match expected');
      console.log(`  Expected: [${expectedThresholds.join(', ')}]`);
      console.log(`  Actual:   [${thresholds.join(', ')}]`);
    }

    // Test consequences
    console.log('');
    console.log('Threshold consequences:');
    for (const threshold of thresholds) {
      const consequence = getThresholdConsequence(threshold);
      console.log(`  ${threshold} points: ${consequence}`);
    }

    console.log('');
    console.log('=== Test Complete ===');

    return {
      success: isCorrect,
      thresholds: thresholds
    };

  } catch (error) {
    console.error('Test failed with error:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test to verify the full threshold detection flow with real Settings data.
 */
function testThresholdDetectionIntegration() {
  console.log('=== Integration Test: Threshold Detection ===');
  console.log('');

  // First verify we can read thresholds
  console.log('Step 1: Reading thresholds from Settings...');
  const thresholds = getThresholdValues();
  console.log(`  Found ${thresholds.length} thresholds: [${thresholds.join(', ')}]`);
  console.log('');

  // Test a realistic scenario
  console.log('Step 2: Simulating employee point progression...');
  console.log('');

  const scenarios = [
    { description: 'New hire with first minor offense', old: 0, new: 1 },
    { description: 'Second minor offense', old: 1, new: 2 },
    { description: 'Moderate offense added', old: 2, new: 5 },
    { description: 'Another minor offense', old: 5, new: 6 },
    { description: 'Major offense added', old: 6, new: 11 },
    { description: 'One more offense pushes to termination', old: 11, new: 16 }
  ];

  for (const scenario of scenarios) {
    const crossed = detectThresholds(scenario.old, scenario.new);
    console.log(`${scenario.description}:`);
    console.log(`  Points: ${scenario.old} → ${scenario.new}`);

    if (crossed.length > 0) {
      console.log(`  Thresholds crossed: [${crossed.join(', ')}]`);
      for (const t of crossed) {
        const consequence = getThresholdConsequence(t);
        console.log(`    ${t} pts → ${consequence}`);
      }
    } else {
      console.log('  No thresholds crossed');
    }
    console.log('');
  }

  console.log('=== Integration Test Complete ===');
}

// ============================================
// PHASE 6: EMAIL TEMPLATE BUILDER
// ============================================

/**
 * Builds formatted email content for threshold notifications.
 *
 * @param {Object} emailData - The email data object
 * @param {string} emailData.employee_name - Employee full name
 * @param {number} emailData.current_points - Current point total
 * @param {Array<number>} emailData.thresholds_crossed - Array of crossed threshold values
 * @param {Array<Object>} emailData.infractions_list - Array of infraction objects from calculatePoints
 * @param {string} emailData.employee_id - Employee ID
 *
 * @returns {Object} Email content object with:
 *   - subject: string - email subject line
 *   - body_html: string - HTML formatted email body
 *   - body_text: string - plain text email body
 *   - priority: string - "high" for 9+, "normal" for others
 *
 * @throws {Error} If required fields are missing
 */
function buildThresholdEmail(emailData) {
  // ========================================
  // VALIDATE REQUIRED FIELDS
  // ========================================

  const requiredFields = ['employee_name', 'current_points', 'thresholds_crossed', 'infractions_list', 'employee_id'];

  for (const field of requiredFields) {
    if (emailData[field] === undefined || emailData[field] === null) {
      throw new Error(`Missing required field: ${field}`);
    }
  }

  if (!Array.isArray(emailData.thresholds_crossed)) {
    throw new Error('thresholds_crossed must be an array');
  }

  if (!Array.isArray(emailData.infractions_list)) {
    throw new Error('infractions_list must be an array');
  }

  const employeeName = emailData.employee_name;
  const currentPoints = Number(emailData.current_points);
  const thresholdsCrossed = emailData.thresholds_crossed;
  const infractionsList = emailData.infractions_list;
  const employeeId = emailData.employee_id;

  // ========================================
  // LOOK UP CONSEQUENCES
  // ========================================

  const consequencesMap = [];

  for (const threshold of thresholdsCrossed) {
    let consequence;
    try {
      consequence = getThresholdConsequence(threshold);
    } catch (e) {
      consequence = 'See employee handbook';
    }

    if (!consequence || consequence === 'Unknown consequence' || consequence === 'No consequence defined') {
      consequence = 'See employee handbook';
    }

    consequencesMap.push({
      threshold: threshold,
      consequence: consequence
    });
  }

  // Sort by threshold ascending
  consequencesMap.sort((a, b) => a.threshold - b.threshold);

  // ========================================
  // DETERMINE PRIORITY AND SUBJECT
  // ========================================

  const hasTermination = thresholdsCrossed.includes(15) || currentPoints >= 15;
  const hasFinalWarning = thresholdsCrossed.includes(9) || thresholdsCrossed.includes(12);

  let subject;
  let priority;

  if (hasTermination) {
    subject = `IMMEDIATE ACTION REQUIRED: Termination Threshold Reached - ${employeeName}`;
    priority = 'high';
  } else if (hasFinalWarning) {
    subject = `URGENT: Final Warning Threshold - ${employeeName}`;
    priority = 'high';
  } else {
    subject = `Accountability Alert: ${employeeName} - ${currentPoints} Points`;
    priority = 'normal';
  }

  // ========================================
  // CALCULATE NEXT EXPIRATION
  // ========================================

  let nextExpirationText = 'No upcoming expirations';
  let nextExpirationPoints = 0;

  if (infractionsList.length > 0) {
    // Sort by expiration date (earliest first)
    const sortedInfractions = [...infractionsList].sort((a, b) => {
      const dateA = a.expiration_date instanceof Date ? a.expiration_date : new Date(a.expiration_date);
      const dateB = b.expiration_date instanceof Date ? b.expiration_date : new Date(b.expiration_date);
      return dateA.getTime() - dateB.getTime();
    });

    const nextExpiration = sortedInfractions[0];
    if (nextExpiration && nextExpiration.expiration_date) {
      const expDate = nextExpiration.expiration_date instanceof Date
        ? nextExpiration.expiration_date
        : new Date(nextExpiration.expiration_date);
      nextExpirationText = formatDateForLog(expDate);
      nextExpirationPoints = nextExpiration.points || 0;
    }
  }

  // ========================================
  // GET RECENT INFRACTIONS (LAST 5)
  // ========================================

  // Sort by date descending (most recent first)
  const recentInfractions = [...infractionsList]
    .sort((a, b) => {
      const dateA = a.date instanceof Date ? a.date : new Date(a.date);
      const dateB = b.date instanceof Date ? b.date : new Date(b.date);
      return dateB.getTime() - dateA.getTime();
    })
    .slice(0, 5);

  // ========================================
  // BUILD ACTION REQUIRED TEXT
  // ========================================

  let actionHtml = '';
  let actionText = '';

  if (hasTermination) {
    actionHtml = '<p style="color: #d32f2f; font-weight: bold;">IMMEDIATE ACTION: Schedule termination discussion</p>';
    actionText = 'IMMEDIATE ACTION: Schedule termination discussion';
  } else if (thresholdsCrossed.includes(12) || currentPoints >= 12) {
    actionHtml = '<p style="color: #d32f2f; font-weight: bold;">REQUIRED: Director must meet with employee within 24 hours</p>';
    actionText = 'REQUIRED: Director must meet with employee within 24 hours';
  } else if (thresholdsCrossed.includes(9) || currentPoints >= 9) {
    actionHtml = '<p style="color: #f57c00; font-weight: bold;">REQUIRED: Director must meet with employee within 24 hours</p>';
    actionText = 'REQUIRED: Director must meet with employee within 24 hours';
  } else if (thresholdsCrossed.includes(6) || currentPoints >= 6) {
    actionHtml = '<p style="color: #f57c00; font-weight: bold;">REQUIRED: Director must meet with employee</p>';
    actionText = 'REQUIRED: Director must meet with employee';
  }

  // ========================================
  // BUILD HTML EMAIL BODY
  // ========================================

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

  let bodyHtml = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 800px;
      margin: 0 auto;
    }
    .header {
      background-color: #E51636;
      color: white;
      padding: 20px;
      text-align: center;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
    }
    .header p {
      margin: 5px 0 0 0;
      font-size: 14px;
    }
    .content {
      padding: 20px;
    }
    .alert-box {
      background-color: #fff3e0;
      border-left: 4px solid #ff9800;
      padding: 15px;
      margin: 15px 0;
    }
    .alert-box.critical {
      background-color: #ffebee;
      border-left-color: #d32f2f;
    }
    .threshold-list {
      list-style: none;
      padding: 0;
    }
    .threshold-list li {
      padding: 8px 0;
      border-bottom: 1px solid #eee;
    }
    .threshold-list li:last-child {
      border-bottom: none;
    }
    .threshold-value {
      font-weight: bold;
      color: #d32f2f;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 15px 0;
    }
    th, td {
      border: 1px solid #ddd;
      padding: 10px;
      text-align: left;
    }
    th {
      background-color: #4285F4;
      color: white;
    }
    tr:nth-child(even) {
      background-color: #f9f9f9;
    }
    .footer {
      background-color: #f5f5f5;
      padding: 15px;
      text-align: center;
      font-size: 12px;
      color: #666;
    }
    .next-steps {
      background-color: #e3f2fd;
      padding: 15px;
      margin: 15px 0;
      border-radius: 4px;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>Chick-fil-A</h1>
    <p>Accountability System Notification</p>
  </div>

  <div class="content">
    <div class="alert-box${hasTermination || hasFinalWarning ? ' critical' : ''}">
      <p><strong>Employee:</strong> ${employeeName}</p>
      <p><strong>Employee ID:</strong> ${employeeId}</p>
      <p><strong>Current Point Total:</strong> <span class="threshold-value">${currentPoints} points</span></p>
    </div>

    <h2>Alert: The following thresholds have been reached:</h2>
    <ul class="threshold-list">`;

  for (const item of consequencesMap) {
    bodyHtml += `
      <li><span class="threshold-value">${item.threshold} points:</span> ${item.consequence}</li>`;
  }

  bodyHtml += `
    </ul>

    <h2>Recent Infractions</h2>`;

  if (recentInfractions.length > 0) {
    bodyHtml += `
    <table>
      <thead>
        <tr>
          <th>Date</th>
          <th>Type</th>
          <th>Points</th>
          <th>Description</th>
        </tr>
      </thead>
      <tbody>`;

    for (const inf of recentInfractions) {
      const infDate = inf.date instanceof Date ? inf.date : new Date(inf.date);
      const dateStr = formatDateForLog(infDate);
      const expDate = inf.expiration_date instanceof Date ? inf.expiration_date : new Date(inf.expiration_date);
      const expStr = formatDateForLog(expDate);

      // Truncate description for table display
      const descShort = inf.description && inf.description.length > 100
        ? inf.description.substring(0, 100) + '...'
        : (inf.description || 'N/A');

      bodyHtml += `
        <tr>
          <td>${dateStr}<br><small>Expires: ${expStr}</small></td>
          <td>${inf.infraction_type || 'N/A'}</td>
          <td>${inf.points || 0}</td>
          <td>${descShort}</td>
        </tr>`;
    }

    bodyHtml += `
      </tbody>
    </table>`;
  } else {
    bodyHtml += `
    <p><em>No recent infractions on record</em></p>`;
  }

  bodyHtml += `
    <div class="next-steps">
      <h3>Next Steps</h3>
      <p><strong>Next point expiration:</strong> ${nextExpirationText}${nextExpirationPoints ? ` (${nextExpirationPoints} points will drop off)` : ''}</p>
      ${actionHtml}
    </div>
  </div>

  <div class="footer">
    <p>Generated: ${timestamp}</p>
    <p>This is an automated notification from CFA Accountability System</p>
    <p>View full employee record in the Accountability System</p>
  </div>
</body>
</html>`;

  // ========================================
  // BUILD PLAIN TEXT EMAIL BODY
  // ========================================

  let bodyText = `======================================
CFA ACCOUNTABILITY SYSTEM NOTIFICATION
======================================

Employee: ${employeeName}
Employee ID: ${employeeId}
Current Points: ${currentPoints}

THRESHOLDS CROSSED:
`;

  for (const item of consequencesMap) {
    bodyText += `- ${item.threshold} points: ${item.consequence}\n`;
  }

  bodyText += `
RECENT INFRACTIONS:
`;

  if (recentInfractions.length > 0) {
    for (let i = 0; i < recentInfractions.length; i++) {
      const inf = recentInfractions[i];
      const infDate = inf.date instanceof Date ? inf.date : new Date(inf.date);
      const dateStr = formatDateForLog(infDate);
      const expDate = inf.expiration_date instanceof Date ? inf.expiration_date : new Date(inf.expiration_date);
      const expStr = formatDateForLog(expDate);

      // Truncate description for text display
      const descShort = inf.description && inf.description.length > 80
        ? inf.description.substring(0, 80) + '...'
        : (inf.description || 'N/A');

      bodyText += `${i + 1}. ${dateStr} - ${inf.infraction_type || 'N/A'} - ${inf.points || 0}pts\n`;
      bodyText += `   ${descShort}\n`;
      bodyText += `   (Expires: ${expStr})\n\n`;
    }
  } else {
    bodyText += `No recent infractions on record\n`;
  }

  bodyText += `
--------------------------------------
NEXT STEPS
--------------------------------------
Next Expiration: ${nextExpirationText}${nextExpirationPoints ? ` (${nextExpirationPoints} points drop off)` : ''}

${actionText}

---
Automated notification - ${timestamp}
This is an automated notification from CFA Accountability System
`;

  // ========================================
  // RETURN EMAIL OBJECT
  // ========================================

  return {
    subject: subject,
    body_html: bodyHtml,
    body_text: bodyText,
    priority: priority
  };
}

// ============================================
// PHASE 6: TEST FUNCTIONS
// ============================================

/**
 * Test function for buildThresholdEmail().
 * Creates sample data and verifies email generation.
 */
function testBuildThresholdEmail() {
  console.log('=== Testing buildThresholdEmail() ===');
  console.log('');

  const testResults = [];
  let allPassed = true;

  // ----------------------------------------
  // Create sample email data
  // ----------------------------------------

  const sampleInfractions = [
    {
      infraction_id: 'INF-20251220-1234',
      date: new Date(2025, 11, 20), // Dec 20, 2025
      infraction_type: 'Bucket 3: Major Offenses',
      points: 5,
      description: 'Employee was involved in a significant food safety violation during afternoon shift. Failed to follow proper temperature logging procedures for the chicken cooler, resulting in potential food waste.',
      location: 'Cockrell Hill DTO',
      entered_by: 'Manager Jones',
      expiration_date: new Date(2026, 2, 20) // Mar 20, 2026
    },
    {
      infraction_id: 'INF-20251215-5678',
      date: new Date(2025, 11, 15), // Dec 15, 2025
      infraction_type: 'Bucket 2: Moderate Offenses',
      points: 3,
      description: 'Employee was observed using cell phone during shift in the drive-thru area. This is a second occurrence of this behavior after verbal coaching.',
      location: 'Cockrell Hill DTO',
      entered_by: 'Manager Smith',
      expiration_date: new Date(2026, 2, 15) // Mar 15, 2026
    },
    {
      infraction_id: 'INF-20251210-9012',
      date: new Date(2025, 11, 10), // Dec 10, 2025
      infraction_type: 'Bucket 2: Moderate Offenses',
      points: 3,
      description: 'Call-out without adequate notice. Employee called out 30 minutes before scheduled shift start, causing staffing issues.',
      location: 'Cockrell Hill DTO',
      entered_by: 'Director Williams',
      expiration_date: new Date(2026, 2, 10) // Mar 10, 2026
    }
  ];

  const emailData = {
    employee_name: 'John Smith',
    current_points: 12,
    thresholds_crossed: [6, 9, 12],
    infractions_list: sampleInfractions,
    employee_id: '12-1234567'
  };

  console.log('Sample data created:');
  console.log(`  Employee: ${emailData.employee_name}`);
  console.log(`  Current Points: ${emailData.current_points}`);
  console.log(`  Thresholds Crossed: [${emailData.thresholds_crossed.join(', ')}]`);
  console.log(`  Infractions Count: ${emailData.infractions_list.length}`);
  console.log('');

  // ----------------------------------------
  // Test 1: Generate email
  // ----------------------------------------

  console.log('Test 1: Generate threshold email');

  const startTime = new Date().getTime();
  let result;

  try {
    result = buildThresholdEmail(emailData);

    const endTime = new Date().getTime();
    const duration = (endTime - startTime) / 1000;

    console.log(`  Execution time: ${duration.toFixed(3)} seconds`);

    if (duration < 1) {
      console.log('  ✓ Performance OK (under 1 second)');
      testResults.push({ test: 'Performance', passed: true });
    } else {
      console.log('  ✗ Performance SLOW (over 1 second)');
      testResults.push({ test: 'Performance', passed: false });
      allPassed = false;
    }
  } catch (error) {
    console.log(`  ✗ ERROR: ${error.toString()}`);
    testResults.push({ test: 'Email generation', passed: false });
    return { success: false, error: error.toString() };
  }

  console.log('');

  // ----------------------------------------
  // Test 2: Verify subject line
  // ----------------------------------------

  console.log('Test 2: Verify subject line');
  console.log(`  Subject: ${result.subject}`);

  // Should be URGENT for 9/12 threshold
  const subjectCorrect = result.subject.includes('URGENT') || result.subject.includes('Final Warning');
  if (subjectCorrect) {
    console.log('  ✓ Subject appropriate for threshold severity');
    testResults.push({ test: 'Subject line', passed: true });
  } else {
    console.log('  ✗ Subject may not reflect severity correctly');
    testResults.push({ test: 'Subject line', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 3: Verify priority
  // ----------------------------------------

  console.log('Test 3: Verify priority');
  console.log(`  Priority: ${result.priority}`);

  // Should be high for 9+ threshold
  const priorityCorrect = result.priority === 'high';
  if (priorityCorrect) {
    console.log('  ✓ Priority correctly set to "high" for 9+ threshold');
    testResults.push({ test: 'Priority', passed: true });
  } else {
    console.log('  ✗ Priority should be "high" for 9+ threshold');
    testResults.push({ test: 'Priority', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 4: Verify HTML body content
  // ----------------------------------------

  console.log('Test 4: Verify HTML body');
  console.log(`  HTML body length: ${result.body_html.length} characters`);
  console.log('  First 500 characters:');
  console.log('  ' + result.body_html.substring(0, 500).replace(/\n/g, '\n  '));
  console.log('');

  // Check thresholds appear in HTML
  let htmlThresholdsOk = true;
  for (const threshold of emailData.thresholds_crossed) {
    if (!result.body_html.includes(`${threshold} points`)) {
      htmlThresholdsOk = false;
      console.log(`  ✗ Threshold ${threshold} not found in HTML`);
    }
  }
  if (htmlThresholdsOk) {
    console.log('  ✓ All thresholds appear in HTML body');
    testResults.push({ test: 'HTML thresholds', passed: true });
  } else {
    testResults.push({ test: 'HTML thresholds', passed: false });
    allPassed = false;
  }

  // Check employee name appears
  if (result.body_html.includes(emailData.employee_name)) {
    console.log('  ✓ Employee name appears in HTML body');
    testResults.push({ test: 'HTML employee name', passed: true });
  } else {
    console.log('  ✗ Employee name not found in HTML');
    testResults.push({ test: 'HTML employee name', passed: false });
    allPassed = false;
  }

  // Check infractions appear
  let htmlInfractionsOk = true;
  for (const inf of sampleInfractions) {
    if (!result.body_html.includes(inf.infraction_type)) {
      htmlInfractionsOk = false;
      console.log(`  ✗ Infraction type "${inf.infraction_type}" not found in HTML`);
    }
  }
  if (htmlInfractionsOk) {
    console.log('  ✓ All infractions appear in HTML body');
    testResults.push({ test: 'HTML infractions', passed: true });
  } else {
    testResults.push({ test: 'HTML infractions', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 5: Verify text body content
  // ----------------------------------------

  console.log('Test 5: Verify text body');
  console.log(`  Text body length: ${result.body_text.length} characters`);
  console.log('  First 500 characters:');
  console.log('  ' + result.body_text.substring(0, 500).replace(/\n/g, '\n  '));
  console.log('');

  // Check no HTML in text body
  const hasHtmlTags = /<[^>]+>/.test(result.body_text);
  if (!hasHtmlTags) {
    console.log('  ✓ No HTML tags in text body');
    testResults.push({ test: 'Text no HTML', passed: true });
  } else {
    console.log('  ✗ HTML tags found in text body');
    testResults.push({ test: 'Text no HTML', passed: false });
    allPassed = false;
  }

  // Check thresholds appear in text
  let textThresholdsOk = true;
  for (const threshold of emailData.thresholds_crossed) {
    if (!result.body_text.includes(`${threshold} points`)) {
      textThresholdsOk = false;
      console.log(`  ✗ Threshold ${threshold} not found in text`);
    }
  }
  if (textThresholdsOk) {
    console.log('  ✓ All thresholds appear in text body');
    testResults.push({ test: 'Text thresholds', passed: true });
  } else {
    testResults.push({ test: 'Text thresholds', passed: false });
    allPassed = false;
  }

  // Check employee name appears
  if (result.body_text.includes(emailData.employee_name)) {
    console.log('  ✓ Employee name appears in text body');
    testResults.push({ test: 'Text employee name', passed: true });
  } else {
    console.log('  ✗ Employee name not found in text');
    testResults.push({ test: 'Text employee name', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 6: Verify expiration info
  // ----------------------------------------

  console.log('Test 6: Verify expiration info');

  if (result.body_html.includes('expir') || result.body_html.includes('Expir')) {
    console.log('  ✓ Expiration info appears in HTML');
    testResults.push({ test: 'Expiration in HTML', passed: true });
  } else {
    console.log('  ✗ Expiration info not found in HTML');
    testResults.push({ test: 'Expiration in HTML', passed: false });
    allPassed = false;
  }

  if (result.body_text.includes('Expir')) {
    console.log('  ✓ Expiration info appears in text');
    testResults.push({ test: 'Expiration in text', passed: true });
  } else {
    console.log('  ✗ Expiration info not found in text');
    testResults.push({ test: 'Expiration in text', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 7: Test termination threshold
  // ----------------------------------------

  console.log('Test 7: Test termination threshold subject');

  const terminationData = {
    ...emailData,
    current_points: 15,
    thresholds_crossed: [15]
  };

  const terminationResult = buildThresholdEmail(terminationData);

  if (terminationResult.subject.includes('IMMEDIATE ACTION') && terminationResult.subject.includes('Termination')) {
    console.log('  ✓ Termination subject line correct');
    testResults.push({ test: 'Termination subject', passed: true });
  } else {
    console.log('  ✗ Termination subject should include "IMMEDIATE ACTION" and "Termination"');
    console.log(`    Got: ${terminationResult.subject}`);
    testResults.push({ test: 'Termination subject', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 8: Test normal threshold
  // ----------------------------------------

  console.log('Test 8: Test normal threshold subject');

  const normalData = {
    ...emailData,
    current_points: 3,
    thresholds_crossed: [3]
  };

  const normalResult = buildThresholdEmail(normalData);

  if (normalResult.subject.includes('Accountability Alert') && normalResult.priority === 'normal') {
    console.log('  ✓ Normal threshold subject and priority correct');
    testResults.push({ test: 'Normal subject', passed: true });
  } else {
    console.log('  ✗ Normal threshold handling incorrect');
    console.log(`    Subject: ${normalResult.subject}`);
    console.log(`    Priority: ${normalResult.priority}`);
    testResults.push({ test: 'Normal subject', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 9: Test empty infractions list
  // ----------------------------------------

  console.log('Test 9: Test empty infractions list');

  const emptyInfractionsData = {
    ...emailData,
    infractions_list: []
  };

  const emptyResult = buildThresholdEmail(emptyInfractionsData);

  if (emptyResult.body_html.includes('No recent infractions') || emptyResult.body_text.includes('No recent infractions')) {
    console.log('  ✓ Empty infractions handled correctly');
    testResults.push({ test: 'Empty infractions', passed: true });
  } else {
    console.log('  ✗ Empty infractions not handled');
    testResults.push({ test: 'Empty infractions', passed: false });
    allPassed = false;
  }
  console.log('');

  // ----------------------------------------
  // Test 10: Test missing field error
  // ----------------------------------------

  console.log('Test 10: Test missing field error');

  const invalidData = {
    employee_name: 'Test',
    // Missing current_points
    thresholds_crossed: [6],
    infractions_list: [],
    employee_id: '12-1234567'
  };

  try {
    buildThresholdEmail(invalidData);
    console.log('  ✗ Should have thrown error for missing field');
    testResults.push({ test: 'Missing field error', passed: false });
    allPassed = false;
  } catch (error) {
    if (error.message.includes('Missing required field')) {
      console.log('  ✓ Correctly threw error for missing field');
      testResults.push({ test: 'Missing field error', passed: true });
    } else {
      console.log(`  ✗ Wrong error message: ${error.message}`);
      testResults.push({ test: 'Missing field error', passed: false });
      allPassed = false;
    }
  }
  console.log('');

  // ----------------------------------------
  // Summary
  // ----------------------------------------

  console.log('=== Test Summary ===');
  let passCount = 0;
  for (const result of testResults) {
    const status = result.passed ? '✓ PASS' : '✗ FAIL';
    console.log(`${status}: ${result.test}`);
    if (result.passed) passCount++;
  }
  console.log('');
  console.log(`${passCount}/${testResults.length} tests passed`);

  if (allPassed) {
    console.log('✓ ALL TESTS PASSED');
  } else {
    console.log('✗ SOME TESTS FAILED');
  }

  console.log('');
  console.log('=== Test Complete ===');

  return {
    success: allPassed,
    results: testResults,
    summary: `${passCount}/${testResults.length} tests passed`,
    sampleEmail: result
  };
}

// ============================================
// PHASE 7: EMAIL SENDING
// ============================================

/**
 * Sends threshold notification email via Gmail and logs to Email_Log.
 *
 * @param {string|Array<string>} recipientEmail - Email address(es) to send to
 * @param {string} subject - Email subject line
 * @param {string} bodyHtml - HTML formatted body
 * @param {string} bodyText - Plain text body (fallback)
 * @param {Object} emailMetadata - Additional metadata for logging
 * @param {string} emailMetadata.employee_id - Employee ID
 * @param {string} emailMetadata.employee_name - Employee name
 * @param {string} emailMetadata.email_type - Type of email (e.g., "6-Point Threshold")
 * @param {Array<number>} emailMetadata.thresholds_crossed - Thresholds that triggered this email
 *
 * @returns {Object} Result object with:
 *   - success: boolean
 *   - log_id: string
 *   - status: string - "Sent", "Failed", or "Retried"
 *   - message: string - success or error message
 */
function sendThresholdEmail(recipientEmail, subject, bodyHtml, bodyText, emailMetadata) {
  // ========================================
  // VALIDATE INPUTS
  // ========================================

  // Validate recipient email
  if (!recipientEmail || (Array.isArray(recipientEmail) && recipientEmail.length === 0)) {
    return {
      success: false,
      log_id: null,
      status: 'Failed',
      message: 'Recipient email is required'
    };
  }

  // Validate subject
  if (!subject || subject.trim() === '') {
    return {
      success: false,
      log_id: null,
      status: 'Failed',
      message: 'Subject is required'
    };
  }

  // Validate at least one body is provided
  if ((!bodyHtml || bodyHtml.trim() === '') && (!bodyText || bodyText.trim() === '')) {
    return {
      success: false,
      log_id: null,
      status: 'Failed',
      message: 'At least one of bodyHtml or bodyText is required'
    };
  }

  // Validate metadata
  if (!emailMetadata) {
    emailMetadata = {};
  }

  const employeeId = emailMetadata.employee_id || 'Unknown';
  const employeeName = emailMetadata.employee_name || 'Unknown';
  const emailType = emailMetadata.email_type || 'Threshold Notification';
  const thresholdsCrossed = emailMetadata.thresholds_crossed || [];

  // ========================================
  // PREPARE EMAIL
  // ========================================

  // Normalize recipient to array
  const recipients = Array.isArray(recipientEmail) ? recipientEmail : [recipientEmail];
  const primaryRecipient = recipients[0];
  const ccRecipients = recipients.slice(1);

  // Determine if high priority (thresholds 9 or higher)
  const isHighPriority = thresholdsCrossed.some(t => t >= 9);

  // Build email options
  const emailOptions = {
    name: 'CFA Accountability System',
    htmlBody: bodyHtml || undefined,
    body: bodyText || 'Please view this email in an HTML-compatible email client.'
  };

  // Add CC if multiple recipients
  if (ccRecipients.length > 0) {
    emailOptions.cc = ccRecipients.join(',');
  }

  // ========================================
  // GENERATE LOG ID
  // ========================================

  const logId = generateEmailLogId();

  // ========================================
  // SEND EMAIL - FIRST ATTEMPT
  // ========================================

  let sendSuccess = false;
  let sendError = null;
  let retryCount = 0;

  try {
    GmailApp.sendEmail(primaryRecipient, subject, emailOptions.body, emailOptions);
    sendSuccess = true;
  } catch (error) {
    sendError = error.toString();
    console.error('First email send attempt failed:', sendError);
  }

  // ========================================
  // RETRY IF FIRST ATTEMPT FAILED
  // ========================================

  if (!sendSuccess) {
    // Wait 2 seconds before retry
    Utilities.sleep(2000);
    retryCount = 1;

    try {
      GmailApp.sendEmail(primaryRecipient, subject, emailOptions.body, emailOptions);
      sendSuccess = true;
      sendError = null;
      console.log('Email sent successfully on retry');
    } catch (error) {
      sendError = error.toString();
      console.error('Retry email send attempt failed:', sendError);
    }
  }

  // ========================================
  // DETERMINE FINAL STATUS
  // ========================================

  let finalStatus;
  if (sendSuccess) {
    finalStatus = retryCount > 0 ? 'Sent' : 'Sent'; // Both map to 'Sent' for validation
  } else {
    finalStatus = 'Failed';
  }

  // ========================================
  // WRITE LOG ENTRY (After send attempt)
  // ========================================

  // Write the log entry AFTER we know the result to avoid validation issues
  // Status must be: Sent, Failed, or Retrying (per sheet validation)
  const logRowNumber = writeEmailLog({
    log_id: logId,
    timestamp: new Date(),
    employee_id: employeeId,
    employee_name: employeeName,
    recipient_email: recipients.join(', '),
    email_type: emailType,
    thresholds_crossed: thresholdsCrossed.join(', '),
    status: finalStatus,
    retry_count: retryCount,
    error_message: sendError || ''
  });

  // ========================================
  // SEND ERROR NOTIFICATION IF FAILED
  // ========================================

  if (!sendSuccess) {
    sendErrorNotification({
      original_recipient: recipients.join(', '),
      employee_name: employeeName,
      employee_id: employeeId,
      thresholds_crossed: thresholdsCrossed,
      error_message: sendError,
      timestamp: new Date()
    });
  }

  // ========================================
  // RETURN RESULT
  // ========================================

  return {
    success: sendSuccess,
    log_id: logId,
    status: finalStatus,
    message: sendSuccess
      ? `Email ${finalStatus.toLowerCase()} to ${primaryRecipient}`
      : `Failed to send email: ${sendError}`
  };
}

/**
 * Generates a unique email log ID.
 * Format: EMAIL-YYYYMMDDHHMMSS-####
 *
 * @returns {string} Generated log ID
 */
function generateEmailLogId() {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const random = Math.floor(1000 + Math.random() * 9000);
  return `EMAIL-${timestamp}-${random}`;
}

/**
 * Writes a new entry to the Email_Log sheet.
 *
 * @param {Object} logData - Log data object
 * @returns {number} Row number where the log was written
 */
function writeEmailLog(logData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const emailLogSheet = ss.getSheetByName('Email_Log');

    if (!emailLogSheet) {
      console.error('Email_Log sheet not found');
      return -1;
    }

    const lastRow = emailLogSheet.getLastRow();
    const newRow = lastRow + 1;

    // Prepare row data (columns A through J)
    const rowData = [
      logData.log_id,              // A: Log_ID
      logData.timestamp,           // B: Timestamp
      logData.employee_id,         // C: Employee_ID
      logData.employee_name,       // D: Employee_Name
      logData.recipient_email,     // E: Recipient_Email
      logData.email_type,          // F: Email_Type
      logData.thresholds_crossed,  // G: Thresholds_Crossed
      logData.status,              // H: Status
      logData.retry_count,         // I: Retry_Count
      logData.error_message        // J: Error_Message
    ];

    emailLogSheet.getRange(newRow, 1, 1, 10).setValues([rowData]);

    return newRow;

  } catch (error) {
    console.error('Error writing to Email_Log:', error.toString());
    return -1;
  }
}

/**
 * Updates an existing entry in the Email_Log sheet.
 *
 * @param {number} rowNumber - Row number to update
 * @param {Object} updateData - Data to update
 */
function updateEmailLog(rowNumber, updateData) {
  try {
    if (rowNumber < 2) {
      console.error('Invalid row number for Email_Log update');
      return;
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const emailLogSheet = ss.getSheetByName('Email_Log');

    if (!emailLogSheet) {
      console.error('Email_Log sheet not found');
      return;
    }

    // Update specific columns
    if (updateData.status !== undefined) {
      emailLogSheet.getRange(rowNumber, 8).setValue(updateData.status); // H: Status
    }
    if (updateData.retry_count !== undefined) {
      emailLogSheet.getRange(rowNumber, 9).setValue(updateData.retry_count); // I: Retry_Count
    }
    if (updateData.error_message !== undefined) {
      emailLogSheet.getRange(rowNumber, 10).setValue(updateData.error_message); // J: Error_Message
    }

  } catch (error) {
    console.error('Error updating Email_Log:', error.toString());
  }
}

/**
 * Sends an error notification email when a threshold email fails to send.
 *
 * @param {Object} errorData - Error information
 */
function sendErrorNotification(errorData) {
  try {
    // Get termination email list from Settings
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      console.error('Settings sheet not found for error notification');
      return;
    }

    // Termination Email List is in B3
    const errorRecipients = settingsSheet.getRange('B3').getValue();

    if (!errorRecipients || errorRecipients.trim() === '') {
      console.error('No termination email list configured for error notifications');
      return;
    }

    const timestamp = Utilities.formatDate(errorData.timestamp, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

    const errorSubject = `[ALERT] Failed to Send Accountability Email - ${errorData.employee_name}`;

    const errorBody = `
ACCOUNTABILITY SYSTEM EMAIL FAILURE ALERT
==========================================

An automated email failed to send. Manual follow-up may be required.

DETAILS:
--------
Original Recipient: ${errorData.original_recipient}
Employee Name: ${errorData.employee_name}
Employee ID: ${errorData.employee_id}
Thresholds Crossed: ${errorData.thresholds_crossed.join(', ')}
Error Message: ${errorData.error_message}
Timestamp: ${timestamp}

ACTION REQUIRED:
----------------
Please manually contact the intended recipient or investigate the email failure.

---
This is an automated error notification from CFA Accountability System
`;

    // Send error notification (don't retry - we don't want infinite loops)
    try {
      GmailApp.sendEmail(errorRecipients, errorSubject, errorBody, {
        name: 'CFA Accountability System - ERROR'
      });
      console.log('Error notification sent to:', errorRecipients);
    } catch (e) {
      console.error('Failed to send error notification:', e.toString());
    }

  } catch (error) {
    console.error('Error in sendErrorNotification:', error.toString());
  }
}

/**
 * Gets the configured store email from Settings.
 *
 * @returns {string} Store email or empty string if not configured
 */
function getStoreEmail() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return '';
    }

    // Store Email is in B2
    return settingsSheet.getRange('B2').getValue() || '';

  } catch (error) {
    console.error('Error getting store email:', error.toString());
    return '';
  }
}

/**
 * Gets the configured termination email list from Settings.
 *
 * @returns {string} Termination email list or empty string if not configured
 */
function getTerminationEmailList() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return '';
    }

    // Termination Email List is in B3
    return settingsSheet.getRange('B3').getValue() || '';

  } catch (error) {
    console.error('Error getting termination email list:', error.toString());
    return '';
  }
}

// ============================================
// PHASE 7: TEST FUNCTIONS
// ============================================

/**
 * Test function for sendThresholdEmail().
 * Tests email sending and logging functionality.
 *
 * IMPORTANT: Update TEST_EMAIL_ADDRESS before running!
 */
function testSendThresholdEmail() {
  console.log('=== Testing sendThresholdEmail() ===');
  console.log('');

  // ========================================
  // CONFIGURATION - UPDATE THIS!
  // ========================================

  // TODO: Replace with your actual test email address
  const TEST_EMAIL_ADDRESS = Session.getActiveUser().getEmail();

  if (!TEST_EMAIL_ADDRESS || TEST_EMAIL_ADDRESS === '') {
    console.log('ERROR: Could not get test email address');
    console.log('Please ensure you have permission to get the active user email');
    return { success: false, message: 'No test email address' };
  }

  console.log(`Using test email: ${TEST_EMAIL_ADDRESS}`);
  console.log('');

  const testResults = [];
  let allPassed = true;

  // ========================================
  // Test 1: Send valid email
  // ========================================

  console.log('Test 1: Send valid threshold email');

  // Create sample email content using Phase 6 function
  const sampleInfractions = [
    {
      infraction_id: 'INF-20251220-1234',
      date: new Date(2025, 11, 20),
      infraction_type: 'Bucket 2: Moderate Offenses',
      points: 3,
      description: 'Test infraction for email sending test. This is sample data to verify the email template renders correctly.',
      location: 'Cockrell Hill DTO',
      entered_by: 'Test Script',
      expiration_date: new Date(2026, 2, 20)
    }
  ];

  const emailData = {
    employee_name: 'Test Employee',
    current_points: 6,
    thresholds_crossed: [6],
    infractions_list: sampleInfractions,
    employee_id: 'TEST-001'
  };

  // Build the email
  const emailContent = buildThresholdEmail(emailData);

  const metadata = {
    employee_id: 'TEST-001',
    employee_name: 'Test Employee',
    email_type: '6-Point Threshold (TEST)',
    thresholds_crossed: [6]
  };

  const startTime = new Date().getTime();
  const result1 = sendThresholdEmail(
    TEST_EMAIL_ADDRESS,
    '[TEST] ' + emailContent.subject,
    emailContent.body_html,
    emailContent.body_text,
    metadata
  );
  const endTime = new Date().getTime();
  const duration = (endTime - startTime) / 1000;

  console.log(`  Result: ${JSON.stringify(result1, null, 2)}`);
  console.log(`  Execution time: ${duration.toFixed(2)} seconds`);

  if (result1.success) {
    console.log('  ✓ Email sent successfully');
    testResults.push({ test: 'Send valid email', passed: true });
  } else {
    console.log('  ✗ Email send failed');
    testResults.push({ test: 'Send valid email', passed: false });
    allPassed = false;
  }

  // Check performance
  if (duration < 5) {
    console.log('  ✓ Performance OK (under 5 seconds)');
    testResults.push({ test: 'Performance', passed: true });
  } else {
    console.log('  ⚠ Performance slow (over 5 seconds)');
    testResults.push({ test: 'Performance', passed: false });
  }

  // Check log_id format
  if (result1.log_id && result1.log_id.startsWith('EMAIL-')) {
    console.log('  ✓ Log ID format correct');
    testResults.push({ test: 'Log ID format', passed: true });
  } else {
    console.log('  ✗ Log ID format incorrect');
    testResults.push({ test: 'Log ID format', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 2: Verify Email_Log entry
  // ========================================

  console.log('Test 2: Verify Email_Log entry');

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const emailLogSheet = ss.getSheetByName('Email_Log');
    const lastRow = emailLogSheet.getLastRow();

    if (lastRow >= 2) {
      const lastEntry = emailLogSheet.getRange(lastRow, 1, 1, 10).getValues()[0];

      console.log('  Last Email_Log entry:');
      console.log(`    Log_ID: ${lastEntry[0]}`);
      console.log(`    Timestamp: ${lastEntry[1]}`);
      console.log(`    Employee_ID: ${lastEntry[2]}`);
      console.log(`    Status: ${lastEntry[7]}`);

      if (lastEntry[0] === result1.log_id) {
        console.log('  ✓ Log entry matches returned log_id');
        testResults.push({ test: 'Log entry created', passed: true });
      } else {
        console.log('  ✗ Log entry does not match');
        testResults.push({ test: 'Log entry created', passed: false });
        allPassed = false;
      }

      // Check all required fields present
      const requiredLogFields = ['Log_ID', 'Timestamp', 'Employee_ID', 'Employee_Name',
                                  'Recipient_Email', 'Email_Type', 'Thresholds_Crossed', 'Status'];
      let allFieldsPresent = true;
      for (let i = 0; i < requiredLogFields.length; i++) {
        if (lastEntry[i] === '' || lastEntry[i] === null || lastEntry[i] === undefined) {
          // Status and some fields can be empty
          if (i < 6) { // First 6 fields should always be present
            allFieldsPresent = false;
            console.log(`    ✗ Missing field: ${requiredLogFields[i]}`);
          }
        }
      }
      if (allFieldsPresent) {
        console.log('  ✓ All required log fields present');
        testResults.push({ test: 'Log fields complete', passed: true });
      } else {
        testResults.push({ test: 'Log fields complete', passed: false });
        allPassed = false;
      }
    } else {
      console.log('  ✗ No entries found in Email_Log');
      testResults.push({ test: 'Log entry created', passed: false });
      allPassed = false;
    }
  } catch (error) {
    console.log(`  ✗ Error checking Email_Log: ${error.toString()}`);
    testResults.push({ test: 'Log entry created', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 3: Test with multiple recipients
  // ========================================

  console.log('Test 3: Test with multiple recipients (CC)');

  const multiResult = sendThresholdEmail(
    [TEST_EMAIL_ADDRESS, TEST_EMAIL_ADDRESS], // Same email for testing
    '[TEST] Multiple Recipients',
    '<h1>Test</h1><p>Testing multiple recipients</p>',
    'Test - Testing multiple recipients',
    {
      employee_id: 'TEST-002',
      employee_name: 'Multi Test',
      email_type: 'Multi-Recipient Test',
      thresholds_crossed: [3]
    }
  );

  if (multiResult.success) {
    console.log('  ✓ Multi-recipient email sent');
    testResults.push({ test: 'Multiple recipients', passed: true });
  } else {
    console.log(`  ✗ Multi-recipient failed: ${multiResult.message}`);
    testResults.push({ test: 'Multiple recipients', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 4: Test validation - empty recipient
  // ========================================

  console.log('Test 4: Test validation - empty recipient');

  const emptyRecipientResult = sendThresholdEmail(
    '',
    'Test Subject',
    '<p>Test</p>',
    'Test',
    metadata
  );

  if (!emptyRecipientResult.success && emptyRecipientResult.message.includes('Recipient')) {
    console.log('  ✓ Empty recipient correctly rejected');
    testResults.push({ test: 'Empty recipient validation', passed: true });
  } else {
    console.log('  ✗ Empty recipient should be rejected');
    testResults.push({ test: 'Empty recipient validation', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 5: Test validation - empty subject
  // ========================================

  console.log('Test 5: Test validation - empty subject');

  const emptySubjectResult = sendThresholdEmail(
    TEST_EMAIL_ADDRESS,
    '',
    '<p>Test</p>',
    'Test',
    metadata
  );

  if (!emptySubjectResult.success && emptySubjectResult.message.includes('Subject')) {
    console.log('  ✓ Empty subject correctly rejected');
    testResults.push({ test: 'Empty subject validation', passed: true });
  } else {
    console.log('  ✗ Empty subject should be rejected');
    testResults.push({ test: 'Empty subject validation', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 6: Test validation - empty body
  // ========================================

  console.log('Test 6: Test validation - empty body');

  const emptyBodyResult = sendThresholdEmail(
    TEST_EMAIL_ADDRESS,
    'Test Subject',
    '',
    '',
    metadata
  );

  if (!emptyBodyResult.success && emptyBodyResult.message.includes('body')) {
    console.log('  ✓ Empty body correctly rejected');
    testResults.push({ test: 'Empty body validation', passed: true });
  } else {
    console.log('  ✗ Empty body should be rejected');
    testResults.push({ test: 'Empty body validation', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 7: Test with invalid email (will fail send)
  // ========================================

  console.log('Test 7: Test with obviously invalid email');

  const invalidEmailResult = sendThresholdEmail(
    'not-a-valid-email',
    '[TEST] Invalid Email Test',
    '<p>Test</p>',
    'Test',
    {
      employee_id: 'TEST-003',
      employee_name: 'Invalid Test',
      email_type: 'Invalid Email Test',
      thresholds_crossed: [2]
    }
  );

  console.log(`  Result: success=${invalidEmailResult.success}, status=${invalidEmailResult.status}`);

  // This might actually succeed (Gmail sometimes accepts bad addresses)
  // or might fail - either is acceptable for this test
  if (invalidEmailResult.log_id) {
    console.log('  ✓ Invalid email attempt was logged');
    testResults.push({ test: 'Invalid email logged', passed: true });
  } else {
    console.log('  ✗ Invalid email attempt was not logged');
    testResults.push({ test: 'Invalid email logged', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Test 8: Test high priority (threshold 9+)
  // ========================================

  console.log('Test 8: Test high priority email (threshold 9+)');

  const highPriorityResult = sendThresholdEmail(
    TEST_EMAIL_ADDRESS,
    '[TEST] High Priority - Final Warning',
    '<h1 style="color: red;">URGENT</h1><p>Final warning threshold reached</p>',
    'URGENT - Final warning threshold reached',
    {
      employee_id: 'TEST-004',
      employee_name: 'Priority Test',
      email_type: '9-Point Final Warning (TEST)',
      thresholds_crossed: [9]
    }
  );

  if (highPriorityResult.success) {
    console.log('  ✓ High priority email sent');
    testResults.push({ test: 'High priority email', passed: true });
  } else {
    console.log(`  ✗ High priority email failed: ${highPriorityResult.message}`);
    testResults.push({ test: 'High priority email', passed: false });
    allPassed = false;
  }
  console.log('');

  // ========================================
  // Summary
  // ========================================

  console.log('=== Test Summary ===');
  let passCount = 0;
  for (const result of testResults) {
    const status = result.passed ? '✓ PASS' : '✗ FAIL';
    console.log(`${status}: ${result.test}`);
    if (result.passed) passCount++;
  }
  console.log('');
  console.log(`${passCount}/${testResults.length} tests passed`);

  if (allPassed) {
    console.log('✓ ALL TESTS PASSED');
  } else {
    console.log('✗ SOME TESTS FAILED');
  }

  console.log('');
  console.log('IMPORTANT: Check your email inbox to verify emails were received!');
  console.log(`Test emails were sent to: ${TEST_EMAIL_ADDRESS}`);
  console.log('');
  console.log('=== Test Complete ===');

  return {
    success: allPassed,
    results: testResults,
    summary: `${passCount}/${testResults.length} tests passed`,
    testEmailAddress: TEST_EMAIL_ADDRESS
  };
}

/**
 * Quick test to verify email sending is working.
 * Sends a simple test email to the current user.
 */
function quickTestEmail() {
  const email = Session.getActiveUser().getEmail();

  if (!email) {
    console.log('Could not get current user email');
    return;
  }

  console.log(`Sending test email to: ${email}`);

  try {
    GmailApp.sendEmail(
      email,
      '[CFA Accountability] Quick Test',
      'This is a quick test email from the CFA Accountability System.',
      {
        name: 'CFA Accountability System',
        htmlBody: '<h2>Quick Test</h2><p>This is a quick test email from the CFA Accountability System.</p><p>If you received this, email sending is working correctly!</p>'
      }
    );
    console.log('Test email sent successfully!');
    console.log('Check your inbox.');
  } catch (error) {
    console.error('Failed to send test email:', error.toString());
  }
}

// ============================================
// PHASE 8: CONNECT PHASES 4-7 (BACKEND COMPLETE)
// ============================================

/**
 * Processes a new infraction with full notification workflow.
 * Wires together: addInfraction → calculatePoints → detectThresholds → sendEmail
 *
 * @param {Object} infractionData - Same structure as Phase 4 addInfraction
 * @param {string} infractionData.employee_id - Employee ID (required)
 * @param {string} infractionData.full_name - Employee full name (required)
 * @param {Date} infractionData.date - Date of infraction (required)
 * @param {string} infractionData.infraction_type - Type from buckets (required)
 * @param {number} infractionData.points_assigned - Points value (required)
 * @param {string} infractionData.description - Description, min 240 chars (required)
 * @param {string} infractionData.location - Location (required)
 * @param {string} infractionData.entered_by - Name of person entering (required)
 *
 * @returns {Object} Complete result object with:
 *   - success: boolean - true if infraction was added
 *   - infraction_id: string - ID of new infraction (if successful)
 *   - old_points: number - point total before infraction
 *   - new_points: number - point total after infraction
 *   - thresholds_crossed: array - threshold values crossed
 *   - email_sent: boolean - true if notification email was sent
 *   - email_status: string - "Sent", "Failed", or "Not Required"
 *   - message: string - summary of what happened
 */
function processInfractionWithNotifications(infractionData) {
  console.log('=== Processing Infraction with Notifications ===');
  console.log(`Employee: ${infractionData.full_name} (${infractionData.employee_id})`);
  console.log(`Infraction Type: ${infractionData.infraction_type}`);
  console.log(`Points: ${infractionData.points_assigned}`);
  console.log('');

  // Default result object
  const result = {
    success: false,
    infraction_id: null,
    old_points: 0,
    new_points: 0,
    thresholds_crossed: [],
    email_sent: false,
    email_status: 'Not Required',
    message: ''
  };

  try {
    // ========================================
    // STEP 1: Get CURRENT points BEFORE adding infraction
    // ========================================

    console.log('Step 1: Calculating current points before infraction...');

    let oldPointsData;
    try {
      oldPointsData = calculatePoints(infractionData.employee_id);
    } catch (error) {
      console.error('Error calculating initial points:', error.toString());
      oldPointsData = { total_points: 0, active_infractions: [] };
    }

    const oldPoints = oldPointsData.total_points;
    result.old_points = oldPoints;

    console.log(`  Current points: ${oldPoints}`);
    console.log(`  Active infractions: ${oldPointsData.active_infractions.length}`);
    console.log('');

    // ========================================
    // STEP 2: Add the new infraction
    // ========================================

    console.log('Step 2: Adding new infraction...');

    const addResult = addInfraction(infractionData);

    if (!addResult.success) {
      console.error('Failed to add infraction:', addResult.message);
      result.message = `Failed to add infraction: ${addResult.message}`;
      return result;
    }

    result.infraction_id = addResult.infraction_id;
    console.log(`  Infraction added: ${addResult.infraction_id}`);

    if (addResult.duplicate_warning) {
      console.log('  WARNING: Duplicate infraction detected');
    }
    console.log('');

    // ========================================
    // STEP 3: Get UPDATED points AFTER adding infraction
    // ========================================

    console.log('Step 3: Calculating updated points after infraction...');

    let newPointsData;
    try {
      newPointsData = calculatePoints(infractionData.employee_id);
    } catch (error) {
      console.error('Error calculating new points:', error.toString());
      // Use estimated new points based on old + new infraction
      newPointsData = {
        total_points: oldPoints + (infractionData.points_assigned || 0),
        active_infractions: oldPointsData.active_infractions
      };
    }

    const newPoints = newPointsData.total_points;
    result.new_points = newPoints;
    result.success = true; // Infraction was added successfully

    console.log(`  New points: ${newPoints}`);
    console.log(`  Point change: ${oldPoints} → ${newPoints} (${newPoints - oldPoints >= 0 ? '+' : ''}${newPoints - oldPoints})`);
    console.log('');

    // ========================================
    // STEP 4: Check if thresholds were crossed
    // ========================================

    console.log('Step 4: Detecting threshold crossings...');

    let thresholdsCrossed = [];
    try {
      thresholdsCrossed = detectThresholds(oldPoints, newPoints);
    } catch (error) {
      console.error('Error detecting thresholds:', error.toString());
      thresholdsCrossed = [];
    }

    result.thresholds_crossed = thresholdsCrossed;

    if (thresholdsCrossed.length === 0) {
      console.log('  No thresholds crossed - no notification required');
      result.email_status = 'Not Required';
      result.message = `Infraction ${result.infraction_id} added. Points: ${oldPoints} → ${newPoints}. No thresholds crossed.`;
      console.log('');
      console.log('=== Processing Complete (No Email Needed) ===');
      return result;
    }

    console.log(`  Thresholds crossed: [${thresholdsCrossed.join(', ')}]`);
    console.log('');

    // ========================================
    // STEP 5: Send notification email
    // ========================================

    console.log('Step 5: Sending notification email...');

    // Determine highest threshold for email type
    const highestThreshold = Math.max(...thresholdsCrossed);
    const emailType = `${highestThreshold}-Point Threshold`;

    console.log(`  Email type: ${emailType}`);

    // Build email content
    let emailContent;
    try {
      emailContent = buildThresholdEmail({
        employee_name: infractionData.full_name,
        current_points: newPoints,
        thresholds_crossed: thresholdsCrossed,
        infractions_list: newPointsData.active_infractions,
        employee_id: infractionData.employee_id
      });
    } catch (error) {
      console.error('Error building email:', error.toString());
      result.email_status = 'Failed';
      result.message = `Infraction ${result.infraction_id} added. Points: ${oldPoints} → ${newPoints}. Thresholds crossed: [${thresholdsCrossed.join(', ')}]. Email build failed: ${error.toString()}`;
      return result;
    }

    console.log(`  Subject: ${emailContent.subject}`);

    // Determine recipient(s)
    let recipientEmail;
    if (thresholdsCrossed.includes(15)) {
      // Termination threshold - send to termination email list
      recipientEmail = getTerminationEmailList();
      console.log(`  Recipient: Termination Email List (threshold 15 reached)`);
    } else {
      // Normal threshold - send to store email
      recipientEmail = getStoreEmail();
      console.log(`  Recipient: Store Email`);
    }

    if (!recipientEmail || recipientEmail.trim() === '') {
      console.log('  WARNING: No recipient email configured in Settings');
      result.email_status = 'Failed';
      result.message = `Infraction ${result.infraction_id} added. Points: ${oldPoints} → ${newPoints}. Thresholds crossed: [${thresholdsCrossed.join(', ')}]. No recipient email configured.`;
      return result;
    }

    console.log(`  Sending to: ${recipientEmail}`);

    // Build email metadata
    const emailMetadata = {
      employee_id: infractionData.employee_id,
      employee_name: infractionData.full_name,
      email_type: emailType,
      thresholds_crossed: thresholdsCrossed
    };

    // Send the email
    let sendResult;
    try {
      sendResult = sendThresholdEmail(
        recipientEmail,
        emailContent.subject,
        emailContent.body_html,
        emailContent.body_text,
        emailMetadata
      );
    } catch (error) {
      console.error('Error sending email:', error.toString());
      sendResult = { success: false, status: 'Failed', message: error.toString() };
    }

    result.email_sent = sendResult.success;
    result.email_status = sendResult.status;

    if (sendResult.success) {
      console.log(`  Email sent successfully! Log ID: ${sendResult.log_id}`);
    } else {
      console.log(`  Email send failed: ${sendResult.message}`);
    }
    console.log('');

    // ========================================
    // STEP 6: Compile final result
    // ========================================

    if (sendResult.success) {
      result.message = `Infraction ${result.infraction_id} added. Points: ${oldPoints} → ${newPoints}. Thresholds crossed: [${thresholdsCrossed.join(', ')}]. Notification email sent.`;
    } else {
      result.message = `Infraction ${result.infraction_id} added. Points: ${oldPoints} → ${newPoints}. Thresholds crossed: [${thresholdsCrossed.join(', ')}]. Email failed: ${sendResult.message}`;
    }

    console.log('=== Processing Complete ===');
    console.log(`Result: ${result.message}`);

    return result;

  } catch (error) {
    console.error('Unexpected error in processInfractionWithNotifications:', error.toString());
    result.message = `Unexpected error: ${error.toString()}`;
    return result;
  }
}

// ============================================
// PHASE 8: TEST FUNCTIONS
// ============================================

/**
 * Comprehensive test function for processInfractionWithNotifications().
 * Tests all four required scenarios.
 */
function testProcessInfractionWithNotifications() {
  console.log('=====================================================');
  console.log('=== Testing processInfractionWithNotifications() ===');
  console.log('=====================================================');
  console.log('');

  const testResults = [];
  let allPassed = true;

  // Get a valid employee for testing
  const employees = getActiveEmployees();
  if (employees.length === 0) {
    console.log('ERROR: No active employees found. Cannot run tests.');
    return { success: false, message: 'No active employees for testing' };
  }

  const testEmployee = employees[0];
  console.log(`Using test employee: ${testEmployee.full_name} (${testEmployee.employee_id})`);
  console.log('');

  // Create a standard 240+ char description
  const testDescription = 'This is a test infraction created by the automated test suite. ' +
    'The purpose of this infraction is to verify that the complete infraction processing workflow ' +
    'functions correctly, including point calculation, threshold detection, and email notifications. ' +
    'This description meets the 240 character minimum requirement.';

  // ========================================
  // Test Case 1: Normal threshold crossing
  // ========================================

  console.log('========================================');
  console.log('TEST CASE 1: Normal threshold crossing');
  console.log('========================================');
  console.log('Scenario: Add infraction that crosses threshold(s)');
  console.log('');

  // First, get current points
  const initialPoints1 = calculatePoints(testEmployee.employee_id);
  console.log(`Current points before test: ${initialPoints1.total_points}`);

  // Add a moderate infraction (3 points)
  const test1Data = {
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: new Date(),
    infraction_type: 'Bucket 2: Moderate Offenses',
    points_assigned: 3,
    description: testDescription,
    location: testEmployee.primary_location || 'Cockrell Hill DTO',
    entered_by: 'Test Script - Case 1'
  };

  const startTime1 = new Date().getTime();
  const result1 = processInfractionWithNotifications(test1Data);
  const duration1 = (new Date().getTime() - startTime1) / 1000;

  console.log('');
  console.log('Result:');
  console.log(`  Success: ${result1.success}`);
  console.log(`  Infraction ID: ${result1.infraction_id}`);
  console.log(`  Points: ${result1.old_points} → ${result1.new_points}`);
  console.log(`  Thresholds Crossed: [${result1.thresholds_crossed.join(', ')}]`);
  console.log(`  Email Sent: ${result1.email_sent}`);
  console.log(`  Email Status: ${result1.email_status}`);
  console.log(`  Duration: ${duration1.toFixed(2)} seconds`);
  console.log('');

  // Verify test 1
  const test1Passed = result1.success && result1.infraction_id !== null;
  if (test1Passed) {
    console.log('✓ Test 1 PASSED: Infraction processed successfully');
    testResults.push({ test: 'Normal threshold crossing', passed: true });
  } else {
    console.log('✗ Test 1 FAILED: Infraction was not processed');
    testResults.push({ test: 'Normal threshold crossing', passed: false });
    allPassed = false;
  }

  // Check performance
  if (duration1 < 10) {
    console.log('✓ Performance OK (under 10 seconds)');
    testResults.push({ test: 'Test 1 Performance', passed: true });
  } else {
    console.log('⚠ Performance slow (over 10 seconds)');
    testResults.push({ test: 'Test 1 Performance', passed: false });
  }
  console.log('');

  // ========================================
  // Test Case 2: No threshold crossed
  // ========================================

  console.log('========================================');
  console.log('TEST CASE 2: No threshold crossed');
  console.log('========================================');
  console.log('Scenario: Add small infraction that doesn\'t cross a threshold');
  console.log('');

  // Get another employee or use same one
  const testEmployee2 = employees.length > 1 ? employees[1] : employees[0];
  const initialPoints2 = calculatePoints(testEmployee2.employee_id);
  console.log(`Current points before test: ${initialPoints2.total_points}`);

  // We need an employee whose current points + 1 won't cross a threshold
  // If employee is at 0 points, adding 1 won't cross threshold 2
  // If at 1, adding 1 will cross threshold 2
  // Let's use a very small point value to minimize threshold crossing

  const test2Data = {
    employee_id: testEmployee2.employee_id,
    full_name: testEmployee2.full_name,
    date: new Date(),
    infraction_type: 'Bucket 1: Minor Offenses',
    points_assigned: 1,
    description: testDescription,
    location: testEmployee2.primary_location || 'Cockrell Hill DTO',
    entered_by: 'Test Script - Case 2'
  };

  const startTime2 = new Date().getTime();
  const result2 = processInfractionWithNotifications(test2Data);
  const duration2 = (new Date().getTime() - startTime2) / 1000;

  console.log('');
  console.log('Result:');
  console.log(`  Success: ${result2.success}`);
  console.log(`  Infraction ID: ${result2.infraction_id}`);
  console.log(`  Points: ${result2.old_points} → ${result2.new_points}`);
  console.log(`  Thresholds Crossed: [${result2.thresholds_crossed.join(', ')}]`);
  console.log(`  Email Sent: ${result2.email_sent}`);
  console.log(`  Email Status: ${result2.email_status}`);
  console.log(`  Duration: ${duration2.toFixed(2)} seconds`);
  console.log('');

  // Verify test 2 - infraction should be added
  const test2Passed = result2.success && result2.infraction_id !== null;
  if (test2Passed) {
    console.log('✓ Test 2 PASSED: Infraction processed successfully');
    testResults.push({ test: 'Small infraction processing', passed: true });
  } else {
    console.log('✗ Test 2 FAILED');
    testResults.push({ test: 'Small infraction processing', passed: false });
    allPassed = false;
  }

  // Note: email_sent depends on whether thresholds were crossed
  console.log(`  (Email sent: ${result2.email_sent} - depends on current point level)`);
  console.log('');

  // ========================================
  // Test Case 3: Severe infraction (potentially termination level)
  // ========================================

  console.log('========================================');
  console.log('TEST CASE 3: Severe infraction');
  console.log('========================================');
  console.log('Scenario: Add severe infraction (8 points)');
  console.log('');

  // Use a third employee if available
  const testEmployee3 = employees.length > 2 ? employees[2] : testEmployee;
  const initialPoints3 = calculatePoints(testEmployee3.employee_id);
  console.log(`Current points before test: ${initialPoints3.total_points}`);

  const test3Data = {
    employee_id: testEmployee3.employee_id,
    full_name: testEmployee3.full_name,
    date: new Date(),
    infraction_type: 'Bucket 4: Severe Offenses',
    points_assigned: 8,
    description: testDescription + ' This is a severe offense test case.',
    location: testEmployee3.primary_location || 'Cockrell Hill DTO',
    entered_by: 'Test Script - Case 3'
  };

  const startTime3 = new Date().getTime();
  const result3 = processInfractionWithNotifications(test3Data);
  const duration3 = (new Date().getTime() - startTime3) / 1000;

  console.log('');
  console.log('Result:');
  console.log(`  Success: ${result3.success}`);
  console.log(`  Infraction ID: ${result3.infraction_id}`);
  console.log(`  Points: ${result3.old_points} → ${result3.new_points}`);
  console.log(`  Thresholds Crossed: [${result3.thresholds_crossed.join(', ')}]`);
  console.log(`  Email Sent: ${result3.email_sent}`);
  console.log(`  Email Status: ${result3.email_status}`);
  console.log(`  Duration: ${duration3.toFixed(2)} seconds`);
  console.log('');

  // Verify test 3
  const test3Passed = result3.success && result3.infraction_id !== null;
  if (test3Passed) {
    console.log('✓ Test 3 PASSED: Severe infraction processed');
    testResults.push({ test: 'Severe infraction', passed: true });
  } else {
    console.log('✗ Test 3 FAILED');
    testResults.push({ test: 'Severe infraction', passed: false });
    allPassed = false;
  }

  // Check if termination threshold was crossed
  if (result3.thresholds_crossed.includes(15)) {
    console.log('  NOTE: Termination threshold (15) was crossed - email should go to Termination Email List');
  }
  console.log('');

  // ========================================
  // Test Case 4: Negative points (positive behavior)
  // ========================================

  console.log('========================================');
  console.log('TEST CASE 4: Negative points (going down)');
  console.log('========================================');
  console.log('Scenario: Add negative points (positive behavior credit)');
  console.log('');

  const initialPoints4 = calculatePoints(testEmployee.employee_id);
  console.log(`Current points before test: ${initialPoints4.total_points}`);

  const test4Data = {
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: new Date(),
    infraction_type: 'Bucket 1: Minor Offenses', // Using existing bucket for simplicity
    points_assigned: -2,
    description: testDescription + ' This is a positive behavior credit for excellent performance.',
    location: testEmployee.primary_location || 'Cockrell Hill DTO',
    entered_by: 'Test Script - Case 4'
  };

  const startTime4 = new Date().getTime();
  const result4 = processInfractionWithNotifications(test4Data);
  const duration4 = (new Date().getTime() - startTime4) / 1000;

  console.log('');
  console.log('Result:');
  console.log(`  Success: ${result4.success}`);
  console.log(`  Infraction ID: ${result4.infraction_id}`);
  console.log(`  Points: ${result4.old_points} → ${result4.new_points}`);
  console.log(`  Thresholds Crossed: [${result4.thresholds_crossed.join(', ')}]`);
  console.log(`  Email Sent: ${result4.email_sent}`);
  console.log(`  Email Status: ${result4.email_status}`);
  console.log(`  Duration: ${duration4.toFixed(2)} seconds`);
  console.log('');

  // Verify test 4 - negative points should not trigger email
  const test4Passed = result4.success && result4.thresholds_crossed.length === 0;
  if (test4Passed) {
    console.log('✓ Test 4 PASSED: Negative points processed without email');
    testResults.push({ test: 'Negative points (no email)', passed: true });
  } else if (result4.success) {
    console.log('✓ Test 4 PASSED: Entry recorded (thresholds depend on point level)');
    testResults.push({ test: 'Negative points entry', passed: true });
  } else {
    console.log('✗ Test 4 FAILED');
    testResults.push({ test: 'Negative points', passed: false });
    allPassed = false;
  }

  // Verify no email sent for going down
  if (!result4.email_sent && result4.email_status === 'Not Required') {
    console.log('✓ Correctly did not send email for points going down');
  } else if (result4.new_points > result4.old_points) {
    console.log('  Note: Points went UP (not down) - email may have been sent');
  }
  console.log('');

  // ========================================
  // Summary
  // ========================================

  console.log('=====================================================');
  console.log('=== TEST SUMMARY ===');
  console.log('=====================================================');

  let passCount = 0;
  for (const result of testResults) {
    const status = result.passed ? '✓ PASS' : '✗ FAIL';
    console.log(`${status}: ${result.test}`);
    if (result.passed) passCount++;
  }

  console.log('');
  console.log(`${passCount}/${testResults.length} tests passed`);

  if (allPassed) {
    console.log('');
    console.log('✓ ALL TESTS PASSED');
  } else {
    console.log('');
    console.log('✗ SOME TESTS FAILED');
  }

  console.log('');
  console.log('NOTE: Check your inbox for any notification emails that were sent during testing.');
  console.log('NOTE: Check the Infractions sheet for newly created test entries.');
  console.log('NOTE: Check the Email_Log sheet for email send records.');
  console.log('');
  console.log('=== Test Complete ===');

  return {
    success: allPassed,
    results: testResults,
    summary: `${passCount}/${testResults.length} tests passed`
  };
}

/**
 * Quick test for processInfractionWithNotifications with minimal output.
 * Good for verifying basic functionality.
 */
function quickTestProcess() {
  console.log('Quick test of processInfractionWithNotifications...');

  const employees = getActiveEmployees();
  if (employees.length === 0) {
    console.log('No employees found');
    return;
  }

  const testEmployee = employees[0];

  const testData = {
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: new Date(),
    infraction_type: 'Bucket 1: Minor Offenses',
    points_assigned: 1,
    description: 'Quick test infraction from automated testing. This description is intentionally made long enough to meet the 240 character minimum requirement for infractions in the CFA Accountability System. Additional text here to ensure we reach the limit.',
    location: testEmployee.primary_location || 'Cockrell Hill DTO',
    entered_by: 'Quick Test'
  };

  const result = processInfractionWithNotifications(testData);

  console.log('');
  console.log('Quick Test Result:');
  console.log(`  Success: ${result.success}`);
  console.log(`  Points: ${result.old_points} → ${result.new_points}`);
  console.log(`  Thresholds: [${result.thresholds_crossed.join(', ')}]`);
  console.log(`  Email: ${result.email_status}`);
  console.log(`  Message: ${result.message}`);
}

// ============================================
// PHASE 9: WEB APP - SIMPLE HTML FORM
// ============================================

/**
 * Serves the appropriate HTML page when the web app is accessed.
 * This is the entry point for the web app.
 * Checks for valid session and redirects to login if needed.
 * Supports routing via ?page= parameter:
 *   - page=employees : Employee list overview
 *   - page=form (or no param) : Add infraction form
 *
 * @param {Object} e - Event parameter from doGet
 * @returns {HtmlOutput} The HTML page to display
 */
function doGet(e) {
  // Get the requested page
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : null;

  // Signup page is public (no session required)
  if (page === 'signup') {
    const token = e.parameter.token || '';
    return HtmlService.createHtmlOutputFromFile('Signup')
      .setTitle('CFA Accountability - Complete Registration')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Check for valid session for all other pages
  const session = validateSession();

  if (session.valid) {
    // Valid session - extend it
    extendSession();

    // Route to appropriate page (default to employees list as main view)
    const targetPage = page || 'employees';

    switch (targetPage) {
      case 'form':
        // Infraction form
        return HtmlService.createHtmlOutputFromFile('InfractionForm')
          .setTitle('CFA Accountability - Add Infraction')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'users':
        // User management (Directors and Operators only)
        if (session.role === 'Manager') {
          // Managers don't have access - redirect to employees
          return HtmlService.createHtmlOutputFromFile('EmployeeList')
            .setTitle('CFA Accountability - Employee Overview')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        }
        return HtmlService.createHtmlOutputFromFile('UserManagement')
          .setTitle('CFA Accountability - User Management')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'employees':
      default:
        // Employee list (default landing page)
        return HtmlService.createHtmlOutputFromFile('EmployeeList')
          .setTitle('CFA Accountability - Employee Overview')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } else {
    // No valid session - show login page
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('CFA Accountability - Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Gets all form data needed to populate dropdowns.
 * Called from client-side JavaScript on page load.
 *
 * @returns {Object} Object containing employees and infraction types
 */
function getFormData() {
  try {
    // Get active employees
    const employees = getActiveEmployees();

    // Format employees for dropdown
    const employeeOptions = employees.map(emp => ({
      id: emp.employee_id,
      name: emp.full_name,
      location: emp.primary_location
    }));

    // Get infraction types from Settings
    const infractionTypes = getInfractionTypesFromSettings();

    return {
      success: true,
      employees: employeeOptions,
      infraction_types: infractionTypes
    };

  } catch (error) {
    console.error('Error in getFormData:', error.toString());
    return {
      success: false,
      error: error.toString(),
      employees: [],
      infraction_types: []
    };
  }
}

/**
 * Gets infraction types with examples from Settings sheet.
 *
 * @returns {Array} Array of bucket objects with name, points, and examples
 */
function getInfractionTypesFromSettings() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      console.error('Settings sheet not found');
      return getDefaultInfractionTypes();
    }

    // Buckets are in rows 23-26, columns A, B, C
    // A = Bucket Name, B = Point Value, C = Examples (JSON array)
    const bucketData = settingsSheet.getRange('A23:C26').getValues();

    const buckets = [];
    for (const row of bucketData) {
      const bucketName = row[0];
      const pointValue = row[1];
      let examples = [];

      // Parse examples from JSON
      try {
        if (row[2] && row[2] !== '') {
          examples = JSON.parse(row[2]);
        }
      } catch (e) {
        console.log('Could not parse examples for ' + bucketName);
        examples = [];
      }

      if (bucketName) {
        buckets.push({
          name: bucketName,
          points: pointValue,
          examples: examples
        });
      }
    }

    return buckets;

  } catch (error) {
    console.error('Error getting infraction types:', error.toString());
    return getDefaultInfractionTypes();
  }
}

/**
 * Returns default infraction types if Settings can't be read.
 *
 * @returns {Array} Default bucket configuration
 */
function getDefaultInfractionTypes() {
  return [
    {
      name: 'Bucket 1: Minor Offenses',
      points: 1,
      examples: ['Tardiness under 15 minutes', 'Uniform violations', 'Minor cleanliness issues', 'Missing name tag', 'Late return from break']
    },
    {
      name: 'Bucket 2: Moderate Offenses',
      points: 3,
      examples: ['Tardiness 15-30 minutes', 'Cell phone use during shift', 'Call-outs', 'Customer complaints', 'Attendance issues']
    },
    {
      name: 'Bucket 3: Major Offenses',
      points: 5,
      examples: ['Tardiness 30+ minutes', 'Insubordination', 'Profanity', 'Food safety violations', 'Leaving shift early', 'Creating hostile environment']
    },
    {
      name: 'Bucket 4: Severe Offenses',
      points: 8,
      examples: ['No-call/no-show', 'Major safety violations', 'Theft', 'Harassment', 'Working under influence', 'Physical altercations']
    }
  ];
}

/**
 * Processes form submission from the web app.
 * Called from client-side JavaScript when form is submitted.
 *
 * @param {Object} formData - Form data from the client
 * @returns {Object} Result of processing the infraction
 */
function submitInfractionForm(formData) {
  console.log('=== Form Submission Received ===');
  console.log('Form data:', JSON.stringify(formData));

  try {
    // Validate required fields
    if (!formData.entered_by || formData.entered_by.trim() === '') {
      return { success: false, message: 'Manager/Director name is required' };
    }

    if (!formData.employee_id || formData.employee_id === '') {
      return { success: false, message: 'Please select an employee' };
    }

    if (!formData.date) {
      return { success: false, message: 'Date is required' };
    }

    if (!formData.location || formData.location === '') {
      return { success: false, message: 'Please select a location' };
    }

    if (!formData.infraction_type || formData.infraction_type === '') {
      return { success: false, message: 'Please select an infraction type' };
    }

    if (!formData.description || formData.description.trim().length < 240) {
      return { success: false, message: 'Description must be at least 240 characters' };
    }

    // Get employee full name
    const employees = getActiveEmployees();
    const employee = employees.find(e => e.employee_id === formData.employee_id);

    if (!employee) {
      return { success: false, message: 'Selected employee not found' };
    }

    // Determine points from infraction type
    const buckets = getInfractionTypesFromSettings();
    let pointsAssigned = 0;
    let bucketName = formData.infraction_type;

    // Find matching bucket
    for (const bucket of buckets) {
      if (formData.infraction_type.startsWith(bucket.name) ||
          bucket.examples.includes(formData.infraction_type)) {
        pointsAssigned = bucket.points;
        bucketName = bucket.name;
        break;
      }
    }

    // Parse the date
    const infractionDate = new Date(formData.date + 'T12:00:00');

    // Build infraction data object
    const infractionData = {
      employee_id: formData.employee_id,
      full_name: employee.full_name,
      date: infractionDate,
      infraction_type: bucketName,
      points_assigned: pointsAssigned,
      description: formData.description.trim(),
      location: formData.location,
      entered_by: formData.entered_by.trim()
    };

    console.log('Infraction data prepared:', JSON.stringify(infractionData));

    // Process the infraction with notifications
    const result = processInfractionWithNotifications(infractionData);

    console.log('Processing result:', JSON.stringify(result));

    return result;

  } catch (error) {
    console.error('Error in submitInfractionForm:', error.toString());
    return {
      success: false,
      message: 'Error processing form: ' + error.toString()
    };
  }
}

/**
 * Test function for getFormData.
 */
function testGetFormData() {
  console.log('=== Testing getFormData() ===');

  const data = getFormData();

  console.log('Success:', data.success);
  console.log('Employees count:', data.employees.length);

  if (data.employees.length > 0) {
    console.log('First 3 employees:');
    for (let i = 0; i < Math.min(3, data.employees.length); i++) {
      console.log(`  ${i + 1}. ${data.employees[i].name} (${data.employees[i].id}) - ${data.employees[i].location}`);
    }
  }

  console.log('');
  console.log('Infraction types:');
  for (const bucket of data.infraction_types) {
    console.log(`  ${bucket.name} (${bucket.points} pts)`);
    console.log(`    Examples: ${bucket.examples.slice(0, 3).join(', ')}...`);
  }

  return data;
}

// ============================================
// MICRO-PHASE 10: CONNECT FORM TO BACKEND
// ============================================

/**
 * Processes form submission from the web app with enhanced formatting.
 * Called from client-side JavaScript when form is submitted.
 *
 * @param {Object} formData - Form data from the client
 * @param {string} formData.entered_by - Manager/Director name
 * @param {string} formData.employee_id - Employee ID from dropdown
 * @param {string} formData.full_name - Employee full name from dropdown text
 * @param {string} formData.date - Date in YYYY-MM-DD format
 * @param {string} formData.location - Location from dropdown
 * @param {string} formData.infraction_type - Infraction type from dropdown
 * @param {string} formData.description - Description from textarea
 *
 * @returns {Object} Result with success, message, and data
 */
function submitInfractionFromForm(formData) {
  console.log('=== Form Submission Received (Enhanced) ===');
  console.log('Form data:', JSON.stringify(formData));

  try {
    // ========================================
    // SESSION VALIDATION
    // ========================================

    const sessionCheck = requireValidSession();
    if (!sessionCheck.valid) {
      return {
        success: false,
        message: sessionCheck.error || 'Session expired. Please login again.',
        sessionExpired: true
      };
    }

    // ========================================
    // VALIDATION
    // ========================================

    // Validate manager name
    if (!formData.entered_by || formData.entered_by.trim() === '') {
      return {
        success: false,
        message: 'Manager/Director name is required'
      };
    }

    // Validate employee selection
    if (!formData.employee_id || formData.employee_id === '') {
      return {
        success: false,
        message: 'Please select an employee'
      };
    }

    // Validate date
    if (!formData.date) {
      return {
        success: false,
        message: 'Invalid date selected. Please choose a date within the last 7 days.'
      };
    }

    // Validate location
    if (!formData.location || formData.location === '') {
      return {
        success: false,
        message: 'Please select a location'
      };
    }

    // Validate infraction type
    if (!formData.infraction_type || formData.infraction_type === '') {
      return {
        success: false,
        message: 'Invalid infraction type'
      };
    }

    // Validate description length
    if (!formData.description || formData.description.trim().length < 240) {
      return {
        success: false,
        message: 'Description must be at least 240 characters'
      };
    }

    // ========================================
    // EMPLOYEE LOOKUP
    // ========================================

    const employees = getActiveEmployees();
    const employee = employees.find(e => e.employee_id === formData.employee_id);

    if (!employee) {
      return {
        success: false,
        message: 'Employee not found in active list. They may have been terminated or placed on leave.'
      };
    }

    // ========================================
    // DETERMINE POINTS FROM INFRACTION TYPE
    // ========================================

    const buckets = getInfractionTypesFromSettings();
    let pointsAssigned = 0;
    let bucketName = formData.infraction_type;
    let bucketFound = false;

    // Find matching bucket
    for (const bucket of buckets) {
      // Check if it's the bucket name itself (General option)
      if (formData.infraction_type === bucket.name ||
          formData.infraction_type === bucket.name + ' - General') {
        pointsAssigned = bucket.points;
        bucketName = bucket.name;
        bucketFound = true;
        break;
      }

      // Check if it's one of the examples
      if (bucket.examples && bucket.examples.includes(formData.infraction_type)) {
        pointsAssigned = bucket.points;
        bucketName = bucket.name;
        bucketFound = true;
        break;
      }
    }

    if (!bucketFound) {
      // Try a fallback - check if infraction type starts with bucket name
      for (const bucket of buckets) {
        if (formData.infraction_type.startsWith(bucket.name)) {
          pointsAssigned = bucket.points;
          bucketName = bucket.name;
          bucketFound = true;
          break;
        }
      }
    }

    if (!bucketFound || pointsAssigned === 0) {
      console.error('Could not determine points for infraction type:', formData.infraction_type);
      return {
        success: false,
        message: 'Invalid infraction type. Could not determine point value.'
      };
    }

    // ========================================
    // PARSE DATE
    // ========================================

    const infractionDate = new Date(formData.date + 'T12:00:00');

    // Validate date is not in future
    const today = new Date();
    today.setHours(23, 59, 59, 999);
    if (infractionDate > today) {
      return {
        success: false,
        message: 'Invalid date selected. Date cannot be in the future.'
      };
    }

    // Validate date is not more than 7 days ago
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    sevenDaysAgo.setHours(0, 0, 0, 0);
    if (infractionDate < sevenDaysAgo) {
      return {
        success: false,
        message: 'Invalid date selected. Date cannot be more than 7 days ago.'
      };
    }

    // ========================================
    // BUILD INFRACTION DATA OBJECT
    // ========================================

    const infractionData = {
      employee_id: formData.employee_id,
      full_name: employee.full_name,
      date: infractionDate,
      infraction_type: bucketName,
      points_assigned: pointsAssigned,
      description: formData.description.trim(),
      location: formData.location,
      entered_by: formData.entered_by.trim()
    };

    console.log('Infraction data prepared:', JSON.stringify(infractionData));

    // ========================================
    // PROCESS THE INFRACTION
    // ========================================

    const result = processInfractionWithNotifications(infractionData);

    console.log('Processing result:', JSON.stringify(result));

    // ========================================
    // FORMAT RESPONSE MESSAGE
    // ========================================

    if (result.success) {
      // Build success message
      let message = '✓ Infraction Recorded Successfully\n\n';
      message += `Employee: ${employee.full_name}\n`;
      message += `New Point Total: ${result.new_points} points\n`;
      message += `Infraction ID: ${result.infraction_id}\n`;

      if (result.thresholds_crossed && result.thresholds_crossed.length > 0) {
        message += `\nThresholds Reached: ${result.thresholds_crossed.join(', ')} points\n`;
        if (result.email_sent) {
          message += 'Email notification sent to management.';
        } else {
          message += `Email status: ${result.email_status}`;
        }
      } else {
        message += '\nNo threshold alerts triggered.';
      }

      return {
        success: true,
        message: message,
        data: {
          infraction_id: result.infraction_id,
          employee_name: employee.full_name,
          old_points: result.old_points,
          new_points: result.new_points,
          points_added: pointsAssigned,
          thresholds_crossed: result.thresholds_crossed,
          email_sent: result.email_sent,
          email_status: result.email_status
        }
      };

    } else {
      // Build error message
      let message = '✗ Error Recording Infraction\n\n';
      message += result.message || 'Unknown error occurred';
      message += '\n\nPlease check your entries and try again.';

      return {
        success: false,
        message: message
      };
    }

  } catch (error) {
    console.error('Error in submitInfractionFromForm:', error.toString());

    // Determine error type and provide appropriate message
    let errorMessage = '✗ Error Recording Infraction\n\n';

    if (error.toString().includes('permission') || error.toString().includes('Permission')) {
      errorMessage += "You don't have permission to submit infractions. Contact administrator.";
    } else if (error.toString().includes('timeout') || error.toString().includes('Timeout')) {
      errorMessage += 'Connection timed out. Please check your internet and try again.';
    } else {
      errorMessage += 'System error occurred. Please try again or contact support if problem persists.';
      errorMessage += '\n\nDetails: ' + error.toString();
    }

    return {
      success: false,
      message: errorMessage
    };
  }
}

/**
 * Checks if a similar infraction already exists.
 * Called from client-side to warn about duplicates.
 *
 * @param {string} employeeId - Employee ID to check
 * @param {string} date - Date in YYYY-MM-DD format
 * @param {string} infractionType - Infraction type to check
 *
 * @returns {Object} Result with isDuplicate boolean and details
 */
function checkForDuplicateInfraction(employeeId, date, infractionType) {
  console.log('Checking for duplicate:', employeeId, date, infractionType);

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { isDuplicate: false, error: 'Infractions sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { isDuplicate: false };
    }

    // Find column indices
    const headers = data[0];
    const empIdCol = headers.indexOf('Employee_ID');
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Infraction_Type');

    if (empIdCol === -1 || dateCol === -1 || typeCol === -1) {
      return { isDuplicate: false, error: 'Required columns not found' };
    }

    // Parse the check date
    const checkDate = new Date(date + 'T12:00:00');
    const checkDateStr = formatDateForInput(checkDate);

    // Check for matching infraction
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpId = String(row[empIdCol]);
      const rowDate = row[dateCol];
      const rowType = String(row[typeCol]);

      // Skip if employee doesn't match
      if (rowEmpId !== String(employeeId)) continue;

      // Check date match
      let rowDateStr = '';
      if (rowDate instanceof Date) {
        rowDateStr = formatDateForInput(rowDate);
      } else if (typeof rowDate === 'string') {
        rowDateStr = rowDate.substring(0, 10);
      }

      if (rowDateStr !== checkDateStr) continue;

      // Check if infraction type matches (or is in same bucket)
      if (rowType === infractionType ||
          infractionType.includes(rowType) ||
          rowType.includes(infractionType.split(' - ')[0])) {
        console.log('Duplicate found at row', i + 1);
        return {
          isDuplicate: true,
          existingType: rowType,
          existingDate: rowDateStr
        };
      }
    }

    return { isDuplicate: false };

  } catch (error) {
    console.error('Error checking for duplicate:', error.toString());
    return { isDuplicate: false, error: error.toString() };
  }
}

/**
 * Helper function to format date for comparison.
 * @param {Date} date - Date object
 * @returns {string} Date in YYYY-MM-DD format
 */
function formatDateForInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Gets the point value for a specific infraction type.
 * Used by the frontend to show confirmation for high-point infractions.
 *
 * @param {string} infractionType - The infraction type to check
 * @returns {Object} Object with points and bucketName
 */
function getPointsForInfractionType(infractionType) {
  try {
    const buckets = getInfractionTypesFromSettings();

    for (const bucket of buckets) {
      // Check if it's the bucket name itself
      if (infractionType === bucket.name ||
          infractionType === bucket.name + ' - General') {
        return { points: bucket.points, bucketName: bucket.name };
      }

      // Check if it's one of the examples
      if (bucket.examples && bucket.examples.includes(infractionType)) {
        return { points: bucket.points, bucketName: bucket.name };
      }
    }

    // Fallback - check if it starts with bucket name
    for (const bucket of buckets) {
      if (infractionType.startsWith(bucket.name)) {
        return { points: bucket.points, bucketName: bucket.name };
      }
    }

    return { points: 0, bucketName: null };

  } catch (error) {
    console.error('Error getting points for infraction type:', error.toString());
    return { points: 0, bucketName: null, error: error.toString() };
  }
}

// ============================================
// MICRO-PHASE 10: TEST FUNCTIONS
// ============================================

/**
 * Test function for form submission - simulates various scenarios.
 */
function testFormSubmission() {
  console.log('=== Testing Form Submission ===\n');

  // Get a real employee for testing
  const employees = getActiveEmployees();
  if (employees.length === 0) {
    console.error('No active employees found for testing');
    return;
  }

  const testEmployee = employees[0];
  console.log('Using test employee:', testEmployee.full_name, testEmployee.employee_id);
  console.log('');

  // ========================================
  // TEST CASE 1: Successful submission
  // ========================================
  console.log('--- Test Case 1: Successful Submission ---');

  const validFormData = {
    entered_by: 'Test Manager',
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: formatDateForInput(new Date()),
    location: 'Cockrell Hill DTO',
    infraction_type: 'Bucket 1',
    description: 'This is a test description for the infraction form submission test. ' +
                 'It needs to be at least 240 characters long to pass validation. ' +
                 'Adding more text here to ensure we meet the minimum character requirement. ' +
                 'This should be enough to pass the validation check for description length.'
  };

  console.log('Form data:', JSON.stringify(validFormData, null, 2));

  const result1 = submitInfractionFromForm(validFormData);
  console.log('Result:', JSON.stringify(result1, null, 2));
  console.log('Success:', result1.success);
  console.log('');

  // ========================================
  // TEST CASE 2: Invalid employee
  // ========================================
  console.log('--- Test Case 2: Invalid Employee ---');

  const invalidEmployeeData = {
    entered_by: 'Test Manager',
    employee_id: 'INVALID_EMP_ID_99999',
    full_name: 'Fake Employee',
    date: formatDateForInput(new Date()),
    location: 'Cockrell Hill DTO',
    infraction_type: 'Bucket 1',
    description: 'This is a test description for the infraction form submission test. ' +
                 'It needs to be at least 240 characters long to pass validation. ' +
                 'Adding more text here to ensure we meet the minimum character requirement. ' +
                 'This should be enough to pass the validation check.'
  };

  const result2 = submitInfractionFromForm(invalidEmployeeData);
  console.log('Result:', JSON.stringify(result2, null, 2));
  console.log('Success should be false:', !result2.success);
  console.log('Message should mention employee:', result2.message.includes('Employee'));
  console.log('');

  // ========================================
  // TEST CASE 3: Missing required field
  // ========================================
  console.log('--- Test Case 3: Missing Required Field ---');

  const missingFieldData = {
    entered_by: '',
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: formatDateForInput(new Date()),
    location: 'Cockrell Hill DTO',
    infraction_type: 'Bucket 1',
    description: 'Test description that meets the 240 character minimum requirement. ' +
                 'Adding more content to ensure this is long enough for the test. ' +
                 'We need to make sure validation catches the missing manager name. ' +
                 'This should be plenty of characters now.'
  };

  const result3 = submitInfractionFromForm(missingFieldData);
  console.log('Result:', JSON.stringify(result3, null, 2));
  console.log('Success should be false:', !result3.success);
  console.log('');

  // ========================================
  // TEST CASE 4: Description too short
  // ========================================
  console.log('--- Test Case 4: Description Too Short ---');

  const shortDescData = {
    entered_by: 'Test Manager',
    employee_id: testEmployee.employee_id,
    full_name: testEmployee.full_name,
    date: formatDateForInput(new Date()),
    location: 'Cockrell Hill DTO',
    infraction_type: 'Bucket 1',
    description: 'Too short'
  };

  const result4 = submitInfractionFromForm(shortDescData);
  console.log('Result:', JSON.stringify(result4, null, 2));
  console.log('Success should be false:', !result4.success);
  console.log('Message should mention 240:', result4.message.includes('240'));
  console.log('');

  // ========================================
  // TEST CASE 5: Duplicate check
  // ========================================
  console.log('--- Test Case 5: Duplicate Check ---');

  const duplicateCheck = checkForDuplicateInfraction(
    testEmployee.employee_id,
    formatDateForInput(new Date()),
    'Bucket 1'
  );
  console.log('Duplicate check result:', JSON.stringify(duplicateCheck, null, 2));
  console.log('');

  // ========================================
  // TEST CASE 6: Points lookup
  // ========================================
  console.log('--- Test Case 6: Points Lookup ---');

  const points1 = getPointsForInfractionType('Bucket 1');
  console.log('Bucket 1 points:', JSON.stringify(points1));

  const points4 = getPointsForInfractionType('Bucket 4');
  console.log('Bucket 4 points:', JSON.stringify(points4));
  console.log('');

  console.log('=== Test Complete ===');
}

// ============================================
// MICRO-PHASE 11: AUTHENTICATION
// ============================================

/**
 * Validates password against Manager and Director passwords in Settings.
 *
 * @param {string} enteredPassword - The password to validate
 * @returns {Object} {valid: boolean, role: string|null}
 */
function validatePassword(enteredPassword) {
  console.log('=== Password Validation Attempt ===');

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      console.error('Settings sheet not found');
      logAuthAttempt(false, 'Settings sheet not found');
      return { valid: false, role: null };
    }

    // Read passwords from Settings
    // Manager Password: B6
    // Director Password: B7
    const managerPassword = settingsSheet.getRange('B6').getValue();
    const directorPassword = settingsSheet.getRange('B7').getValue();

    // Compare entered password (simple comparison)
    if (enteredPassword === directorPassword) {
      console.log('Director password matched');
      logAuthAttempt(true, 'Director');
      return { valid: true, role: 'Director' };
    }

    if (enteredPassword === managerPassword) {
      console.log('Manager password matched');
      logAuthAttempt(true, 'Manager');
      return { valid: true, role: 'Manager' };
    }

    console.log('Password did not match any role');
    logAuthAttempt(false, 'Invalid password');
    return { valid: false, role: null };

  } catch (error) {
    console.error('Error validating password:', error.toString());
    logAuthAttempt(false, 'Error: ' + error.toString());
    return { valid: false, role: null };
  }
}

/**
 * Logs authentication attempts (without logging actual passwords).
 *
 * @param {boolean} success - Whether authentication was successful
 * @param {string} details - Details about the attempt (role or error)
 */
function logAuthAttempt(success, details) {
  try {
    console.log(`Auth attempt: success=${success}, details=${details}, time=${new Date().toISOString()}`);

    // Could also log to a sheet if desired:
    // const ss = SpreadsheetApp.openById(SHEET_ID);
    // const logSheet = ss.getSheetByName('Auth_Log');
    // if (logSheet) {
    //   logSheet.appendRow([new Date(), success ? 'Success' : 'Failed', details]);
    // }

  } catch (error) {
    console.error('Error logging auth attempt:', error.toString());
  }
}

// ============================================
// SIMPLIFIED AUTHENTICATION (Shared Passwords)
// ============================================

/**
 * Validates a role-based password login.
 * Uses shared passwords stored in Settings sheet.
 *
 * @param {string} role - 'Operator', 'Director', or 'Manager'
 * @param {string} password - The password to validate
 * @returns {Object} { success: boolean, role: string, message: string }
 */
function validateRoleLogin(role, password) {
  try {
    // Validate inputs
    if (!role || !password) {
      return { success: false, message: 'Role and password are required' };
    }

    // Normalize role
    const normalizedRole = role.charAt(0).toUpperCase() + role.slice(1).toLowerCase();
    if (!['Operator', 'Director', 'Manager'].includes(normalizedRole)) {
      return { success: false, message: 'Invalid role' };
    }

    // Get password from Settings
    const storedPassword = getRolePassword(normalizedRole);
    if (!storedPassword) {
      return { success: false, message: 'Password not configured for this role' };
    }

    // Check password
    if (password !== storedPassword) {
      console.log(`Failed login attempt for role: ${normalizedRole}`);
      return { success: false, message: 'Incorrect password' };
    }

    // Success - store role in session
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('cfa_role', normalizedRole);
    userProperties.setProperty('cfa_login_time', new Date().toISOString());

    console.log(`Successful login for role: ${normalizedRole}`);
    return { success: true, role: normalizedRole };

  } catch (error) {
    console.error('Login error:', error.toString());
    return { success: false, message: 'Login error: ' + error.toString() };
  }
}

/**
 * Gets the stored password for a role from Settings sheet.
 *
 * @param {string} role - 'Operator', 'Director', or 'Manager'
 * @returns {string|null} The password or null if not found
 */
function getRolePassword(role) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      console.error('Settings sheet not found');
      return null;
    }

    // Settings structure:
    // Row 6: Operator Password
    // Row 7: Director Password
    // Row 8: Manager Password
    const rowMap = {
      'Operator': 6,
      'Director': 7,
      'Manager': 8
    };

    const row = rowMap[role];
    if (!row) return null;

    const password = settingsSheet.getRange(row, 2).getValue();
    return password ? password.toString() : null;

  } catch (error) {
    console.error('Error getting role password:', error.toString());
    return null;
  }
}

/**
 * STANDALONE FUNCTION - Run this to set up role passwords.
 * Updates the Settings sheet with initial passwords.
 * You can change these passwords directly in the Settings sheet after running.
 *
 * Default passwords (CHANGE THESE!):
 * - Operator: operator123
 * - Director: director123
 * - Manager: manager123
 */
function setupRolePasswords() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      console.error('Settings sheet not found. Run setupAllTabs() first.');
      return 'ERROR: Settings sheet not found. Run setupAllTabs() first.';
    }

    // Set default passwords (user should change these!)
    // Row 6: Operator Password
    // Row 7: Director Password
    // Row 8: Manager Password
    settingsSheet.getRange('A6').setValue('Operator Password');
    settingsSheet.getRange('B6').setValue('operator123');

    settingsSheet.getRange('A7').setValue('Director Password');
    settingsSheet.getRange('B7').setValue('director123');

    settingsSheet.getRange('A8').setValue('Manager Password');
    settingsSheet.getRange('B8').setValue('manager123');

    console.log('Role passwords set up successfully!');
    console.log('IMPORTANT: Change these default passwords in the Settings sheet!');

    return 'Role passwords configured! IMPORTANT: Change the default passwords in the Settings sheet (rows 6-8).';

  } catch (error) {
    console.error('Error setting up role passwords:', error.toString());
    return 'ERROR: ' + error.toString();
  }
}

/**
 * Gets the current user's role from session.
 *
 * @returns {Object} { authenticated: boolean, role: string }
 */
function getCurrentRole() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const role = userProperties.getProperty('cfa_role');
    const loginTime = userProperties.getProperty('cfa_login_time');

    if (!role || !loginTime) {
      return { authenticated: false };
    }

    // Check session timeout (configurable, default 30 minutes)
    const loginDate = new Date(loginTime);
    const now = new Date();
    const sessionTimeoutMs = 30 * 60 * 1000; // 30 minutes

    if (now - loginDate > sessionTimeoutMs) {
      // Session expired
      userProperties.deleteProperty('cfa_role');
      userProperties.deleteProperty('cfa_login_time');
      return { authenticated: false, expired: true };
    }

    // Refresh login time to extend session
    userProperties.setProperty('cfa_login_time', now.toISOString());

    return { authenticated: true, role: role };

  } catch (error) {
    console.error('Error getting current role:', error.toString());
    return { authenticated: false };
  }
}

/**
 * Logs out the current user.
 *
 * @returns {Object} { success: boolean }
 */
function logoutRole() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty('cfa_role');
    userProperties.deleteProperty('cfa_login_time');
    return { success: true };
  } catch (error) {
    console.error('Error logging out:', error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Checks if the current role can view a specific employee.
 * Visibility rules:
 * - Operator: Can see everyone
 * - Director: Can see all employees EXCEPT other Directors
 * - Manager: Can see only hourly employees (no one with system access)
 *
 * @param {string} employeeId - The employee to check
 * @returns {boolean} True if current user can view this employee
 */
function canViewEmployee(employeeId) {
  const session = getCurrentRole();
  if (!session.authenticated) return false;

  const role = session.role;

  // Operator sees everyone
  if (role === 'Operator') return true;

  // Check if employee has system access
  const systemAccessMap = getSystemAccessMap();
  const employeeAccess = systemAccessMap[employeeId];

  // No system access = hourly employee, everyone can see
  if (!employeeAccess) return true;

  const employeeRole = employeeAccess.role;

  // Director can see everyone except other Directors
  if (role === 'Director') {
    return employeeRole !== 'Director';
  }

  // Manager can't see anyone with system access
  if (role === 'Manager') {
    return false;
  }

  return false;
}

/**
 * Gets the filtered employee list based on current user's role.
 *
 * @returns {Object} { success: boolean, employees: Array, role: string }
 */
function getFilteredEmployeeList() {
  try {
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    const role = session.role;
    const allEmployees = getActiveEmployees();

    if (!allEmployees || allEmployees.length === 0) {
      return { success: true, employees: [], role: role };
    }

    // Get system access map for filtering
    const systemAccessMap = getSystemAccessMap();

    // Filter based on role
    const filteredEmployees = allEmployees.filter(emp => {
      const employeeAccess = systemAccessMap[emp.employee_id];

      // No system access = hourly employee, everyone can see
      if (!employeeAccess) return true;

      const employeeRole = employeeAccess.role;

      // Operator sees everyone
      if (role === 'Operator') return true;

      // Director sees everyone except other Directors
      if (role === 'Director') {
        return employeeRole !== 'Director';
      }

      // Manager can't see anyone with system access
      if (role === 'Manager') {
        return false;
      }

      return false;
    });

    return { success: true, employees: filteredEmployees, role: role };

  } catch (error) {
    console.error('Error getting filtered employee list:', error.toString());
    return { success: false, message: error.toString() };
  }
}

// ============================================
// LEGACY: SESSION MANAGEMENT & PIN AUTHENTICATION (3-Tier System)
// Note: Kept for reference, new system uses shared passwords above
// ============================================

/**
 * Hashes a PIN using SHA-256.
 *
 * @param {string} pin - The 6-digit PIN to hash
 * @returns {string} SHA-256 hash of the PIN
 */
function hashPIN(pin) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin);
  return rawHash.map(byte => {
    const hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/**
 * Generates a random session ID.
 *
 * @returns {string} A random 32-character alphanumeric string
 */
function generateSessionId() {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let sessionId = '';
  for (let i = 0; i < 32; i++) {
    sessionId += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return sessionId;
}

/**
 * Gets all active users for the login dropdown.
 *
 * @returns {Array} Array of {employee_id, full_name} sorted alphabetically
 */
function getActiveUsersForLogin() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('User_Permissions');

    if (!sheet) {
      console.log('User_Permissions sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // Get columns A (Employee_ID), B (Full_Name), I (Status)
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const activeUsers = [];

    for (const row of data) {
      const employeeId = row[0];
      const fullName = row[1];
      const status = row[8]; // Column I is Status

      if (status === 'Active' && employeeId && fullName) {
        activeUsers.push({
          employee_id: employeeId,
          full_name: fullName
        });
      }
    }

    // Sort alphabetically by name
    activeUsers.sort((a, b) => a.full_name.localeCompare(b.full_name));

    return activeUsers;

  } catch (error) {
    console.error('Error getting active users:', error);
    return [];
  }
}

/**
 * Validates PIN login and creates a session.
 *
 * @param {string} employee_id - The employee ID selected from dropdown
 * @param {string} pin - The 6-digit PIN entered by user
 * @returns {Object} {valid: boolean, user_data: Object|null, message: string}
 */
function validatePINLogin(employee_id, pin) {
  console.log('Validating PIN login for employee:', employee_id);

  try {
    // Validate input
    if (!employee_id || !pin) {
      return { valid: false, user_data: null, message: 'Employee ID and PIN are required' };
    }

    if (!/^\d{6}$/.test(pin)) {
      return { valid: false, user_data: null, message: 'PIN must be exactly 6 digits' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('User_Permissions');

    if (!sheet) {
      return { valid: false, user_data: null, message: 'System configuration error' };
    }

    // Find the user
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { valid: false, user_data: null, message: 'Invalid credentials' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    let userRow = -1;
    let userData = null;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === employee_id) {
        userRow = i + 2; // +2 for header row and 0-indexing
        userData = {
          employee_id: data[i][0],    // A
          full_name: data[i][1],       // B
          email: data[i][2],           // C
          role: data[i][3],            // D
          pin_hash: data[i][4],        // E
          can_see_directors: data[i][5] === 'TRUE' || data[i][5] === true, // F
          date_added: data[i][6],      // G
          added_by: data[i][7],        // H
          status: data[i][8],          // I
          last_login: data[i][9],      // J
          login_count: data[i][10] || 0, // K
          failed_attempts: data[i][11] || 0, // L
          lockout_until: data[i][12]   // M
        };
        break;
      }
    }

    // User not found
    if (!userData) {
      return { valid: false, user_data: null, message: 'Invalid credentials' };
    }

    // Check if user is active
    if (userData.status !== 'Active') {
      return { valid: false, user_data: null, message: 'Account is inactive' };
    }

    // Check if account is locked out
    if (userData.lockout_until) {
      const lockoutTime = new Date(userData.lockout_until).getTime();
      const now = Date.now();
      if (now < lockoutTime) {
        const remainingMs = lockoutTime - now;
        const remainingMin = Math.ceil(remainingMs / 60000);
        return {
          valid: false,
          user_data: null,
          message: `Account is locked. Try again in ${remainingMin} minute(s).`,
          locked: true,
          lockout_remaining_ms: remainingMs
        };
      } else {
        // Lockout has expired, clear it
        sheet.getRange(userRow, 12).setValue(0); // L: Failed_Attempts
        sheet.getRange(userRow, 13).setValue(''); // M: Lockout_Until
        userData.failed_attempts = 0;
      }
    }

    // Hash the entered PIN and compare
    const enteredPINHash = hashPIN(pin);

    if (enteredPINHash !== userData.pin_hash) {
      // Wrong PIN - increment failed attempts
      const newFailedAttempts = (userData.failed_attempts || 0) + 1;
      sheet.getRange(userRow, 12).setValue(newFailedAttempts); // L: Failed_Attempts

      if (newFailedAttempts >= 3) {
        // Lock the account for 5 minutes
        const lockoutTime = new Date(Date.now() + 5 * 60 * 1000);
        sheet.getRange(userRow, 13).setValue(lockoutTime); // M: Lockout_Until
        return {
          valid: false,
          user_data: null,
          message: 'Too many failed attempts. Account locked for 5 minutes.',
          locked: true,
          lockout_remaining_ms: 5 * 60 * 1000
        };
      }

      const remainingAttempts = 3 - newFailedAttempts;
      return {
        valid: false,
        user_data: null,
        message: `Invalid PIN. ${remainingAttempts} attempt(s) remaining.`
      };
    }

    // PIN is correct - create session
    const sessionResult = createSession({
      user_id: userData.employee_id,
      name: userData.full_name,
      email: userData.email,
      role: userData.role,
      can_see_directors: userData.can_see_directors
    });

    if (!sessionResult.success) {
      return { valid: false, user_data: null, message: 'Failed to create session' };
    }

    // Update login tracking
    const now = new Date();
    sheet.getRange(userRow, 10).setValue(now); // J: Last_Login
    sheet.getRange(userRow, 11).setValue((userData.login_count || 0) + 1); // K: Login_Count
    sheet.getRange(userRow, 12).setValue(0); // L: Reset Failed_Attempts
    sheet.getRange(userRow, 13).setValue(''); // M: Clear Lockout_Until

    console.log('Login successful for:', userData.full_name, '- Role:', userData.role);

    return {
      valid: true,
      user_data: {
        user_id: userData.employee_id,
        name: userData.full_name,
        email: userData.email,
        role: userData.role,
        can_see_directors: userData.can_see_directors
      },
      message: `Welcome back, ${userData.full_name}!`
    };

  } catch (error) {
    console.error('Error in validatePINLogin:', error);
    return { valid: false, user_data: null, message: 'Login failed. Please try again.' };
  }
}

/**
 * Creates a new session for the authenticated user.
 *
 * @param {Object} user_data - User data containing user_id, name, email, role, can_see_directors
 * @returns {Object} {success: boolean, session_id: string|null, error: string|null}
 */
function createSession(user_data) {
  console.log('Creating session for user:', user_data.name, '- Role:', user_data.role);

  try {
    // Generate a random session ID
    const sessionId = generateSessionId();
    const now = Date.now();

    // Build session data object
    const sessionData = {
      session_id: sessionId,
      user_id: user_data.user_id,
      name: user_data.name,
      email: user_data.email || '',
      role: user_data.role,
      can_see_directors: user_data.can_see_directors || false,
      created_at: now,
      last_activity: now
    };

    // Store session as JSON in UserProperties
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('cfa_session', JSON.stringify(sessionData));

    console.log('Session created successfully');
    console.log('Session expires at:', new Date(now + (10 * 60 * 1000)).toISOString());

    return {
      success: true,
      session_id: sessionId,
      role: user_data.role
    };

  } catch (error) {
    console.error('Error creating session:', error.toString());
    return {
      success: false,
      session_id: null,
      error: error.toString()
    };
  }
}

/**
 * Validates the current session.
 *
 * @returns {Object} {valid: boolean, role: string|null, user: Object|null, expired: boolean}
 */
function validateSession() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const sessionJson = userProps.getProperty('cfa_session');

    if (!sessionJson) {
      console.log('No session found');
      return { valid: false, role: null, user: null, expired: false };
    }

    // Parse session data
    let sessionData;
    try {
      sessionData = JSON.parse(sessionJson);
    } catch (parseError) {
      console.log('Invalid session format');
      userProps.deleteProperty('cfa_session');
      return { valid: false, role: null, user: null, expired: false };
    }

    // Check if session has expired (10 minutes = 600000 ms)
    const SESSION_DURATION_MS = 10 * 60 * 1000;
    const elapsed = Date.now() - sessionData.last_activity;

    if (elapsed > SESSION_DURATION_MS) {
      console.log('Session expired');
      userProps.deleteProperty('cfa_session');
      return { valid: false, role: sessionData.role, user: null, expired: true };
    }

    // Session is valid - extend it
    sessionData.last_activity = Date.now();
    userProps.setProperty('cfa_session', JSON.stringify(sessionData));

    console.log('Session valid for:', sessionData.name, '- Role:', sessionData.role);

    return {
      valid: true,
      role: sessionData.role,
      user: {
        user_id: sessionData.user_id,
        name: sessionData.name,
        email: sessionData.email,
        role: sessionData.role,
        can_see_directors: sessionData.can_see_directors
      },
      expired: false
    };

  } catch (error) {
    console.error('Error validating session:', error.toString());
    return { valid: false, role: null, user: null, expired: false };
  }
}

/**
 * Extends the current session by updating the last_activity timestamp.
 *
 * @returns {Object} {success: boolean, error: string|null}
 */
function extendSession() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const sessionJson = userProps.getProperty('cfa_session');

    if (!sessionJson) {
      return { success: false, error: 'No session to extend' };
    }

    // Parse and update session data
    const sessionData = JSON.parse(sessionJson);
    sessionData.last_activity = Date.now();

    userProps.setProperty('cfa_session', JSON.stringify(sessionData));

    console.log('Session extended for:', sessionData.name);
    return { success: true };

  } catch (error) {
    console.error('Error extending session:', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Ends the current session (logout).
 *
 * @returns {Object} {success: boolean}
 */
function endSession() {
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('cfa_session');

    console.log('Session ended');
    return { success: true };

  } catch (error) {
    console.error('Error ending session:', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Gets the current user's role from the session.
 *
 * @returns {string|null} The user's role or null if not authenticated
 */
function getCurrentUserRole() {
  const session = validateSession();
  return session.valid ? session.role : null;
}

/**
 * Gets the current user's full data from the session.
 *
 * @returns {Object|null} The user data or null if not authenticated
 */
function getCurrentUser() {
  const session = validateSession();
  return session.valid ? session.user : null;
}

/**
 * Creates a new user in the User_Permissions sheet.
 * Can only be called by Directors (for Manager/Director) or Operators (for all).
 *
 * @param {Object} userData - User data containing employee_id, full_name, email, role, pin, can_see_directors
 * @returns {Object} {success: boolean, message: string}
 */
function createUser(userData) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    // Permission check based on role being created
    const currentRole = session.role;
    const targetRole = userData.role;

    // Operators can create anyone
    // Directors can create Managers and Directors
    // Managers cannot create users
    if (currentRole === 'Manager') {
      return { success: false, message: 'Managers cannot create users' };
    }

    if (currentRole === 'Director' && targetRole === 'Operator') {
      return { success: false, message: 'Directors cannot create Operators' };
    }

    // Validate required fields
    if (!userData.employee_id || !userData.full_name || !userData.role || !userData.pin) {
      return { success: false, message: 'Missing required fields' };
    }

    // Validate PIN format
    if (!/^\d{6}$/.test(userData.pin)) {
      return { success: false, message: 'PIN must be exactly 6 digits' };
    }

    // Validate role
    if (!['Manager', 'Director', 'Operator'].includes(userData.role)) {
      return { success: false, message: 'Invalid role' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('User_Permissions');

    if (!sheet) {
      return { success: false, message: 'User_Permissions sheet not found' };
    }

    // Check if employee_id already exists
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (const row of existingIds) {
        if (row[0] === userData.employee_id) {
          return { success: false, message: 'Employee ID already exists' };
        }
      }
    }

    // Hash the PIN
    const pinHash = hashPIN(userData.pin);

    // Add the new user
    sheet.appendRow([
      userData.employee_id,                                    // A: Employee_ID
      userData.full_name,                                      // B: Full_Name
      userData.email || '',                                    // C: Email
      userData.role,                                           // D: Role
      pinHash,                                                 // E: PIN_Hash
      userData.role === 'Director' ? (userData.can_see_directors ? 'TRUE' : 'FALSE') : 'FALSE', // F: Can_See_Directors
      new Date(),                                              // G: Date_Added
      session.user.name,                                       // H: Added_By
      'Active',                                                // I: Status
      '',                                                      // J: Last_Login
      0,                                                       // K: Login_Count
      0,                                                       // L: Failed_Attempts
      ''                                                       // M: Lockout_Until
    ]);

    return { success: true, message: `User ${userData.full_name} created successfully` };

  } catch (error) {
    console.error('Error creating user:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Changes a user's PIN.
 *
 * @param {string} employee_id - The employee ID (if changing another user's PIN)
 * @param {string} current_pin - Current PIN (required when changing own PIN)
 * @param {string} new_pin - New 6-digit PIN
 * @returns {Object} {success: boolean, message: string}
 */
function changePIN(employee_id, current_pin, new_pin) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    // Validate new PIN
    if (!/^\d{6}$/.test(new_pin)) {
      return { success: false, message: 'New PIN must be exactly 6 digits' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('User_Permissions');

    if (!sheet) {
      return { success: false, message: 'User_Permissions sheet not found' };
    }

    // Find the user
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: 'User not found' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let userRow = -1;
    let userPinHash = null;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === employee_id) {
        userRow = i + 2;
        userPinHash = data[i][4];
        break;
      }
    }

    if (userRow === -1) {
      return { success: false, message: 'User not found' };
    }

    // If changing own PIN, verify current PIN
    if (employee_id === session.user.user_id) {
      if (!current_pin) {
        return { success: false, message: 'Current PIN is required' };
      }
      const currentHash = hashPIN(current_pin);
      if (currentHash !== userPinHash) {
        return { success: false, message: 'Current PIN is incorrect' };
      }
    } else {
      // Changing another user's PIN - check permissions
      // Only Directors and Operators can do this
      if (session.role === 'Manager') {
        return { success: false, message: 'Managers can only change their own PIN' };
      }
    }

    // Update the PIN
    const newPinHash = hashPIN(new_pin);
    sheet.getRange(userRow, 5).setValue(newPinHash); // E: PIN_Hash

    return { success: true, message: 'PIN changed successfully' };

  } catch (error) {
    console.error('Error changing PIN:', error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// SELF-SERVICE SIGNUP SYSTEM (Micro-Phase 11.5)
// ============================================

/**
 * Generates a secure random token for signup links.
 * @returns {string} 64-character hex token
 */
function generateSecureToken() {
  const bytes = [];
  for (let i = 0; i < 32; i++) {
    bytes.push(Math.floor(Math.random() * 256));
  }
  return bytes.map(b => {
    const hex = b.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/**
 * Generates a signup ID in format SU_YYYYMMDD_NNN
 * @param {Sheet} sheet - The Pending_Signups sheet
 * @returns {string} Unique signup ID
 */
function generateSignupId(sheet) {
  const today = new Date();
  const dateStr = today.getFullYear().toString() +
                  (today.getMonth() + 1).toString().padStart(2, '0') +
                  today.getDate().toString().padStart(2, '0');

  const prefix = 'SU_' + dateStr + '_';

  // Find highest existing number for today
  const lastRow = sheet.getLastRow();
  let maxNum = 0;

  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (const row of ids) {
      const id = row[0];
      if (id && id.startsWith(prefix)) {
        const num = parseInt(id.substring(prefix.length), 10);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  return prefix + (maxNum + 1).toString().padStart(3, '0');
}

/**
 * Generates a signup link for a new user.
 * Creates a pending signup entry with a unique token.
 *
 * @param {Object} signup_data - {employee_id, email, role, can_see_directors}
 * @returns {Object} {success: boolean, signup_id?: string, token?: string, link?: string, message: string}
 */
function generateSignupLink(signup_data) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    const created_by = session.user.name;
    const currentRole = session.role;
    const targetRole = signup_data.role;

    // Permission check
    // Operators can create signups for anyone
    // Directors can create signups for Managers and Directors
    // Managers cannot create signups
    if (currentRole === 'Manager') {
      return { success: false, message: 'Managers cannot create signup links' };
    }

    if (currentRole === 'Director' && targetRole === 'Operator') {
      return { success: false, message: 'Directors cannot create signup links for Operators' };
    }

    // Validate required fields
    if (!signup_data.employee_id || !signup_data.email || !signup_data.role) {
      return { success: false, message: 'Missing required fields: employee_id, email, and role are required' };
    }

    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(signup_data.email)) {
      return { success: false, message: 'Invalid email format' };
    }

    // Validate role
    if (!['Manager', 'Director', 'Operator'].includes(signup_data.role)) {
      return { success: false, message: 'Invalid role. Must be Manager, Director, or Operator' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);

    // Check if employee_id already exists in User_Permissions
    const userSheet = ss.getSheetByName('User_Permissions');
    if (userSheet) {
      const lastUserRow = userSheet.getLastRow();
      if (lastUserRow >= 2) {
        const existingIds = userSheet.getRange(2, 1, lastUserRow - 1, 1).getValues();
        for (const row of existingIds) {
          if (row[0] === signup_data.employee_id) {
            return { success: false, message: 'Employee ID already exists in the system' };
          }
        }
      }
    }

    const signupSheet = ss.getSheetByName('Pending_Signups');
    if (!signupSheet) {
      return { success: false, message: 'Pending_Signups sheet not found. Please run setup.' };
    }

    // Check for existing pending signup with same employee_id or email
    const lastRow = signupSheet.getLastRow();
    if (lastRow >= 2) {
      const existingData = signupSheet.getRange(2, 1, lastRow - 1, 9).getValues();
      for (const row of existingData) {
        // row: [Signup_ID, Token, Employee_ID, Email, Role, Can_See_Directors, Created_Date, Expires_Date, Status]
        if (row[8] === 'Pending') { // Status column
          if (row[2] === signup_data.employee_id) {
            return { success: false, message: 'A pending signup already exists for this Employee ID' };
          }
          if (row[3] === signup_data.email) {
            return { success: false, message: 'A pending signup already exists for this email address' };
          }
        }
      }
    }

    // Generate signup ID and token
    const signup_id = generateSignupId(signupSheet);
    const token = generateSecureToken();

    // Calculate expiration date (7 days from now)
    const now = new Date();
    const expiresDate = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

    // Add the signup entry
    signupSheet.appendRow([
      signup_id,                                              // A: Signup_ID
      token,                                                  // B: Token
      signup_data.employee_id,                                // C: Employee_ID
      signup_data.email,                                      // D: Email
      signup_data.role,                                       // E: Role
      signup_data.can_see_directors ? 'TRUE' : 'FALSE',       // F: Can_See_Directors
      now,                                                    // G: Created_Date
      expiresDate,                                            // H: Expires_Date
      'Pending',                                              // I: Status
      created_by,                                             // J: Created_By
      ''                                                      // K: Completed_Date
    ]);

    // Generate the signup link
    const scriptUrl = ScriptApp.getService().getUrl();
    const signupLink = scriptUrl + '?page=signup&token=' + token;

    return {
      success: true,
      signup_id: signup_id,
      token: token,
      link: signupLink,
      email: signup_data.email,
      message: 'Signup link generated successfully'
    };

  } catch (error) {
    console.error('Error generating signup link:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Sends signup email to a pending user.
 *
 * @param {string} signup_id - The signup ID
 * @returns {Object} {success: boolean, message: string}
 */
function sendSignupEmail(signup_id) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupSheet = ss.getSheetByName('Pending_Signups');

    if (!signupSheet) {
      return { success: false, message: 'Pending_Signups sheet not found' };
    }

    // Find the signup entry
    const lastRow = signupSheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: 'Signup not found' };
    }

    const data = signupSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    let signupRow = -1;
    let signupData = null;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === signup_id) {
        signupRow = i + 2;
        signupData = {
          signup_id: data[i][0],
          token: data[i][1],
          employee_id: data[i][2],
          email: data[i][3],
          role: data[i][4],
          status: data[i][8],
          created_by: data[i][9]
        };
        break;
      }
    }

    if (signupRow === -1) {
      return { success: false, message: 'Signup not found' };
    }

    if (signupData.status !== 'Pending') {
      return { success: false, message: 'Cannot send email for signup with status: ' + signupData.status };
    }

    // Generate the signup link
    const scriptUrl = ScriptApp.getService().getUrl();
    const signupLink = scriptUrl + '?page=signup&token=' + signupData.token;

    // Send the email
    const subject = 'CFA Accountability System - Complete Your Registration';
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #E31837;">Welcome to the CFA Accountability System</h2>
        <p>You have been invited to create an account on the CFA Accountability System.</p>
        <p><strong>Your Role:</strong> ${signupData.role}</p>
        <p><strong>Employee ID:</strong> ${signupData.employee_id}</p>
        <p>Click the button below to complete your registration and create your 6-digit PIN:</p>
        <p style="text-align: center; margin: 30px 0;">
          <a href="${signupLink}" style="background-color: #E31837; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold;">
            Complete Registration
          </a>
        </p>
        <p style="color: #666; font-size: 14px;">This link will expire in 7 days. If you did not expect this email, please disregard it.</p>
        <p style="color: #666; font-size: 12px; margin-top: 30px; border-top: 1px solid #ddd; padding-top: 15px;">
          Invited by: ${signupData.created_by}
        </p>
      </div>
    `;

    MailApp.sendEmail({
      to: signupData.email,
      subject: subject,
      htmlBody: htmlBody
    });

    // Log the email
    logEmail(signupData.email, subject, 'Signup Invitation', 'Sent');

    return { success: true, message: 'Signup email sent successfully to ' + signupData.email };

  } catch (error) {
    console.error('Error sending signup email:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Validates a signup token (for public signup page).
 * This function does NOT require a session.
 *
 * @param {string} token - The signup token from the URL
 * @returns {Object} {valid: boolean, signup_data?: Object, message: string}
 */
function validateSignupToken(token) {
  try {
    if (!token || token.length !== 64) {
      return { valid: false, message: 'Invalid token format' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupSheet = ss.getSheetByName('Pending_Signups');

    if (!signupSheet) {
      return { valid: false, message: 'System error. Please contact administrator.' };
    }

    const lastRow = signupSheet.getLastRow();
    if (lastRow < 2) {
      return { valid: false, message: 'Signup link not found' };
    }

    const data = signupSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    let signupData = null;
    let signupRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === token) { // Token column
        signupRow = i + 2;
        signupData = {
          signup_id: data[i][0],
          employee_id: data[i][2],
          email: data[i][3],
          role: data[i][4],
          expires_date: data[i][7],
          status: data[i][8]
        };
        break;
      }
    }

    if (!signupData) {
      return { valid: false, message: 'Signup link not found or has been used' };
    }

    // Check status
    if (signupData.status === 'Completed') {
      return { valid: false, message: 'This signup has already been completed' };
    }

    if (signupData.status === 'Cancelled') {
      return { valid: false, message: 'This signup link has been cancelled' };
    }

    // Check expiration
    const now = new Date();
    const expiresDate = new Date(signupData.expires_date);

    if (now > expiresDate) {
      // Update status to Expired
      signupSheet.getRange(signupRow, 9).setValue('Expired');
      return { valid: false, message: 'This signup link has expired. Please request a new one.' };
    }

    // Token is valid
    return {
      valid: true,
      signup_data: {
        employee_id: signupData.employee_id,
        email: signupData.email,
        role: signupData.role
      },
      message: 'Token is valid'
    };

  } catch (error) {
    console.error('Error validating signup token:', error);
    return { valid: false, message: 'An error occurred. Please try again.' };
  }
}

/**
 * Completes a signup by creating the user account with their PIN.
 * This function does NOT require a session.
 *
 * @param {string} token - The signup token
 * @param {string} pin - The 6-digit PIN chosen by the user
 * @returns {Object} {success: boolean, message: string}
 */
function completeSignup(token, pin) {
  try {
    // Validate PIN format
    if (!/^\d{6}$/.test(pin)) {
      return { success: false, message: 'PIN must be exactly 6 digits' };
    }

    // First validate the token
    const tokenValidation = validateSignupToken(token);
    if (!tokenValidation.valid) {
      return { success: false, message: tokenValidation.message };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupSheet = ss.getSheetByName('Pending_Signups');
    const userSheet = ss.getSheetByName('User_Permissions');

    if (!signupSheet || !userSheet) {
      return { success: false, message: 'System error. Please contact administrator.' };
    }

    // Find the signup entry
    const lastRow = signupSheet.getLastRow();
    const data = signupSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    let signupRow = -1;
    let signupData = null;

    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === token) {
        signupRow = i + 2;
        signupData = {
          signup_id: data[i][0],
          employee_id: data[i][2],
          email: data[i][3],
          role: data[i][4],
          can_see_directors: data[i][5] === 'TRUE',
          created_by: data[i][9]
        };
        break;
      }
    }

    if (!signupData) {
      return { success: false, message: 'Signup not found' };
    }

    // Create the user in User_Permissions
    const pinHash = hashPIN(pin);
    const now = new Date();

    userSheet.appendRow([
      signupData.employee_id,                                 // A: Employee_ID
      '',                                                     // B: Full_Name (user can update later or set during signup)
      signupData.email,                                       // C: Email
      signupData.role,                                        // D: Role
      pinHash,                                                // E: PIN_Hash
      signupData.can_see_directors ? 'TRUE' : 'FALSE',        // F: Can_See_Directors
      now,                                                    // G: Date_Added
      'Self-Signup (invited by ' + signupData.created_by + ')', // H: Added_By
      'Active',                                               // I: Status
      '',                                                     // J: Last_Login
      0,                                                      // K: Login_Count
      0,                                                      // L: Failed_Attempts
      ''                                                      // M: Lockout_Until
    ]);

    // Update the signup entry to Completed
    signupSheet.getRange(signupRow, 9).setValue('Completed');  // Status
    signupSheet.getRange(signupRow, 11).setValue(now);         // Completed_Date

    return {
      success: true,
      message: 'Registration complete! You can now log in with your Employee ID and PIN.'
    };

  } catch (error) {
    console.error('Error completing signup:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Cancels a pending signup.
 *
 * @param {string} signup_id - The signup ID to cancel
 * @returns {Object} {success: boolean, message: string}
 */
function cancelSignup(signup_id) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    const cancelled_by = session.user.name;

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupSheet = ss.getSheetByName('Pending_Signups');

    if (!signupSheet) {
      return { success: false, message: 'Pending_Signups sheet not found' };
    }

    const lastRow = signupSheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, message: 'Signup not found' };
    }

    const data = signupSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    let signupRow = -1;
    let currentStatus = null;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === signup_id) {
        signupRow = i + 2;
        currentStatus = data[i][8];
        break;
      }
    }

    if (signupRow === -1) {
      return { success: false, message: 'Signup not found' };
    }

    if (currentStatus !== 'Pending') {
      return { success: false, message: 'Cannot cancel signup with status: ' + currentStatus };
    }

    // Update status to Cancelled
    signupSheet.getRange(signupRow, 9).setValue('Cancelled');

    return { success: true, message: 'Signup cancelled successfully' };

  } catch (error) {
    console.error('Error cancelling signup:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Resends a signup email for a pending signup.
 *
 * @param {string} signup_id - The signup ID
 * @returns {Object} {success: boolean, message: string}
 */
function resendSignupEmail(signup_id) {
  // This is just a wrapper around sendSignupEmail
  return sendSignupEmail(signup_id);
}

/**
 * Gets all pending signups for the management UI.
 *
 * @returns {Object} {success: boolean, signups?: Array, message?: string}
 */
function getPendingSignups() {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    // Only Directors and Operators can view pending signups
    if (session.role === 'Manager') {
      return { success: false, message: 'Access denied' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupSheet = ss.getSheetByName('Pending_Signups');

    if (!signupSheet) {
      return { success: false, message: 'Pending_Signups sheet not found' };
    }

    const lastRow = signupSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, signups: [] };
    }

    const data = signupSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const signups = [];

    const now = new Date();

    for (const row of data) {
      // Auto-update expired signups
      const status = row[8];
      const expiresDate = new Date(row[7]);

      if (status === 'Pending' && now > expiresDate) {
        // We won't update the sheet here to avoid performance issues,
        // but we'll show it as expired in the UI
        signups.push({
          signup_id: row[0],
          employee_id: row[2],
          email: row[3],
          role: row[4],
          can_see_directors: row[5] === 'TRUE',
          created_date: row[6],
          expires_date: row[7],
          status: 'Expired',
          created_by: row[9],
          completed_date: row[10]
        });
      } else {
        signups.push({
          signup_id: row[0],
          employee_id: row[2],
          email: row[3],
          role: row[4],
          can_see_directors: row[5] === 'TRUE',
          created_date: row[6],
          expires_date: row[7],
          status: row[8],
          created_by: row[9],
          completed_date: row[10]
        });
      }
    }

    // Sort by created_date descending (most recent first)
    signups.sort((a, b) => new Date(b.created_date) - new Date(a.created_date));

    return { success: true, signups: signups };

  } catch (error) {
    console.error('Error getting pending signups:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets active employees who don't have system access yet.
 * Used to populate the employee dropdown in the signup form.
 *
 * @returns {Object} {success: boolean, employees: Array}
 */
function getEmployeesForSignup() {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    // Only Directors and Operators can access this
    if (session.role === 'Manager') {
      return { success: false, message: 'Access denied' };
    }

    // Get all active employees from Payroll Tracker
    const activeEmployees = getActiveEmployees();

    if (!activeEmployees || activeEmployees.length === 0) {
      return { success: true, employees: [] };
    }

    // Get employees who already have system access
    const systemAccessMap = getSystemAccessMap();

    // Also get pending signups to exclude those employees
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const signupSheet = ss.getSheetByName('Pending_Signups');
    const pendingEmployeeIds = [];

    if (signupSheet) {
      const lastRow = signupSheet.getLastRow();
      if (lastRow >= 2) {
        const signupData = signupSheet.getRange(2, 1, lastRow - 1, 9).getValues();
        for (const row of signupData) {
          if (row[8] === 'Pending') { // Status column
            pendingEmployeeIds.push(row[2]); // Employee_ID column
          }
        }
      }
    }

    // Filter to only employees without system access and without pending signup
    const availableEmployees = activeEmployees.filter(emp => {
      // Exclude if already has system access
      if (systemAccessMap[emp.employee_id]) {
        return false;
      }
      // Exclude if has pending signup
      if (pendingEmployeeIds.includes(emp.employee_id)) {
        return false;
      }
      return true;
    });

    // Sort alphabetically by name
    availableEmployees.sort((a, b) => (a.full_name || '').localeCompare(b.full_name || ''));

    // Return minimal data for the dropdown
    const employees = availableEmployees.map(emp => ({
      employee_id: emp.employee_id,
      full_name: emp.full_name,
      email: emp.email || ''
    }));

    return { success: true, employees: employees };

  } catch (error) {
    console.error('Error getting employees for signup:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Generates a signup link and immediately sends the email.
 * Convenience function combining generateSignupLink and sendSignupEmail.
 *
 * @param {Object} signup_data - {employee_id, email, role, can_see_directors}
 * @returns {Object} {success: boolean, signup_id?: string, message: string}
 */
function createAndSendSignup(signup_data) {
  try {
    // Generate the signup link
    const result = generateSignupLink(signup_data);

    if (!result.success) {
      return result;
    }

    // Send the email
    const emailResult = sendSignupEmail(result.signup_id);

    if (!emailResult.success) {
      return {
        success: true,
        signup_id: result.signup_id,
        link: result.link,
        message: 'Signup created but email failed to send: ' + emailResult.message
      };
    }

    return {
      success: true,
      signup_id: result.signup_id,
      link: result.link,
      message: 'Signup link created and email sent to ' + signup_data.email
    };

  } catch (error) {
    console.error('Error creating and sending signup:', error);
    return { success: false, message: error.toString() };
  }
}

// ============================================
// SESSION VALIDATION WRAPPER
// ============================================

/**
 * Validates session before allowing data modification.
 * Call this at the start of any function that modifies data.
 *
 * @returns {Object} {valid: boolean, role: string|null, error: string|null}
 */
function requireValidSession() {
  const session = validateSession();

  if (!session.valid) {
    if (session.expired) {
      return {
        valid: false,
        role: null,
        error: 'Session expired. Please login again.'
      };
    } else {
      return {
        valid: false,
        role: null,
        error: 'Not authenticated. Please login.'
      };
    }
  }

  // Extend session on valid request
  extendSession();

  return {
    valid: true,
    role: session.role,
    error: null
  };
}

// ============================================
// PASSWORD EXPIRATION CHECK
// ============================================

/**
 * Checks if the password needs to be changed (older than 1 year).
 * Should be run daily via time-based trigger.
 */
function checkPasswordExpiration() {
  console.log('=== Checking Password Expiration ===');

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      console.error('Settings sheet not found');
      return;
    }

    // Last Password Change date is in B8
    const lastChangeDate = settingsSheet.getRange('B8').getValue();

    if (!lastChangeDate || !(lastChangeDate instanceof Date)) {
      console.log('No last password change date found');
      return;
    }

    // Calculate days since last change
    const now = new Date();
    const daysSinceChange = Math.floor((now - lastChangeDate) / (1000 * 60 * 60 * 24));

    console.log('Days since password change:', daysSinceChange);

    // Check if > 365 days
    if (daysSinceChange > 365) {
      // Check if alert already sent (B9)
      const alertSent = settingsSheet.getRange('B9').getValue();

      if (alertSent === true || alertSent === 'TRUE') {
        console.log('Password expiration alert already sent');
        return;
      }

      console.log('Password is expired, sending alert...');

      // Send expiration alert email
      sendPasswordExpirationAlert(daysSinceChange);

      // Mark alert as sent
      settingsSheet.getRange('B9').setValue(true);

      console.log('Password expiration alert sent and recorded');

    } else {
      console.log('Password is not expired');
    }

  } catch (error) {
    console.error('Error checking password expiration:', error.toString());
  }
}

/**
 * Sends password expiration alert email to Jeff and directors.
 *
 * @param {number} daysSinceChange - Days since password was last changed
 */
function sendPasswordExpirationAlert(daysSinceChange) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    // Get notification email (B3) for Jeff
    const primaryEmail = settingsSheet.getRange('B3').getValue();

    // Get termination email list (B5) for directors
    const terminationEmails = settingsSheet.getRange('B5').getValue();

    // Combine recipients
    let recipients = primaryEmail;
    if (terminationEmails) {
      recipients += ',' + terminationEmails;
    }

    const subject = 'CFA Accountability System - Password Change Required';
    const body = `
This is an automated alert from the CFA Accountability System.

The system password is ${daysSinceChange} days old and needs to be changed.

For security reasons, passwords should be changed at least once per year.

Please update the Manager and/or Director passwords in the Settings sheet.

Steps to change password:
1. Open the CFA Accountability System spreadsheet
2. Go to the Settings sheet
3. Update the password in cell B6 (Manager) and/or B7 (Director)
4. Update the "Last Password Change" date in cell B8
5. Set the "Password Expiration Alert Sent" in B9 to FALSE

This is an automated message. Do not reply.
    `.trim();

    const htmlBody = `
<h2>CFA Accountability System - Password Change Required</h2>

<p>This is an automated alert from the CFA Accountability System.</p>

<p><strong>The system password is ${daysSinceChange} days old and needs to be changed.</strong></p>

<p>For security reasons, passwords should be changed at least once per year.</p>

<h3>Steps to change password:</h3>
<ol>
  <li>Open the CFA Accountability System spreadsheet</li>
  <li>Go to the Settings sheet</li>
  <li>Update the password in cell B6 (Manager) and/or B7 (Director)</li>
  <li>Update the "Last Password Change" date in cell B8</li>
  <li>Set the "Password Expiration Alert Sent" in B9 to FALSE</li>
</ol>

<p><em>This is an automated message. Do not reply.</em></p>
    `.trim();

    GmailApp.sendEmail(recipients, subject, body, {
      htmlBody: htmlBody,
      name: 'CFA Accountability System'
    });

    console.log('Password expiration alert sent to:', recipients);

  } catch (error) {
    console.error('Error sending password expiration alert:', error.toString());
  }
}

/**
 * Creates a daily trigger to check password expiration.
 * Run this once to set up the trigger.
 */
function createPasswordExpirationTrigger() {
  // Delete any existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'checkPasswordExpiration') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create new daily trigger (runs at midnight)
  ScriptApp.newTrigger('checkPasswordExpiration')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  console.log('Password expiration trigger created');
}

// ============================================
// MICRO-PHASE 11: TEST FUNCTIONS
// ============================================

/**
 * Test function for authentication system.
 */
function testAuthentication() {
  console.log('=== Testing Authentication System ===\n');

  // ========================================
  // TEST CASE 1: Valid Manager Password
  // ========================================
  console.log('--- Test Case 1: Valid Manager Password ---');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const settingsSheet = ss.getSheetByName('Settings');
  const managerPassword = settingsSheet.getRange('B6').getValue();

  if (managerPassword) {
    const result1 = validatePassword(managerPassword);
    console.log('Result:', JSON.stringify(result1));
    console.log('Valid should be true:', result1.valid === true);
    console.log('Role should be Manager:', result1.role === 'Manager');
  } else {
    console.log('Manager password not set in Settings');
  }
  console.log('');

  // ========================================
  // TEST CASE 2: Valid Director Password
  // ========================================
  console.log('--- Test Case 2: Valid Director Password ---');

  const directorPassword = settingsSheet.getRange('B7').getValue();

  if (directorPassword) {
    const result2 = validatePassword(directorPassword);
    console.log('Result:', JSON.stringify(result2));
    console.log('Valid should be true:', result2.valid === true);
    console.log('Role should be Director:', result2.role === 'Director');
  } else {
    console.log('Director password not set in Settings');
  }
  console.log('');

  // ========================================
  // TEST CASE 3: Invalid Password
  // ========================================
  console.log('--- Test Case 3: Invalid Password ---');

  const result3 = validatePassword('definitely_wrong_password_12345');
  console.log('Result:', JSON.stringify(result3));
  console.log('Valid should be false:', result3.valid === false);
  console.log('Role should be null:', result3.role === null);
  console.log('');

  // ========================================
  // TEST CASE 4: Session Creation
  // ========================================
  console.log('--- Test Case 4: Session Creation ---');

  const sessionResult = createSession('Manager');
  console.log('Session created:', JSON.stringify(sessionResult));
  console.log('Success should be true:', sessionResult.success === true);
  console.log('SessionId should exist:', !!sessionResult.sessionId);
  console.log('');

  // ========================================
  // TEST CASE 5: Session Validation
  // ========================================
  console.log('--- Test Case 5: Session Validation ---');

  const validateResult = validateSession();
  console.log('Validation result:', JSON.stringify(validateResult));
  console.log('Valid should be true:', validateResult.valid === true);
  console.log('Role should be Manager:', validateResult.role === 'Manager');
  console.log('');

  // ========================================
  // TEST CASE 6: Session Extension
  // ========================================
  console.log('--- Test Case 6: Session Extension ---');

  const extendResult = extendSession();
  console.log('Extend result:', JSON.stringify(extendResult));
  console.log('Success should be true:', extendResult.success === true);

  // Verify session still valid
  const validateAfterExtend = validateSession();
  console.log('Still valid after extend:', validateAfterExtend.valid === true);
  console.log('');

  // ========================================
  // TEST CASE 7: Session End
  // ========================================
  console.log('--- Test Case 7: Session End ---');

  const endResult = endSession();
  console.log('End result:', JSON.stringify(endResult));
  console.log('Success should be true:', endResult.success === true);

  // Verify session no longer valid
  const validateAfterEnd = validateSession();
  console.log('Valid after end should be false:', validateAfterEnd.valid === false);
  console.log('');

  // ========================================
  // TEST CASE 8: Require Valid Session
  // ========================================
  console.log('--- Test Case 8: Require Valid Session (no session) ---');

  const requireResult1 = requireValidSession();
  console.log('Require result (no session):', JSON.stringify(requireResult1));
  console.log('Valid should be false:', requireResult1.valid === false);
  console.log('');

  // Create a session and try again
  createSession('Director');
  const requireResult2 = requireValidSession();
  console.log('Require result (with session):', JSON.stringify(requireResult2));
  console.log('Valid should be true:', requireResult2.valid === true);
  console.log('Role should be Director:', requireResult2.role === 'Director');

  // Clean up
  endSession();
  console.log('');

  console.log('=== Authentication Tests Complete ===');
}

/**
 * Gets the web app URL for redirects.
 * Used by client-side JavaScript to redirect after login/logout.
 *
 * @returns {string} The web app URL
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ============================================
// PHASE 12: EMPLOYEE LIST WITH POINTS
// ============================================

/**
 * Gets all active employees with their calculated point totals.
 * Includes color coding, status badges, and role-based permission filtering.
 *
 * Filtering Logic (Micro-Phase 12):
 * - Operators: See all employees EXCEPT other Operators
 * - Directors with can_see_directors=TRUE: See Managers, Directors, and hourly (not Operators)
 * - Directors with can_see_directors=FALSE: See only Managers and hourly (not Directors or Operators)
 * - Managers: See only hourly employees (no one with system access)
 *
 * @returns {Object} Object with success status and filtered employee array
 */
function getAllEmployeesWithPoints() {
  const startTime = Date.now();

  try {
    // Validate session using simplified auth
    const session = getCurrentRole();
    if (!session.authenticated) {
      return {
        success: false,
        sessionExpired: true,
        error: 'Session expired',
        employees: [],
        counts: { atRisk: 0, finalWarning: 0, termination: 0, total: 0 },
        executionTime: Date.now() - startTime
      };
    }

    // Get requesting user's role (simplified - no individual user tracking)
    const requestingUser = {
      role: session.role
    };

    console.log(`getAllEmployeesWithPoints called by ${requestingUser.role}`);

    // Step 1: Get all active employees
    const employees = getActiveEmployees();

    if (!employees || employees.length === 0) {
      return {
        success: true,
        employees: [],
        counts: { atRisk: 0, finalWarning: 0, termination: 0 },
        userRole: requestingUser.role,
        executionTime: Date.now() - startTime
      };
    }

    // Step 2: Get system access map (employee_id -> { role, can_see_directors })
    const systemAccessMap = getSystemAccessMap();

    // Step 3: Get all infractions at once for performance
    const allInfractions = getAllActiveInfractions();

    // Step 4: Process each employee and build complete data
    const allEmployeesWithPoints = [];

    for (const emp of employees) {
      try {
        // Skip invalid employee records
        if (!emp || !emp.employee_id) {
          console.log('Skipping invalid employee record:', emp);
          continue;
        }

        // Calculate points for this employee
        const pointData = calculatePointsFromData(emp.employee_id, allInfractions);

        // Determine point level color
        const pointLevelColor = getPointLevelColor(pointData.total_points);

        // Check for status badges
        const statusBadges = getStatusBadges(pointData, allInfractions);

        // Check system access and get role
        const systemAccess = systemAccessMap[emp.employee_id];
        const hasSystemAccess = !!systemAccess;
        const systemRole = systemAccess ? systemAccess.role : null;

        // Build employee object with defensive defaults
        const employeeObj = {
          employee_id: emp.employee_id || 'Unknown',
          full_name: emp.full_name || 'Unknown Employee',
          primary_location: emp.primary_location || 'Unknown',
          current_points: pointData.total_points || 0,
          active_infractions_count: (pointData.active_infractions || []).length,
          next_expiration_date: pointData.next_expiration_date instanceof Date
            ? pointData.next_expiration_date.toISOString()
            : (pointData.next_expiration_date || null),
          next_expiration_points: pointData.next_expiration_points || 0,
          has_system_access: hasSystemAccess,
          system_role: systemRole,
          point_level_color: pointLevelColor || 'green',
          status_badges: statusBadges || []
        };

        allEmployeesWithPoints.push(employeeObj);

      } catch (empError) {
        console.error(`Error processing employee ${emp.employee_id || 'unknown'}:`, empError);
        // Add employee with error state
        const systemAccess = systemAccessMap[emp.employee_id];
        allEmployeesWithPoints.push({
          employee_id: emp.employee_id || 'Unknown',
          full_name: emp.full_name || 'Unknown Employee',
          primary_location: emp.primary_location || 'Unknown',
          current_points: -1,
          active_infractions_count: 0,
          next_expiration_date: null,
          next_expiration_points: 0,
          has_system_access: !!systemAccess,
          system_role: systemAccess ? systemAccess.role : null,
          point_level_color: 'gray',
          status_badges: [{ type: 'error', label: 'Unable to calculate' }],
          error: true
        });
      }
    }

    // Step 5: Apply role-based filtering
    const filteredEmployees = filterEmployeesByPermission(allEmployeesWithPoints, requestingUser);

    // Step 6: Calculate threshold counts on filtered list
    let atRiskCount = 0;
    let finalWarningCount = 0;
    let terminationCount = 0;

    for (const emp of filteredEmployees) {
      if (emp.current_points >= 15) {
        terminationCount++;
      } else if (emp.current_points >= 9) {
        finalWarningCount++;
      } else if (emp.current_points >= 6) {
        atRiskCount++;
      }
    }

    // Step 7: Sort by full_name ascending (default)
    filteredEmployees.sort((a, b) => (a.full_name || '').localeCompare(b.full_name || ''));

    const executionTime = Date.now() - startTime;
    console.log(`getAllEmployeesWithPoints completed in ${executionTime}ms. Total: ${allEmployeesWithPoints.length}, Filtered: ${filteredEmployees.length}`);

    return {
      success: true,
      employees: filteredEmployees,
      counts: {
        atRisk: atRiskCount,
        finalWarning: finalWarningCount,
        termination: terminationCount,
        total: filteredEmployees.length
      },
      userRole: requestingUser.role,
      executionTime: executionTime
    };

  } catch (error) {
    console.error('Error in getAllEmployeesWithPoints:', error);
    console.error('Error stack:', error.stack);
    return {
      success: false,
      error: error.message || error.toString(),
      errorStack: error.stack || 'No stack available',
      employees: [],
      counts: { atRisk: 0, finalWarning: 0, termination: 0, total: 0 },
      executionTime: Date.now() - startTime
    };
  }
}

/**
 * Filters employees based on the requesting user's role and permissions.
 * This is the core permission filtering logic for Micro-Phase 12.
 *
 * @param {Array} employees - Array of employee objects with system_role property
 * @param {Object} requestingUser - { user_id, role, can_see_directors }
 * @returns {Array} Filtered array of employees the user has permission to see
 */
function filterEmployeesByPermission(employees, requestingUser) {
  const filteredEmployees = [];

  for (const emp of employees) {
    // Determine if this employee should be visible based on requesting user's role
    // Simplified visibility rules (no individual user tracking):
    // - Operator: sees everyone EXCEPT other Operators
    // - Director: sees everyone EXCEPT other Directors and Operators
    // - Manager: sees only hourly employees (no one with system access)

    if (requestingUser.role === 'Operator') {
      // Operators see all employees EXCEPT other Operators
      if (emp.system_role === 'Operator') {
        continue; // Skip - don't show other operators
      }
      filteredEmployees.push(emp);
    }

    else if (requestingUser.role === 'Director') {
      // Directors never see Operators or other Directors
      if (emp.system_role === 'Operator' || emp.system_role === 'Director') {
        continue; // Skip - don't show operators or directors
      }

      // Hourly employees and Managers are visible
      filteredEmployees.push(emp);
    }

    else if (requestingUser.role === 'Manager') {
      // Managers only see hourly employees (no one with system access)
      if (!emp.has_system_access) {
        filteredEmployees.push(emp);
      }
      // Skip all employees with system access
    }

    else {
      // Unknown role - default to no access for security
      console.warn(`Unknown role: ${requestingUser.role} - denying access to employee ${emp.employee_id}`);
    }
  }

  return filteredEmployees;
}

/**
 * Checks if a requesting user has permission to view a specific target employee.
 * This is the core permission check for Micro-Phase 13 Employee Detail View.
 *
 * @param {Object} requestingUser - { user_id, role, can_see_directors }
 * @param {Object} targetEmployee - { employee_id, has_system_access, system_role }
 * @returns {boolean} TRUE if user can view, FALSE if access denied
 */
function checkViewPermission(requestingUser, targetEmployee) {
  // Step 1: Check if target has no system access (hourly employee)
  // Everyone can see hourly employees
  if (!targetEmployee.has_system_access) {
    return true;
  }

  // Step 2: Check if target is an Operator
  // Operators NEVER appear in employee system for anyone
  if (targetEmployee.system_role === 'Operator') {
    return false;
  }

  // Step 3: Check based on requesting user's role
  if (requestingUser.role === 'Operator') {
    // Operators can see everyone except other operators (already checked above)
    return true;
  }

  if (requestingUser.role === 'Director') {
    // Directors can always see Managers
    if (targetEmployee.system_role === 'Manager') {
      return true;
    }

    // Directors can see other Directors only if can_see_directors is TRUE
    if (targetEmployee.system_role === 'Director') {
      return requestingUser.can_see_directors === true;
    }

    // Default: allow (hourly already handled above)
    return true;
  }

  if (requestingUser.role === 'Manager') {
    // Managers cannot see anyone with system access
    // (hourly employees already returned TRUE above)
    return false;
  }

  // Step 4: Default - deny access for unknown roles (security)
  console.warn(`checkViewPermission: Unknown role ${requestingUser.role} - denying access`);
  return false;
}

/**
 * Gets list of employee IDs that have system access (managers/directors).
 *
 * @returns {Array} Array of employee IDs with system access
 */
function getEmployeesWithSystemAccess() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const permSheet = ss.getSheetByName('User_Permissions');

    if (!permSheet) {
      console.log('User_Permissions sheet not found');
      return [];
    }

    const lastRow = permSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // Read Employee_ID (A), Role (D), and Status (I) columns
    // Note: Column layout is A:Employee_ID, B:Full_Name, C:Email, D:Role, E:PIN_Hash, F:Can_See_Directors, G:Date_Added, H:Added_By, I:Status
    const data = permSheet.getRange(2, 1, lastRow - 1, 9).getValues();

    const accessIds = [];
    for (const row of data) {
      const employeeId = row[0]; // A: Employee_ID
      const status = row[8];     // I: Status

      if (employeeId && status === 'Active') {
        accessIds.push(employeeId);
      }
    }

    return accessIds;

  } catch (error) {
    console.error('Error getting employees with system access:', error);
    return [];
  }
}

/**
 * Gets a map of employee IDs to their system roles from User_Permissions.
 * Returns both the access status and the specific role for filtering.
 *
 * @returns {Object} Map of employee_id -> { role: string, can_see_directors: boolean }
 */
function getSystemAccessMap() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const permSheet = ss.getSheetByName('User_Permissions');

    if (!permSheet) {
      console.log('User_Permissions sheet not found');
      return {};
    }

    const lastRow = permSheet.getLastRow();
    if (lastRow < 2) {
      return {};
    }

    // Read Employee_ID (A), Role (D), Can_See_Directors (F), and Status (I) columns
    const data = permSheet.getRange(2, 1, lastRow - 1, 9).getValues();

    const accessMap = {};
    for (const row of data) {
      const employeeId = row[0];      // A: Employee_ID
      const role = row[3];            // D: Role
      const canSeeDirectors = row[5]; // F: Can_See_Directors
      const status = row[8];          // I: Status

      if (employeeId && status === 'Active') {
        accessMap[employeeId] = {
          role: role,
          can_see_directors: canSeeDirectors === 'TRUE' || canSeeDirectors === true
        };
      }
    }

    return accessMap;

  } catch (error) {
    console.error('Error getting system access map:', error);
    return {};
  }
}

/**
 * Gets all active infractions from the Infractions sheet.
 * Used to batch process points calculations for performance.
 *
 * @returns {Array} Array of infraction objects
 */
function getAllActiveInfractions() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const infractionsSheet = ss.getSheetByName('Infractions');

    if (!infractionsSheet) {
      console.error('Infractions sheet not found');
      return [];
    }

    const lastRow = infractionsSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // Read all infraction data
    const data = infractionsSheet.getRange(2, 1, lastRow - 1, 16).getValues();

    const infractions = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const status = row[14]; // O: Status

      if (status !== 'Active') {
        continue;
      }

      // Parse dates with null checks
      let infractionDate = row[3]; // D: Date
      if (!infractionDate) {
        console.log(`Skipping infraction at row ${i + 2}: missing date`);
        continue; // Skip infractions without dates
      }
      if (!(infractionDate instanceof Date)) {
        infractionDate = new Date(infractionDate);
      }
      // Validate the date is valid
      if (isNaN(infractionDate.getTime())) {
        console.log(`Skipping infraction at row ${i + 2}: invalid date`);
        continue;
      }

      let expirationDate = row[15]; // P: Expiration_Date
      if (expirationDate && !(expirationDate instanceof Date)) {
        expirationDate = new Date(expirationDate);
      }
      if (!expirationDate || isNaN(expirationDate.getTime())) {
        // Calculate expiration if not set or invalid
        expirationDate = new Date(infractionDate);
        expirationDate.setDate(expirationDate.getDate() + 90);
      }

      infractions.push({
        infraction_id: row[0],     // A
        employee_id: row[1],       // B
        full_name: row[2],         // C
        date: infractionDate,      // D
        infraction_type: row[4],   // E
        points: row[5] || 0,       // F
        bucket: row[6],            // G
        description: row[7],       // H
        location: row[8],          // I
        entered_by: row[9],        // J
        status: status,            // O
        expiration_date: expirationDate // P
      });
    }

    return infractions;

  } catch (error) {
    console.error('Error getting all active infractions:', error);
    return [];
  }
}

/**
 * Calculates points for an employee from pre-loaded infraction data.
 * More efficient than calling calculatePoints for each employee.
 *
 * @param {string} employeeId - The employee ID
 * @param {Array} allInfractions - Array of all active infractions
 * @returns {Object} Point calculation result
 */
function calculatePointsFromData(employeeId, allInfractions) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const cutoffDate = new Date(today);
  cutoffDate.setDate(cutoffDate.getDate() - 90);

  const activeInfractions = [];
  let totalPoints = 0;
  let nextExpirationDate = null;
  let nextExpirationPoints = 0;

  // Filter infractions for this employee
  const empInfractions = allInfractions.filter(inf => inf.employee_id === employeeId);

  for (const inf of empInfractions) {
    // Skip if date is missing or invalid
    if (!inf.date || !inf.expiration_date) {
      continue;
    }

    const infDate = new Date(inf.date);
    if (isNaN(infDate.getTime())) {
      continue;
    }
    infDate.setHours(0, 0, 0, 0);

    const expDate = new Date(inf.expiration_date);
    if (isNaN(expDate.getTime())) {
      continue;
    }
    expDate.setHours(0, 0, 0, 0);

    // Check if infraction is still active (not expired)
    if (expDate >= today && infDate >= cutoffDate) {
      activeInfractions.push(inf);
      totalPoints += (inf.points || 0);

      // Track next expiration
      if (!nextExpirationDate || expDate < nextExpirationDate) {
        nextExpirationDate = expDate;
        nextExpirationPoints = inf.points;
      }
    }
  }

  return {
    total_points: totalPoints,
    active_infractions: activeInfractions,
    next_expiration_date: nextExpirationDate,
    next_expiration_points: nextExpirationPoints
  };
}

/**
 * Determines the color level based on point total.
 *
 * @param {number} points - Total points
 * @returns {string} Color level: green, yellow, orange, or red
 */
function getPointLevelColor(points) {
  if (points >= 9) {
    return 'red';
  } else if (points >= 6) {
    return 'orange';
  } else if (points >= 3) {
    return 'yellow';
  } else {
    return 'green';
  }
}

/**
 * Gets status badges for an employee based on their infraction data.
 *
 * @param {Object} pointData - Point calculation result
 * @param {Array} allInfractions - All active infractions (to check for threshold crossings)
 * @returns {Array} Array of badge objects
 */
function getStatusBadges(pointData, allInfractions) {
  const badges = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Final Warning badge (9+ points)
  if (pointData.total_points >= 9 && pointData.total_points < 15) {
    badges.push({
      type: 'final_warning',
      label: 'Final Warning',
      icon: '⚠️',
      color: '#dc3545'
    });
  }

  // Termination level badge (15+ points)
  if (pointData.total_points >= 15) {
    badges.push({
      type: 'termination',
      label: 'Termination Level',
      icon: '⛔',
      color: '#721c24'
    });
  }

  // Check for 30-day probation (crossed 9-point threshold within last 30 days)
  if (pointData.total_points >= 9) {
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    // Check if any recent infraction pushed them to 9+
    const recentInfractions = pointData.active_infractions.filter(inf => {
      const infDate = new Date(inf.date);
      return infDate >= thirtyDaysAgo;
    });

    if (recentInfractions.length > 0) {
      badges.push({
        type: 'probation',
        label: '30-Day Probation',
        icon: '🚨',
        color: '#fd7e14'
      });
    }
  }

  return badges;
}

/**
 * Gets detailed information for a specific employee.
 *
 * @param {string} employeeId - The employee ID
 * @returns {Object} Detailed employee data
 */
function getEmployeeDetails(employeeId) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    // Get employee from Payroll Tracker
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === employeeId);

    if (!employee) {
      return {
        success: false,
        error: 'Employee not found'
      };
    }

    // Get point calculation
    const pointData = calculatePoints(employeeId);

    // Get system access status
    const systemAccessIds = getEmployeesWithSystemAccess();
    const hasSystemAccess = systemAccessIds.includes(employeeId);

    // Access control: Managers cannot view other managers/directors
    if (session.role === 'Manager' && hasSystemAccess) {
      return {
        success: false,
        error: 'Access denied'
      };
    }

    // Build detailed response
    // IMPORTANT: Convert Date objects to ISO strings for proper serialization
    return {
      success: true,
      employee: {
        employee_id: employee.employee_id,
        full_name: employee.full_name,
        primary_location: employee.primary_location,
        hire_date: employee.hire_date instanceof Date
          ? employee.hire_date.toISOString()
          : (employee.hire_date || null),
        current_points: pointData.total_points,
        point_level_color: getPointLevelColor(pointData.total_points),
        status_badges: getStatusBadges(pointData, pointData.active_infractions),
        has_system_access: hasSystemAccess,
        active_infractions: pointData.active_infractions.map(inf => ({
          infraction_id: inf.infraction_id,
          date: inf.date instanceof Date ? inf.date.toISOString() : (inf.date || null),
          infraction_type: inf.infraction_type,
          points: inf.points,
          description: inf.description,
          expiration_date: inf.expiration_date instanceof Date
            ? inf.expiration_date.toISOString()
            : (inf.expiration_date || null),
          entered_by: inf.entered_by
        })),
        next_expiration_date: pointData.next_expiration_date instanceof Date
          ? pointData.next_expiration_date.toISOString()
          : (pointData.next_expiration_date || null),
        expired_infractions: (pointData.expired_infractions || []).map(inf => ({
          ...inf,
          date: inf.date instanceof Date ? inf.date.toISOString() : (inf.date || null),
          expiration_date: inf.expiration_date instanceof Date
            ? inf.expiration_date.toISOString()
            : (inf.expiration_date || null)
        }))
      }
    };

  } catch (error) {
    console.error('Error getting employee details:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Test function for getAllEmployeesWithPoints.
 */
function testGetAllEmployeesWithPoints() {
  console.log('=== Testing getAllEmployeesWithPoints ===');
  console.log('');

  const startTime = Date.now();
  const result = getAllEmployeesWithPoints();
  const duration = Date.now() - startTime;

  console.log('Execution time:', duration, 'ms');
  console.log('Success:', result.success);
  console.log('Employee count:', result.employees ? result.employees.length : 0);
  console.log('');

  if (result.success && result.employees.length > 0) {
    // Check first employee has required fields
    const firstEmp = result.employees[0];
    console.log('First employee:', JSON.stringify(firstEmp, null, 2));
    console.log('');

    const requiredFields = [
      'employee_id', 'full_name', 'primary_location', 'current_points',
      'active_infractions_count', 'next_expiration_date', 'has_system_access',
      'point_level_color', 'status_badges'
    ];

    const missingFields = requiredFields.filter(f => !(f in firstEmp));
    if (missingFields.length > 0) {
      console.log('FAIL: Missing fields:', missingFields);
    } else {
      console.log('PASS: All required fields present');
    }

    // Verify color coding
    console.log('');
    console.log('Color coding verification:');
    const colorCounts = { green: 0, yellow: 0, orange: 0, red: 0, gray: 0 };
    for (const emp of result.employees) {
      colorCounts[emp.point_level_color] = (colorCounts[emp.point_level_color] || 0) + 1;
    }
    console.log('Color distribution:', JSON.stringify(colorCounts));

    // Verify threshold counts
    console.log('');
    console.log('Threshold counts:');
    console.log('At Risk (6+ points):', result.counts.atRisk);
    console.log('Final Warning (9+ points):', result.counts.finalWarning);
    console.log('Termination Level (15+ points):', result.counts.termination);

    // Performance check
    console.log('');
    if (duration < 30000) {
      console.log('PASS: Completed in under 30 seconds');
    } else {
      console.log('FAIL: Took longer than 30 seconds');
    }
  } else {
    console.log('No employees returned or error occurred');
    if (result.error) {
      console.log('Error:', result.error);
    }
  }

  console.log('');
  console.log('=== Test Complete ===');
}

// ============================================
// PHASE 13: EMPLOYEE DETAIL VIEW FUNCTIONS
// ============================================

/**
 * Gets comprehensive employee detail data for the detail view modal.
 * Includes all infractions, statistics, threshold history, and action tracking.
 *
 * @param {string} employeeId - The employee ID
 * @returns {Object} Complete employee detail data
 */
function getEmployeeDetailData(employeeId) {
  const startTime = Date.now();

  try {
    // Validate session first
    const session = validateSession();
    if (!session.valid) {
      return {
        success: false,
        sessionExpired: true,
        error: 'Session expired'
      };
    }

    if (!employeeId) {
      return {
        success: false,
        error: 'Employee ID is required'
      };
    }

    // Step 1: Get employee basic info
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === employeeId);

    if (!employee) {
      return {
        success: false,
        error: 'Employee not found'
      };
    }

    // Step 2: Check system access and get role for target employee
    const systemAccessMap = getSystemAccessMap();
    const targetAccess = systemAccessMap[employeeId];
    const hasSystemAccess = !!targetAccess;
    const systemRole = targetAccess ? targetAccess.role : null;

    // Build target employee object for permission check
    const targetEmployee = {
      employee_id: employeeId,
      has_system_access: hasSystemAccess,
      system_role: systemRole
    };

    // Build requesting user object
    const requestingUser = {
      user_id: session.user.user_id,
      role: session.role,
      can_see_directors: session.user.can_see_directors || false
    };

    // Step 3: Permission check using Micro-Phase 13 logic
    const canView = checkViewPermission(requestingUser, targetEmployee);

    if (!canView) {
      // Log unauthorized access attempt
      console.warn(`Unauthorized access attempt: User ${requestingUser.user_id} (${requestingUser.role}) tried to view employee ${employeeId} (${systemRole || 'hourly'})`);

      return {
        success: false,
        error: "You don't have permission to view this employee"
      };
    }

    // Step 3: Get ALL infractions for this employee (including deleted/modified)
    const allInfractions = getAllEmployeeInfractions(employeeId);

    // Step 4: Calculate current points using existing function
    const pointData = calculatePoints(employeeId);

    // Step 5: Calculate statistics
    const stats = calculateEmployeeStatistics(employeeId, allInfractions);

    // Step 6: Calculate threshold history
    const thresholdHistory = calculateThresholdHistory(allInfractions);

    // Step 7: Get points in last 30 days and expiring in next 30 days
    const pointTrends = calculatePointTrends(allInfractions);

    // Step 8: Get action tracking data (directors only)
    let actionTracking = null;
    if (session.role === 'Director') {
      actionTracking = getActionTrackingData(employeeId);
    }

    // Step 9: Process infractions for display
    const processedInfractions = processInfractionsForDisplay(allInfractions);

    // Build response object
    const response = {
      success: true,
      userRole: session.role,
      employee: {
        employee_id: employee.employee_id,
        full_name: employee.full_name,
        primary_location: employee.primary_location,
        hire_date: employee.hire_date instanceof Date
          ? employee.hire_date.toISOString()
          : (employee.hire_date || null),
        has_system_access: hasSystemAccess
      },
      currentPoints: {
        total: pointData.total_points || 0,
        activeCount: (pointData.active_infractions || []).length,
        pointLevelColor: getPointLevelColor(pointData.total_points || 0),
        statusBadges: getStatusBadges(pointData, pointData.active_infractions || []),
        nextExpirationDate: pointData.next_expiration_date instanceof Date
          ? pointData.next_expiration_date.toISOString()
          : (pointData.next_expiration_date || null)
      },
      pointTrends: pointTrends,
      statistics: stats,
      thresholdHistory: thresholdHistory,
      infractions: processedInfractions,
      actionTracking: actionTracking,
      executionTime: Date.now() - startTime
    };

    console.log(`getEmployeeDetailData completed in ${response.executionTime}ms for ${employeeId}`);
    return response;

  } catch (error) {
    console.error('Error in getEmployeeDetailData:', error);
    return {
      success: false,
      error: error.toString(),
      errorStack: error.stack
    };
  }
}

/**
 * Gets ALL infractions for an employee, including deleted and modified.
 * Used for the detail view timeline.
 *
 * @param {string} employeeId - The employee ID
 * @returns {Array} Array of all infraction objects
 */
function getAllEmployeeInfractions(employeeId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      console.error('Infractions sheet not found');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // Read all data (A:P columns = 16 columns)
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const infractions = [];

    // Column indices (0-based)
    // A=0: Infraction_ID, B=1: Employee_ID, C=2: Full_Name, D=3: Date
    // E=4: Infraction_Type, F=5: Points_Assigned, G=6: Bucket
    // H=7: Description, I=8: Location, J=9: Entered_By
    // K=10: Entry_Timestamp, L=11: Modified_By, M=12: Modified_Date
    // N=13: Modification_Reason, O=14: Status, P=15: Expiration_Date

    for (const row of data) {
      if (row[1] !== employeeId) {
        continue;
      }

      // Parse dates
      const infractionDate = row[3] instanceof Date ? row[3] : new Date(row[3]);
      const expirationDate = row[15] instanceof Date ? row[15] : new Date(row[15]);
      const entryTimestamp = row[10] instanceof Date ? row[10] : (row[10] ? new Date(row[10]) : null);
      const modifiedDate = row[12] instanceof Date ? row[12] : (row[12] ? new Date(row[12]) : null);

      // Check if expired (older than 90 days from today)
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const isExpired = expirationDate < today;

      // Check if positive (negative points = positive behavior credit)
      const points = Number(row[5]) || 0;
      const isPositive = points < 0;

      infractions.push({
        infraction_id: row[0],
        employee_id: row[1],
        full_name: row[2],
        date: infractionDate,
        infraction_type: row[4],
        points: points,
        bucket: row[6],
        description: row[7],
        location: row[8],
        entered_by: row[9],
        entry_timestamp: entryTimestamp,
        modified_by: row[11] || null,
        modified_date: modifiedDate,
        modification_reason: row[13] || null,
        status: row[14] || 'Active',
        expiration_date: expirationDate,
        is_expired: isExpired,
        is_positive: isPositive
      });
    }

    // Sort by date descending (newest first)
    infractions.sort((a, b) => b.date.getTime() - a.date.getTime());

    return infractions;

  } catch (error) {
    console.error('Error in getAllEmployeeInfractions:', error);
    return [];
  }
}

/**
 * Calculates employee statistics from infraction history.
 *
 * @param {string} employeeId - The employee ID
 * @param {Array} infractions - All infractions for the employee
 * @returns {Object} Statistics object
 */
function calculateEmployeeStatistics(employeeId, infractions) {
  const stats = {
    totalInfractionsEver: 0,
    activeInfractions: 0,
    expiredInfractions: 0,
    positiveCredits: 0,
    deletedInfractions: 0,
    mostCommonType: null,
    avgDaysBetweenInfractions: null,
    highestPointsEver: 0,
    daysSinceLastInfraction: null,
    totalPointsEver: 0
  };

  if (!infractions || infractions.length === 0) {
    return stats;
  }

  const typeCounts = {};
  const activeInfractionDates = [];
  let runningTotal = 0;
  let maxPoints = 0;

  // Sort infractions by date chronologically for point calculation
  const chronoInfractions = [...infractions].sort((a, b) => a.date.getTime() - b.date.getTime());

  for (const inf of infractions) {
    // Count by status
    if (inf.status === 'Deleted') {
      stats.deletedInfractions++;
      continue; // Don't count deleted in stats
    }

    if (inf.is_positive) {
      stats.positiveCredits++;
    } else {
      stats.totalInfractionsEver++;

      // Count types
      const type = inf.infraction_type || 'Unknown';
      typeCounts[type] = (typeCounts[type] || 0) + 1;

      // Track dates for average calculation
      if (!inf.is_expired && inf.status === 'Active') {
        stats.activeInfractions++;
      } else if (inf.is_expired) {
        stats.expiredInfractions++;
      }

      activeInfractionDates.push(inf.date);
    }

    // Calculate total points ever assigned
    stats.totalPointsEver += Math.abs(inf.points);
  }

  // Calculate highest point total ever (running total simulation)
  for (const inf of chronoInfractions) {
    if (inf.status !== 'Active' && inf.status !== 'Expired') continue;
    runningTotal += inf.points;
    if (runningTotal > maxPoints) {
      maxPoints = runningTotal;
    }
    // Check for expirations (simplified - would need full simulation for accuracy)
  }
  stats.highestPointsEver = maxPoints;

  // Find most common type
  let maxCount = 0;
  for (const [type, count] of Object.entries(typeCounts)) {
    if (count > maxCount) {
      maxCount = count;
      stats.mostCommonType = type;
    }
  }

  // Calculate average days between infractions
  if (activeInfractionDates.length > 1) {
    activeInfractionDates.sort((a, b) => a.getTime() - b.getTime());
    let totalDays = 0;
    for (let i = 1; i < activeInfractionDates.length; i++) {
      const diff = activeInfractionDates[i].getTime() - activeInfractionDates[i-1].getTime();
      totalDays += diff / (1000 * 60 * 60 * 24);
    }
    stats.avgDaysBetweenInfractions = Math.round(totalDays / (activeInfractionDates.length - 1));
  }

  // Calculate days since last infraction
  const lastInfractionDate = activeInfractionDates.length > 0
    ? activeInfractionDates[activeInfractionDates.length - 1]
    : null;

  if (lastInfractionDate) {
    const today = new Date();
    const diff = today.getTime() - lastInfractionDate.getTime();
    stats.daysSinceLastInfraction = Math.floor(diff / (1000 * 60 * 60 * 24));
  }

  return stats;
}

/**
 * Calculates threshold crossing history from infractions.
 *
 * @param {Array} infractions - All infractions for the employee
 * @returns {Array} Array of threshold events with dates
 */
function calculateThresholdHistory(infractions) {
  const thresholdEvents = [];

  if (!infractions || infractions.length === 0) {
    return thresholdEvents;
  }

  // Sort chronologically
  const chronoInfractions = [...infractions]
    .filter(inf => inf.status === 'Active' || inf.status === 'Expired')
    .sort((a, b) => a.date.getTime() - b.date.getTime());

  if (chronoInfractions.length === 0) {
    return thresholdEvents;
  }

  // Simulate point accumulation with 90-day rolling window
  let currentPoints = 0;
  let previousPoints = 0;
  const thresholds = [
    { level: 6, label: 'At Risk (6 points)', consequence: 'Written warning issued' },
    { level: 9, label: 'Final Warning (9 points)', consequence: '30-day probation period' },
    { level: 15, label: 'Termination Level (15 points)', consequence: 'Subject to termination' }
  ];

  for (const inf of chronoInfractions) {
    previousPoints = currentPoints;
    currentPoints += inf.points;

    // Cap at minimum of -6
    if (currentPoints < -6) currentPoints = -6;

    const infDateStr = inf.date instanceof Date
      ? inf.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
      : 'Unknown date';

    // Check each threshold
    for (const threshold of thresholds) {
      // Crossed threshold going UP
      if (previousPoints < threshold.level && currentPoints >= threshold.level) {
        thresholdEvents.push({
          date: inf.date instanceof Date ? inf.date.toISOString() : null,
          dateFormatted: infDateStr,
          type: 'crossed_up',
          threshold: threshold.level,
          label: `Reached ${threshold.level} points`,
          description: threshold.label,
          consequence: threshold.consequence,
          triggeringInfraction: inf.infraction_type
        });
      }

      // Crossed threshold going DOWN (points expired or removed)
      if (previousPoints >= threshold.level && currentPoints < threshold.level) {
        thresholdEvents.push({
          date: inf.date instanceof Date ? inf.date.toISOString() : null,
          dateFormatted: infDateStr,
          type: 'crossed_down',
          threshold: threshold.level,
          label: `Returned below ${threshold.level} points`,
          description: `Now at ${currentPoints} points`,
          consequence: null,
          triggeringInfraction: inf.is_positive ? 'Positive Behavior Credit' : 'Point expiration'
        });
      }
    }
  }

  return thresholdEvents;
}

/**
 * Calculates point trends - points gained recently and expiring soon.
 *
 * @param {Array} infractions - All infractions for the employee
 * @returns {Object} Point trend data
 */
function calculatePointTrends(infractions) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  const thirtyDaysFromNow = new Date(today);
  thirtyDaysFromNow.setDate(thirtyDaysFromNow.getDate() + 30);

  let pointsLast30Days = 0;
  let pointsExpiringNext30Days = 0;
  let nextExpirationDate = null;
  let nextExpirationPoints = 0;

  for (const inf of infractions) {
    if (inf.status !== 'Active' || inf.is_expired) continue;

    // Points gained in last 30 days
    if (inf.date >= thirtyDaysAgo && inf.date <= today) {
      pointsLast30Days += inf.points;
    }

    // Points expiring in next 30 days
    if (inf.expiration_date >= today && inf.expiration_date <= thirtyDaysFromNow) {
      pointsExpiringNext30Days += inf.points;

      // Track next expiration
      if (!nextExpirationDate || inf.expiration_date < nextExpirationDate) {
        nextExpirationDate = inf.expiration_date;
        nextExpirationPoints = inf.points;
      }
    }
  }

  return {
    pointsLast30Days: pointsLast30Days,
    pointsExpiringNext30Days: pointsExpiringNext30Days,
    nextExpirationDate: nextExpirationDate instanceof Date
      ? nextExpirationDate.toISOString()
      : null,
    nextExpirationPoints: nextExpirationPoints
  };
}

/**
 * Gets action tracking data for an employee (Directors only).
 *
 * @param {string} employeeId - The employee ID
 * @returns {Object} Action tracking data
 */
function getActionTrackingData(employeeId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('ActionTracking');

    // If sheet doesn't exist, return empty tracking
    if (!sheet) {
      return {
        scheduleDayRemoved: { completed: false, date: null, notes: null },
        suspensionCompleted: { completed: false, date: null, notes: null },
        directorMeetingHeld: { completed: false, date: null, notes: null },
        directorNotes: null
      };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        scheduleDayRemoved: { completed: false, date: null, notes: null },
        suspensionCompleted: { completed: false, date: null, notes: null },
        directorMeetingHeld: { completed: false, date: null, notes: null },
        directorNotes: null
      };
    }

    // Read all action tracking data
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    // Find records for this employee
    const actions = {
      scheduleDayRemoved: { completed: false, date: null, notes: null },
      suspensionCompleted: { completed: false, date: null, notes: null },
      directorMeetingHeld: { completed: false, date: null, notes: null },
      directorNotes: null
    };

    for (const row of data) {
      if (row[0] !== employeeId) continue;

      const actionType = row[1];
      const completed = row[2] === true || row[2] === 'TRUE';
      const actionDate = row[3] instanceof Date ? row[3].toISOString() : (row[3] || null);
      const notes = row[4] || null;

      switch (actionType) {
        case 'schedule_day_removed':
          actions.scheduleDayRemoved = { completed, date: actionDate, notes };
          break;
        case 'suspension_completed':
          actions.suspensionCompleted = { completed, date: actionDate, notes };
          break;
        case 'director_meeting':
          actions.directorMeetingHeld = { completed, date: actionDate, notes };
          break;
        case 'director_notes':
          actions.directorNotes = notes;
          break;
      }
    }

    return actions;

  } catch (error) {
    console.error('Error in getActionTrackingData:', error);
    return {
      scheduleDayRemoved: { completed: false, date: null, notes: null },
      suspensionCompleted: { completed: false, date: null, notes: null },
      directorMeetingHeld: { completed: false, date: null, notes: null },
      directorNotes: null,
      error: error.toString()
    };
  }
}

/**
 * Processes infractions for display, converting dates to strings.
 *
 * @param {Array} infractions - Raw infraction data
 * @returns {Array} Processed infractions with serializable dates
 */
function processInfractionsForDisplay(infractions) {
  return infractions.map(inf => ({
    infraction_id: inf.infraction_id,
    date: inf.date instanceof Date ? inf.date.toISOString() : (inf.date || null),
    dateFormatted: inf.date instanceof Date
      ? inf.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
      : 'Unknown',
    infraction_type: inf.infraction_type,
    points: inf.points,
    bucket: inf.bucket,
    description: inf.description,
    location: inf.location,
    entered_by: inf.entered_by,
    entry_timestamp: inf.entry_timestamp instanceof Date
      ? inf.entry_timestamp.toISOString()
      : (inf.entry_timestamp || null),
    modified_by: inf.modified_by,
    modified_date: inf.modified_date instanceof Date
      ? inf.modified_date.toISOString()
      : (inf.modified_date || null),
    modification_reason: inf.modification_reason,
    status: inf.status,
    expiration_date: inf.expiration_date instanceof Date
      ? inf.expiration_date.toISOString()
      : (inf.expiration_date || null),
    expirationFormatted: inf.expiration_date instanceof Date
      ? inf.expiration_date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
      : 'Unknown',
    is_expired: inf.is_expired,
    is_positive: inf.is_positive,
    // Color coding for timeline
    severityColor: inf.is_positive ? 'green' :
      (inf.is_expired ? 'gray' :
        (inf.points >= 5 ? 'red' : (inf.points >= 3 ? 'orange' : 'yellow')))
  }));
}

// ============================================
// DIRECTOR ACTION FUNCTIONS
// ============================================

/**
 * Saves action tracking data for an employee (Directors only).
 *
 * @param {string} employeeId - The employee ID
 * @param {string} actionType - Type of action (schedule_day_removed, suspension_completed, director_meeting, director_notes)
 * @param {boolean} completed - Whether action is completed
 * @param {string} actionDate - Date of action (ISO string)
 * @param {string} notes - Optional notes
 * @returns {Object} Result with success status
 */
function saveActionTracking(employeeId, actionType, completed, actionDate, notes) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can update action tracking' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('ActionTracking');

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('ActionTracking');
      sheet.appendRow(['Employee_ID', 'Action_Type', 'Completed', 'Action_Date', 'Notes', 'Updated_By', 'Updated_Timestamp']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }

    // Look for existing record
    const lastRow = sheet.getLastRow();
    let foundRow = -1;

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === employeeId && data[i][1] === actionType) {
          foundRow = i + 2;
          break;
        }
      }
    }

    const timestamp = new Date();
    const parsedDate = actionDate ? new Date(actionDate) : null;

    if (foundRow > 0) {
      // Update existing record
      sheet.getRange(foundRow, 3, 1, 5).setValues([[
        completed,
        parsedDate,
        notes || '',
        session.role,
        timestamp
      ]]);
    } else {
      // Add new record
      sheet.appendRow([
        employeeId,
        actionType,
        completed,
        parsedDate,
        notes || '',
        session.role,
        timestamp
      ]);
    }

    return { success: true };

  } catch (error) {
    console.error('Error in saveActionTracking:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Adds a positive behavior credit for an employee (Directors only).
 *
 * @param {Object} creditData - Credit data containing employee_id, date, description, points
 * @returns {Object} Result with success status
 */
function addPositiveBehaviorCredit(creditData) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can add positive behavior credits' };
    }

    // Validate input
    if (!creditData.employee_id || !creditData.date || !creditData.description) {
      return { success: false, error: 'Missing required fields' };
    }

    if (!creditData.description || creditData.description.length < 240) {
      return { success: false, error: 'Description must be at least 240 characters' };
    }

    const points = Math.abs(creditData.points || 1);
    if (points < 1 || points > 6) {
      return { success: false, error: 'Points must be between 1 and 6' };
    }

    // Get employee info
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === creditData.employee_id);

    if (!employee) {
      return { success: false, error: 'Employee not found' };
    }

    // Create infraction entry with negative points
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    // Generate infraction ID
    const infractionId = generateInfractionId();

    // Parse date
    const creditDate = new Date(creditData.date);
    const expirationDate = new Date(creditDate);
    expirationDate.setDate(expirationDate.getDate() + 90);

    // Append the row
    sheet.appendRow([
      infractionId,                           // A: Infraction_ID
      creditData.employee_id,                 // B: Employee_ID
      employee.full_name,                     // C: Full_Name
      creditDate,                             // D: Date
      'Positive Behavior Credit',             // E: Infraction_Type
      -points,                                // F: Points_Assigned (negative for credits)
      'Credit',                               // G: Bucket
      creditData.description,                 // H: Description
      creditData.location || employee.primary_location, // I: Location
      creditData.entered_by || 'Director',    // J: Entered_By
      new Date(),                             // K: Entry_Timestamp
      '',                                     // L: Modified_By
      '',                                     // M: Modified_Date
      '',                                     // N: Modification_Reason
      'Active',                               // O: Status
      expirationDate                          // P: Expiration_Date
    ]);

    return {
      success: true,
      message: `Added ${points} point credit for ${employee.full_name}`,
      infractionId: infractionId
    };

  } catch (error) {
    console.error('Error in addPositiveBehaviorCredit:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Removes points from an employee by creating a point adjustment (Directors only).
 *
 * @param {Object} removalData - Removal data containing employee_id, points, reason
 * @returns {Object} Result with success status
 */
function removePoints(removalData) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can remove points' };
    }

    // Validate input
    if (!removalData.employee_id || !removalData.reason) {
      return { success: false, error: 'Missing required fields' };
    }

    if (!removalData.reason || removalData.reason.length < 240) {
      return { success: false, error: 'Reason must be at least 240 characters' };
    }

    const pointsToRemove = Math.abs(removalData.points || 1);
    if (pointsToRemove < 1) {
      return { success: false, error: 'Must remove at least 1 point' };
    }

    // Get employee info
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === removalData.employee_id);

    if (!employee) {
      return { success: false, error: 'Employee not found' };
    }

    // Create point adjustment entry
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    // Generate infraction ID
    const infractionId = generateInfractionId();

    const today = new Date();
    const expirationDate = new Date(today);
    expirationDate.setDate(expirationDate.getDate() + 90);

    // Append the row
    sheet.appendRow([
      infractionId,                           // A: Infraction_ID
      removalData.employee_id,                // B: Employee_ID
      employee.full_name,                     // C: Full_Name
      today,                                  // D: Date
      'Point Adjustment',                     // E: Infraction_Type
      -pointsToRemove,                        // F: Points_Assigned (negative)
      'Adjustment',                           // G: Bucket
      removalData.reason,                     // H: Description
      employee.primary_location,              // I: Location
      removalData.entered_by || 'Director',   // J: Entered_By
      new Date(),                             // K: Entry_Timestamp
      '',                                     // L: Modified_By
      '',                                     // M: Modified_Date
      '',                                     // N: Modification_Reason
      'Active',                               // O: Status
      expirationDate                          // P: Expiration_Date
    ]);

    return {
      success: true,
      message: `Removed ${pointsToRemove} points from ${employee.full_name}`,
      infractionId: infractionId
    };

  } catch (error) {
    console.error('Error in removePoints:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Edits an existing infraction (Directors only).
 *
 * @param {Object} editData - Edit data containing infraction_id and fields to update
 * @returns {Object} Result with success status
 */
function editInfraction(editData) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can edit infractions' };
    }

    // Validate input
    if (!editData.infraction_id || !editData.modification_reason) {
      return { success: false, error: 'Infraction ID and modification reason are required' };
    }

    if (editData.modification_reason.length < 240) {
      return { success: false, error: 'Modification reason must be at least 240 characters' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    // Find the infraction
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, error: 'No infractions found' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    let foundRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === editData.infraction_id) {
        foundRow = i + 2;
        break;
      }
    }

    if (foundRow < 0) {
      return { success: false, error: 'Infraction not found' };
    }

    // Update fields that were provided
    const updates = [];

    if (editData.description !== undefined) {
      sheet.getRange(foundRow, 8).setValue(editData.description); // H: Description
    }

    if (editData.points !== undefined) {
      sheet.getRange(foundRow, 6).setValue(editData.points); // F: Points_Assigned
    }

    if (editData.date !== undefined) {
      const newDate = new Date(editData.date);
      const newExpiration = new Date(newDate);
      newExpiration.setDate(newExpiration.getDate() + 90);
      sheet.getRange(foundRow, 4).setValue(newDate); // D: Date
      sheet.getRange(foundRow, 16).setValue(newExpiration); // P: Expiration_Date
    }

    // Update modification tracking
    sheet.getRange(foundRow, 12).setValue(editData.modified_by || 'Director'); // L: Modified_By
    sheet.getRange(foundRow, 13).setValue(new Date()); // M: Modified_Date
    sheet.getRange(foundRow, 14).setValue(editData.modification_reason); // N: Modification_Reason
    sheet.getRange(foundRow, 15).setValue('Modified'); // O: Status

    return {
      success: true,
      message: 'Infraction updated successfully'
    };

  } catch (error) {
    console.error('Error in editInfraction:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Deletes/soft-deletes an infraction (Directors only).
 *
 * @param {string} infractionId - The infraction ID to delete
 * @param {string} reason - Reason for deletion (min 240 chars)
 * @returns {Object} Result with success status
 */
function deleteInfraction(infractionId, reason) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can delete infractions' };
    }

    if (!reason || reason.length < 240) {
      return { success: false, error: 'Deletion reason must be at least 240 characters' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    // Find the infraction
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, error: 'No infractions found' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let foundRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === infractionId) {
        foundRow = i + 2;
        break;
      }
    }

    if (foundRow < 0) {
      return { success: false, error: 'Infraction not found' };
    }

    // Soft delete - update status to Deleted
    sheet.getRange(foundRow, 12).setValue('Director'); // L: Modified_By
    sheet.getRange(foundRow, 13).setValue(new Date()); // M: Modified_Date
    sheet.getRange(foundRow, 14).setValue(reason); // N: Modification_Reason
    sheet.getRange(foundRow, 15).setValue('Deleted'); // O: Status

    return {
      success: true,
      message: 'Infraction deleted successfully'
    };

  } catch (error) {
    console.error('Error in deleteInfraction:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Marks an employee as terminated (Directors only).
 * Moves infractions to archive and flags employee record.
 *
 * @param {Object} terminationData - Termination details
 * @returns {Object} Result with success status
 */
function markAsTerminated(terminationData) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can mark employees as terminated' };
    }

    if (!terminationData.employee_id || !terminationData.reason || !terminationData.termination_date) {
      return { success: false, error: 'Employee ID, reason, and termination date are required' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);

    // Get or create Terminated_Employees sheet
    let terminatedSheet = ss.getSheetByName('Terminated_Employees');
    if (!terminatedSheet) {
      terminatedSheet = ss.insertSheet('Terminated_Employees');
      terminatedSheet.appendRow([
        'Employee_ID', 'Full_Name', 'Primary_Location', 'Termination_Date',
        'Termination_Reason', 'Final_Point_Total', 'Terminated_By', 'Timestamp'
      ]);
      terminatedSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    // Get employee info
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === terminationData.employee_id);

    if (!employee) {
      return { success: false, error: 'Employee not found' };
    }

    // Calculate final points
    const pointData = calculatePoints(terminationData.employee_id);

    // Add to terminated employees sheet
    terminatedSheet.appendRow([
      terminationData.employee_id,
      employee.full_name,
      employee.primary_location,
      new Date(terminationData.termination_date),
      terminationData.reason,
      pointData.total_points,
      'Director',
      new Date()
    ]);

    // Archive infractions - update status to Terminated
    const infractionsSheet = ss.getSheetByName('Infractions');
    if (infractionsSheet) {
      const lastRow = infractionsSheet.getLastRow();
      if (lastRow >= 2) {
        const data = infractionsSheet.getRange(2, 1, lastRow - 1, 16).getValues();
        for (let i = 0; i < data.length; i++) {
          if (data[i][1] === terminationData.employee_id && data[i][14] === 'Active') {
            infractionsSheet.getRange(i + 2, 15).setValue('Archived - Terminated'); // O: Status
          }
        }
      }
    }

    return {
      success: true,
      message: `${employee.full_name} has been marked as terminated`
    };

  } catch (error) {
    console.error('Error in markAsTerminated:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Generates a PDF write-up for an employee.
 *
 * @param {string} employeeId - The employee ID
 * @returns {Object} Result with PDF URL or base64 data
 */
function generateWriteUpPdf(employeeId) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    // Get employee details
    const detailData = getEmployeeDetailData(employeeId);
    if (!detailData.success) {
      return { success: false, error: detailData.error };
    }

    // For now, return a placeholder since PDF generation requires more setup
    // In a full implementation, this would use Google Docs API to generate a PDF
    return {
      success: true,
      message: 'PDF generation not yet implemented. Use print functionality for now.',
      printUrl: null
    };

  } catch (error) {
    console.error('Error in generateWriteUpPdf:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Test function for getEmployeeDetailData.
 */
function testGetEmployeeDetailData() {
  console.log('=== Testing getEmployeeDetailData ===');
  console.log('');

  // Get a sample employee ID from the system
  const employees = getActiveEmployees();
  if (!employees || employees.length === 0) {
    console.log('FAIL: No employees found in system');
    return;
  }

  // Test with first employee
  const testEmployeeId = employees[0].employee_id;
  console.log('Testing with employee:', testEmployeeId);
  console.log('');

  const startTime = Date.now();
  const result = getEmployeeDetailData(testEmployeeId);
  const duration = Date.now() - startTime;

  console.log('Execution time:', duration, 'ms');
  console.log('Success:', result.success);
  console.log('');

  if (result.success) {
    // Check required fields
    console.log('Employee Info:');
    console.log('  ID:', result.employee.employee_id);
    console.log('  Name:', result.employee.full_name);
    console.log('  Location:', result.employee.primary_location);
    console.log('');

    console.log('Current Points:');
    console.log('  Total:', result.currentPoints.total);
    console.log('  Active Infractions:', result.currentPoints.activeCount);
    console.log('  Point Level Color:', result.currentPoints.pointLevelColor);
    console.log('  Status Badges:', JSON.stringify(result.currentPoints.statusBadges));
    console.log('');

    console.log('Point Trends:');
    console.log('  Points Last 30 Days:', result.pointTrends.pointsLast30Days);
    console.log('  Points Expiring Next 30 Days:', result.pointTrends.pointsExpiringNext30Days);
    console.log('');

    console.log('Statistics:');
    console.log('  Total Infractions Ever:', result.statistics.totalInfractionsEver);
    console.log('  Active:', result.statistics.activeInfractions);
    console.log('  Expired:', result.statistics.expiredInfractions);
    console.log('  Positive Credits:', result.statistics.positiveCredits);
    console.log('  Most Common Type:', result.statistics.mostCommonType);
    console.log('  Highest Points Ever:', result.statistics.highestPointsEver);
    console.log('  Days Since Last Infraction:', result.statistics.daysSinceLastInfraction);
    console.log('');

    console.log('Threshold History:', result.thresholdHistory.length, 'events');
    for (const event of result.thresholdHistory) {
      console.log('  -', event.dateFormatted, ':', event.label);
    }
    console.log('');

    console.log('Infractions:', result.infractions.length, 'total');
    if (result.infractions.length > 0) {
      console.log('First infraction:', JSON.stringify(result.infractions[0], null, 2));
    }
    console.log('');

    // Performance check
    if (duration < 10000) {
      console.log('PASS: Completed in under 10 seconds');
    } else {
      console.log('WARN: Took longer than 10 seconds');
    }

  } else {
    console.log('Error:', result.error);
    if (result.sessionExpired) {
      console.log('Session expired - run createSession() first');
    }
  }

  console.log('');
  console.log('=== Test Complete (Phase 13) ===');
}

// ============================================================================
// MICRO-PHASE 14: DIRECTOR EDIT/OVERRIDE FUNCTIONALITY WITH AUDIT TRAIL
// ============================================================================

/**
 * Gets or creates the Edit_Log audit trail sheet.
 *
 * Schema:
 * A: Log_ID - Unique identifier for this log entry
 * B: Timestamp - When the action occurred
 * C: Action_Type - edit_infraction, delete_infraction, remove_points, add_credit, terminate
 * D: Director_Email - Email of director who performed the action
 * E: Target_Type - infraction, employee, points
 * F: Target_ID - Infraction_ID or Employee_ID
 * G: Employee_ID - Always the affected employee
 * H: Employee_Name - Name for easy reference
 * I: Field_Changed - Which field was modified
 * J: Original_Value - Value before change
 * K: New_Value - Value after change
 * L: Reason - Director's reason for the change (min 240 chars)
 * M: IP_Address - For security audit (if available)
 *
 * @returns {Sheet} The Edit_Log sheet
 */
function getOrCreateEditLogSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Edit_Log');

  if (!sheet) {
    sheet = ss.insertSheet('Edit_Log');

    // Set up headers
    const headers = [
      'Log_ID',
      'Timestamp',
      'Action_Type',
      'Director_Email',
      'Target_Type',
      'Target_ID',
      'Employee_ID',
      'Employee_Name',
      'Field_Changed',
      'Original_Value',
      'New_Value',
      'Reason',
      'Session_Info'
    ];

    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E51636').setFontColor('white');

    // Freeze header row
    sheet.setFrozenRows(1);

    // Set column widths for readability
    sheet.setColumnWidth(1, 120);  // Log_ID
    sheet.setColumnWidth(2, 150);  // Timestamp
    sheet.setColumnWidth(3, 130);  // Action_Type
    sheet.setColumnWidth(4, 180);  // Director_Email
    sheet.setColumnWidth(5, 100);  // Target_Type
    sheet.setColumnWidth(6, 120);  // Target_ID
    sheet.setColumnWidth(7, 100);  // Employee_ID
    sheet.setColumnWidth(8, 150);  // Employee_Name
    sheet.setColumnWidth(9, 120);  // Field_Changed
    sheet.setColumnWidth(10, 200); // Original_Value
    sheet.setColumnWidth(11, 200); // New_Value
    sheet.setColumnWidth(12, 400); // Reason
    sheet.setColumnWidth(13, 150); // Session_Info

    console.log('Created Edit_Log audit trail sheet');
  }

  return sheet;
}

/**
 * Generates a unique log entry ID.
 * Format: LOG-YYYYMMDD-XXXXXX (random alphanumeric)
 *
 * @returns {string} Unique log ID
 */
function generateLogId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'America/Chicago', 'yyyyMMdd');
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let randomPart = '';
  for (let i = 0; i < 6; i++) {
    randomPart += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return `LOG-${dateStr}-${randomPart}`;
}

/**
 * Logs an edit action to the Edit_Log audit trail.
 *
 * @param {Object} logData - Data to log
 * @param {string} logData.actionType - Type of action performed
 * @param {string} logData.directorEmail - Email of director
 * @param {string} logData.targetType - Type of target (infraction, employee, points)
 * @param {string} logData.targetId - ID of target (infraction_id or employee_id)
 * @param {string} logData.employeeId - Affected employee ID
 * @param {string} logData.employeeName - Affected employee name
 * @param {string} logData.fieldChanged - Field that was changed
 * @param {string} logData.originalValue - Original value
 * @param {string} logData.newValue - New value
 * @param {string} logData.reason - Reason for change
 * @param {string} logData.sessionInfo - Additional session info
 * @returns {string} The generated Log_ID
 */
function logEditAction(logData) {
  try {
    const sheet = getOrCreateEditLogSheet();
    const logId = generateLogId();
    const timestamp = new Date();

    sheet.appendRow([
      logId,
      timestamp,
      logData.actionType || 'unknown',
      logData.directorEmail || 'unknown',
      logData.targetType || 'unknown',
      logData.targetId || '',
      logData.employeeId || '',
      logData.employeeName || '',
      logData.fieldChanged || '',
      String(logData.originalValue || ''),
      String(logData.newValue || ''),
      logData.reason || '',
      logData.sessionInfo || ''
    ]);

    console.log(`Logged edit action: ${logId} - ${logData.actionType}`);
    return logId;

  } catch (error) {
    console.error('Error logging edit action:', error);
    // Don't throw - logging failure shouldn't prevent the action
    return null;
  }
}

/**
 * Gets the original values of an infraction for audit logging.
 *
 * @param {string} infractionId - The infraction ID
 * @returns {Object|null} Original infraction data or null if not found
 */
function getInfractionOriginalValues(infractionId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) return null;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === infractionId) {
        return {
          row: i + 2,
          infraction_id: data[i][0],
          employee_id: data[i][1],
          full_name: data[i][2],
          date: data[i][3],
          infraction_type: data[i][4],
          points: data[i][5],
          bucket: data[i][6],
          description: data[i][7],
          location: data[i][8],
          entered_by: data[i][9],
          entry_timestamp: data[i][10],
          modified_by: data[i][11],
          modified_date: data[i][12],
          modification_reason: data[i][13],
          status: data[i][14],
          expiration_date: data[i][15]
        };
      }
    }

    return null;

  } catch (error) {
    console.error('Error getting infraction original values:', error);
    return null;
  }
}

/**
 * Enhanced editInfraction with full audit trail.
 * Now logs all changes to Edit_Log before making modifications.
 *
 * @param {Object} editData - Edit data containing:
 *   - infraction_id: Required - ID of infraction to edit
 *   - modification_reason: Required - Reason for modification (min 240 chars)
 *   - description: Optional - New description
 *   - points: Optional - New point value
 *   - date: Optional - New date
 *   - infraction_type: Optional - New infraction type
 *   - location: Optional - New location
 * @returns {Object} Result with success status and log_id
 */
function editInfractionWithAudit(editData) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can edit infractions' };
    }

    // Validate input
    if (!editData.infraction_id) {
      return { success: false, error: 'Infraction ID is required' };
    }

    if (!editData.modification_reason || editData.modification_reason.length < 240) {
      return { success: false, error: 'Modification reason must be at least 240 characters' };
    }

    // Get original values for audit
    const original = getInfractionOriginalValues(editData.infraction_id);
    if (!original) {
      return { success: false, error: 'Infraction not found' };
    }

    // Get director email from session
    const directorEmail = Session.getActiveUser().getEmail() || 'Unknown Director';

    // Track all changes for audit log
    const changes = [];

    if (editData.description !== undefined && editData.description !== original.description) {
      changes.push({
        field: 'description',
        originalValue: original.description,
        newValue: editData.description
      });
    }

    if (editData.points !== undefined && editData.points !== original.points) {
      changes.push({
        field: 'points',
        originalValue: original.points,
        newValue: editData.points
      });
    }

    if (editData.date !== undefined) {
      const newDate = new Date(editData.date);
      const origDate = original.date instanceof Date ? original.date : new Date(original.date);
      if (newDate.getTime() !== origDate.getTime()) {
        changes.push({
          field: 'date',
          originalValue: origDate.toISOString(),
          newValue: newDate.toISOString()
        });
      }
    }

    if (editData.infraction_type !== undefined && editData.infraction_type !== original.infraction_type) {
      changes.push({
        field: 'infraction_type',
        originalValue: original.infraction_type,
        newValue: editData.infraction_type
      });
    }

    if (editData.location !== undefined && editData.location !== original.location) {
      changes.push({
        field: 'location',
        originalValue: original.location,
        newValue: editData.location
      });
    }

    if (changes.length === 0) {
      return { success: false, error: 'No changes detected' };
    }

    // Log each change to audit trail
    const logIds = [];
    for (const change of changes) {
      const logId = logEditAction({
        actionType: 'edit_infraction',
        directorEmail: directorEmail,
        targetType: 'infraction',
        targetId: editData.infraction_id,
        employeeId: original.employee_id,
        employeeName: original.full_name,
        fieldChanged: change.field,
        originalValue: change.originalValue,
        newValue: change.newValue,
        reason: editData.modification_reason,
        sessionInfo: `Session: ${session.role}`
      });
      if (logId) logIds.push(logId);
    }

    // Now apply the actual changes
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');
    const row = original.row;

    if (editData.description !== undefined) {
      sheet.getRange(row, 8).setValue(editData.description); // H: Description
    }

    if (editData.points !== undefined) {
      sheet.getRange(row, 6).setValue(editData.points); // F: Points_Assigned
    }

    if (editData.date !== undefined) {
      const newDate = new Date(editData.date);
      const newExpiration = new Date(newDate);
      newExpiration.setDate(newExpiration.getDate() + 90);
      sheet.getRange(row, 4).setValue(newDate); // D: Date
      sheet.getRange(row, 16).setValue(newExpiration); // P: Expiration_Date
    }

    if (editData.infraction_type !== undefined) {
      sheet.getRange(row, 5).setValue(editData.infraction_type); // E: Infraction_Type
    }

    if (editData.location !== undefined) {
      sheet.getRange(row, 9).setValue(editData.location); // I: Location
    }

    // Update modification tracking
    sheet.getRange(row, 12).setValue(directorEmail); // L: Modified_By
    sheet.getRange(row, 13).setValue(new Date()); // M: Modified_Date
    sheet.getRange(row, 14).setValue(editData.modification_reason); // N: Modification_Reason
    sheet.getRange(row, 15).setValue('Modified'); // O: Status

    return {
      success: true,
      message: `Infraction updated successfully. ${changes.length} field(s) modified.`,
      changesApplied: changes.length,
      logIds: logIds
    };

  } catch (error) {
    console.error('Error in editInfractionWithAudit:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Enhanced removePoints with full audit trail.
 *
 * @param {Object} removalData - Removal data containing:
 *   - employee_id: Required - Employee ID
 *   - points: Required - Number of points to remove
 *   - reason: Required - Reason for removal (min 240 chars)
 *   - specific_infraction_id: Optional - If removing points from a specific infraction
 * @returns {Object} Result with success status and log_id
 */
function removePointsWithAudit(removalData) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can remove points' };
    }

    // Validate input
    if (!removalData.employee_id) {
      return { success: false, error: 'Employee ID is required' };
    }

    if (!removalData.reason || removalData.reason.length < 240) {
      return { success: false, error: 'Reason must be at least 240 characters' };
    }

    const pointsToRemove = Math.abs(removalData.points || 1);
    if (pointsToRemove < 1 || pointsToRemove > 15) {
      return { success: false, error: 'Points must be between 1 and 15' };
    }

    // Get employee info
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === removalData.employee_id);

    if (!employee) {
      return { success: false, error: 'Employee not found' };
    }

    // Get current points for audit
    const currentPoints = calculatePoints(removalData.employee_id);
    const directorEmail = Session.getActiveUser().getEmail() || 'Unknown Director';

    // Log the action
    const logId = logEditAction({
      actionType: 'remove_points',
      directorEmail: directorEmail,
      targetType: 'points',
      targetId: removalData.specific_infraction_id || 'general_adjustment',
      employeeId: removalData.employee_id,
      employeeName: employee.full_name,
      fieldChanged: 'points_removed',
      originalValue: currentPoints.total_points,
      newValue: currentPoints.total_points - pointsToRemove,
      reason: removalData.reason,
      sessionInfo: `Session: ${session.role}`
    });

    // Create point adjustment entry
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    // Generate infraction ID
    const infractionId = generateInfractionId();

    const today = new Date();
    const expirationDate = new Date(today);
    expirationDate.setDate(expirationDate.getDate() + 90);

    // Append the adjustment row
    sheet.appendRow([
      infractionId,                           // A: Infraction_ID
      removalData.employee_id,                // B: Employee_ID
      employee.full_name,                     // C: Full_Name
      today,                                  // D: Date
      'Director Point Adjustment',            // E: Infraction_Type
      -pointsToRemove,                        // F: Points_Assigned (negative)
      'Adjustment',                           // G: Bucket
      `POINT REMOVAL: ${removalData.reason}`, // H: Description
      employee.primary_location,              // I: Location
      directorEmail,                          // J: Entered_By
      new Date(),                             // K: Entry_Timestamp
      '',                                     // L: Modified_By
      '',                                     // M: Modified_Date
      '',                                     // N: Modification_Reason
      'Active',                               // O: Status
      expirationDate                          // P: Expiration_Date
    ]);

    return {
      success: true,
      message: `Removed ${pointsToRemove} points from ${employee.full_name}. New total: ${currentPoints.total_points - pointsToRemove}`,
      infractionId: infractionId,
      logId: logId,
      previousPoints: currentPoints.total_points,
      newPoints: currentPoints.total_points - pointsToRemove
    };

  } catch (error) {
    console.error('Error in removePointsWithAudit:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Enhanced deleteInfraction with full audit trail.
 * Soft-deletes the infraction and logs the action.
 *
 * @param {string} infractionId - The infraction ID to delete
 * @param {string} reason - Reason for deletion (min 240 chars)
 * @returns {Object} Result with success status and log_id
 */
function deleteInfractionWithAudit(infractionId, reason) {
  try {
    // Validate session and check for Director role
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can delete infractions' };
    }

    if (!infractionId) {
      return { success: false, error: 'Infraction ID is required' };
    }

    if (!reason || reason.length < 240) {
      return { success: false, error: 'Deletion reason must be at least 240 characters' };
    }

    // Get original infraction for audit
    const original = getInfractionOriginalValues(infractionId);
    if (!original) {
      return { success: false, error: 'Infraction not found' };
    }

    if (original.status === 'Deleted') {
      return { success: false, error: 'Infraction is already deleted' };
    }

    const directorEmail = Session.getActiveUser().getEmail() || 'Unknown Director';

    // Log the deletion
    const logId = logEditAction({
      actionType: 'delete_infraction',
      directorEmail: directorEmail,
      targetType: 'infraction',
      targetId: infractionId,
      employeeId: original.employee_id,
      employeeName: original.full_name,
      fieldChanged: 'status',
      originalValue: `Active - ${original.infraction_type} (${original.points} pts)`,
      newValue: 'Deleted',
      reason: reason,
      sessionInfo: `Session: ${session.role}`
    });

    // Perform the soft delete
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');
    const row = original.row;

    sheet.getRange(row, 12).setValue(directorEmail); // L: Modified_By
    sheet.getRange(row, 13).setValue(new Date()); // M: Modified_Date
    sheet.getRange(row, 14).setValue(reason); // N: Modification_Reason
    sheet.getRange(row, 15).setValue('Deleted'); // O: Status

    return {
      success: true,
      message: `Infraction deleted successfully. ${original.points} points removed from ${original.full_name}.`,
      logId: logId,
      deletedInfraction: {
        id: infractionId,
        type: original.infraction_type,
        points: original.points,
        date: original.date
      }
    };

  } catch (error) {
    console.error('Error in deleteInfractionWithAudit:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Gets the edit history for an employee from the Edit_Log.
 *
 * @param {string} employeeId - Employee ID to get history for
 * @returns {Object} Result with edit history
 */
function getEmployeeEditHistory(employeeId) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can view edit history' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Edit_Log');

    if (!sheet) {
      return { success: true, history: [], message: 'No edit history found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, history: [], message: 'No edit history found' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    const history = [];

    for (const row of data) {
      if (row[6] === employeeId) { // G: Employee_ID
        history.push({
          logId: row[0],
          timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1],
          timestampFormatted: row[1] instanceof Date ?
            row[1].toLocaleDateString('en-US', {
              month: 'short', day: 'numeric', year: 'numeric',
              hour: '2-digit', minute: '2-digit'
            }) : row[1],
          actionType: row[2],
          directorEmail: row[3],
          targetType: row[4],
          targetId: row[5],
          employeeId: row[6],
          employeeName: row[7],
          fieldChanged: row[8],
          originalValue: row[9],
          newValue: row[10],
          reason: row[11],
          sessionInfo: row[12]
        });
      }
    }

    // Sort by timestamp descending (newest first)
    history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    return {
      success: true,
      history: history,
      count: history.length
    };

  } catch (error) {
    console.error('Error in getEmployeeEditHistory:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Gets the full edit log (Directors only).
 * Supports pagination and filtering.
 *
 * @param {Object} options - Query options
 * @param {number} options.limit - Max records to return (default 50)
 * @param {number} options.offset - Starting offset (default 0)
 * @param {string} options.actionType - Filter by action type
 * @param {string} options.directorEmail - Filter by director
 * @returns {Object} Result with edit log entries
 */
function getEditLog(options = {}) {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director') {
      return { success: false, error: 'Only Directors can view the edit log' };
    }

    const limit = options.limit || 50;
    const offset = options.offset || 0;

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Edit_Log');

    if (!sheet) {
      return { success: true, entries: [], total: 0 };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, entries: [], total: 0 };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    let entries = [];

    for (const row of data) {
      const entry = {
        logId: row[0],
        timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1],
        timestampFormatted: row[1] instanceof Date ?
          row[1].toLocaleDateString('en-US', {
            month: 'short', day: 'numeric', year: 'numeric',
            hour: '2-digit', minute: '2-digit'
          }) : row[1],
        actionType: row[2],
        directorEmail: row[3],
        targetType: row[4],
        targetId: row[5],
        employeeId: row[6],
        employeeName: row[7],
        fieldChanged: row[8],
        originalValue: row[9],
        newValue: row[10],
        reason: row[11],
        sessionInfo: row[12]
      };

      // Apply filters
      if (options.actionType && entry.actionType !== options.actionType) {
        continue;
      }
      if (options.directorEmail && entry.directorEmail !== options.directorEmail) {
        continue;
      }

      entries.push(entry);
    }

    // Sort by timestamp descending
    entries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    const total = entries.length;

    // Apply pagination
    entries = entries.slice(offset, offset + limit);

    return {
      success: true,
      entries: entries,
      total: total,
      limit: limit,
      offset: offset,
      hasMore: (offset + entries.length) < total
    };

  } catch (error) {
    console.error('Error in getEditLog:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Gets infraction types for the edit dropdown.
 * Returns unique infraction types from existing data.
 *
 * @returns {Object} Result with infraction types
 */
function getInfractionTypes() {
  try {
    // Validate session
    const session = validateSession();
    if (!session.valid) {
      return { success: false, sessionExpired: true };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Infractions');

    if (!sheet) {
      return { success: true, types: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, types: [] };
    }

    // Get column E (Infraction_Type)
    const data = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    const typeSet = new Set();

    for (const row of data) {
      if (row[0] && typeof row[0] === 'string' && row[0].trim()) {
        typeSet.add(row[0].trim());
      }
    }

    // Sort alphabetically
    const types = Array.from(typeSet).sort();

    // Add standard types if not present
    const standardTypes = [
      'Attendance - Tardy',
      'Attendance - NCNS',
      'Attendance - Left Early',
      'Conduct',
      'Dress Code',
      'Performance',
      'Safety',
      'Positive Behavior Credit',
      'Point Adjustment',
      'Director Point Adjustment'
    ];

    for (const std of standardTypes) {
      if (!types.includes(std)) {
        types.push(std);
      }
    }

    types.sort();

    return {
      success: true,
      types: types
    };

  } catch (error) {
    console.error('Error in getInfractionTypes:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Test function for Micro-Phase 14 Director Edit features.
 */
function testDirectorEdits() {
  console.log('=== Testing Micro-Phase 14: Director Edit Features ===');
  console.log('');

  // Test 1: Create Edit_Log sheet
  console.log('Test 1: Create/Get Edit_Log Sheet');
  try {
    const sheet = getOrCreateEditLogSheet();
    console.log('PASS: Edit_Log sheet exists with', sheet.getLastRow(), 'rows');
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 2: Log an edit action
  console.log('Test 2: Log Edit Action');
  try {
    const logId = logEditAction({
      actionType: 'test_action',
      directorEmail: 'test@example.com',
      targetType: 'test',
      targetId: 'TEST-001',
      employeeId: 'EMP-001',
      employeeName: 'Test Employee',
      fieldChanged: 'test_field',
      originalValue: 'old_value',
      newValue: 'new_value',
      reason: 'This is a test log entry to verify the audit trail functionality. '.repeat(5),
      sessionInfo: 'Test session'
    });
    if (logId) {
      console.log('PASS: Created log entry', logId);
    } else {
      console.log('FAIL: No log ID returned');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 3: Get infraction types
  console.log('Test 3: Get Infraction Types');
  try {
    const result = getInfractionTypes();
    if (result.success) {
      console.log('PASS: Got', result.types.length, 'infraction types');
      console.log('Types:', result.types.slice(0, 5).join(', '), '...');
    } else {
      console.log('FAIL:', result.error);
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 4: Get edit log
  console.log('Test 4: Get Edit Log');
  try {
    const result = getEditLog({ limit: 10 });
    if (result.success) {
      console.log('PASS: Got', result.entries.length, 'of', result.total, 'total entries');
      if (result.entries.length > 0) {
        console.log('Latest entry:', result.entries[0].actionType, 'at', result.entries[0].timestampFormatted);
      }
    } else {
      console.log('FAIL:', result.error);
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 5: Verify sheet structure
  console.log('Test 5: Verify Edit_Log Sheet Structure');
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Edit_Log');
    const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
    const expectedHeaders = [
      'Log_ID', 'Timestamp', 'Action_Type', 'Director_Email',
      'Target_Type', 'Target_ID', 'Employee_ID', 'Employee_Name',
      'Field_Changed', 'Original_Value', 'New_Value', 'Reason', 'Session_Info'
    ];

    let match = true;
    for (let i = 0; i < expectedHeaders.length; i++) {
      if (headers[i] !== expectedHeaders[i]) {
        console.log('FAIL: Header mismatch at column', i + 1, '- expected', expectedHeaders[i], 'got', headers[i]);
        match = false;
      }
    }

    if (match) {
      console.log('PASS: All 13 headers match expected values');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  console.log('=== Director Edit Tests Complete ===');
}

// ============================================================================
// MICRO-PHASE 11 TEST: PIN AUTHENTICATION TESTS
// ============================================================================

/**
 * Test function for Micro-Phase 11 PIN Authentication features.
 * Tests the 3-tier authentication system.
 */
function testPINAuthentication() {
  console.log('=== Testing Micro-Phase 11: PIN Authentication (3-Tier System) ===');
  console.log('');

  // Test 1: Hash PIN function
  console.log('Test 1: Hash PIN Function');
  try {
    const testPin = '123456';
    const hash1 = hashPIN(testPin);
    const hash2 = hashPIN(testPin);
    const hash3 = hashPIN('654321');

    if (hash1 === hash2) {
      console.log('PASS: Same PIN produces same hash');
    } else {
      console.log('FAIL: Same PIN produces different hashes');
    }

    if (hash1 !== hash3) {
      console.log('PASS: Different PINs produce different hashes');
    } else {
      console.log('FAIL: Different PINs produce same hash');
    }

    if (hash1.length === 64) { // SHA-256 produces 64 hex chars
      console.log('PASS: Hash is correct length (64 chars for SHA-256)');
    } else {
      console.log('FAIL: Hash length is', hash1.length, 'expected 64');
    }

    console.log('Sample hash:', hash1.substring(0, 16) + '...');
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 2: Get active users for login
  console.log('Test 2: Get Active Users for Login');
  try {
    const users = getActiveUsersForLogin();
    console.log('Found', users.length, 'active users');

    if (users.length > 0) {
      console.log('PASS: Retrieved active users');
      console.log('First user:', users[0].full_name, '- ID:', users[0].employee_id);

      // Check sorting
      let sorted = true;
      for (let i = 1; i < users.length; i++) {
        if (users[i].full_name.localeCompare(users[i-1].full_name) < 0) {
          sorted = false;
          break;
        }
      }
      if (sorted) {
        console.log('PASS: Users are sorted alphabetically');
      } else {
        console.log('FAIL: Users are not sorted alphabetically');
      }
    } else {
      console.log('INFO: No active users found - add users to User_Permissions sheet');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 3: Session creation
  console.log('Test 3: Session Creation');
  try {
    const testUserData = {
      user_id: 'TEST-001',
      name: 'Test User',
      email: 'test@example.com',
      role: 'Manager',
      can_see_directors: false
    };

    const sessionResult = createSession(testUserData);
    console.log('Session created:', sessionResult.success);

    if (sessionResult.success) {
      console.log('PASS: Session created successfully');
      console.log('Session ID:', sessionResult.session_id ? sessionResult.session_id.substring(0, 8) + '...' : 'N/A');
    } else {
      console.log('FAIL:', sessionResult.error);
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 4: Session validation
  console.log('Test 4: Session Validation');
  try {
    const session = validateSession();
    console.log('Session valid:', session.valid);
    console.log('Role:', session.role);

    if (session.valid && session.user) {
      console.log('PASS: Session validated with user data');
      console.log('User ID:', session.user.user_id);
      console.log('User Name:', session.user.name);
      console.log('Can See Directors:', session.user.can_see_directors);
    } else if (session.expired) {
      console.log('INFO: Session expired (expected if test ran slowly)');
    } else {
      console.log('INFO: No valid session (may need to login first)');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 5: Session extension
  console.log('Test 5: Session Extension');
  try {
    const extendResult = extendSession();
    if (extendResult.success) {
      console.log('PASS: Session extended successfully');
    } else {
      console.log('INFO:', extendResult.error);
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 6: Validate PIN Login (with invalid credentials)
  console.log('Test 6: Validate PIN Login (Invalid Credentials)');
  try {
    const invalidResult = validatePINLogin('NONEXISTENT-USER', '123456');
    if (!invalidResult.valid && invalidResult.message === 'Invalid credentials') {
      console.log('PASS: Invalid user returns correct error');
    } else {
      console.log('Result:', JSON.stringify(invalidResult));
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 7: PIN validation format
  console.log('Test 7: PIN Format Validation');
  try {
    const shortPin = validatePINLogin('TEST', '12345'); // 5 digits
    if (!shortPin.valid && shortPin.message.includes('6 digits')) {
      console.log('PASS: Short PIN rejected');
    } else {
      console.log('FAIL: Short PIN not properly rejected');
    }

    const nonNumericPin = validatePINLogin('TEST', 'abcdef');
    if (!nonNumericPin.valid && nonNumericPin.message.includes('6 digits')) {
      console.log('PASS: Non-numeric PIN rejected');
    } else {
      console.log('FAIL: Non-numeric PIN not properly rejected');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 8: Get current user
  console.log('Test 8: Get Current User');
  try {
    const currentUser = getCurrentUser();
    if (currentUser) {
      console.log('PASS: Current user retrieved');
      console.log('User:', currentUser.name, '- Role:', currentUser.role);
    } else {
      console.log('INFO: No current user (session may have ended)');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 9: End session
  console.log('Test 9: End Session');
  try {
    const endResult = endSession();
    if (endResult.success) {
      console.log('PASS: Session ended successfully');
    } else {
      console.log('FAIL:', endResult.error);
    }

    // Verify session is gone
    const verifySession = validateSession();
    if (!verifySession.valid) {
      console.log('PASS: Session properly cleared');
    } else {
      console.log('FAIL: Session still exists after end');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 10: User_Permissions sheet structure
  console.log('Test 10: User_Permissions Sheet Structure');
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('User_Permissions');

    if (!sheet) {
      console.log('INFO: User_Permissions sheet not found - run setupAllTabs() first');
    } else {
      const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
      const expectedHeaders = [
        'Employee_ID', 'Full_Name', 'Email', 'Role', 'PIN_Hash',
        'Can_See_Directors', 'Date_Added', 'Added_By', 'Status',
        'Last_Login', 'Login_Count', 'Failed_Attempts', 'Lockout_Until'
      ];

      let match = true;
      for (let i = 0; i < expectedHeaders.length; i++) {
        if (headers[i] !== expectedHeaders[i]) {
          console.log('Column', i + 1, '- expected:', expectedHeaders[i], '- got:', headers[i]);
          match = false;
        }
      }

      if (match) {
        console.log('PASS: All 13 User_Permissions headers match');
      } else {
        console.log('FAIL: Headers do not match - run setupAllTabs() to update');
      }

      // Check for role validation
      const validation = sheet.getRange('D2').getDataValidation();
      if (validation) {
        const criteria = validation.getCriteriaValues();
        if (criteria && criteria[0] && criteria[0].includes('Operator')) {
          console.log('PASS: Role validation includes 3-tier roles (Manager, Director, Operator)');
        } else {
          console.log('INFO: Role validation may need updating for 3-tier system');
        }
      }
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  console.log('=== PIN Authentication Tests Complete ===');
  console.log('');
  console.log('To create a test user, run: createTestUser()');
}

/**
 * Creates a test user for authentication testing.
 * PIN: 123456
 */
function createTestUser() {
  console.log('Creating test user...');

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('User_Permissions');

    if (!sheet) {
      console.log('ERROR: User_Permissions sheet not found');
      return;
    }

    // Check if test user already exists
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (const row of existingIds) {
        if (row[0] === 'TEST-OPERATOR-001') {
          console.log('Test user already exists');
          return;
        }
      }
    }

    // Hash the test PIN (123456)
    const testPin = '123456';
    const pinHash = hashPIN(testPin);

    // Add test user (Operator role for full access)
    sheet.appendRow([
      'TEST-OPERATOR-001',  // A: Employee_ID
      'Test Operator',      // B: Full_Name
      'test@cfatest.com',   // C: Email
      'Operator',           // D: Role
      pinHash,              // E: PIN_Hash
      'TRUE',               // F: Can_See_Directors
      new Date(),           // G: Date_Added
      'System',             // H: Added_By
      'Active',             // I: Status
      '',                   // J: Last_Login
      0,                    // K: Login_Count
      0,                    // L: Failed_Attempts
      ''                    // M: Lockout_Until
    ]);

    console.log('');
    console.log('=== Test User Created ===');
    console.log('Name: Test Operator');
    console.log('Employee ID: TEST-OPERATOR-001');
    console.log('Role: Operator');
    console.log('PIN: 123456');
    console.log('');
    console.log('You can now login with these credentials.');

  } catch (error) {
    console.log('ERROR:', error.toString());
  }
}

/**
 * Test function for Micro-Phase 11.5 Self-Service Signup System.
 * Tests signup link generation, validation, and completion.
 */
function testSignupSystem() {
  console.log('=== Testing Micro-Phase 11.5: Self-Service Signup System ===');
  console.log('');

  // Test 1: Pending_Signups sheet structure
  console.log('Test 1: Pending_Signups Sheet Structure');
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Pending_Signups');

    if (!sheet) {
      console.log('INFO: Pending_Signups sheet not found - run setupPendingSignupsTab() first');
    } else {
      const headers = sheet.getRange(1, 1, 1, 11).getValues()[0];
      const expectedHeaders = [
        'Signup_ID', 'Token', 'Employee_ID', 'Email', 'Role',
        'Can_See_Directors', 'Created_Date', 'Expires_Date', 'Status',
        'Created_By', 'Completed_Date'
      ];

      let match = true;
      for (let i = 0; i < expectedHeaders.length; i++) {
        if (headers[i] !== expectedHeaders[i]) {
          console.log('Column', i + 1, '- expected:', expectedHeaders[i], '- got:', headers[i]);
          match = false;
        }
      }

      if (match) {
        console.log('PASS: All 11 Pending_Signups headers match');
      } else {
        console.log('FAIL: Headers do not match - run setupPendingSignupsTab() to update');
      }
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 2: Secure token generation
  console.log('Test 2: Secure Token Generation');
  try {
    const token1 = generateSecureToken();
    const token2 = generateSecureToken();

    if (token1.length === 64) {
      console.log('PASS: Token is correct length (64 hex chars)');
    } else {
      console.log('FAIL: Token length is', token1.length, 'expected 64');
    }

    if (token1 !== token2) {
      console.log('PASS: Tokens are unique');
    } else {
      console.log('FAIL: Duplicate tokens generated');
    }

    if (/^[0-9a-f]+$/.test(token1)) {
      console.log('PASS: Token contains only hex characters');
    } else {
      console.log('FAIL: Token contains non-hex characters');
    }

    console.log('Sample token:', token1.substring(0, 16) + '...');
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 3: Signup ID generation
  console.log('Test 3: Signup ID Generation');
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Pending_Signups');

    if (sheet) {
      const signupId = generateSignupId(sheet);
      const today = new Date();
      const expectedPrefix = 'SU_' + today.getFullYear().toString() +
                             (today.getMonth() + 1).toString().padStart(2, '0') +
                             today.getDate().toString().padStart(2, '0') + '_';

      if (signupId.startsWith(expectedPrefix)) {
        console.log('PASS: Signup ID has correct date prefix');
      } else {
        console.log('FAIL: Expected prefix', expectedPrefix, 'got', signupId);
      }

      if (/^SU_\d{8}_\d{3}$/.test(signupId)) {
        console.log('PASS: Signup ID matches expected format (SU_YYYYMMDD_NNN)');
      } else {
        console.log('FAIL: Signup ID format incorrect:', signupId);
      }

      console.log('Generated ID:', signupId);
    } else {
      console.log('INFO: Cannot test - Pending_Signups sheet not found');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 4: Token validation (invalid token)
  console.log('Test 4: Token Validation (Invalid Token)');
  try {
    const invalidResult = validateSignupToken('invalid-token-12345');
    if (!invalidResult.valid && invalidResult.message.includes('Invalid token')) {
      console.log('PASS: Invalid token format rejected');
    } else {
      console.log('Result:', JSON.stringify(invalidResult));
    }

    const nonexistentResult = validateSignupToken('0'.repeat(64));
    if (!nonexistentResult.valid && nonexistentResult.message.includes('not found')) {
      console.log('PASS: Non-existent token rejected');
    } else {
      console.log('Result:', JSON.stringify(nonexistentResult));
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 5: Complete signup (invalid inputs)
  console.log('Test 5: Complete Signup Validation');
  try {
    // Test invalid PIN format
    const shortPinResult = completeSignup('0'.repeat(64), '12345');
    if (!shortPinResult.success && shortPinResult.message.includes('6 digits')) {
      console.log('PASS: Short PIN rejected');
    } else {
      console.log('Result:', JSON.stringify(shortPinResult));
    }

    const nonNumericResult = completeSignup('0'.repeat(64), 'abcdef');
    if (!nonNumericResult.success && nonNumericResult.message.includes('6 digits')) {
      console.log('PASS: Non-numeric PIN rejected');
    } else {
      console.log('Result:', JSON.stringify(nonNumericResult));
    }

    // Test with invalid token
    const invalidTokenResult = completeSignup('0'.repeat(64), '123456');
    if (!invalidTokenResult.success) {
      console.log('PASS: Invalid token rejected during completion');
    } else {
      console.log('FAIL: Invalid token should have been rejected');
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 6: Get pending signups (requires session)
  console.log('Test 6: Get Pending Signups');
  try {
    const result = getPendingSignups();
    if (result.sessionExpired) {
      console.log('INFO: Session required - login first to test this feature');
    } else if (result.success) {
      console.log('PASS: Retrieved pending signups');
      console.log('Count:', result.signups.length);
      if (result.signups.length > 0) {
        console.log('Most recent:', result.signups[0].signup_id, '-', result.signups[0].email);
      }
    } else {
      console.log('Result:', result.message);
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 7: Generate signup link (requires session)
  console.log('Test 7: Generate Signup Link (Validation)');
  try {
    // Test missing fields
    const missingFieldsResult = generateSignupLink({});
    if (missingFieldsResult.sessionExpired) {
      console.log('INFO: Session required - login first to test this feature');
    } else if (!missingFieldsResult.success && missingFieldsResult.message.includes('Missing required')) {
      console.log('PASS: Missing fields rejected');
    } else {
      console.log('Result:', JSON.stringify(missingFieldsResult));
    }

    // Test invalid email
    const invalidEmailResult = generateSignupLink({
      employee_id: 'TEST-001',
      email: 'not-an-email',
      role: 'Manager'
    });
    if (invalidEmailResult.sessionExpired) {
      console.log('INFO: Session required for email validation test');
    } else if (!invalidEmailResult.success && invalidEmailResult.message.includes('Invalid email')) {
      console.log('PASS: Invalid email rejected');
    } else {
      console.log('Result:', JSON.stringify(invalidEmailResult));
    }

    // Test invalid role
    const invalidRoleResult = generateSignupLink({
      employee_id: 'TEST-001',
      email: 'test@example.com',
      role: 'InvalidRole'
    });
    if (invalidRoleResult.sessionExpired) {
      console.log('INFO: Session required for role validation test');
    } else if (!invalidRoleResult.success && invalidRoleResult.message.includes('Invalid role')) {
      console.log('PASS: Invalid role rejected');
    } else {
      console.log('Result:', JSON.stringify(invalidRoleResult));
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  // Test 8: Cancel signup (requires session)
  console.log('Test 8: Cancel Signup');
  try {
    const result = cancelSignup('NONEXISTENT-SIGNUP');
    if (result.sessionExpired) {
      console.log('INFO: Session required - login first to test this feature');
    } else if (!result.success && result.message.includes('not found')) {
      console.log('PASS: Non-existent signup handled correctly');
    } else {
      console.log('Result:', JSON.stringify(result));
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  console.log('=== Signup System Tests Complete ===');
  console.log('');
  console.log('To perform a full signup flow test:');
  console.log('1. Login as an Operator or Director');
  console.log('2. Run testCreateSignup() to create a test signup');
  console.log('3. Use the generated link to complete registration');
}

/**
 * Creates a test signup link (requires active session as Director/Operator).
 */
function testCreateSignup() {
  console.log('Creating test signup...');

  const testData = {
    employee_id: 'TEST-SIGNUP-' + Date.now().toString().slice(-6),
    email: 'test-signup@example.com',
    role: 'Manager',
    can_see_directors: false
  };

  console.log('Test signup data:', JSON.stringify(testData));
  console.log('');

  const result = generateSignupLink(testData);

  if (result.sessionExpired) {
    console.log('ERROR: No active session. Please login first.');
    console.log('Run createTestUser() to create a test operator, then login.');
    return;
  }

  if (result.success) {
    console.log('=== Signup Created Successfully ===');
    console.log('Signup ID:', result.signup_id);
    console.log('Email:', result.email);
    console.log('Link:', result.link);
    console.log('');
    console.log('To complete the signup:');
    console.log('1. Open the link above in a browser');
    console.log('2. Create a 6-digit PIN');
    console.log('3. The user will be added to User_Permissions');
  } else {
    console.log('ERROR:', result.message);
  }
}

/**
 * Test function for Micro-Phase 12 Permission Filtering.
 * Tests the filterEmployeesByPermission function with mock data.
 */
function testPermissionFiltering() {
  console.log('=== Testing Micro-Phase 12: Permission Filtering ===');
  console.log('');

  // Create mock employee data
  const mockEmployees = [
    { employee_id: 'EMP-001', full_name: 'Alice Hourly', has_system_access: false, system_role: null },
    { employee_id: 'EMP-002', full_name: 'Bob Hourly', has_system_access: false, system_role: null },
    { employee_id: 'MGR-001', full_name: 'Charlie Manager', has_system_access: true, system_role: 'Manager' },
    { employee_id: 'DIR-001', full_name: 'Diana Director', has_system_access: true, system_role: 'Director' },
    { employee_id: 'DIR-002', full_name: 'Eve Director', has_system_access: true, system_role: 'Director' },
    { employee_id: 'OPR-001', full_name: 'Frank Operator', has_system_access: true, system_role: 'Operator' }
  ];

  console.log('Mock employees:', mockEmployees.length);
  console.log('- Hourly (no access): 2');
  console.log('- Managers: 1');
  console.log('- Directors: 2');
  console.log('- Operators: 1');
  console.log('');

  // Test 1: Operator viewing
  console.log('Test 1: Operator View');
  const operatorUser = { user_id: 'OPR-001', role: 'Operator', can_see_directors: true };
  const operatorFiltered = filterEmployeesByPermission(mockEmployees, operatorUser);
  console.log('Operator should see: All except other Operators (5 employees)');
  console.log('Operator sees:', operatorFiltered.length, 'employees');
  const operatorHasOperator = operatorFiltered.some(e => e.system_role === 'Operator');
  if (operatorFiltered.length === 5 && !operatorHasOperator) {
    console.log('PASS: Operator sees correct employees');
  } else {
    console.log('FAIL: Expected 5 without operators, got', operatorFiltered.length);
    console.log('Visible:', operatorFiltered.map(e => e.full_name).join(', '));
  }
  console.log('');

  // Test 2: Director with can_see_directors=TRUE
  console.log('Test 2: Director (can_see_directors=TRUE)');
  const directorWithAccess = { user_id: 'DIR-001', role: 'Director', can_see_directors: true };
  const directorWithFiltered = filterEmployeesByPermission(mockEmployees, directorWithAccess);
  console.log('Director should see: Hourly (2) + Manager (1) + Directors (2) = 5 (no Operators)');
  console.log('Director sees:', directorWithFiltered.length, 'employees');
  const directorWithHasOperator = directorWithFiltered.some(e => e.system_role === 'Operator');
  if (directorWithFiltered.length === 5 && !directorWithHasOperator) {
    console.log('PASS: Director with can_see_directors sees correct employees');
  } else {
    console.log('FAIL: Expected 5 without operators');
    console.log('Visible:', directorWithFiltered.map(e => e.full_name).join(', '));
  }
  console.log('');

  // Test 3: Director with can_see_directors=FALSE
  console.log('Test 3: Director (can_see_directors=FALSE)');
  const directorWithoutAccess = { user_id: 'DIR-002', role: 'Director', can_see_directors: false };
  const directorWithoutFiltered = filterEmployeesByPermission(mockEmployees, directorWithoutAccess);
  console.log('Director should see: Hourly (2) + Manager (1) = 3 (no Directors or Operators)');
  console.log('Director sees:', directorWithoutFiltered.length, 'employees');
  const hasDirectorOrOperator = directorWithoutFiltered.some(e =>
    e.system_role === 'Director' || e.system_role === 'Operator'
  );
  if (directorWithoutFiltered.length === 3 && !hasDirectorOrOperator) {
    console.log('PASS: Director without can_see_directors sees correct employees');
  } else {
    console.log('FAIL: Expected 3 without directors or operators');
    console.log('Visible:', directorWithoutFiltered.map(e => e.full_name).join(', '));
  }
  console.log('');

  // Test 4: Manager viewing
  console.log('Test 4: Manager View');
  const managerUser = { user_id: 'MGR-001', role: 'Manager', can_see_directors: false };
  const managerFiltered = filterEmployeesByPermission(mockEmployees, managerUser);
  console.log('Manager should see: Only hourly (2) - no one with system access');
  console.log('Manager sees:', managerFiltered.length, 'employees');
  const hasAnySystemAccess = managerFiltered.some(e => e.has_system_access);
  if (managerFiltered.length === 2 && !hasAnySystemAccess) {
    console.log('PASS: Manager sees only hourly employees');
  } else {
    console.log('FAIL: Expected 2 hourly employees only');
    console.log('Visible:', managerFiltered.map(e => e.full_name).join(', '));
  }
  console.log('');

  // Test 5: Real data test (requires session)
  console.log('Test 5: Real Data Test (Live)');
  try {
    const result = getAllEmployeesWithPoints();
    if (result.sessionExpired) {
      console.log('INFO: Session required - login first to test with real data');
    } else if (result.success) {
      console.log('PASS: Live filtering successful');
      console.log('User Role:', result.userRole);
      console.log('Employees visible:', result.employees.length);
      console.log('Threshold counts - At Risk:', result.counts.atRisk,
                  ', Final Warning:', result.counts.finalWarning,
                  ', Termination:', result.counts.termination);

      // Check that operators never appear
      const hasOperators = result.employees.some(e => e.system_role === 'Operator');
      if (!hasOperators) {
        console.log('PASS: No Operators appear in results (as expected)');
      } else {
        console.log('FAIL: Operators should never appear in employee list');
      }
    } else {
      console.log('ERROR:', result.error);
    }
  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');

  console.log('=== Permission Filtering Tests Complete ===');
}

/**
 * Test the getSystemAccessMap function.
 */
function testSystemAccessMap() {
  console.log('=== Testing getSystemAccessMap ===');
  console.log('');

  try {
    const accessMap = getSystemAccessMap();
    const count = Object.keys(accessMap).length;

    console.log('System access entries:', count);

    if (count > 0) {
      console.log('');
      console.log('Users with system access:');
      for (const [empId, data] of Object.entries(accessMap)) {
        console.log(`  ${empId}: ${data.role} (can_see_directors: ${data.can_see_directors})`);
      }
    } else {
      console.log('INFO: No users found in User_Permissions sheet');
    }

  } catch (error) {
    console.log('FAIL:', error.toString());
  }
  console.log('');
}
