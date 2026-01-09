/**
 * Payroll Review - Google Apps Script Backend
 * 
 * This file handles all server-side logic:
 * - Serving the web app
 * - Saving/retrieving data from Google Sheets
 * - Settings management
 * - PDF export
 * 
 * SETUP: This file goes in the Google Apps Script editor (Extensions > Apps Script)
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Sheet names - these will be created automatically if they don't exist
const SHEET_NAMES = {
  OT_HISTORY: 'OT_History',
  SETTINGS: 'Settings',
  EMPLOYEES: 'Employees',
  UNIFORM_CATALOG: 'Uniform_Catalog',
  UNIFORM_ORDERS: 'Uniform_Orders',
  UNIFORM_ORDER_ITEMS: 'Uniform_Order_Items',
  PTO: 'PTO',
  PAYROLL_SETTINGS: 'Payroll_Settings',
  SYSTEM_COUNTERS: 'System_Counters'
};

// Default settings
const DEFAULT_SETTINGS = {
  hourlyWage: 16.5,
  otMultiplier: 1.5,
  moderateThreshold: 5,
  highThreshold: 10,
  reallyHighThreshold: 15,
  // Location settings (1-3 locations supported)
  numberOfLocations: 2,
  location1Name: 'Cockrell Hill DTO',
  location2Name: 'DBU',
  location3Name: '',
  consecutiveAlertPeriods: 3,
  monthlyIncreaseAlertPercent: 25,
  // Alert thresholds
  pendingPTOAlertThreshold: 5,
  payrollUrgencyDays: 2,
  // Notification settings
  notificationsEnabled: false,
  adminEmails: '',
  notifyOnPTORequest: true,
  notifyOnHighOT: true,
  notifyOnPayrollDue: true,
  notifyOnUniformOrder: false,
  // Uniform settings
  paydayReference: '2024-11-29',  // Friday - bi-weekly pay dates
  weeklyEmailRecipients: ''
};

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================

/**
 * Serves the web app HTML
 * This is called automatically when someone visits the web app URL
 */
// Code version for debugging deployment issues
const CODE_VERSION = 'v2024.12.20.2';

function getCodeVersion() {
  return CODE_VERSION;
}

/**
 * Wrapper function to test if new code is deployed
 * Call this from console: google.script.run.withSuccessHandler(console.log).testNewDeployment()
 */
function testNewDeployment() {
  return {
    version: CODE_VERSION,
    timestamp: new Date().toISOString(),
    message: 'New deployment is working!'
  };
}

/**
 * Alternative function name to bypass deployment caching
 * This is identical to getEmployeesNeedingReview but with a new name
 */
function fetchEmployeesForReview() {
  return getEmployeesNeedingReview();
}

function doGet(e) {
  try {
    // Initialize sheets if they don't exist
    initializeSheets();
    
    // Check for URL parameters for routing
    const view = e && e.parameter && e.parameter.view ? e.parameter.view : 'default';
    
    // Public routes (no auth required)
    if (view === 'pto-request') {
      // Serve the employee PTO request form (public access via QR code)
      return HtmlService.createTemplateFromFile('EmployeePTORequest')
        .evaluate()
        .setTitle('Request Time Off / Solicitar Tiempo Libre')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
    }
    
    if (view === 'uniform-request') {
      // Serve the employee uniform request form (public access via QR code)
      return HtmlService.createTemplateFromFile('EmployeeUniformRequest')
        .evaluate()
        .setTitle('Request Uniforms / Solicitar Uniformes')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
    }
    
    // Admin routes - require Google authentication
    // Check for valid session (created after OAuth via google.script.run)
    const hasSession = checkAuthSession();
    const auth = isUserAuthenticated();
    
    if (!hasSession && !auth.authenticated) {
      // User is not signed in - show login prompt
      return HtmlService.createTemplateFromFile('LoginPrompt')
        .evaluate()
        .setTitle('Sign In - Payroll Review')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    // User is authenticated (either via session or direct auth) - serve admin dashboard
    return HtmlService.createTemplateFromFile('MainApp')
      .evaluate()
      .setTitle('Payroll Review')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    console.error('Error in doGet:', error);
    return HtmlService.createHtmlOutput('<h2>Error loading application</h2><p>' + error.message + '</p>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * Checks if the current user is authenticated with a Google account
 * Used by doGet() for initial page load check
 * @returns {Object} { authenticated: boolean, email: string|null }
 */
function isUserAuthenticated() {
  try {
    // Try to get the active user's email
    const email = Session.getActiveUser().getEmail();
    
    // If email is empty or null, user is not authenticated
    if (!email || email === '') {
      return { authenticated: false, email: null };
    }
    
    return { authenticated: true, email: email };
  } catch (error) {
    console.log('Auth check error:', error);
    return { authenticated: false, email: null };
  }
}

/**
 * Authenticates the user via google.script.run call
 * This triggers the OAuth consent flow when called from client-side
 * @returns {Object} { authenticated: boolean, email: string|null }
 */
function authenticateUser() {
  try {
    // When called via google.script.run, this triggers OAuth consent
    // and getActiveUser().getEmail() will return the user's email
    const email = Session.getActiveUser().getEmail();
    
    if (!email || email === '') {
      // Try effective user as fallback
      const effectiveEmail = Session.getEffectiveUser().getEmail();
      if (effectiveEmail && effectiveEmail !== '') {
        return { authenticated: true, email: effectiveEmail };
      }
      return { authenticated: false, email: null };
    }
    
    return { authenticated: true, email: email };
  } catch (error) {
    console.error('Authentication error:', error);
    return { authenticated: false, email: null };
  }
}

/**
 * Creates an authentication session for the user
 * Uses UserProperties which are unique per user (based on their temporary session key)
 * @param {string} email - The authenticated user's email
 */
function createAuthSession(email) {
  try {
    const userProps = PropertiesService.getUserProperties();
    const sessionData = {
      email: email,
      created: Date.now(),
      expires: Date.now() + (8 * 60 * 60 * 1000) // 8 hours
    };
    userProps.setProperty('auth_session', JSON.stringify(sessionData));
    return true;
  } catch (error) {
    console.error('Error creating auth session:', error);
    return false;
  }
}

/**
 * Checks if the user has a valid authentication session
 * @returns {boolean} True if session is valid
 */
function checkAuthSession() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const sessionStr = userProps.getProperty('auth_session');
    
    if (!sessionStr) {
      return false;
    }
    
    const session = JSON.parse(sessionStr);
    
    // Check if session is expired
    if (Date.now() > session.expires) {
      userProps.deleteProperty('auth_session');
      return false;
    }
    
    return true;
  } catch (error) {
    console.error('Error checking auth session:', error);
    return false;
  }
}

/**
 * Gets the current session email (for activity logging)
 * @returns {string|null} The session email or null
 */
function getSessionEmail() {
  try {
    // First try active user
    const activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail && activeEmail !== '') {
      return activeEmail;
    }
    
    // Fall back to session storage
    const userProps = PropertiesService.getUserProperties();
    const sessionStr = userProps.getProperty('auth_session');
    
    if (sessionStr) {
      const session = JSON.parse(sessionStr);
      if (Date.now() <= session.expires) {
        return session.email;
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * Clears the authentication session (logout)
 */
function clearAuthSession() {
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('auth_session');
    return true;
  } catch (error) {
    console.error('Error clearing auth session:', error);
    return false;
  }
}

/**
 * Returns HTML content for a view template
 * Used by the MainApp to load views dynamically
 * @param {string} templateName - Name of the template file (without .html)
 * @returns {string} HTML content
 */
function getViewTemplate(templateName) {
  try {
    // Security: only allow known template names
    const allowedTemplates = [
      'Dashboard',
      'OTUpload',
      'OTHistory',
      'OTEmployee',
      'OTTrends',
      'UniformOrders',
      'UniformDeductions',
      'UniformCatalog',
      'PTORecords',
      'PTOSummary',
      'PayrollProcessing',
      'PayrollCalendar',
      'UniformSummary',
      'SettingsPage',
      'SystemHealth',
      'Help',
      'YearEndWizard'
    ];
    
    if (!allowedTemplates.includes(templateName)) {
      throw new Error('Invalid template name: ' + templateName);
    }
    
    // Load the template (files are named View_TemplateName.html)
    const template = HtmlService.createTemplateFromFile('View_' + templateName);
    return template.evaluate().getContent();
  } catch (error) {
    console.error('Error loading template ' + templateName + ':', error);
    return '<div class="empty-state"><span class="material-icons-round">error_outline</span><h3>View Not Found</h3><p>' + error.message + '</p></div>';
  }
}

/**
 * Includes another HTML file - used for modular HTML/CSS/JS
 * @param {string} filename - Name of the file to include (without .html)
 * @returns {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
// INITIALIZATION
// ============================================================================

/**
 * Creates required sheets if they don't exist
 * Called automatically when the app loads
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create OT_History sheet if it doesn't exist
  let historySheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_NAMES.OT_HISTORY);
    // Add headers (19 columns: A-S, with Employee_ID as column S)
    const headers = [
      'Period End', 'Employee Name', 'Match Key', 'Location', 
      'CH Hours', 'DBU Hours', 'Total Hours', 'Regular Hours',
      'Week 1 Hours', 'Week 2 Hours', 'Week 1 OT', 'Week 2 OT',
      'Total OT', 'OT Cost', 'Flag', 'Is Multi-Location',
      'Import Date', 'Imported By', 'Employee_ID'
    ];
    historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    historySheet.setFrozenRows(1);
    historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    // Check if Employee_ID column exists (column S = 19), add if missing
    const lastCol = historySheet.getLastColumn();
    if (lastCol < 19) {
      historySheet.getRange(1, 19).setValue('Employee_ID');
      historySheet.getRange(1, 19).setFontWeight('bold');
    }
  }
  
  // Create Settings sheet if it doesn't exist
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    // Add default settings
    const settingsData = [
      ['Setting', 'Value'],
      ['hourlyWage', DEFAULT_SETTINGS.hourlyWage],
      ['otMultiplier', DEFAULT_SETTINGS.otMultiplier],
      ['moderateThreshold', DEFAULT_SETTINGS.moderateThreshold],
      ['highThreshold', DEFAULT_SETTINGS.highThreshold],
      ['reallyHighThreshold', DEFAULT_SETTINGS.reallyHighThreshold],
      ['numberOfLocations', DEFAULT_SETTINGS.numberOfLocations],
      ['location1Name', DEFAULT_SETTINGS.location1Name],
      ['location2Name', DEFAULT_SETTINGS.location2Name],
      ['location3Name', DEFAULT_SETTINGS.location3Name],
      ['consecutiveAlertPeriods', DEFAULT_SETTINGS.consecutiveAlertPeriods],
      ['monthlyIncreaseAlertPercent', DEFAULT_SETTINGS.monthlyIncreaseAlertPercent],
      ['notificationsEnabled', DEFAULT_SETTINGS.notificationsEnabled],
      ['adminEmails', DEFAULT_SETTINGS.adminEmails]
    ];
    settingsSheet.getRange(1, 1, settingsData.length, 2).setValues(settingsData);
    settingsSheet.setFrozenRows(1);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  
  // Create Employees sheet if it doesn't exist
  initializeEmployeesSheet();
  
  // Create Uniform sheets if they don't exist
  initializeUniformSheets();
  
  // Create PTO sheet if it doesn't exist
  initializePTOTab();
  
  // Create Payroll_Settings sheet if it doesn't exist
  initializePayrollSettings();
}

/**
 * Creates Employees sheet if it doesn't exist
 * Columns: Employee_ID, Full_Name, Match_Key, Primary_Location, Status, First_Seen, Last_Seen, Last_Period_End
 */
function initializeEmployeesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (!empSheet) {
    empSheet = ss.insertSheet(SHEET_NAMES.EMPLOYEES);
    const headers = [
      'Employee_ID', 'Full_Name', 'Match_Key', 'Primary_Location',
      'Status', 'First_Seen', 'Last_Seen', 'Last_Period_End'
    ];
    empSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    empSheet.setFrozenRows(1);
    empSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Set column widths for readability
    empSheet.setColumnWidth(1, 120); // Employee_ID
    empSheet.setColumnWidth(2, 180); // Full_Name
    empSheet.setColumnWidth(3, 180); // Match_Key
    empSheet.setColumnWidth(4, 140); // Primary_Location
    empSheet.setColumnWidth(5, 80);  // Status
    empSheet.setColumnWidth(6, 100); // First_Seen
    empSheet.setColumnWidth(7, 100); // Last_Seen
    empSheet.setColumnWidth(8, 120); // Last_Period_End
  }
  
  return empSheet;
}

/**
 * Creates Uniform-related sheets if they don't exist
 * - Uniform_Catalog: Master list of all uniform items
 * - Uniform_Orders: Order headers with payment tracking
 * - Uniform_Order_Items: Line items for each order
 */
function initializeUniformSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Uniform_Catalog sheet
  let catalogSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_CATALOG);
  if (!catalogSheet) {
    catalogSheet = ss.insertSheet(SHEET_NAMES.UNIFORM_CATALOG);
    const headers = ['Item_ID', 'Item_Name', 'Category', 'Available_Sizes', 'Price', 'Active'];
    catalogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    catalogSheet.setFrozenRows(1);
    catalogSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Set column widths
    catalogSheet.setColumnWidth(1, 100);  // Item_ID
    catalogSheet.setColumnWidth(2, 250);  // Item_Name
    catalogSheet.setColumnWidth(3, 150);  // Category
    catalogSheet.setColumnWidth(4, 200);  // Available_Sizes
    catalogSheet.setColumnWidth(5, 80);   // Price
    catalogSheet.setColumnWidth(6, 60);   // Active
    
    // Pre-populate with catalog items
    populateUniformCatalog(catalogSheet);
  }
  
  // Create Uniform_Orders sheet
  let ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet(SHEET_NAMES.UNIFORM_ORDERS);
    const headers = [
      'Order_ID', 'Employee_ID', 'Employee_Name', 'Location', 'Order_Date', 'Total_Amount',
      'Payment_Plan', 'Amount_Per_Paycheck', 'First_Deduction_Date',
      'Payments_Made', 'Amount_Paid', 'Amount_Remaining', 'Status', 'Notes',
      'Created_By', 'Created_Date', 'Received_Date'
    ];
    ordersSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    ordersSheet.setFrozenRows(1);
    ordersSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Set column widths
    ordersSheet.setColumnWidth(1, 130);  // Order_ID
    ordersSheet.setColumnWidth(2, 100);  // Employee_ID
    ordersSheet.setColumnWidth(3, 150);  // Employee_Name
    ordersSheet.setColumnWidth(4, 120);  // Location
    ordersSheet.setColumnWidth(5, 100);  // Order_Date
    ordersSheet.setColumnWidth(6, 100);  // Total_Amount
    ordersSheet.setColumnWidth(7, 100);  // Payment_Plan
    ordersSheet.setColumnWidth(8, 130);  // Amount_Per_Paycheck
    ordersSheet.setColumnWidth(9, 140);  // First_Deduction_Date
    ordersSheet.setColumnWidth(10, 110); // Payments_Made
    ordersSheet.setColumnWidth(11, 100); // Amount_Paid
    ordersSheet.setColumnWidth(12, 120); // Amount_Remaining
    ordersSheet.setColumnWidth(13, 90);  // Status
    ordersSheet.setColumnWidth(14, 200); // Notes
    ordersSheet.setColumnWidth(15, 120); // Created_By
    ordersSheet.setColumnWidth(16, 100); // Created_Date
    ordersSheet.setColumnWidth(17, 100); // Received_Date
  }
  
  // Create Uniform_Order_Items sheet
  let itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
  if (!itemsSheet) {
    itemsSheet = ss.insertSheet(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    const headers = ['Line_ID', 'Order_ID', 'Item_ID', 'Item_Name', 'Size', 'Quantity', 'Unit_Price', 'Line_Total', 'Is_Replacement'];
    itemsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    itemsSheet.setFrozenRows(1);
    itemsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Set column widths
    itemsSheet.setColumnWidth(1, 100);  // Line_ID
    itemsSheet.setColumnWidth(2, 130);  // Order_ID
    itemsSheet.setColumnWidth(3, 100);  // Item_ID
    itemsSheet.setColumnWidth(4, 250);  // Item_Name
    itemsSheet.setColumnWidth(5, 80);   // Size
    itemsSheet.setColumnWidth(6, 70);   // Quantity
    itemsSheet.setColumnWidth(7, 90);   // Unit_Price
    itemsSheet.setColumnWidth(8, 90);   // Line_Total
    itemsSheet.setColumnWidth(9, 100);  // Is_Replacement
  }
}

/**
 * Pre-populates the Uniform Catalog with all standard items
 */
function populateUniformCatalog(sheet) {
  const catalogItems = [
    // POLOS
    ['SHAE', 'Team Member Shae Polo', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 19.00, true],
    ['IVY', 'Ivy Green Dot Polo', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 27.00, true],
    ['SERRAMONTE', 'Serramonte Polo (Manager)', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 27.00, true],
    ['HOWELL', 'Howell Striped Polo', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 19.00, true],
    ['TOGETHER', 'Together Polo', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 40.00, true],
    ['LENEXA', 'Lenexa Team Lead Polo', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 28.00, true],
    ['KENWOOD', 'Kenwood Polo (Manager)', 'Polos', 'XS,S,M,L,XL,2XL,3XL', 28.00, true],
    
    // JACKETS & OUTERWEAR
    ['BLUEMOUND', 'Bluemound Soft Shell Jacket', 'Jackets & Outerwear', 'XS,S,M,L,XL,2XL,3XL', 53.25, true],
    ['HAYDEN-JKT', 'Hayden Jacket', 'Jackets & Outerwear', 'XS,S,M,L,XL,2XL,3XL', 71.00, true],
    ['NE8', 'Northeast 8 Jacket', 'Jackets & Outerwear', 'XS,S,M,L,XL,2XL,3XL', 72.75, true],
    ['TANASBOURNE', 'Tanasbourne Rain Jacket', 'Jackets & Outerwear', 'XS,S,M,L,XL,2XL,3XL', 70.00, true],
    ['HAYDEN-VEST', 'Hayden Vest', 'Jackets & Outerwear', 'XS,S,M,L,XL,2XL,3XL', 60.00, true],
    
    // FLEECES & PULLOVERS
    ['KNOX', 'Knox Fleece', 'Fleeces & Pullovers', 'XS,S,M,L,XL,2XL,3XL', 32.25, true],
    ['PALOMAR', 'Palomar Pullover', 'Fleeces & Pullovers', 'XS,S,M,L,XL,2XL,3XL', 30.00, true],
    ['DURANT', 'Durant Pullover', 'Fleeces & Pullovers', 'XS,S,M,L,XL,2XL,3XL', 39.95, true],
    ['CHAPEL', 'Chapel Fleece', 'Fleeces & Pullovers', 'XS,S,M,L,XL,2XL,3XL', 37.50, true],
    ['SPRINGFIELD', 'Springfield 3Q Sleeve Cardigan', 'Fleeces & Pullovers', 'XS,S,M,L,XL,2XL,3XL', 26.99, true],
    
    // BASE LAYERS
    ['BASETOP', 'Base Layer Top', 'Base Layers', 'XS,S,M,L,XL,2XL,3XL', 17.15, true],
    ['BASEBOTTOM', 'HW Base Layer Bottom', 'Base Layers', 'XS,S,M,L,XL,2XL,3XL', 18.25, true],
    
    // CHEF WEAR
    ['CHEFCOAT', 'Chef Coat', 'Chef Wear', 'XS,S,M,L,XL,2XL,3XL', 30.00, true],
    
    // PANTS
    ['PANTS', 'Pants', 'Pants', '24,26,28,30,32,34,36,38,40,42', 28.85, true],
    ['MATERNITY', 'Maternity Pelham Pants', 'Pants', 'XS,S,M,L,XL,2XL', 27.00, true],
    
    // SHORTS
    ['SHORTS', 'Shorts', 'Shorts', '24,26,28,30,32,34,36,38,40,42', 25.25, true],
    
    // SKIRTS
    ['SPARKS', 'Sparks Skirt', 'Skirts', 'XS,S,M,L,XL,2XL', 30.00, true],
    
    // ACCESSORIES (Gloves, Performance Gear, Aprons, Belts, Hair, Name Tags, Cooling Towels)
    ['SOFTSHELL-GLV', 'Softshell Tech Touch Glove', 'Accessories', 'S/M,L/XL', 19.75, true],
    ['FITTED-GLV', 'Fitted Tech Touch Gloves', 'Accessories', 'S/M,L/XL', 11.85, true],
    ['SLEEVES', 'Performance Sleeves', 'Accessories', 'S/M,L/XL', 10.75, true],
    ['CICERO', 'Cicero Apron', 'Accessories', 'One Size', 11.50, true],
    ['BELT', 'Belt', 'Accessories', 'S (28-32),M (32-36),L (36-40),XL (40-44)', 12.00, true],
    ['TAYLORS-BELT', 'Unisex Taylors Belt', 'Accessories', 'S (28-32),M (32-36),L (36-40),XL (40-44)', 12.00, true],
    ['SAVANNAH', 'Savannah Bow Ponytail Band', 'Accessories', 'One Size', 2.99, true],
    ['NAMETAG', 'Name Tag', 'Accessories', 'No Size', 5.00, true],
    ['HYDROCHILL-TWL', 'Hydrochill Cooling Towel', 'Accessories', 'One Size', 7.25, true],
    ['NEON', 'Neon Shirt', 'Other', 'S,M,L,XL,2XL', 7.00, true],
    
    // HATS & HEADWEAR
    ['POMPOM', 'Pom-Pom Beanie', 'Hats & Headwear', 'One Size', 13.50, true],
    ['BROOKFIELD', 'Brookfield Beanie', 'Hats & Headwear', 'One Size', 8.50, true],
    ['COMMACK', 'Commack Cable Knit Beanie', 'Hats & Headwear', 'One Size', 8.75, true],
    ['HYDROCHILL-CAP', 'Hydrochill Crown Cap', 'Hats & Headwear', 'One Size', 8.00, true],
    ['FLAGLER-VISOR', 'Flagler Visor', 'Hats & Headwear', 'One Size', 6.85, true],
    ['HIALEAH', 'Hialeah Headband', 'Hats & Headwear', 'One Size', 6.00, true],
    ['CHERRYDALE', 'Cherrydale Hat', 'Hats & Headwear', 'One Size', 19.25, true],
    ['FLAGLER-HAT', 'Flagler Hat', 'Hats & Headwear', 'One Size', 6.85, true],
    
    // BUNDLES
    ['WINTER-BUNDLE', 'Winter Bundle (Jacket + Hat + Gloves)', 'Bundles', 'Size Selection Required', 64.95, true],
    ['FIRST-UNIFORM', 'First Uniform (Store Paid)', 'Bundles', 'Size Selection Required', 0.00, true]
  ];
  
  if (catalogItems.length > 0) {
    sheet.getRange(2, 1, catalogItems.length, 6).setValues(catalogItems);
  }
}

/**
 * Manual initialization function - run this once after setup
 * You can run this from the Apps Script editor to test that everything works
 */
function manualInit() {
  initializeSheets();
  Logger.log('Sheets initialized successfully!');
  Logger.log('OT_History sheet: ' + (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.OT_HISTORY) ? 'EXISTS' : 'MISSING'));
  Logger.log('Settings sheet: ' + (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS) ? 'EXISTS' : 'MISSING'));
  Logger.log('Employees sheet: ' + (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.EMPLOYEES) ? 'EXISTS' : 'MISSING'));
  Logger.log('Uniform_Catalog sheet: ' + (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.UNIFORM_CATALOG) ? 'EXISTS' : 'MISSING'));
  Logger.log('Uniform_Orders sheet: ' + (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.UNIFORM_ORDERS) ? 'EXISTS' : 'MISSING'));
  Logger.log('Uniform_Order_Items sheet: ' + (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS) ? 'EXISTS' : 'MISSING'));
}

/**
 * One-time function to consolidate catalog categories into "Accessories"
 * Run this once from Apps Script editor to fix existing catalog
 */
function consolidateCatalogCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_CATALOG);
  
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('No catalog data to update');
    return;
  }
  
  const categoriesToMerge = ['Gloves', 'Performance Gear', 'Aprons', 'Belts', 'Hair Accessories', 'Other'];
  const newCategory = 'Accessories';
  
  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues(); // Column C = Category
  let updated = 0;
  
  for (let i = 0; i < data.length; i++) {
    const currentCategory = data[i][0];
    if (categoriesToMerge.includes(currentCategory)) {
      sheet.getRange(i + 2, 3).setValue(newCategory);
      updated++;
    }
  }
  
  Logger.log('Updated ' + updated + ' items to "Accessories" category');
}

/**
 * One-time migration function to add new columns to existing Uniform sheets
 * Run this from Apps Script editor if you have existing order data
 */
function migrateUniformSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Migrate Uniform_Orders sheet - add Location and Received_Date columns
  const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
  if (ordersSheet && ordersSheet.getLastColumn() < 17) {
    Logger.log('Migrating Uniform_Orders sheet...');
    
    // Check if we need to insert Location column (D)
    const headers = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
    
    if (!headers.includes('Location')) {
      // Insert column D for Location
      ordersSheet.insertColumnAfter(3);
      ordersSheet.getRange(1, 4).setValue('Location');
      ordersSheet.getRange(1, 4).setFontWeight('bold');
      ordersSheet.setColumnWidth(4, 120);
      Logger.log('Added Location column');
    }
    
    if (!headers.includes('Received_Date')) {
      // Add Received_Date column at the end
      const lastCol = ordersSheet.getLastColumn();
      ordersSheet.getRange(1, lastCol + 1).setValue('Received_Date');
      ordersSheet.getRange(1, lastCol + 1).setFontWeight('bold');
      ordersSheet.setColumnWidth(lastCol + 1, 100);
      Logger.log('Added Received_Date column');
    }
    
    // Update existing "Active" orders to stay Active (they were created before Pending workflow)
    // New orders will start as "Pending"
    Logger.log('Migration complete - existing Active orders remain Active');
  }
  
  // Migrate Uniform_Order_Items sheet - add Is_Replacement column
  const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
  if (itemsSheet && itemsSheet.getLastColumn() < 9) {
    Logger.log('Migrating Uniform_Order_Items sheet...');
    
    const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
    
    if (!headers.includes('Is_Replacement')) {
      const lastCol = itemsSheet.getLastColumn();
      itemsSheet.getRange(1, lastCol + 1).setValue('Is_Replacement');
      itemsSheet.getRange(1, lastCol + 1).setFontWeight('bold');
      itemsSheet.setColumnWidth(lastCol + 1, 100);
      
      // Set existing items to false (not replacement)
      if (itemsSheet.getLastRow() > 1) {
        const numRows = itemsSheet.getLastRow() - 1;
        const falseValues = Array(numRows).fill([false]);
        itemsSheet.getRange(2, lastCol + 1, numRows, 1).setValues(falseValues);
      }
      Logger.log('Added Is_Replacement column');
    }
  }
  
  // Add new catalog items if missing
  const catalogSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_CATALOG);
  if (catalogSheet) {
    const existingIds = catalogSheet.getRange(2, 1, Math.max(1, catalogSheet.getLastRow() - 1), 1).getValues().flat();
    
    const newItems = [
      ['SPRINGFIELD', 'Springfield 3Q Sleeve Cardigan', 'Fleeces & Pullovers', 'XS,S,M,L,XL,2XL,3XL', 26.99, true],
      ['TAYLORS-BELT', 'Unisex Taylors Belt', 'Accessories', 'S (28-32),M (32-36),L (36-40),XL (40-44)', 12.00, true],
      ['NEON', 'Neon Shirt', 'Other', 'S,M,L,XL,2XL', 7.00, true]
    ];
    
    for (const item of newItems) {
      if (!existingIds.includes(item[0])) {
        catalogSheet.appendRow(item);
        Logger.log('Added catalog item: ' + item[1]);
      }
    }
    
    // Update Belt sizes if needed
    for (let i = 2; i <= catalogSheet.getLastRow(); i++) {
      const itemId = catalogSheet.getRange(i, 1).getValue();
      if (itemId === 'BELT') {
        const currentSizes = catalogSheet.getRange(i, 4).getValue();
        if (currentSizes === 'S,M,L') {
          catalogSheet.getRange(i, 4).setValue('S (28-32),M (32-36),L (36-40),XL (40-44)');
          Logger.log('Updated Belt sizes');
        }
      }
    }
  }
  
  Logger.log('Migration complete!');
}

/**
 * Migration function to add receiving columns to Uniform_Order_Items
 * Adds: Item_Received, Item_Received_Date, Item_Received_By, Item_Status
 * Run this once from Apps Script editor
 */
function migrateUniformItemsForReceiving() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
  const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
  
  if (!itemsSheet) {
    Logger.log('Uniform_Order_Items sheet not found');
    return;
  }
  
  const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
  Logger.log('Current headers: ' + headers.join(', '));
  
  // Add new columns if they don't exist
  const newColumns = ['Item_Received', 'Item_Received_Date', 'Item_Received_By', 'Item_Status'];
  let columnsAdded = 0;
  
  for (const colName of newColumns) {
    if (!headers.includes(colName)) {
      const newCol = itemsSheet.getLastColumn() + 1;
      itemsSheet.getRange(1, newCol).setValue(colName);
      itemsSheet.getRange(1, newCol).setFontWeight('bold');
      itemsSheet.setColumnWidth(newCol, colName === 'Item_Received_By' ? 150 : 120);
      columnsAdded++;
      Logger.log('Added column: ' + colName);
    }
  }
  
  if (columnsAdded === 0) {
    Logger.log('All receiving columns already exist');
    return;
  }
  
  // Now populate existing items based on their order status
  if (itemsSheet.getLastRow() < 2) {
    Logger.log('No existing items to migrate');
    return;
  }
  
  // Get all orders to check their status
  const ordersData = ordersSheet ? ordersSheet.getRange(2, 1, Math.max(1, ordersSheet.getLastRow() - 1), 13).getValues() : [];
  const orderStatusMap = {};
  for (const row of ordersData) {
    const orderId = row[0];
    const status = row[12]; // Status column (M)
    if (orderId) {
      orderStatusMap[orderId] = status;
    }
  }
  
  // Get current column indices after adding new columns
  const updatedHeaders = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
  const receivedCol = updatedHeaders.indexOf('Item_Received') + 1;
  const receivedDateCol = updatedHeaders.indexOf('Item_Received_Date') + 1;
  const receivedByCol = updatedHeaders.indexOf('Item_Received_By') + 1;
  const statusCol = updatedHeaders.indexOf('Item_Status') + 1;
  
  // Iterate through all items and set values based on order status
  const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 2).getValues();
  
  for (let i = 0; i < itemsData.length; i++) {
    const rowNum = i + 2;
    const orderId = itemsData[i][1]; // Order_ID is column B (index 1)
    const orderStatus = orderStatusMap[orderId] || 'Pending';
    
    let itemReceived = false;
    let itemStatus = 'Pending';
    
    // Determine item status based on order status
    if (orderStatus === 'Active' || orderStatus === 'Completed' || orderStatus === 'Store Paid') {
      itemReceived = true;
      itemStatus = 'Received';
    } else if (orderStatus === 'Cancelled') {
      itemReceived = false;
      itemStatus = 'Cancelled';
    }
    // Pending orders keep items as Pending
    
    itemsSheet.getRange(rowNum, receivedCol).setValue(itemReceived);
    itemsSheet.getRange(rowNum, statusCol).setValue(itemStatus);
    // Leave Item_Received_Date and Item_Received_By blank for existing items
  }
  
  Logger.log('Migration complete! Processed ' + itemsData.length + ' items');
}

/**
 * Adds manager receiving passcode to Payroll_Settings if not exists
 * Run this once from Apps Script editor
 */
function addManagerReceivingPasscode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Payroll_Settings');
  
  if (!sheet) {
    Logger.log('Payroll_Settings sheet not found');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  // Check if setting already exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'manager_receiving_passcode') {
      Logger.log('Manager receiving passcode already exists');
      return;
    }
  }
  
  // Add the setting
  const now = new Date();
  sheet.appendRow(['manager_receiving_passcode', '03177', 'Passcode for manager receiving section on uniform request form', now]);
  Logger.log('Added manager_receiving_passcode setting with default value 03177');
}

// ============================================================================
// MANAGER RECEIVING PASSCODE FUNCTIONS
// ============================================================================

/**
 * Gets the manager receiving passcode
 * @returns {string} The passcode
 */
function getManagerReceivingPasscode() {
  const passcode = getPayrollSetting('manager_receiving_passcode');
  return passcode || '03177'; // Default fallback
}

/**
 * Validates the manager receiving passcode
 * @param {string} inputCode - The code entered by the user
 * @returns {boolean} True if valid
 */
function validateManagerPasscode(inputCode) {
  const correctCode = getManagerReceivingPasscode();
  
  // Normalize both values: trim whitespace and convert to string for comparison
  const normalizedInput = String(inputCode || '').trim();
  const normalizedCorrect = String(correctCode || '').trim();
  
  // Direct string comparison first
  if (normalizedInput === normalizedCorrect) {
    return true;
  }
  
  // Handle case where spreadsheet stored number without leading zeros
  // E.g., user types "03177" but spreadsheet has 3177 (as number)
  // Compare numeric values if both are numeric
  const inputNum = parseInt(normalizedInput, 10);
  const correctNum = parseInt(normalizedCorrect, 10);
  
  if (!isNaN(inputNum) && !isNaN(correctNum) && inputNum === correctNum) {
    return true;
  }
  
  return false;
}

/**
 * Updates the manager receiving passcode
 * @param {string} newPasscode - The new passcode
 * @returns {Object} Result
 */
function setManagerReceivingPasscode(newPasscode) {
  if (!newPasscode || newPasscode.length < 4) {
    return { success: false, error: 'Passcode must be at least 4 characters' };
  }
  
  const result = updatePayrollSetting('manager_receiving_passcode', newPasscode);
  
  if (result.success) {
    logActivity('UPDATE', 'SETTINGS', 'Updated manager receiving passcode');
  }
  
  return result;
}

// ============================================================================
// SETTINGS FUNCTIONS
// ============================================================================

/**
 * Gets all settings from the Settings sheet
 * @returns {Object} Settings object
 */
function getSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      return DEFAULT_SETTINGS;
    }
    
    const data = sheet.getDataRange().getValues();
    const settings = {};
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      let value = data[i][1];
      
      // Convert numeric values
      if (['hourlyWage', 'otMultiplier', 'moderateThreshold', 'highThreshold', 'reallyHighThreshold',
           'consecutiveAlertPeriods', 'monthlyIncreaseAlertPercent', 'pendingPTOAlertThreshold',
           'payrollUrgencyDays', 'numberOfLocations'].includes(key)) {
        value = parseFloat(value) || DEFAULT_SETTINGS[key];
      }
      
      // Convert boolean values
      if (['notificationsEnabled', 'notifyOnPTORequest', 'notifyOnHighOT', 
           'notifyOnPayrollDue', 'notifyOnUniformOrder'].includes(key)) {
        value = value === true || value === 'true' || value === 'TRUE';
      }
      
      settings[key] = value;
    }
    
    // Fill in any missing settings with defaults
    for (const key in DEFAULT_SETTINGS) {
      if (settings[key] === undefined) {
        settings[key] = DEFAULT_SETTINGS[key];
      }
    }
    
    // Also get the next payroll date from Payroll_Settings
    try {
      const payrollSheet = ss.getSheetByName('Payroll_Settings');
      if (payrollSheet) {
        const payrollData = payrollSheet.getDataRange().getValues();
        for (let i = 1; i < payrollData.length; i++) {
          if (payrollData[i][0] === 'next_payroll_date' && payrollData[i][1]) {
            settings.nextPayrollDate = payrollData[i][1];
            break;
          }
        }
      }
    } catch (e) {
      console.log('Could not get payroll date:', e);
    }
    
    return settings;
  } catch (error) {
    console.error('Error getting settings:', error);
    return DEFAULT_SETTINGS;
  }
}

/**
 * Gets the list of active/configured locations
 * Returns only locations that are configured (1-3 based on numberOfLocations setting)
 * @returns {Array} Array of location objects with id and name
 */
function getActiveLocations() {
  try {
    const settings = getSettings();
    const count = Math.min(Math.max(parseInt(settings.numberOfLocations) || 2, 1), 3);
    const locations = [];
    
    if (count >= 1 && settings.location1Name) {
      locations.push({
        id: 'location1',
        name: settings.location1Name
      });
    }
    
    if (count >= 2 && settings.location2Name) {
      locations.push({
        id: 'location2',
        name: settings.location2Name
      });
    }
    
    if (count >= 3 && settings.location3Name) {
      locations.push({
        id: 'location3',
        name: settings.location3Name
      });
    }
    
    // If no locations configured, return a default
    if (locations.length === 0) {
      locations.push({
        id: 'location1',
        name: 'Main Location'
      });
    }
    
    return locations;
  } catch (error) {
    console.error('Error getting active locations:', error);
    return [{ id: 'location1', name: 'Main Location' }];
  }
}

/**
 * Gets just the location names as a simple array
 * Useful for dropdown options
 * @returns {Array} Array of location name strings
 */
function getLocationNames() {
  const locations = getActiveLocations();
  return locations.map(loc => loc.name);
}

/**
 * Updates settings in the Settings sheet
 * @param {Object} newSettings - Object with setting keys and values to update
 * @returns {Object} Result object
 */
function updateSettings(newSettings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Update each setting
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      if (newSettings[key] !== undefined) {
        sheet.getRange(i + 1, 2).setValue(newSettings[key]);
      }
    }
    
    return { success: true };
  } catch (error) {
    console.error('Error updating settings:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Saves all settings (updates existing and adds new)
 * @param {Object} newSettings - Object with all setting keys and values
 * @returns {Object} Result object
 */
function saveSettings(newSettings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    
    const data = sheet.getDataRange().getValues();
    const existingKeys = new Set();
    
    // Update existing settings
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      existingKeys.add(key);
      if (newSettings[key] !== undefined) {
        sheet.getRange(i + 1, 2).setValue(newSettings[key]);
      }
    }
    
    // Add new settings that don't exist yet
    const newRows = [];
    for (const key in newSettings) {
      if (!existingKeys.has(key) && newSettings[key] !== undefined) {
        newRows.push([key, newSettings[key]]);
      }
    }
    
    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 2).setValues(newRows);
    }
    
    // Log the settings change
    logActivity('UPDATE', 'SETTINGS', 'Settings updated', '');
    
    return { success: true };
  } catch (error) {
    console.error('Error saving settings:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// UNIFORM SIZE CONFIGURATION
// =====================================================

/**
 * Default size configuration for pants
 * Used if no configuration exists in Settings
 */
const DEFAULT_SIZE_CONFIG = {
  pants_male_waist: ['28', '30', '32', '34', '36', '38', '40', '42', '44', '46', '48', '50', '52', '54', '56', '58', '60'],
  pants_male_inseam: ['28', '30', '32', '34', '36'],
  pants_female_waist: ['00', '0', '2', '4', '6', '8', '10', '12', '14', '16', '18', '20', '22', '24', '26', '28', '30', '32', '34', '36'],
  pants_female_inseam: ['29', '31', '33', '35'],
  shorts_waist: ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL'],
  shorts_inseam: ['30', '32', '34']
};

/**
 * Gets the uniform size configuration
 * @returns {Object} Size configuration object
 */
function getSizeConfiguration() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      return { success: true, config: DEFAULT_SIZE_CONFIG };
    }
    
    const data = sheet.getDataRange().getValues();
    const config = {};
    
    // Look for size configuration keys
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      
      if (key && key.startsWith('sizes_')) {
        // Remove 'sizes_' prefix for the config key
        const configKey = key.replace('sizes_', '');
        try {
          // Parse JSON array
          config[configKey] = JSON.parse(value);
        } catch (e) {
          // If not valid JSON, try comma-separated
          config[configKey] = value.split(',').map(s => s.trim()).filter(s => s);
        }
      }
    }
    
    // Merge with defaults for any missing keys
    for (const key in DEFAULT_SIZE_CONFIG) {
      if (!config[key] || config[key].length === 0) {
        config[key] = DEFAULT_SIZE_CONFIG[key];
      }
    }
    
    return { success: true, config: config };
  } catch (error) {
    console.error('Error getting size configuration:', error);
    return { success: false, error: error.message, config: DEFAULT_SIZE_CONFIG };
  }
}

/**
 * Saves the uniform size configuration
 * @param {Object} config - Size configuration object with arrays of sizes
 * @returns {Object} Result object
 */
function saveSizeConfiguration(config) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    }
    
    const data = sheet.getDataRange().getValues();
    const existingRows = {};
    
    // Find existing size config rows
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      if (key && key.startsWith('sizes_')) {
        existingRows[key] = i + 1; // 1-indexed row number
      }
    }
    
    // Update or add each config key
    for (const key in config) {
      const settingsKey = 'sizes_' + key;
      const value = JSON.stringify(config[key]);
      
      if (existingRows[settingsKey]) {
        // Update existing row
        sheet.getRange(existingRows[settingsKey], 2).setValue(value);
      } else {
        // Add new row
        sheet.appendRow([settingsKey, value]);
      }
    }
    
    // Log the change
    logActivity('UPDATE', 'SETTINGS', 'Uniform size configuration updated', '');
    
    return { success: true };
  } catch (error) {
    console.error('Error saving size configuration:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Ensures all location settings exist in the Settings sheet
 * Run this manually if dynamic locations aren't working
 */
function ensureLocationSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    
    if (!sheet) {
      console.log('Settings sheet not found');
      return { success: false, error: 'Settings sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const existingKeys = new Set();
    
    // Get existing keys
    for (let i = 1; i < data.length; i++) {
      existingKeys.add(data[i][0]);
    }
    
    // Settings that must exist for dynamic locations
    const requiredSettings = {
      'numberOfLocations': DEFAULT_SETTINGS.numberOfLocations,
      'location1Name': DEFAULT_SETTINGS.location1Name,
      'location2Name': DEFAULT_SETTINGS.location2Name,
      'location3Name': DEFAULT_SETTINGS.location3Name
    };
    
    // Add missing settings
    const newRows = [];
    for (const key in requiredSettings) {
      if (!existingKeys.has(key)) {
        newRows.push([key, requiredSettings[key]]);
        console.log('Adding missing setting:', key);
      }
    }
    
    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 2).setValues(newRows);
      console.log('Added', newRows.length, 'missing location settings');
      return { success: true, added: newRows.length };
    } else {
      console.log('All location settings already exist');
      return { success: true, added: 0, message: 'All settings already exist' };
    }
  } catch (error) {
    console.error('Error ensuring location settings:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// EMPLOYEE MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Upserts a batch of employees (creates new or updates existing)
 * @param {Array} employees - Array of employee objects with: employeeId, displayName, matchKey, location
 * @param {Date} periodEnd - The period end date for Last_Period_End
 * @returns {Object} Result with counts
 */
function upsertEmployeesBatch(employees, periodEnd) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!empSheet) {
      empSheet = initializeEmployeesSheet();
    }
    
    const now = new Date();
    const periodDate = new Date(periodEnd);
    
    // Get existing employees
    const lastRow = empSheet.getLastRow();
    let existingData = [];
    let existingMap = new Map(); // matchKey -> row index (1-based)
    
    if (lastRow > 1) {
      existingData = empSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      existingData.forEach((row, idx) => {
        const matchKey = row[2]; // Column C = Match_Key
        if (matchKey) {
          existingMap.set(matchKey.toLowerCase(), idx + 2); // 1-based row number
        }
      });
    }
    
    let created = 0;
    let updated = 0;
    const newRows = [];
    const updateBatches = [];
    
    for (const emp of employees) {
      if (!emp.matchKey) continue;
      
      const matchKeyLower = emp.matchKey.toLowerCase();
      const existingRow = existingMap.get(matchKeyLower);
      
      if (existingRow) {
        // Update existing employee
        // Get current data for this row
        const currentData = existingData[existingRow - 2];
        const currentStatus = currentData[4] || 'Active';
        
        // Update: Last_Seen, Last_Period_End, Status (reactivate if was inactive)
        // Also update Employee_ID if we now have it and didn't before
        const currentEmployeeId = currentData[0];
        const newEmployeeId = emp.employeeId || currentEmployeeId || '';
        
        updateBatches.push({
          row: existingRow,
          values: [
            newEmployeeId,                    // A: Employee_ID (update if we have new one)
            emp.displayName,                  // B: Full_Name (might have changed format)
            emp.matchKey,                     // C: Match_Key
            emp.location || currentData[3],   // D: Primary_Location
            'Active',                         // E: Status (reactivate)
            currentData[5],                   // F: First_Seen (keep original)
            now,                              // G: Last_Seen
            periodDate                        // H: Last_Period_End
          ]
        });
        updated++;
      } else {
        // Create new employee
        newRows.push([
          emp.employeeId || '',    // A: Employee_ID
          emp.displayName,         // B: Full_Name
          emp.matchKey,            // C: Match_Key
          emp.location || '',      // D: Primary_Location
          'Active',                // E: Status
          now,                     // F: First_Seen
          now,                     // G: Last_Seen
          periodDate               // H: Last_Period_End
        ]);
        created++;
        
        // Add to existingMap to prevent duplicates in same batch
        existingMap.set(matchKeyLower, -1);
      }
    }
    
    // Batch update existing employees
    for (const update of updateBatches) {
      empSheet.getRange(update.row, 1, 1, 8).setValues([update.values]);
    }
    
    // Batch insert new employees
    if (newRows.length > 0) {
      const insertRow = empSheet.getLastRow() + 1;
      empSheet.getRange(insertRow, 1, newRows.length, 8).setValues(newRows);
    }
    
    console.log(`Employees upserted: ${created} created, ${updated} updated`);
    
    return { success: true, created, updated };
    
  } catch (error) {
    console.error('Error upserting employees:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets all employees with optional filters
 * @param {Object} filters - Optional: { status: 'Active'|'Inactive', location: string }
 * @returns {Array} Array of employee objects
 */
function getEmployees(filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!empSheet || empSheet.getLastRow() < 2) {
      return [];
    }
    
    const data = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 8).getValues();
    
    let employees = data.map(row => ({
      employeeId: row[0] || '',
      fullName: row[1] || '',
      matchKey: row[2] || '',
      primaryLocation: row[3] || '',
      status: row[4] || 'Active',
      firstSeen: row[5] ? new Date(row[5]).toISOString() : null,
      lastSeen: row[6] ? new Date(row[6]).toISOString() : null,
      lastPeriodEnd: row[7] ? new Date(row[7]).toISOString() : null
    })).filter(e => e.matchKey); // Filter out empty rows
    
    // Apply filters
    if (filters.status) {
      employees = employees.filter(e => e.status === filters.status);
    }
    
    if (filters.location) {
      employees = employees.filter(e => e.primaryLocation === filters.location);
    }
    
    // Sort by name
    employees.sort((a, b) => a.fullName.localeCompare(b.fullName));
    
    return employees;
    
  } catch (error) {
    console.error('Error getting employees:', error);
    return [];
  }
}

/**
 * Gets a single employee by their Employee_ID
 * @param {string} employeeId - The Employee_ID to look up
 * @returns {Object|null} Employee object or null if not found
 */
function getEmployeeById(employeeId) {
  try {
    const employees = getEmployees();
    return employees.find(e => e.employeeId === employeeId) || null;
  } catch (error) {
    console.error('Error getting employee by ID:', error);
    return null;
  }
}

/**
 * Gets a single employee by their Match_Key (lowercase name)
 * @param {string} matchKey - The Match_Key to look up
 * @returns {Object|null} Employee object or null if not found
 */
function getEmployeeByMatchKey(matchKey) {
  try {
    const employees = getEmployees();
    const keyLower = matchKey.toLowerCase();
    return employees.find(e => e.matchKey.toLowerCase() === keyLower) || null;
  } catch (error) {
    console.error('Error getting employee by match key:', error);
    return null;
  }
}

/**
 * Updates an employee's status
 * @param {string} employeeId - The Employee_ID
 * @param {string} status - 'Active' or 'Inactive'
 * @returns {Object} Result
 */
function updateEmployeeStatus(employeeId, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!empSheet || empSheet.getLastRow() < 2) {
      return { success: false, error: 'No employees found' };
    }
    
    const data = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        empSheet.getRange(i + 2, 5).setValue(status); // Column E = Status
        return { success: true };
      }
    }
    
    return { success: false, error: 'Employee not found' };
    
  } catch (error) {
    console.error('Error updating employee status:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// NEW EMPLOYEE CREATION & SEARCH (with Safeguards)
// ============================================================================

/**
 * Normalizes a name - converts "Last, First" to "First Last" and cleans whitespace
 * @param {string} name - Raw name input
 * @returns {string} Normalized name
 */
function normalizeEmployeeName(name) {
  if (!name) return '';
  
  let normalized = name.trim();
  
  // Convert "Last, First" to "First Last"
  if (normalized.includes(',')) {
    const parts = normalized.split(',').map(p => p.trim());
    if (parts.length >= 2) {
      normalized = `${parts[1]} ${parts[0]}`;
    }
  }
  
  // Collapse multiple spaces to single space
  normalized = normalized.replace(/\s+/g, ' ');
  
  // Proper case (first letter of each word uppercase)
  normalized = normalized.split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');
  
  return normalized;
}

/**
 * Generates a match key from a name (lowercase, single spaces)
 * @param {string} name - The name to convert
 * @returns {string} Match key
 */
function generateMatchKey(name) {
  const normalized = normalizeEmployeeName(name);
  return normalized.toLowerCase().replace(/\s+/g, ' ');
}

/**
 * Calculates Levenshtein distance between two strings
 * @param {string} str1 
 * @param {string} str2 
 * @returns {number} Edit distance
 */
function levenshteinDistance(str1, str2) {
  const m = str1.length;
  const n = str2.length;
  
  // Create matrix
  const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));
  
  // Initialize first row and column
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  
  // Fill matrix
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
  }
  
  return dp[m][n];
}

/**
 * Calculates similarity percentage between two names
 * @param {string} name1 
 * @param {string} name2 
 * @returns {number} Similarity percentage (0-100)
 */
function calculateNameSimilarity(name1, name2) {
  const s1 = name1.toLowerCase().trim();
  const s2 = name2.toLowerCase().trim();
  
  if (s1 === s2) return 100;
  
  const distance = levenshteinDistance(s1, s2);
  const maxLen = Math.max(s1.length, s2.length);
  
  if (maxLen === 0) return 100;
  
  const similarity = Math.round((1 - distance / maxLen) * 100);
  return Math.max(0, similarity);
}

/**
 * Gets all employees from both OT_History (recent) and Employees sheet (all active)
 * Used for uniform order dropdowns
 * @returns {Array} Combined employee list with duplicates removed
 */
function getAllEmployeesForUniform() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSet = new Map(); // matchKey -> employee object
    
    // Source 1: Recent employees from OT_History (last 4 pay periods)
    const historySheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    if (historySheet && historySheet.getLastRow() >= 2) {
      const periods = getPayPeriods();
      const recentPeriods = periods.slice(0, 4);
      const recentPeriodTimes = new Set(recentPeriods.map(p => new Date(p).getTime()));
      
      const historyData = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 7).getValues();
      
      for (const row of historyData) {
        const periodEnd = row[0];
        const name = row[1];
        const matchKey = (row[2] || '').toLowerCase();
        const location = row[3] || '';
        const totalHours = parseFloat(row[6]) || 0;
        
        if (!name || !periodEnd || !matchKey) continue;
        
        const periodTime = new Date(periodEnd).getTime();
        if (recentPeriodTimes.has(periodTime) && totalHours > 0) {
          if (!employeeSet.has(matchKey)) {
            employeeSet.set(matchKey, {
              fullName: name,
              matchKey: matchKey,
              location: location,
              source: 'recent'
            });
          }
        }
      }
    }
    
    // Source 2: All active employees from Employees sheet
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    if (empSheet && empSheet.getLastRow() >= 2) {
      const empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 5).getValues();
      
      for (const row of empData) {
        const fullName = row[1] || '';
        const matchKey = (row[2] || '').toLowerCase();
        const location = row[3] || '';
        const status = row[4] || 'Active';
        
        if (!fullName || !matchKey || status !== 'Active') continue;
        
        if (!employeeSet.has(matchKey)) {
          employeeSet.set(matchKey, {
            fullName: fullName,
            matchKey: matchKey,
            location: location,
            source: 'employees'
          });
        }
      }
    }
    
    // Convert to array and sort
    const employees = Array.from(employeeSet.values());
    employees.sort((a, b) => a.fullName.localeCompare(b.fullName));
    
    return employees;
    
  } catch (error) {
    console.error('Error getting employees for uniform:', error);
    return [];
  }
}

/**
 * Searches for employees with fuzzy matching
 * @param {string} searchTerm - The name to search for
 * @returns {Object} { exact: [], similar: [], exactMatch: boolean }
 */
function searchEmployeesFuzzy(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim().length < 2) {
      return { exact: [], similar: [], exactMatch: false };
    }
    
    const normalizedSearch = normalizeEmployeeName(searchTerm);
    const searchKey = generateMatchKey(searchTerm);
    const allEmployees = getAllEmployeesForUniform();
    
    const exact = [];
    const similar = [];
    let exactMatch = false;
    
    for (const emp of allEmployees) {
      const empKey = emp.matchKey.toLowerCase();
      const empName = emp.fullName.toLowerCase();
      const searchLower = searchKey.toLowerCase();
      
      // Check for exact match
      if (empKey === searchLower) {
        exact.push(emp);
        exactMatch = true;
        continue;
      }
      
      // Check for partial/contains match
      if (empName.includes(searchLower) || searchLower.includes(empName)) {
        exact.push(emp);
        continue;
      }
      
      // Check for fuzzy similarity
      const similarity = calculateNameSimilarity(empKey, searchLower);
      if (similarity >= 70) {
        similar.push({
          ...emp,
          similarity: similarity
        });
      }
    }
    
    // Sort similar by similarity descending
    similar.sort((a, b) => b.similarity - a.similarity);
    
    return {
      exact: exact.slice(0, 10),     // Limit to 10 exact/partial matches
      similar: similar.slice(0, 5),  // Limit to 5 similar matches
      exactMatch: exactMatch
    };
    
  } catch (error) {
    console.error('Error in fuzzy search:', error);
    return { exact: [], similar: [], exactMatch: false };
  }
}

/**
 * Creates a new employee record with safeguards
 * @param {Object} employeeData - { fullName, location, createdBy, creationMethod }
 * @returns {Object} { success, employee, warnings }
 */
function addNewEmployee(employeeData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!empSheet) {
      empSheet = initializeEmployeesSheet();
    }
    
    // Normalize the name
    const fullName = normalizeEmployeeName(employeeData.fullName);
    const matchKey = generateMatchKey(fullName);
    
    if (!fullName || fullName.length < 2) {
      return { success: false, error: 'Invalid employee name' };
    }
    
    // Check for exact duplicate
    const existing = getEmployeeByMatchKey(matchKey);
    if (existing) {
      return { 
        success: false, 
        error: `Employee "${existing.fullName}" already exists with this name`,
        existingEmployee: existing
      };
    }
    
    // Check for similar names (warnings only, don't block)
    const searchResults = searchEmployeesFuzzy(fullName);
    const warnings = searchResults.similar.map(emp => ({
      name: emp.fullName,
      similarity: emp.similarity
    }));
    
    // Generate Employee ID
    const employeeId = generateEmployeeId();
    const now = new Date();
    
    // Add to Employees sheet
    const newRow = [
      employeeId,                          // A: Employee_ID
      fullName,                            // B: Full_Name
      matchKey,                            // C: Match_Key
      employeeData.location || '',         // D: Primary_Location
      'Active',                            // E: Status
      now,                                 // F: First_Seen
      now,                                 // G: Last_Seen
      null                                 // H: Last_Period_End (not worked yet)
    ];
    
    empSheet.appendRow(newRow);
    
    // Log the creation
    logEmployeeCreation({
      employeeId: employeeId,
      fullName: fullName,
      matchKey: matchKey,
      location: employeeData.location || '',
      createdBy: employeeData.createdBy || Session.getActiveUser().getEmail() || 'Unknown',
      creationMethod: employeeData.creationMethod || 'Manual',
      similarWarnings: warnings.map(w => w.name).join(', ')
    });
    
    console.log(`New employee created: ${fullName} (${employeeId}) via ${employeeData.creationMethod}`);
    
    return {
      success: true,
      employee: {
        employeeId: employeeId,
        fullName: fullName,
        matchKey: matchKey,
        location: employeeData.location || ''
      },
      warnings: warnings
    };
    
  } catch (error) {
    console.error('Error adding new employee:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Generates a unique Employee ID
 * @returns {string} New employee ID (EMP-XXXXX format)
 */
function generateEmployeeId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let counterSheet = ss.getSheetByName(SHEET_NAMES.SYSTEM_COUNTERS);
  
  if (!counterSheet) {
    counterSheet = ss.insertSheet(SHEET_NAMES.SYSTEM_COUNTERS);
    counterSheet.getRange(1, 1, 1, 2).setValues([['Counter', 'Value']]);
    counterSheet.getRange(2, 1, 1, 2).setValues([['employee_id', 1000]]);
    counterSheet.setFrozenRows(1);
  }
  
  // Get current counter
  const data = counterSheet.getRange(2, 1, counterSheet.getLastRow() - 1, 2).getValues();
  let counterRow = -1;
  let currentValue = 1000;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'employee_id') {
      counterRow = i + 2;
      currentValue = parseInt(data[i][1]) || 1000;
      break;
    }
  }
  
  // Increment counter
  const newValue = currentValue + 1;
  
  if (counterRow > 0) {
    counterSheet.getRange(counterRow, 2).setValue(newValue);
  } else {
    counterSheet.appendRow(['employee_id', newValue]);
  }
  
  return `EMP-${String(newValue).padStart(5, '0')}`;
}

/**
 * Logs employee creation to audit trail
 * @param {Object} data - Creation data
 */
function logEmployeeCreation(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Employee_Audit_Log');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('Employee_Audit_Log');
      const headers = [
        'Timestamp', 'Employee_ID', 'Employee_Name', 'Match_Key', 
        'Location', 'Created_By', 'Creation_Method', 'Similar_Warnings',
        'Reviewed', 'Reviewed_By', 'Review_Date'
      ];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.setFrozenRows(1);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    const row = [
      new Date(),                           // A: Timestamp
      data.employeeId || '',               // B: Employee_ID
      data.fullName || '',                 // C: Employee_Name
      data.matchKey || '',                 // D: Match_Key
      data.location || '',                 // E: Location
      data.createdBy || '',                // F: Created_By
      data.creationMethod || '',           // G: Creation_Method
      data.similarWarnings || '',          // H: Similar_Warnings
      false,                               // I: Reviewed
      '',                                  // J: Reviewed_By
      ''                                   // K: Review_Date
    ];
    
    logSheet.appendRow(row);
    
  } catch (error) {
    console.error('Error logging employee creation:', error);
  }
}

/**
 * Scans for potential duplicate employees based on name similarity
 * @returns {Array} List of potential duplicate pairs
 */
function scanForDuplicateEmployees() {
  try {
    const allEmployees = getAllEmployeesForUniform();
    const duplicates = [];
    
    // Compare each pair of employees
    for (let i = 0; i < allEmployees.length; i++) {
      for (let j = i + 1; j < allEmployees.length; j++) {
        const emp1 = allEmployees[i];
        const emp2 = allEmployees[j];
        
        const similarity = calculateNameSimilarity(emp1.matchKey, emp2.matchKey);
        
        // Flag pairs with >80% similarity (but not exact matches)
        if (similarity >= 80 && similarity < 100) {
          duplicates.push({
            employee1: {
              fullName: emp1.fullName,
              matchKey: emp1.matchKey,
              location: emp1.location
            },
            employee2: {
              fullName: emp2.fullName,
              matchKey: emp2.matchKey,
              location: emp2.location
            },
            similarity: similarity
          });
        }
      }
    }
    
    // Sort by similarity descending
    duplicates.sort((a, b) => b.similarity - a.similarity);
    
    return duplicates;
    
  } catch (error) {
    console.error('Error scanning for duplicates:', error);
    return [];
  }
}

/**
 * Gets preview of what will be affected by a merge
 * @param {string} sourceKey - Match key of employee to merge FROM
 * @param {string} targetKey - Match key of employee to merge INTO
 * @returns {Object} Preview data
 */
function getMergePreview(sourceKey, targetKey) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceKeyLower = sourceKey.toLowerCase();
    const targetKeyLower = targetKey.toLowerCase();
    
    // Get employee details
    const sourceEmp = getEmployeeByMatchKey(sourceKey);
    const targetEmp = getEmployeeByMatchKey(targetKey);
    
    if (!sourceEmp || !targetEmp) {
      return { success: false, error: 'One or both employees not found' };
    }
    
    // Count affected records
    let uniformOrdersAffected = 0;
    let ptoRecordsAffected = 0;
    let otHistoryAffected = 0;
    
    // Check Uniform_Orders
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    if (ordersSheet && ordersSheet.getLastRow() >= 2) {
      const ordersData = ordersSheet.getRange(2, 3, ordersSheet.getLastRow() - 1, 1).getValues();
      uniformOrdersAffected = ordersData.filter(row => 
        (row[0] || '').toLowerCase().replace(/\s+/g, ' ') === sourceKeyLower
      ).length;
    }
    
    // Check PTO
    const ptoSheet = ss.getSheetByName(SHEET_NAMES.PTO);
    if (ptoSheet && ptoSheet.getLastRow() >= 2) {
      const ptoData = ptoSheet.getRange(2, 2, ptoSheet.getLastRow() - 1, 1).getValues();
      ptoRecordsAffected = ptoData.filter(row =>
        (row[0] || '').toLowerCase().replace(/\s+/g, ' ') === sourceKeyLower
      ).length;
    }
    
    // Check OT_History
    const historySheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    if (historySheet && historySheet.getLastRow() >= 2) {
      const historyData = historySheet.getRange(2, 3, historySheet.getLastRow() - 1, 1).getValues();
      otHistoryAffected = historyData.filter(row =>
        (row[0] || '').toLowerCase() === sourceKeyLower
      ).length;
    }
    
    return {
      success: true,
      source: sourceEmp,
      target: targetEmp,
      uniformOrdersAffected: uniformOrdersAffected,
      ptoRecordsAffected: ptoRecordsAffected,
      otHistoryAffected: otHistoryAffected
    };
    
  } catch (error) {
    console.error('Error getting merge preview:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Merges two employee records - moves all data from source to target
 * @param {string} sourceKey - Match key of employee to merge FROM (will be deleted)
 * @param {string} targetKey - Match key of employee to merge INTO (will be kept)
 * @param {string} keepName - 'source' or 'target' - which name to keep
 * @returns {Object} Result with counts
 */
function mergeEmployees(sourceKey, targetKey, keepName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceKeyLower = sourceKey.toLowerCase();
    const targetKeyLower = targetKey.toLowerCase();
    
    if (sourceKeyLower === targetKeyLower) {
      return { success: false, error: 'Cannot merge employee with themselves' };
    }
    
    // Get employee details
    const sourceEmp = getEmployeeByMatchKey(sourceKey);
    const targetEmp = getEmployeeByMatchKey(targetKey);
    
    if (!sourceEmp || !targetEmp) {
      return { success: false, error: 'One or both employees not found' };
    }
    
    // Determine final name
    const finalName = keepName === 'source' ? sourceEmp.fullName : targetEmp.fullName;
    const finalMatchKey = keepName === 'source' ? sourceEmp.matchKey : targetEmp.matchKey;
    
    let uniformOrdersUpdated = 0;
    let ptoRecordsUpdated = 0;
    let otHistoryUpdated = 0;
    
    // Update Uniform_Orders (Column C = Employee_Name)
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    if (ordersSheet && ordersSheet.getLastRow() >= 2) {
      const ordersData = ordersSheet.getRange(2, 3, ordersSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ordersData.length; i++) {
        const empName = (ordersData[i][0] || '').toLowerCase().replace(/\s+/g, ' ');
        if (empName === sourceKeyLower) {
          ordersSheet.getRange(i + 2, 3).setValue(finalName);
          uniformOrdersUpdated++;
        }
      }
    }
    
    // Update PTO (Column B = Employee_Name, need to check Match_Key too)
    const ptoSheet = ss.getSheetByName(SHEET_NAMES.PTO);
    if (ptoSheet && ptoSheet.getLastRow() >= 2) {
      const ptoData = ptoSheet.getRange(2, 2, ptoSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ptoData.length; i++) {
        const empName = (ptoData[i][0] || '').toLowerCase().replace(/\s+/g, ' ');
        if (empName === sourceKeyLower) {
          ptoSheet.getRange(i + 2, 2).setValue(finalName);
          ptoRecordsUpdated++;
        }
      }
    }
    
    // Update OT_History (Column B = Name, Column C = Match_Key)
    const historySheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    if (historySheet && historySheet.getLastRow() >= 2) {
      const historyData = historySheet.getRange(2, 2, historySheet.getLastRow() - 1, 2).getValues();
      for (let i = 0; i < historyData.length; i++) {
        const matchKey = (historyData[i][1] || '').toLowerCase();
        if (matchKey === sourceKeyLower) {
          historySheet.getRange(i + 2, 2).setValue(finalName);       // Name
          historySheet.getRange(i + 2, 3).setValue(finalMatchKey);   // Match_Key
          otHistoryUpdated++;
        }
      }
    }
    
    // Delete source employee from Employees sheet
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    if (empSheet && empSheet.getLastRow() >= 2) {
      const empData = empSheet.getRange(2, 3, empSheet.getLastRow() - 1, 1).getValues();
      for (let i = empData.length - 1; i >= 0; i--) {
        const matchKey = (empData[i][0] || '').toLowerCase();
        if (matchKey === sourceKeyLower) {
          empSheet.deleteRow(i + 2);
          break;
        }
      }
    }
    
    // Update target employee name if keeping source name
    if (keepName === 'source') {
      const empData = empSheet.getRange(2, 2, empSheet.getLastRow() - 1, 2).getValues();
      for (let i = 0; i < empData.length; i++) {
        const matchKey = (empData[i][1] || '').toLowerCase();
        if (matchKey === targetKeyLower) {
          empSheet.getRange(i + 2, 2).setValue(finalName);
          empSheet.getRange(i + 2, 3).setValue(finalMatchKey);
          break;
        }
      }
    }
    
    // Log the merge action
    logEmployeeCreation({
      employeeId: 'MERGE',
      fullName: `${sourceEmp.fullName} → ${targetEmp.fullName}`,
      matchKey: finalMatchKey,
      location: '',
      createdBy: Session.getActiveUser().getEmail() || 'Unknown',
      creationMethod: 'Merge',
      similarWarnings: `Kept: ${finalName}`
    });
    
    console.log(`Merged "${sourceEmp.fullName}" into "${targetEmp.fullName}". Updated: ${uniformOrdersUpdated} orders, ${ptoRecordsUpdated} PTO, ${otHistoryUpdated} OT records`);
    
    return {
      success: true,
      message: `Successfully merged employees`,
      uniformOrdersUpdated: uniformOrdersUpdated,
      ptoRecordsUpdated: ptoRecordsUpdated,
      otHistoryUpdated: otHistoryUpdated,
      finalName: finalName
    };
    
  } catch (error) {
    console.error('Error merging employees:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets employees that were created via uniform orders and need review
 * @returns {Array} Employees needing review
 */
function getEmployeesNeedingReview() {
  // Wrap entire function in explicit try-catch with logging
  try {
    console.log('getEmployeesNeedingReview: Starting...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('getEmployeesNeedingReview: No active spreadsheet!');
      return [];
    }
    
    const logSheet = ss.getSheetByName('Employee_Audit_Log');
    
    if (!logSheet) {
      console.log('Employee_Audit_Log: Sheet not found');
      return [];
    }
    
    const lastRow = logSheet.getLastRow();
    console.log('Employee_Audit_Log: lastRow = ' + lastRow);
    
    if (lastRow < 2) {
      console.log('Employee_Audit_Log: No data rows');
      return [];
    }
    
    const data = logSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const needsReview = [];
    
    console.log('Employee_Audit_Log: Found ' + data.length + ' rows');
    
    for (let i = 0; i < data.length; i++) {
      const creationMethod = (data[i][6] || '').toString().trim();
      const reviewedRaw = data[i][8];
      const reviewed = reviewedRaw === true || reviewedRaw === 'TRUE' || reviewedRaw === 'true';
      
      console.log('Row ' + (i+2) + ': Method="' + creationMethod + '" | Reviewed=' + reviewedRaw + ' (' + typeof reviewedRaw + ')');
      
      // Only show uniform order created employees that haven't been reviewed
      const isUniformOrder = creationMethod.toLowerCase().includes('uniform');
      
      if (isUniformOrder && !reviewed) {
        // Convert Date to string to avoid serialization issues
        const timestampStr = data[i][0] ? new Date(data[i][0]).toISOString() : '';
        
        needsReview.push({
          rowIndex: i + 2,
          timestamp: timestampStr,
          employeeId: String(data[i][1] || ''),
          fullName: String(data[i][2] || ''),
          matchKey: String(data[i][3] || ''),
          location: String(data[i][4] || ''),
          createdBy: String(data[i][5] || ''),
          similarWarnings: String(data[i][7] || '')
        });
      }
    }
    
    console.log('Employees needing review: ' + needsReview.length);
    
    // Explicit return - ensure it's a plain array with no circular refs
    const result = JSON.parse(JSON.stringify(needsReview));
    console.log('getEmployeesNeedingReview: Returning ' + result.length + ' items');
    return result;
    
  } catch (error) {
    console.error('getEmployeesNeedingReview ERROR:', error.toString());
    console.error('Stack:', error.stack);
    // Return empty array on error, not null
    return [];
  }
}

/**
 * DEBUG: Test function to verify getEmployeesNeedingReview works
 * Run this directly from Apps Script editor to see output in Logs
 */
function debugEmployeesNeedingReview() {
  const result = getEmployeesNeedingReview();
  Logger.log('=== DEBUG: Employees Needing Review ===');
  Logger.log('Total count: ' + result.length);
  Logger.log('Full result: ' + JSON.stringify(result, null, 2));
  
  // Also log raw sheet data for comparison
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Employee_Audit_Log');
  if (logSheet && logSheet.getLastRow() > 1) {
    const rawData = logSheet.getRange(1, 1, logSheet.getLastRow(), 11).getValues();
    Logger.log('=== Raw Sheet Data ===');
    Logger.log('Headers: ' + JSON.stringify(rawData[0]));
    for (let i = 1; i < rawData.length; i++) {
      Logger.log('Row ' + (i+1) + ': ' + JSON.stringify({
        timestamp: rawData[i][0],
        employeeId: rawData[i][1],
        name: rawData[i][2],
        creationMethod: rawData[i][6],
        reviewed: rawData[i][8],
        reviewedType: typeof rawData[i][8]
      }));
    }
  }
  
  return result;
}

/**
 * Marks an employee creation as reviewed
 * @param {number} rowIndex - Row in audit log
 * @returns {Object} Result
 */
function markEmployeeReviewed(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Employee_Audit_Log');
    
    if (!logSheet) {
      return { success: false, error: 'Audit log not found' };
    }
    
    logSheet.getRange(rowIndex, 9).setValue(true);  // Reviewed
    logSheet.getRange(rowIndex, 10).setValue(Session.getActiveUser().getEmail() || 'Unknown');  // Reviewed_By
    logSheet.getRange(rowIndex, 11).setValue(new Date());  // Review_Date
    
    return { success: true };
    
  } catch (error) {
    console.error('Error marking employee reviewed:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets recent employee audit log entries
 * @param {number} limit - Max entries to return
 * @returns {Array} Audit log entries
 */
function getEmployeeAuditLog(limit) {
  try {
    limit = limit || 50;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Employee_Audit_Log');
    
    if (!logSheet || logSheet.getLastRow() < 2) {
      return [];
    }
    
    const lastRow = logSheet.getLastRow();
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    
    const data = logSheet.getRange(startRow, 1, numRows, 11).getValues();
    const entries = [];
    
    // Read in reverse order (newest first)
    for (let i = data.length - 1; i >= 0; i--) {
      entries.push({
        timestamp: data[i][0],
        employeeId: data[i][1],
        fullName: data[i][2],
        matchKey: data[i][3],
        location: data[i][4],
        createdBy: data[i][5],
        creationMethod: data[i][6],
        similarWarnings: data[i][7],
        reviewed: data[i][8] === true || data[i][8] === 'TRUE',
        reviewedBy: data[i][9],
        reviewDate: data[i][10]
      });
    }
    
    return entries;
    
  } catch (error) {
    console.error('Error getting audit log:', error);
    return [];
  }
}

/**
 * Updates employee statuses based on the 28-day absence rule
 * Employees who haven't appeared in 28+ days (2 pay periods) are marked Inactive
 * Called after each OT save
 * @param {Date} currentPeriodEnd - The period that was just saved
 * @param {Set} activeMatchKeys - Set of match keys for employees who appeared in this save
 */
function updateEmployeeStatuses(currentPeriodEnd, activeMatchKeys) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!empSheet || empSheet.getLastRow() < 2) {
      return;
    }
    
    const periodDate = new Date(currentPeriodEnd);
    const cutoffDate = new Date(periodDate);
    cutoffDate.setDate(cutoffDate.getDate() - 28); // 28 days = 2 pay periods
    
    const data = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 8).getValues();
    const updates = [];
    
    for (let i = 0; i < data.length; i++) {
      const matchKey = (data[i][2] || '').toLowerCase();
      const currentStatus = data[i][4] || 'Active';
      const lastPeriodEnd = data[i][7] ? new Date(data[i][7]) : null;
      
      // Skip if already inactive (don't repeatedly mark inactive)
      if (currentStatus === 'Inactive') continue;
      
      // Check if this employee appeared in the current save
      if (activeMatchKeys.has(matchKey)) {
        // Employee is active - already updated in upsertEmployeesBatch
        continue;
      }
      
      // Employee did NOT appear in this save
      // Check if their Last_Period_End is more than 28 days before current period
      if (lastPeriodEnd && lastPeriodEnd < cutoffDate) {
        // Mark as inactive
        updates.push({ row: i + 2, status: 'Inactive' });
        console.log(`Marking ${data[i][1]} as Inactive (last seen: ${lastPeriodEnd.toDateString()})`);
      }
    }
    
    // Apply updates
    for (const update of updates) {
      empSheet.getRange(update.row, 5).setValue(update.status);
    }
    
    if (updates.length > 0) {
      console.log(`Updated ${updates.length} employees to Inactive status`);
    }
    
  } catch (error) {
    console.error('Error updating employee statuses:', error);
  }
}

/**
 * Backfills Employee_ID in OT_History for records that are missing it
 * Matches by Match_Key to find the Employee_ID from the Employees sheet
 * Run manually from Apps Script editor or automatically after saves
 * @returns {Object} Result with count of updated records
 */
function backfillEmployeeIds() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!historySheet || !empSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    // Build a map of matchKey -> employeeId from Employees sheet
    const empLastRow = empSheet.getLastRow();
    if (empLastRow < 2) {
      return { success: true, updated: 0, message: 'No employees to match against' };
    }
    
    const empData = empSheet.getRange(2, 1, empLastRow - 1, 3).getValues(); // A, B, C
    const matchKeyToId = new Map();
    
    for (const row of empData) {
      const employeeId = row[0];
      const matchKey = (row[2] || '').toLowerCase();
      if (employeeId && matchKey) {
        matchKeyToId.set(matchKey, employeeId);
      }
    }
    
    if (matchKeyToId.size === 0) {
      return { success: true, updated: 0, message: 'No employee IDs to backfill with' };
    }
    
    // Scan OT_History for records missing Employee_ID (column S = 19)
    const historyLastRow = historySheet.getLastRow();
    if (historyLastRow < 2) {
      return { success: true, updated: 0, message: 'No OT history records' };
    }
    
    // Get columns C (Match_Key) and S (Employee_ID)
    const matchKeys = historySheet.getRange(2, 3, historyLastRow - 1, 1).getValues();
    const employeeIds = historySheet.getRange(2, 19, historyLastRow - 1, 1).getValues();
    
    let updated = 0;
    const updates = [];
    
    for (let i = 0; i < matchKeys.length; i++) {
      const currentId = employeeIds[i][0];
      const matchKey = (matchKeys[i][0] || '').toLowerCase();
      
      // Skip if already has an ID
      if (currentId) continue;
      
      // Look up the employee ID
      const foundId = matchKeyToId.get(matchKey);
      if (foundId) {
        updates.push({ row: i + 2, id: foundId });
        updated++;
      }
    }
    
    // Batch update
    for (const update of updates) {
      historySheet.getRange(update.row, 19).setValue(update.id);
    }
    
    console.log(`Backfilled ${updated} Employee_IDs in OT_History`);
    
    return { success: true, updated, message: `Updated ${updated} records` };
    
  } catch (error) {
    console.error('Error backfilling employee IDs:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// UNIFORM MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Gets all catalog items with optional filters
 * @param {boolean} activeOnly - If true, only returns active items
 * @returns {Array} Array of catalog item objects
 */
function getCatalogItems(activeOnly = true) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_CATALOG);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    // Get up to 8 columns (added Waist_Sizes and Inseam_Sizes for pants)
    const numCols = Math.min(sheet.getLastColumn(), 8);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
    
    let items = data.map((row, index) => ({
      itemId: row[0] || '',
      itemName: row[1] || '',
      category: row[2] || '',
      availableSizes: row[3] ? String(row[3]).split(',').map(s => s.trim()) : [],
      price: parseFloat(row[4]) || 0,
      active: row[5] === true || row[5] === 'TRUE' || row[5] === true,
      // Columns 7 & 8: Optional waist/inseam overrides for pants
      waistSizes: row[6] ? String(row[6]).split(',').map(s => s.trim()) : [],
      inseamSizes: row[7] ? String(row[7]).split(',').map(s => s.trim()) : [],
      rowIndex: index + 2
    })).filter(item => item.itemId);
    
    if (activeOnly) {
      items = items.filter(item => item.active);
    }
    
    return items;
    
  } catch (error) {
    console.error('Error getting catalog items:', error);
    return [];
  }
}

/**
 * Gets catalog items formatted for the employee uniform request form
 * Returns item name and category only (no prices shown to employees)
 * @returns {Array} Array of catalog item objects
 */
function getUniformCatalog() {
  try {
    const items = getCatalogItems(true);
    // Return data for employee form including prices for payment preview
    // Also include waist/inseam overrides for pants
    return items.map(item => ({
      itemName: item.itemName,
      category: item.category,
      itemId: item.itemId,
      price: item.price,
      // Include waist/inseam overrides (empty arrays mean use defaults from Settings)
      waistSizes: item.waistSizes || [],
      inseamSizes: item.inseamSizes || []
    }));
  } catch (error) {
    console.error('Error getting uniform catalog:', error);
    return [];
  }
}

/**
 * Submits a uniform request from the employee form
 * Creates order as Pending with default payment plan (1 check)
 * @param {Object} orderData - { employeeName, location, items[], notes, source }
 * @returns {Object} Result with order ID
 */
function submitEmployeeUniformRequest(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || !itemsSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    // Handle new employee creation if flagged
    if (orderData.isNewEmployee && orderData.employeeName) {
      const empResult = addNewEmployee({
        fullName: orderData.employeeName,
        location: orderData.location || '',
        createdBy: 'Employee Self-Service',
        creationMethod: 'Uniform Order'
      });
      
      if (!empResult.success) {
        // If employee already exists, use them
        if (empResult.existingEmployee) {
          orderData.employeeId = empResult.existingEmployee.matchKey;
          orderData.employeeName = empResult.existingEmployee.fullName;
        } else {
          return { success: false, error: 'Failed to create employee: ' + empResult.error };
        }
      } else {
        // Use normalized name from creation
        orderData.employeeId = empResult.employee.matchKey;
        orderData.employeeName = empResult.employee.fullName;
        console.log('New employee created via self-service: ' + empResult.employee.fullName);
      }
    }
    
    // Validate order data
    if (!orderData.employeeName) {
      return { success: false, error: 'Employee name is required' };
    }
    
    if (!orderData.location) {
      return { success: false, error: 'Location is required' };
    }
    
    if (!orderData.items || orderData.items.length === 0) {
      return { success: false, error: 'At least one item is required' };
    }
    
    if (orderData.items.length > 5) {
      return { success: false, error: 'Maximum 5 items per request' };
    }
    
    // Get catalog to look up prices
    const catalogItems = getCatalogItems(true);
    
    // First, enrich items with prices from catalog
    const enrichedItems = [];
    for (const item of orderData.items) {
      const catalogItem = catalogItems.find(c => c.itemName === item.itemName);
      const unitPrice = catalogItem ? catalogItem.price : 0;
      const isReplacement = item.isReplacement || false;
      const effectivePrice = isReplacement ? 0 : unitPrice;
      const quantity = item.quantity || 1;
      
      enrichedItems.push({
        itemId: catalogItem ? catalogItem.itemId : '',
        itemName: item.itemName,
        size: formatSizeString(item.size, item.gender, item.inseam),
        quantity: quantity,
        unitPrice: effectivePrice,
        lineTotal: effectivePrice * quantity,
        isReplacement: isReplacement
      });
    }
    
    // Calculate total from enriched items (explicit summation)
    let totalAmount = 0;
    for (const item of enrichedItems) {
      totalAmount += item.lineTotal;
    }
    totalAmount = parseFloat(totalAmount.toFixed(2));
    
    console.log('Order total calculated:', totalAmount, 'from', enrichedItems.length, 'items');
    
    // Get payment options from order data
    const payCash = orderData.payCash || false;
    
    // Payment plan: use employee selection (1-3) unless paying cash (0)
    let paymentPlan = 1; // Default to 1 paycheck
    if (payCash) {
      paymentPlan = 0; // Cash payment - no deductions
    } else if (orderData.paymentPlan && [1, 2, 3].includes(parseInt(orderData.paymentPlan))) {
      paymentPlan = parseInt(orderData.paymentPlan);
    }
    
    const amountPerPaycheck = paymentPlan > 0 ? parseFloat((totalAmount / paymentPlan).toFixed(2)) : 0;
    
    // Generate order ID
    const orderId = generateOrderId();
    const now = new Date();
    
    // Determine status based on total and payment method
    let status = 'Pending';
    if (totalAmount === 0) {
      status = 'Store Paid';
    } else if (payCash) {
      status = 'Pending - Cash';
    }
    
    // Create order row (17 columns)
    const orderRow = [
      orderId,                          // A: Order_ID
      '',                               // B: Employee_ID (not available from form)
      orderData.employeeName,           // C: Employee_Name
      orderData.location,               // D: Location
      now,                              // E: Order_Date
      totalAmount,                      // F: Total_Amount
      paymentPlan,                      // G: Payment_Plan
      amountPerPaycheck,                // H: Amount_Per_Paycheck
      null,                             // I: First_Deduction_Date (set when received)
      0,                                // J: Payments_Made
      0,                                // K: Amount_Paid
      totalAmount,                      // L: Amount_Remaining
      status,                           // M: Status
      orderData.notes || '',            // N: Notes
      'Employee Form',                  // O: Created_By (mark as employee form)
      now,                              // P: Created_Date
      null                              // Q: Received_Date
    ];
    
    ordersSheet.appendRow(orderRow);
    
    // Create line item rows (9 columns)
    const lineRows = enrichedItems.map(item => {
      const lineId = generateLineId();
      
      return [
        lineId,
        orderId,
        item.itemId,
        item.itemName,
        item.size,
        item.quantity,
        item.unitPrice,
        parseFloat(item.lineTotal.toFixed(2)),
        item.isReplacement
      ];
    });
    
    if (lineRows.length > 0) {
      itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, lineRows.length, 9).setValues(lineRows);
    }
    
    // Log the activity
    const paymentMethod = payCash ? 'CASH' : `${paymentPlan} paycheck(s)`;
    logActivity('CREATE', 'UNIFORM', 
      `Employee uniform request: ${orderData.employeeName} - $${totalAmount.toFixed(2)} (${orderData.items.length} items) [Payment: ${paymentMethod}]`,
      orderId
    );
    
    // Note: Immediate email notifications disabled - using weekly summary instead (Saturday 1PM)
    // Uniform orders are included in the weekly summary email
    
    return {
      success: true,
      orderId: orderId,
      message: `Order ${orderId} submitted successfully`,
      totalAmount: totalAmount,
      status: status
    };
    
  } catch (error) {
    console.error('Error submitting employee uniform request:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Cancel a pending uniform order (for employee undo)
 * Only works for Pending or Pending - Cash orders
 * @param {string} orderId - The Order_ID to cancel
 * @returns {Object} Result
 */
function cancelUniformOrder(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet) {
      return { success: false, error: 'Orders sheet not found' };
    }
    
    // Find the order
    const ordersData = ordersSheet.getDataRange().getValues();
    const headers = ordersData[0];
    const statusColIdx = headers.indexOf('Status');
    
    let orderRow = null;
    let orderRowNum = -1;
    
    for (let i = 1; i < ordersData.length; i++) {
      if (ordersData[i][0] === orderId) {
        orderRow = ordersData[i];
        orderRowNum = i + 1;
        break;
      }
    }
    
    if (!orderRow) {
      return { success: false, error: 'Order not found' };
    }
    
    const currentStatus = orderRow[statusColIdx];
    
    // Only allow cancellation of pending orders
    if (currentStatus !== 'Pending' && currentStatus !== 'Pending - Cash') {
      return { success: false, error: `Cannot cancel - order status is "${currentStatus}". Only pending orders can be cancelled.` };
    }
    
    // Update order status to Cancelled
    ordersSheet.getRange(orderRowNum, statusColIdx + 1).setValue('Cancelled');
    
    // Also cancel all items in this order
    if (itemsSheet && itemsSheet.getLastRow() > 1) {
      const itemsData = itemsSheet.getDataRange().getValues();
      const itemHeaders = itemsData[0];
      const itemStatusCol = itemHeaders.indexOf('Item_Status') + 1;
      
      if (itemStatusCol > 0) {
        for (let i = 1; i < itemsData.length; i++) {
          if (itemsData[i][1] === orderId) { // Order_ID column
            itemsSheet.getRange(i + 1, itemStatusCol).setValue('Cancelled');
          }
        }
      }
    }
    
    // Log the activity
    logActivity('CANCEL', 'UNIFORM', `Order ${orderId} cancelled by employee (undo)`, orderId);
    
    return {
      success: true,
      message: `Order ${orderId} has been cancelled`,
      orderId: orderId
    };
    
  } catch (error) {
    console.error('Error cancelling uniform order:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Helper to format size string for display
 */
function formatSizeString(size, gender, inseam) {
  let sizeStr = size || '';
  if (gender) {
    sizeStr += ` (${gender === 'male' ? 'M' : 'F'})`;
  }
  if (inseam) {
    sizeStr += ` / ${inseam}" inseam`;
  }
  return sizeStr;
}

/**
 * Gets unique categories from the catalog
 * @returns {Array} Array of category names
 */
function getCatalogCategories() {
  try {
    const items = getCatalogItems(true);
    const categories = [...new Set(items.map(item => item.category))];
    return categories.sort();
  } catch (error) {
    console.error('Error getting categories:', error);
    return [];
  }
}

/**
 * Adds a new item to the catalog
 * @param {Object} item - Item object with itemId, itemName, category, availableSizes, price
 * @returns {Object} Result
 */
function addCatalogItem(item) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_CATALOG);
    
    if (!sheet) {
      return { success: false, error: 'Catalog sheet not found' };
    }
    
    // Ensure headers exist for waist/inseam columns
    ensureCatalogHeaders(sheet);
    
    // Check for duplicate Item_ID
    const existingItems = getCatalogItems(false);
    if (existingItems.some(i => i.itemId === item.itemId)) {
      return { success: false, error: 'Item ID already exists' };
    }
    
    const newRow = [
      item.itemId,
      item.itemName,
      item.category,
      Array.isArray(item.availableSizes) ? item.availableSizes.join(',') : (item.availableSizes || ''),
      item.price,
      true,
      // Waist and Inseam sizes for pants (optional overrides)
      Array.isArray(item.waistSizes) ? item.waistSizes.join(',') : (item.waistSizes || ''),
      Array.isArray(item.inseamSizes) ? item.inseamSizes.join(',') : (item.inseamSizes || '')
    ];
    
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Item added successfully' };
    
  } catch (error) {
    console.error('Error adding catalog item:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Ensures the catalog sheet has headers for waist/inseam columns
 */
function ensureCatalogHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, 8).getValues()[0];
  if (!headers[6]) sheet.getRange(1, 7).setValue('Waist_Sizes');
  if (!headers[7]) sheet.getRange(1, 8).setValue('Inseam_Sizes');
}

/**
 * Updates an existing catalog item
 * @param {string} itemId - The Item_ID to update
 * @param {Object} updates - Object with fields to update
 * @returns {Object} Result
 */
function updateCatalogItem(itemId, updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_CATALOG);
    
    if (!sheet) {
      return { success: false, error: 'Catalog sheet not found' };
    }
    
    // Ensure headers exist for waist/inseam columns
    ensureCatalogHeaders(sheet);
    
    const items = getCatalogItems(false);
    const item = items.find(i => i.itemId === itemId);
    
    if (!item) {
      return { success: false, error: 'Item not found' };
    }
    
    const row = item.rowIndex;
    
    if (updates.itemName !== undefined) sheet.getRange(row, 2).setValue(updates.itemName);
    if (updates.category !== undefined) sheet.getRange(row, 3).setValue(updates.category);
    if (updates.availableSizes !== undefined) {
      const sizes = Array.isArray(updates.availableSizes) ? updates.availableSizes.join(',') : updates.availableSizes;
      sheet.getRange(row, 4).setValue(sizes);
    }
    if (updates.price !== undefined) sheet.getRange(row, 5).setValue(updates.price);
    if (updates.active !== undefined) sheet.getRange(row, 6).setValue(updates.active);
    
    // Waist and Inseam sizes for pants (optional overrides)
    if (updates.waistSizes !== undefined) {
      const waist = Array.isArray(updates.waistSizes) ? updates.waistSizes.join(',') : updates.waistSizes;
      sheet.getRange(row, 7).setValue(waist);
    }
    if (updates.inseamSizes !== undefined) {
      const inseam = Array.isArray(updates.inseamSizes) ? updates.inseamSizes.join(',') : updates.inseamSizes;
      sheet.getRange(row, 8).setValue(inseam);
    }
    
    return { success: true, message: 'Item updated successfully' };
    
  } catch (error) {
    console.error('Error updating catalog item:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Deactivates a catalog item (soft delete)
 * @param {string} itemId - The Item_ID to deactivate
 * @returns {Object} Result
 */
function deactivateCatalogItem(itemId) {
  return updateCatalogItem(itemId, { active: false });
}

// ============================================================================
// ATOMIC ID GENERATION (prevents duplicate IDs from simultaneous requests)
// ============================================================================

/**
 * Initializes or gets the System_Counters sheet
 * Creates it with current max values if it doesn't exist
 * @returns {Sheet} The System_Counters sheet
 */
function getOrCreateCountersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SYSTEM_COUNTERS);
  
  if (!sheet) {
    // Create the sheet
    sheet = ss.insertSheet(SHEET_NAMES.SYSTEM_COUNTERS);
    
    // Set up headers
    sheet.getRange('A1:C1').setValues([['Counter_Name', 'Current_Value', 'Last_Updated']]);
    sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#E51636').setFontColor('white');
    
    // Scan existing data to seed counters with current max values
    const orderMax = scanExistingOrderIds();
    const lineMax = scanExistingLineIds();
    
    // Add counter rows
    const now = new Date();
    sheet.getRange('A2:C3').setValues([
      ['Order_ID_Counter', orderMax, now],
      ['Line_ID_Counter', lineMax, now]
    ]);
    
    // Format
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 180);
    sheet.getRange('C2:C3').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    
    // Protect the sheet (optional warning)
    sheet.protect().setWarningOnly(true);
    
    Logger.log(`System_Counters sheet created. Order counter: ${orderMax}, Line counter: ${lineMax}`);
  }
  
  return sheet;
}

/**
 * Scans existing orders to find the highest order number
 * Used only when initializing the counters sheet
 * @returns {number} The highest order number found
 */
function scanExistingOrderIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }
  
  const orderIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;
  
  for (const row of orderIds) {
    const orderId = row[0];
    if (orderId && typeof orderId === 'string') {
      // Match ORD-YYYY-NNNN pattern and extract the number
      const match = orderId.match(/ORD-\d{4}-(\d+)/);
      if (match) {
        const num = parseInt(match[1]) || 0;
      if (num > maxNum) maxNum = num;
      }
    }
  }
  
  return maxNum;
}

/**
 * Scans existing line items to find the highest line number
 * Used only when initializing the counters sheet
 * @returns {number} The highest line number found
 */
function scanExistingLineIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }
  
  const lineIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let maxNum = 0;
  
  for (const row of lineIds) {
    const lineId = row[0];
    if (lineId && typeof lineId === 'string' && lineId.startsWith('LINE-')) {
      const num = parseInt(lineId.replace('LINE-', '')) || 0;
      if (num > maxNum) maxNum = num;
    }
  }
  
  return maxNum;
}

/**
 * Atomically generates the next order ID using lock service
 * Prevents duplicate IDs when multiple orders are created simultaneously
 * @returns {string} New order ID in format ORD-YYYY-NNNN
 */
function generateOrderId() {
  const lock = LockService.getScriptLock();
  
  try {
    // Wait up to 10 seconds to acquire lock
    lock.waitLock(10000);
    
    const sheet = getOrCreateCountersSheet();
    const year = new Date().getFullYear();
    
    // Find the Order_ID_Counter row (should be row 2)
    const data = sheet.getRange('A2:C3').getValues();
    let counterRow = 2;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'Order_ID_Counter') {
        counterRow = i + 2;
        break;
      }
    }
    
    // Get current value and increment
    const currentValue = parseInt(sheet.getRange(counterRow, 2).getValue()) || 0;
    const newValue = currentValue + 1;
    
    // Write back the new value with timestamp
    sheet.getRange(counterRow, 2).setValue(newValue);
    sheet.getRange(counterRow, 3).setValue(new Date());
    
    // Force the write to complete before releasing lock
    SpreadsheetApp.flush();
    
    return `ORD-${year}-${String(newValue).padStart(4, '0')}`;
    
  } catch (e) {
    Logger.log('Error in generateOrderId: ' + e.message);
    throw new Error('Failed to generate order ID. Please try again.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Atomically generates the next line item ID using lock service
 * Prevents duplicate IDs when multiple items are created simultaneously
 * @returns {string} New line ID in format LINE-NNNNNN
 */
function generateLineId() {
  const lock = LockService.getScriptLock();
  
  try {
    // Wait up to 10 seconds to acquire lock
    lock.waitLock(10000);
    
    const sheet = getOrCreateCountersSheet();
    
    // Find the Line_ID_Counter row (should be row 3)
    const data = sheet.getRange('A2:C3').getValues();
    let counterRow = 3;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'Line_ID_Counter') {
        counterRow = i + 2;
        break;
      }
    }
    
    // Get current value and increment
    const currentValue = parseInt(sheet.getRange(counterRow, 2).getValue()) || 0;
    const newValue = currentValue + 1;
    
    // Write back the new value with timestamp
    sheet.getRange(counterRow, 2).setValue(newValue);
    sheet.getRange(counterRow, 3).setValue(new Date());
    
    // Force the write to complete before releasing lock
    SpreadsheetApp.flush();
    
    return 'LINE-' + String(newValue).padStart(6, '0');
    
  } catch (e) {
    Logger.log('Error in generateLineId: ' + e.message);
    throw new Error('Failed to generate line ID. Please try again.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Utility function to manually initialize or reset the counters sheet
 * Can be run from the script editor if needed
 */
function initializeSystemCounters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Delete existing sheet if present (for reset)
  const existing = ss.getSheetByName(SHEET_NAMES.SYSTEM_COUNTERS);
  if (existing) {
    ss.deleteSheet(existing);
  }
  
  // Create fresh
  getOrCreateCountersSheet();
  
  return { success: true, message: 'System_Counters sheet initialized successfully' };
}

/**
 * Calculates paydays based on reference date (includes historical and future)
 * Reference: Friday, November 29, 2024 - all paydays are bi-weekly Fridays
 * @param {number} futureCount - Number of future paydays to return (default 8)
 * @param {number} historyCount - Number of historical paydays to return (default 26 = ~1 year)
 * @returns {Array} Array of payday dates (Fridays) sorted chronologically as YYYY-MM-DD strings
 */
function getUpcomingPaydays(futureCount = 8, historyCount = 26) {
  // CRITICAL: Use explicit year, month (0-indexed), day to avoid timezone issues
  // November 29, 2024 is a Friday - this is our reference payday
  const REFERENCE_YEAR = 2024;
  const REFERENCE_MONTH = 10; // 0-indexed: 10 = November
  const REFERENCE_DAY = 29;
  
  // Create reference date in LOCAL time (not UTC)
  const referenceDate = new Date(REFERENCE_YEAR, REFERENCE_MONTH, REFERENCE_DAY, 12, 0, 0);
  
  const today = new Date();
  today.setHours(12, 0, 0, 0); // Use noon to avoid any edge cases
  
  const paydays = [];
  
  // Calculate the most recent payday (on or before today)
  let mostRecentPayday = new Date(referenceDate.getTime());
  while (mostRecentPayday < today) {
    mostRecentPayday.setDate(mostRecentPayday.getDate() + 14);
  }
  // If we went past today, go back one period
  if (mostRecentPayday > today) {
    mostRecentPayday.setDate(mostRecentPayday.getDate() - 14);
  }
  
  // Verify mostRecentPayday is a Friday (day 5)
  // If not, something is wrong - log it
  if (mostRecentPayday.getDay() !== 5) {
    console.error('WARNING: mostRecentPayday is not a Friday! Day of week: ' + mostRecentPayday.getDay());
  }
  
  // Generate historical paydays (going backwards)
  for (let i = historyCount - 1; i >= 0; i--) {
    const payday = new Date(mostRecentPayday.getTime());
    payday.setDate(mostRecentPayday.getDate() - (i * 14));
    paydays.push(payday);
  }
  
  // Generate future paydays (starting from next payday after most recent)
  for (let i = 1; i <= futureCount; i++) {
    const payday = new Date(mostRecentPayday.getTime());
    payday.setDate(mostRecentPayday.getDate() + (i * 14));
    paydays.push(payday);
  }
  
  // Sort chronologically
  paydays.sort((a, b) => a - b);
  
  // Convert to YYYY-MM-DD strings using LOCAL date (not UTC)
  return paydays.map(d => {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  });
}

/**
 * Creates a new uniform order with line items
 * Status starts as "Pending" until items are received, then changes to "Active"
 * @param {Object} orderData - Order data including employeeId, employeeName, location, items, paymentPlan, notes
 * @returns {Object} Result with order ID
 */
function createUniformOrder(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || !itemsSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    // Handle new employee creation if flagged
    if (orderData.isNewEmployee && orderData.employeeName) {
      const empResult = addNewEmployee({
        fullName: orderData.employeeName,
        location: orderData.location || '',
        createdBy: Session.getActiveUser().getEmail() || 'Unknown',
        creationMethod: 'Uniform Order'
      });
      
      if (!empResult.success) {
        // If employee already exists, that's actually OK - use them
        if (empResult.existingEmployee) {
          orderData.employeeId = empResult.existingEmployee.matchKey;
          orderData.employeeName = empResult.existingEmployee.fullName;
        } else {
          return { success: false, error: 'Failed to create employee: ' + empResult.error };
        }
      } else {
        // Use the normalized name and match key from employee creation
        orderData.employeeId = empResult.employee.matchKey;
        orderData.employeeName = empResult.employee.fullName;
        console.log('Created new employee for uniform order: ' + empResult.employee.fullName);
      }
    }
    
    // Validate order data
    if (!orderData.employeeId && !orderData.employeeName) {
      return { success: false, error: 'Employee is required' };
    }
    
    if (!orderData.items || orderData.items.length === 0) {
      return { success: false, error: 'At least one item is required' };
    }
    
    // Calculate totals (replacement items are $0)
    let totalAmount = 0;
    for (const item of orderData.items) {
      const isReplacement = item.isReplacement || false;
      const effectivePrice = isReplacement ? 0 : (item.unitPrice || 0);
      totalAmount += effectivePrice * (item.quantity || 1);
    }
    totalAmount = parseFloat(totalAmount.toFixed(2));
    
    // Validate payment plan (only 1, 2, or 3 allowed for new orders)
    const requestedPlan = parseInt(orderData.paymentPlan) || 1;
    if (requestedPlan < 1 || requestedPlan > 3) {
      return { success: false, error: 'Payment plan must be 1, 2, or 3 paychecks. Please select a valid option.' };
    }
    const paymentPlan = requestedPlan;
    const amountPerPaycheck = totalAmount > 0 ? parseFloat((totalAmount / paymentPlan).toFixed(2)) : 0;
    
    // Generate order ID
    const orderId = generateOrderId();
    const now = new Date();
    const userEmail = Session.getActiveUser().getEmail() || 'Unknown';
    
    // Determine status:
    // - "Store Paid" if $0 total (all replacements or free items)
    // - "Pending" otherwise (waiting for items to be received)
    const status = totalAmount === 0 ? 'Store Paid' : 'Pending';
    
    // Create order row (17 columns with Location and Received_Date)
    // First_Deduction_Date is null until order is marked as received
    const orderRow = [
      orderId,                          // A: Order_ID
      orderData.employeeId || '',       // B: Employee_ID
      orderData.employeeName || '',     // C: Employee_Name
      orderData.location || '',         // D: Location (NEW)
      now,                              // E: Order_Date
      totalAmount,                      // F: Total_Amount
      paymentPlan,                      // G: Payment_Plan
      amountPerPaycheck,                // H: Amount_Per_Paycheck
      null,                             // I: First_Deduction_Date (set when received)
      0,                                // J: Payments_Made
      0,                                // K: Amount_Paid
      totalAmount,                      // L: Amount_Remaining
      status,                           // M: Status
      orderData.notes || '',            // N: Notes
      userEmail,                        // O: Created_By
      now,                              // P: Created_Date
      null                              // Q: Received_Date
    ];
    
    ordersSheet.appendRow(orderRow);
    
    // Create line item rows (9 columns with Is_Replacement)
    const lineRows = orderData.items.map(item => {
      const lineId = generateLineId();
      const isReplacement = item.isReplacement || false;
      const effectivePrice = isReplacement ? 0 : (item.unitPrice || 0);
      const lineTotal = parseFloat((effectivePrice * (item.quantity || 1)).toFixed(2));
      
      return [
        lineId,
        orderId,
        item.itemId || '',
        item.itemName || '',
        item.size || '',
        item.quantity || 1,
        effectivePrice,
        lineTotal,
        isReplacement   // NEW: Is_Replacement column
      ];
    });
    
    if (lineRows.length > 0) {
      itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, lineRows.length, 9).setValues(lineRows);
    }
    
    // Log the activity
    logActivity('CREATE', 'UNIFORM', 
      `Uniform order created: ${orderData.employeeName || orderData.employeeId} - $${totalAmount.toFixed(2)} (${orderData.items.length} items)`,
      orderId
    );
    
    // Note: Immediate email notifications disabled - using weekly summary instead (Saturday 1PM)
    // Uniform orders are included in the weekly summary email
    
    return {
      success: true,
      orderId: orderId,
      message: totalAmount === 0 
        ? `Order ${orderId} created (Store Paid/Replacement)` 
        : `Order ${orderId} created - Status: Pending (waiting for items)`,
      totalAmount: totalAmount,
      amountPerPaycheck: amountPerPaycheck,
      status: status
    };
    
  } catch (error) {
    console.error('Error creating uniform order:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets uniform orders with optional filters
 * @param {Object} filters - Optional filters: status, employeeId, employeeName, location
 * @returns {Array} Array of order objects
 */
function getUniformOrders(filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    // Read all 17 columns
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
    
    // Also get line items to calculate correct totals
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    let lineItemTotals = {};
    
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 9).getValues();
      // Sum up line totals by order ID
      for (const row of itemsData) {
        const orderId = row[1];
        const lineTotal = parseFloat(row[7]) || 0;
        if (orderId) {
          lineItemTotals[orderId] = (lineItemTotals[orderId] || 0) + lineTotal;
        }
      }
    }
    
    let orders = data.map((row, index) => {
      const orderId = row[0] || '';
      const storedTotal = parseFloat(row[5]) || 0;
      
      // Use calculated total from line items if available, otherwise use stored value
      const calculatedTotal = lineItemTotals[orderId];
      const actualTotal = calculatedTotal !== undefined ? calculatedTotal : storedTotal;
      
      // Recalculate amount per paycheck and remaining based on correct total
      const paymentPlan = parseInt(row[6]) || 1;
      const paymentsMade = parseInt(row[9]) || 0;
      const amountPaid = parseFloat(row[10]) || 0;
      const amountPerPaycheck = actualTotal > 0 ? parseFloat((actualTotal / paymentPlan).toFixed(2)) : 0;
      const amountRemaining = parseFloat((actualTotal - amountPaid).toFixed(2));
      
      return {
        orderId: orderId,
        employeeId: row[1] || '',
        employeeName: row[2] || '',
        location: row[3] || '',
        orderDate: row[4] ? new Date(row[4]).toISOString() : null,
        totalAmount: actualTotal,
        paymentPlan: paymentPlan,
        amountPerPaycheck: amountPerPaycheck,
        firstDeductionDate: row[8] ? new Date(row[8]).toISOString().split('T')[0] : null,
        paymentsMade: paymentsMade,
        amountPaid: amountPaid,
        amountRemaining: amountRemaining,
        status: row[12] || 'Pending',
        notes: row[13] || '',
        createdBy: row[14] || '',
        createdDate: row[15] ? new Date(row[15]).toISOString() : null,
        receivedDate: row[16] ? new Date(row[16]).toISOString().split('T')[0] : null,
        rowIndex: index + 2
      };
    }).filter(o => o.orderId);
    
    // Apply filters
    if (filters.statuses && Array.isArray(filters.statuses)) {
      // Multiple statuses (e.g., "Current" = Pending + Active)
      orders = orders.filter(o => filters.statuses.includes(o.status));
    } else if (filters.status) {
      // Single status
      orders = orders.filter(o => o.status === filters.status);
    }
    if (filters.employeeId) {
      orders = orders.filter(o => o.employeeId === filters.employeeId);
    }
    if (filters.employeeName) {
      orders = orders.filter(o => o.employeeName.toLowerCase().includes(filters.employeeName.toLowerCase()));
    }
    if (filters.location) {
      orders = orders.filter(o => o.location === filters.location);
    }
    
    // Sort by order date descending
    orders.sort((a, b) => new Date(b.orderDate) - new Date(a.orderDate));
    
    return orders;
    
  } catch (error) {
    console.error('Error getting uniform orders:', error);
    return [];
  }
}

/**
 * Marks an order as received - changes status from "Pending" to "Active"
 * and sets First_Deduction_Date to next payday at least 7 days away
 * @param {string} orderId - The Order_ID
 * @returns {Object} Result
 */
function markOrderReceived(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!sheet) {
      return { success: false, error: 'Orders sheet not found' };
    }
    
    const orders = getUniformOrders();
    const order = orders.find(o => o.orderId === orderId);
    
    if (!order) {
      return { success: false, error: 'Order not found' };
    }
    
    // Allow marking Pending and Store Paid orders as received
    if (order.status !== 'Pending' && order.status !== 'Store Paid') {
      return { success: false, error: `Cannot mark as received - current status is "${order.status}"` };
    }
    
    const row = order.rowIndex;
    const now = new Date();
    
    // Calculate first deduction date based on which pay period the order was received in
    // Orders received during a pay period are deducted on that period's pay date
    const firstDeductionDate = getPaydayForReceivedDate(now);
    
    // Update the order
    sheet.getRange(row, 9).setValue(new Date(firstDeductionDate));  // I: First_Deduction_Date
    sheet.getRange(row, 13).setValue('Active');                      // M: Status
    sheet.getRange(row, 17).setValue(now);                           // Q: Received_Date
    
    return {
      success: true,
      message: `Order ${orderId} marked as received`,
      firstDeductionDate: firstDeductionDate,
      status: 'Active'
    };
    
  } catch (error) {
    console.error('Error marking order received:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// CHUNK 12: BULK RECEIVING MODE
// ============================================================================

/**
 * Bulk activate multiple orders at once
 * Marks all items as received and activates the orders
 * @param {Array} orderIds - Array of Order_IDs to activate
 * @returns {Object} Result with summary
 */
function bulkActivateOrders(orderIds) {
  try {
    if (!orderIds || orderIds.length === 0) {
      return { success: false, error: 'No orders provided' };
    }
    
    console.log('Bulk activating orders:', orderIds);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || !itemsSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    const now = new Date();
    // Calculate first deduction date based on which pay period orders are received in
    const firstDeductionDate = getPaydayForReceivedDate(now);
    
    // CHUNK 19: Collect before state for all orders being activated
    const beforeStateOrders = [];
    for (const oid of orderIds) {
      const state = getOrderStateForUndo(oid);
      if (state) beforeStateOrders.push(state);
    }
    const beforeState = { orders: beforeStateOrders };
    
    // Get all orders
    const allOrders = getUniformOrders();
    const activatedOrders = [];
    let totalAmount = 0;
    let cashTotal = 0;
    let deductionTotal = 0;
    
    for (const orderId of orderIds) {
      const order = allOrders.find(o => o.orderId === orderId);
      
      if (!order) {
        console.log('Order not found:', orderId);
        continue;
      }
      
      // Allow Pending, Pending - Cash, and Store Paid orders to be processed
      if (order.status !== 'Pending' && order.status !== 'Pending - Cash' && order.status !== 'Store Paid') {
        console.log('Order not in receivable status:', orderId, order.status);
        continue;
      }
      
      const isCashOrder = order.status === 'Pending - Cash';
      const isStorePaid = order.status === 'Store Paid';
      const row = order.rowIndex;
      
      // Mark all items as received for this order
      const itemsData = itemsSheet.getDataRange().getValues();
      const headers = itemsData[0];
      const receivedColIdx = headers.indexOf('Item_Received');
      const receivedDateColIdx = headers.indexOf('Item_Received_Date');
      const statusColIdx = headers.indexOf('Item_Status');
      
      for (let i = 1; i < itemsData.length; i++) {
        if (itemsData[i][1] === orderId) {
          const itemStatus = itemsData[i][statusColIdx];
          // Only mark non-cancelled items as received
          if (itemStatus !== 'Cancelled') {
            itemsSheet.getRange(i + 1, receivedColIdx + 1).setValue(true);
            itemsSheet.getRange(i + 1, receivedDateColIdx + 1).setValue(now);
            if (statusColIdx >= 0) {
              itemsSheet.getRange(i + 1, statusColIdx + 1).setValue('Received');
            }
          }
        }
      }
      
      // Calculate order total from items (excluding cancelled)
      let orderTotal = 0;
      for (let i = 1; i < itemsData.length; i++) {
        if (itemsData[i][1] === orderId && itemsData[i][statusColIdx] !== 'Cancelled') {
          orderTotal += parseFloat(itemsData[i][7]) || 0; // Line_Total column
        }
      }
      
      // Update the order status
      if (isCashOrder) {
        // Cash orders become "Completed" since they're paid in full
        ordersSheet.getRange(row, 13).setValue('Completed');  // M: Status
        ordersSheet.getRange(row, 10).setValue(1);            // J: Payments_Made
        ordersSheet.getRange(row, 11).setValue(orderTotal);   // K: Amount_Paid
        ordersSheet.getRange(row, 12).setValue(0);            // L: Amount_Remaining
        cashTotal += orderTotal;
      } else {
        // Regular orders become "Active" for paycheck deductions
        ordersSheet.getRange(row, 9).setValue(new Date(firstDeductionDate));  // I: First_Deduction_Date
        ordersSheet.getRange(row, 13).setValue('Active');                      // M: Status
        deductionTotal += orderTotal;
      }
      ordersSheet.getRange(row, 17).setValue(now);  // Q: Received_Date
      
      totalAmount += orderTotal;
      activatedOrders.push({
        orderId: orderId,
        employeeName: order.employeeName,
        amount: orderTotal,
        isCash: isCashOrder
      });
      
      // Log the activity
      logActivity('BULK_ACTIVATE', 'UNIFORM', 
        `Bulk activated: ${order.employeeName} - $${orderTotal.toFixed(2)}${isCashOrder ? ' (Cash)' : ''}`,
        orderId
      );
    }
    
    // Send confirmation email
    sendBulkActivationEmail(activatedOrders, totalAmount, cashTotal, deductionTotal, firstDeductionDate);
    
    // CHUNK 19: Save action state for undo
    const afterStateOrders = [];
    for (const activated of activatedOrders) {
      const state = getOrderStateForUndo(activated.orderId);
      if (state) afterStateOrders.push(state);
    }
    const afterState = { orders: afterStateOrders };
    const description = `Bulk activated ${activatedOrders.length} orders - $${totalAmount.toFixed(2)} total`;
    const affectedIds = activatedOrders.map(o => o.orderId);
    const actionId = saveActionState('BULK_ACTIVATE', description, affectedIds, beforeState, afterState);
    
    return {
      success: true,
      activatedCount: activatedOrders.length,
      totalAmount: totalAmount,
      cashTotal: cashTotal,
      deductionTotal: deductionTotal,
      activatedOrders: activatedOrders,
      firstDeductionDate: firstDeductionDate,
      actionId: actionId // For undo capability
    };
    
  } catch (error) {
    console.error('Error in bulk activation:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Send email confirmation for bulk activation
 */
function sendBulkActivationEmail(activatedOrders, totalAmount, cashTotal, deductionTotal, firstDeductionDate) {
  try {
    const settings = getSettings();
    const adminEmails = settings.adminEmails;
    
    if (!adminEmails || adminEmails.trim() === '') {
      console.log('No admin emails configured for bulk activation notification');
      return;
    }
    
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM d, yyyy h:mm a');
    
    // Build order list HTML
    const orderListHtml = activatedOrders.map(o => `
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">${o.employeeName}</td>
        <td style="padding: 8px; border-bottom: 1px solid #eee;">${o.orderId}</td>
        <td style="padding: 8px; text-align: right; border-bottom: 1px solid #eee;">$${o.amount.toFixed(2)}</td>
        <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">
          ${o.isCash ? '<span style="background: #F59E0B; color: white; padding: 2px 8px; border-radius: 4px; font-size: 11px;">CASH</span>' : 
                       '<span style="background: #22C55E; color: white; padding: 2px 8px; border-radius: 4px; font-size: 11px;">DEDUCT</span>'}
        </td>
      </tr>
    `).join('');
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #22C55E 0%, #16A34A 100%); color: white; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0; font-size: 20px;">✅ Bulk Order Activation Complete</h2>
          <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">${dateStr}</p>
        </div>
        <div style="background: #ffffff; padding: 24px; border: 1px solid #e0e0e0; border-top: none;">
          
          <div style="display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap;">
            <div style="flex: 1; min-width: 150px; background: #F0FDF4; padding: 16px; border-radius: 8px; text-align: center;">
              <div style="font-size: 2rem; font-weight: 700; color: #22C55E;">${activatedOrders.length}</div>
              <div style="font-size: 0.85rem; color: #166534;">Orders Activated</div>
            </div>
            <div style="flex: 1; min-width: 150px; background: #EEF2FF; padding: 16px; border-radius: 8px; text-align: center;">
              <div style="font-size: 2rem; font-weight: 700; color: #4F46E5;">$${totalAmount.toFixed(2)}</div>
              <div style="font-size: 0.85rem; color: #3730A3;">Total Amount</div>
            </div>
          </div>
          
          ${deductionTotal > 0 ? `
            <div style="background: #F0FDF4; border: 1px solid #86EFAC; padding: 12px 16px; border-radius: 8px; margin-bottom: 16px;">
              <strong style="color: #166534;">Paycheck Deductions Starting:</strong>
              <div style="font-size: 1.1rem; margin-top: 4px;">$${deductionTotal.toFixed(2)} beginning ${firstDeductionDate}</div>
            </div>
          ` : ''}
          
          ${cashTotal > 0 ? `
            <div style="background: #FEF3C7; border: 1px solid #FCD34D; padding: 12px 16px; border-radius: 8px; margin-bottom: 16px;">
              <strong style="color: #92400E;">Cash Payments Collected:</strong>
              <div style="font-size: 1.1rem; margin-top: 4px;">$${cashTotal.toFixed(2)}</div>
            </div>
          ` : ''}
          
          <h3 style="color: #374151; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E5E7EB;">
            Activated Orders
          </h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr style="background: #f5f5f5;">
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Employee</th>
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Order ID</th>
              <th style="padding: 8px; text-align: right; border-bottom: 1px solid #ddd;">Amount</th>
              <th style="padding: 8px; text-align: center; border-bottom: 1px solid #ddd;">Type</th>
            </tr>
            ${orderListHtml}
          </table>
          
          <div style="margin-top: 20px; padding-top: 16px; border-top: 1px solid #eee; text-align: center;">
            <a href="${ScriptApp.getService().getUrl()}" 
               style="background: #E51636; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: 500;">
              Open Payroll Review
            </a>
          </div>
          
        </div>
        <div style="padding: 12px 24px; background: #f5f5f5; border-radius: 0 0 8px 8px; border: 1px solid #e0e0e0; border-top: none;">
          <p style="color: #888; font-size: 11px; margin: 0; text-align: center;">
            Bulk activation performed via Manager Receiving Panel
          </p>
        </div>
      </div>
    `;
    
    const emailList = adminEmails.split(',').map(e => e.trim()).filter(e => e);
    
    MailApp.sendEmail({
      to: emailList.join(','),
      subject: `[Payroll Review] Bulk Activation: ${activatedOrders.length} orders - $${totalAmount.toFixed(2)}`,
      htmlBody: htmlBody
    });
    
    console.log(`Bulk activation email sent to ${emailList.length} recipient(s)`);
    
  } catch (error) {
    console.error('Error sending bulk activation email:', error);
    // Don't throw - email failure shouldn't fail the whole operation
  }
}

/**
 * Gets the next payday that is at least N days away
 * @param {number} minDays - Minimum days buffer (default 7)
 * @returns {string} ISO date string for the payday
 */
/**
 * OLD FUNCTION - kept for backward compatibility
 * Gets the next payday that is at least minDays away
 * @param {number} minDays - Minimum days buffer (default 7)
 * @returns {string} Date string YYYY-MM-DD
 */
function getNextPaydayWithBuffer(minDays = 7) {
  const settings = getSettings();
  const referenceDate = new Date(settings.paydayReference || '2024-11-29');
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Calculate the minimum acceptable date
  const minDate = new Date(today);
  minDate.setDate(minDate.getDate() + minDays);
  
  let currentPayday = new Date(referenceDate);
  
  // Find the next payday on or after the minimum date
  while (currentPayday < minDate) {
    currentPayday.setDate(currentPayday.getDate() + 14);
  }
  
  return currentPayday.toISOString().split('T')[0];
}

/**
 * NEW FUNCTION - Gets the payday for a given received date based on pay period
 * 
 * Logic: Find which pay period the received date falls into, return that period's pay date
 * 
 * Pay periods are bi-weekly:
 * - Period runs Saturday to Friday (14 days)
 * - Pay date is 6 days after period end (following Thursday)
 * 
 * Example: Period 12/7-12/20 → Pay date 12/26
 *   - Received 12/10 → falls in 12/7-12/20 → deduct 12/26
 *   - Received 12/21 → falls in 12/21-1/3 → deduct 1/9
 * 
 * @param {Date|string} receivedDate - The date the order was received
 * @returns {string} Date string YYYY-MM-DD for the payday
 */
function getPaydayForReceivedDate(receivedDate) {
  // CRITICAL: Use explicit date construction to avoid timezone issues
  // Reference: Friday, November 29, 2024 (a known Friday payday)
  const REFERENCE_YEAR = 2024;
  const REFERENCE_MONTH = 10; // 0-indexed: 10 = November  
  const REFERENCE_DAY = 29;
  
  // Create reference payday in LOCAL time (not UTC)
  const referencePayday = new Date(REFERENCE_YEAR, REFERENCE_MONTH, REFERENCE_DAY, 12, 0, 0);
  
  const received = new Date(receivedDate);
  received.setHours(12, 0, 0, 0); // Use noon to avoid edge cases
  
  // Pay structure for FRIDAY paydays:
  // - Period ends on Saturday (6 days before Friday pay date)
  // - Period starts on Sunday (13 days before period end = 19 days before pay date)
  // - Pay date is FRIDAY
  // Example: Period 12/7(Sun)-12/20(Sat), Pay date 12/26(Fri)
  
  const DAYS_BEFORE_PAY = 6; // Period ends 6 days before payday (Saturday before Friday)
  const PERIOD_LENGTH = 14; // 14 day pay periods
  
  // Find which pay period the received date falls into
  let currentPayday = new Date(referencePayday.getTime());
  
  // Calculate period boundaries for reference payday
  let periodEnd = new Date(currentPayday.getTime());
  periodEnd.setDate(periodEnd.getDate() - DAYS_BEFORE_PAY);
  
  let periodStart = new Date(periodEnd.getTime());
  periodStart.setDate(periodStart.getDate() - (PERIOD_LENGTH - 1)); // -13 days to get 14 day period
  
  // Go backward if received date is before our reference period
  while (received < periodStart) {
    currentPayday.setDate(currentPayday.getDate() - 14);
    periodEnd.setDate(periodEnd.getDate() - 14);
    periodStart.setDate(periodStart.getDate() - 14);
  }
  
  // Go forward until we find the period containing the received date
  while (received > periodEnd) {
    currentPayday.setDate(currentPayday.getDate() + 14);
    periodEnd.setDate(periodEnd.getDate() + 14);
    periodStart.setDate(periodStart.getDate() + 14);
  }
  
  // At this point, periodStart <= received <= periodEnd
  // The payday for this period is currentPayday
  
  // Format dates using LOCAL components (not UTC)
  const formatLocal = (d) => {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };
  
  console.log(`Received: ${formatLocal(received)}, Period: ${formatLocal(periodStart)} to ${formatLocal(periodEnd)}, Payday: ${formatLocal(currentPayday)}`);
  
  return formatLocal(currentPayday);
}

/**
 * REPAIR FUNCTION: Fixes First_Deduction_Date for all orders that have a Received_Date
 * This corrects dates that were saved with the timezone bug
 * Run this once after deploying the timezone fix
 * @returns {Object} Summary of repairs made
 */
function repairFirstDeductionDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!sheet) {
      return { success: false, error: 'UNIFORM_ORDERS sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const cols = {
      orderId: headers.indexOf('Order_ID'),
      status: headers.indexOf('Status'),
      firstDeduction: headers.indexOf('First_Deduction_Date'),
      receivedDate: headers.indexOf('Received_Date'),
      totalAmount: headers.indexOf('Total_Amount'),
      paymentPlan: headers.indexOf('Payment_Plan')
    };
    
    if (cols.firstDeduction === -1 || cols.receivedDate === -1) {
      return { success: false, error: 'Required columns not found' };
    }
    
    const repairs = [];
    const skipped = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const orderId = row[cols.orderId];
      const status = row[cols.status];
      const receivedDate = row[cols.receivedDate];
      const currentFirstDeduction = row[cols.firstDeduction];
      const totalAmount = row[cols.totalAmount];
      
      // Skip orders without received date or with $0 total (Store Paid, etc.)
      if (!receivedDate || !orderId) {
        continue;
      }
      
      // Skip Store Paid and Cancelled orders (no deductions needed)
      if (status === 'Store Paid' || status === 'Cancelled') {
        skipped.push({ orderId, reason: `Status is ${status}` });
        continue;
      }
      
      // Skip if total amount is 0 or less
      if (!totalAmount || totalAmount <= 0) {
        skipped.push({ orderId, reason: 'No amount to deduct' });
        continue;
      }
      
      // Calculate the correct First_Deduction_Date
      const correctPayday = getPaydayForReceivedDate(receivedDate);
      
      // Format current date for comparison
      let currentDateStr = '';
      if (currentFirstDeduction) {
        const d = new Date(currentFirstDeduction);
        if (!isNaN(d.getTime())) {
          const year = d.getFullYear();
          const month = String(d.getMonth() + 1).padStart(2, '0');
          const day = String(d.getDate()).padStart(2, '0');
          currentDateStr = `${year}-${month}-${day}`;
        }
      }
      
      // Check if repair is needed
      if (currentDateStr !== correctPayday) {
        // Update the cell
        const rowNum = i + 1; // 1-indexed for sheet
        const colNum = cols.firstDeduction + 1; // 1-indexed for sheet
        
        // Parse the correct date and set it
        const [year, month, day] = correctPayday.split('-').map(Number);
        const correctDate = new Date(year, month - 1, day, 12, 0, 0);
        
        sheet.getRange(rowNum, colNum).setValue(correctDate);
        
        repairs.push({
          orderId,
          receivedDate: receivedDate.toISOString ? receivedDate.toISOString().split('T')[0] : String(receivedDate),
          oldDate: currentDateStr || '(empty)',
          newDate: correctPayday
        });
      }
    }
    
    // Log results
    console.log('=== REPAIR FIRST DEDUCTION DATES ===');
    console.log('Repairs made:', repairs.length);
    repairs.forEach(r => {
      console.log(`  ${r.orderId}: ${r.oldDate} → ${r.newDate} (received: ${r.receivedDate})`);
    });
    console.log('Skipped:', skipped.length);
    
    return {
      success: true,
      repairsCount: repairs.length,
      repairs: repairs,
      skippedCount: skipped.length,
      skipped: skipped,
      message: `Repaired ${repairs.length} orders, skipped ${skipped.length}`
    };
    
  } catch (error) {
    console.error('Error repairing first deduction dates:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets order details including line items with receiving status
 * @param {string} orderId - The Order_ID to look up
 * @returns {Object} Order with items array
 */
function getOrderDetails(orderId) {
  try {
    const orders = getUniformOrders();
    const order = orders.find(o => o.orderId === orderId);
    
    if (!order) {
      return { success: false, error: 'Order not found' };
    }
    
    // Get line items with receiving columns
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!itemsSheet || itemsSheet.getLastRow() < 2) {
      order.items = [];
    } else {
      // Check if receiving columns exist, get headers
      const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
      const hasReceivingCols = headers.includes('Item_Received');
      
      // If no receiving columns, auto-migrate
      if (!hasReceivingCols) {
        console.log('Receiving columns not found in getOrderDetails, running migration...');
        migrateUniformItemsForReceiving();
        addManagerReceivingPasscode();
        // Re-get headers after migration
        const newHeaders = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
      }
      
      // Get column indices
      const numCols = itemsSheet.getLastColumn();
      const receivedColIdx = headers.indexOf('Item_Received');
      const receivedDateColIdx = headers.indexOf('Item_Received_Date');
      const receivedByColIdx = headers.indexOf('Item_Received_By');
      const statusColIdx = headers.indexOf('Item_Status');
      
      const data = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, numCols).getValues();
      order.items = data
        .filter(row => row[1] === orderId)
        .map(row => ({
          lineId: row[0],
          orderId: row[1],
          itemId: row[2],
          itemName: row[3],
          size: row[4],
          quantity: parseInt(row[5]) || 1,
          unitPrice: parseFloat(row[6]) || 0,
          lineTotal: parseFloat(row[7]) || 0,
          isReplacement: row[8] === true || row[8] === 'TRUE',
          itemReceived: receivedColIdx >= 0 ? (row[receivedColIdx] === true || row[receivedColIdx] === 'TRUE') : false,
          itemReceivedDate: receivedDateColIdx >= 0 ? row[receivedDateColIdx] : null,
          itemReceivedBy: receivedByColIdx >= 0 ? row[receivedByColIdx] : '',
          itemStatus: statusColIdx >= 0 ? (row[statusColIdx] || 'Pending') : 'Pending'
        }));
    }
    
    // Calculate receiving stats
    const receivedItems = order.items.filter(i => i.itemReceived);
    const receivedTotal = receivedItems.reduce((sum, i) => sum + i.lineTotal, 0);
    order.receivedCount = receivedItems.length;
    order.totalItemCount = order.items.length;
    order.receivedTotal = parseFloat(receivedTotal.toFixed(2));
    
    // Ensure proper serialization for web app
    const result = JSON.parse(JSON.stringify({ success: true, order: order }));
    return result;
    
  } catch (error) {
    console.error('Error getting order details:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// UNIFORM RECEIVING FUNCTIONS (Manager Receiving Section)
// ============================================================================

/**
 * Gets pending orders for the manager receiving interface
 * @param {string} location - Location filter (optional, '' for all)
 * @returns {Array} Array of pending orders with items
 */
function getPendingOrdersForReceiving(location) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return [];
    }
    
    // Check if receiving columns exist, if not, auto-migrate
    if (itemsSheet) {
      const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Item_Received')) {
        console.log('Receiving columns not found, running migration...');
        migrateUniformItemsForReceiving();
        // Also add the passcode if it doesn't exist
        addManagerReceivingPasscode();
      }
    }
    
    // Get all pending orders
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 17).getValues();
    
    console.log('Total orders in sheet:', ordersData.length);
    
    // Include both "Pending" and "Store Paid" orders that need to be received
    // Store Paid orders are $0 orders that still need to be marked as received/given to employee
    let pendingOrders = ordersData
      .map((row, index) => ({
        rowIndex: index + 2,
        orderId: row[0],
        employeeId: row[1],
        employeeName: row[2],
        location: row[3],
        orderDate: row[4] instanceof Date ? row[4].toISOString() : row[4],
        totalAmount: parseFloat(row[5]) || 0,
        paymentPlan: parseInt(row[6]) || 1,
        status: row[12],
        notes: row[13],
        receivedDate: row[16] // Column Q: Received_Date
      }))
      .filter(order => {
        // Include orders that are:
        // 1. Pending (waiting to be received)
        // 2. Store Paid but NOT yet received (no received date)
        if (!order.orderId) return false;
        if (order.status === 'Pending') return true;
        if (order.status === 'Store Paid' && !order.receivedDate) return true;
        return false;
      });
    
    console.log('Pending/receivable orders found:', pendingOrders.length);
    if (pendingOrders.length > 0) {
      console.log('First pending order:', JSON.stringify(pendingOrders[0]));
    }
    
    // Filter by location if specified
    if (location && location !== '' && location !== 'All') {
      pendingOrders = pendingOrders.filter(o => o.location === location);
    }
    
    // Get items for each order
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const numCols = Math.max(13, itemsSheet.getLastColumn());
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, numCols).getValues();
      
      // Get header row to find column indices
      const headers = itemsSheet.getRange(1, 1, 1, numCols).getValues()[0];
      const receivedCol = headers.indexOf('Item_Received');
      const receivedDateCol = headers.indexOf('Item_Received_Date');
      const receivedByCol = headers.indexOf('Item_Received_By');
      const statusCol = headers.indexOf('Item_Status');
      
      for (const order of pendingOrders) {
        order.items = itemsData
          .filter(row => row[1] === order.orderId)
          .map((row, idx) => {
            // Find the actual row number in the sheet for this item
            let actualRowNum = 2;
            let matchCount = 0;
            for (let i = 0; i < itemsData.length; i++) {
              if (itemsData[i][1] === order.orderId) {
                if (matchCount === idx) {
                  actualRowNum = i + 2;
                  break;
                }
                matchCount++;
              }
            }
            
            return {
              lineId: row[0],
              rowIndex: actualRowNum,
              orderId: row[1],
              itemId: row[2],
              itemName: row[3],
              size: row[4],
              quantity: parseInt(row[5]) || 1,
              unitPrice: parseFloat(row[6]) || 0,
              lineTotal: parseFloat(row[7]) || 0,
              isReplacement: row[8] === true || row[8] === 'TRUE',
              itemReceived: receivedCol >= 0 ? (row[receivedCol] === true || row[receivedCol] === 'TRUE') : false,
              itemReceivedDate: receivedDateCol >= 0 ? row[receivedDateCol] : null,
              itemReceivedBy: receivedByCol >= 0 ? row[receivedByCol] : '',
              itemStatus: statusCol >= 0 ? (row[statusCol] || 'Pending') : 'Pending'
            };
          });
        
        // Calculate receiving stats
        const receivedItems = order.items.filter(i => i.itemReceived && i.itemStatus !== 'Cancelled');
        const activeItems = order.items.filter(i => i.itemStatus !== 'Cancelled');
        order.receivedCount = receivedItems.length;
        order.totalItemCount = activeItems.length;
        order.receivedTotal = parseFloat(receivedItems.reduce((sum, i) => sum + i.lineTotal, 0).toFixed(2));
        order.pendingTotal = parseFloat(activeItems.filter(i => !i.itemReceived).reduce((sum, i) => sum + i.lineTotal, 0).toFixed(2));
      }
    }
    
    // Sort by order date descending
    pendingOrders.sort((a, b) => new Date(b.orderDate) - new Date(a.orderDate));
    
    return pendingOrders;
    
  } catch (error) {
    console.error('Error getting pending orders for receiving:', error);
    return [];
  }
}

/**
 * Updates the received status for multiple items
 * @param {Array} updates - Array of {lineId, received: boolean}
 * @param {string} receivedBy - Who is marking items (name or identifier)
 * @returns {Object} Result
 */
function updateItemsReceivedStatus(updates, receivedBy) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!itemsSheet || itemsSheet.getLastRow() < 2) {
      return { success: false, error: 'No items found' };
    }
    
    // Get header row to find column indices
    const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
    const receivedCol = headers.indexOf('Item_Received') + 1;
    const receivedDateCol = headers.indexOf('Item_Received_Date') + 1;
    const receivedByCol = headers.indexOf('Item_Received_By') + 1;
    const statusCol = headers.indexOf('Item_Status') + 1;
    
    // Check for Received_Quantity column (add if missing)
    let receivedQtyCol = headers.indexOf('Received_Quantity') + 1;
    if (receivedQtyCol === 0) {
      // Add the column
      const lastCol = itemsSheet.getLastColumn();
      itemsSheet.getRange(1, lastCol + 1).setValue('Received_Quantity');
      receivedQtyCol = lastCol + 1;
    }
    
    if (receivedCol === 0) {
      return { success: false, error: 'Item_Received column not found. Please run migration first.' };
    }
    
    // Get all items data
    const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 2).getValues();
    const now = new Date();
    let updatedCount = 0;
    
    for (const update of updates) {
      // Find the row for this lineId
      let rowNum = -1;
      for (let i = 0; i < itemsData.length; i++) {
        if (itemsData[i][0] === update.lineId) {
          rowNum = i + 2;
          break;
        }
      }
      
      if (rowNum === -1) continue;
      
      // Update the item
      itemsSheet.getRange(rowNum, receivedCol).setValue(update.received);
      
      // Store received quantity (for partial quantity support)
      if (receivedQtyCol > 0 && update.receivedQuantity !== undefined) {
        itemsSheet.getRange(rowNum, receivedQtyCol).setValue(update.receivedQuantity);
      }
      
      if (update.received) {
        if (receivedDateCol > 0) itemsSheet.getRange(rowNum, receivedDateCol).setValue(now);
        if (receivedByCol > 0) itemsSheet.getRange(rowNum, receivedByCol).setValue(receivedBy || 'Manager');
        if (statusCol > 0) itemsSheet.getRange(rowNum, statusCol).setValue('Received');
      } else {
        // If unmarking as received
        if (receivedDateCol > 0) itemsSheet.getRange(rowNum, receivedDateCol).setValue('');
        if (receivedByCol > 0) itemsSheet.getRange(rowNum, receivedByCol).setValue('');
        if (statusCol > 0) itemsSheet.getRange(rowNum, statusCol).setValue('Pending');
      }
      
      updatedCount++;
    }
    
    return {
      success: true,
      message: `Updated ${updatedCount} item(s)`,
      updatedCount: updatedCount
    };
    
  } catch (error) {
    console.error('Error updating items received status:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Cancels an individual uniform item
 * @param {string} lineId - The Line_ID to cancel
 * @returns {Object} Result
 */
function cancelUniformItem(lineId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!itemsSheet || itemsSheet.getLastRow() < 2) {
      return { success: false, error: 'No items found' };
    }
    
    // Get header row
    const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('Item_Status') + 1;
    
    if (statusCol === 0) {
      return { success: false, error: 'Item_Status column not found. Please run migration first.' };
    }
    
    // Find the item
    const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 2).getValues();
    let rowNum = -1;
    let orderId = null;
    
    for (let i = 0; i < itemsData.length; i++) {
      if (itemsData[i][0] === lineId) {
        rowNum = i + 2;
        orderId = itemsData[i][1];
        break;
      }
    }
    
    if (rowNum === -1) {
      return { success: false, error: 'Item not found' };
    }
    
    // Mark as cancelled
    itemsSheet.getRange(rowNum, statusCol).setValue('Cancelled');
    
    // Log the activity
    logActivity('CANCEL', 'UNIFORM_ITEM', `Cancelled item ${lineId} from order ${orderId}`);
    
    return {
      success: true,
      message: `Item ${lineId} cancelled`,
      orderId: orderId
    };
    
  } catch (error) {
    console.error('Error cancelling uniform item:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Activates an order with only the received items
 * If there are unreceived items, they are split into a new order
 * @param {string} orderId - The Order_ID to activate
 * @returns {Object} Result with optional newOrderId for split items
 */
function activateOrderWithReceivedItems(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || !itemsSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    // Get order details
    const orderResult = getOrderDetails(orderId);
    if (!orderResult.success) {
      return orderResult;
    }
    
    const order = orderResult.order;
    
    // Allow Pending, Pending - Cash, and Store Paid orders to be activated
    if (order.status !== 'Pending' && order.status !== 'Pending - Cash' && order.status !== 'Store Paid') {
      return { success: false, error: `Cannot activate - order status is "${order.status}"` };
    }
    
    // CHUNK 19: Save state before activation for undo capability
    const beforeState = {
      orders: [getOrderStateForUndo(orderId)]
    };
    
    // Get item sheet headers for column lookups
    const itemHeaders = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
    const lineIdCol = itemHeaders.indexOf('Line_ID') + 1;
    const orderIdCol = itemHeaders.indexOf('Order_ID') + 1;
    const quantityCol = itemHeaders.indexOf('Quantity') + 1;
    const lineTotalCol = itemHeaders.indexOf('Line_Total') + 1;
    const receivedQtyCol = itemHeaders.indexOf('Received_Quantity') + 1;
    
    // Get all item data to read received quantities
    const allItemData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();
    
    // Process items with partial quantity support
    let receivedTotal = 0;
    let unreceivedTotal = 0;
    const itemsToReceive = [];
    const itemsForNewOrder = [];  // Unreceived items/quantities
    const partialItemsToSplit = [];
    
    for (const item of order.items) {
      if (item.itemStatus === 'Cancelled') continue;
      
      const orderedQty = item.quantity || 1;
      
      // Find the item row to get received quantity
      let receivedQty = item.itemReceived ? orderedQty : 0;
      for (let i = 0; i < allItemData.length; i++) {
        if (allItemData[i][lineIdCol - 1] === item.lineId) {
          // Check if there's a Received_Quantity value
          if (receivedQtyCol > 0 && allItemData[i][receivedQtyCol - 1] !== '' && allItemData[i][receivedQtyCol - 1] !== null) {
            receivedQty = parseInt(allItemData[i][receivedQtyCol - 1]) || 0;
          }
          break;
        }
      }
      
      const unitPrice = item.lineTotal / orderedQty;
      const receivedAmount = receivedQty * unitPrice;
      const unreceivedQty = orderedQty - receivedQty;
      const unreceivedAmount = unreceivedQty * unitPrice;
      
      if (receivedQty > 0) {
        receivedTotal += receivedAmount;
        itemsToReceive.push({
          ...item,
          receivedQty: receivedQty,
          receivedAmount: receivedAmount
        });
      }
      
      if (unreceivedQty > 0) {
        unreceivedTotal += unreceivedAmount;
        
        if (receivedQty > 0) {
          // This is a partial receive - need to split the item
          partialItemsToSplit.push({
            ...item,
            receivedQty: receivedQty,
            receivedAmount: receivedAmount,
            unreceivedQty: unreceivedQty,
            unreceivedAmount: unreceivedAmount,
            unitPrice: unitPrice
          });
        } else {
          // Fully unreceived - just move to new order
          itemsForNewOrder.push(item);
        }
      }
    }
    
    if (itemsToReceive.length === 0) {
      return { success: false, error: 'No items have been marked as received' };
    }
    
    // Calculate new total for received items/quantities
    const newTotal = parseFloat(receivedTotal.toFixed(2));
    const paymentPlan = order.paymentPlan;
    const newAmountPerPaycheck = newTotal > 0 ? parseFloat((newTotal / paymentPlan).toFixed(2)) : 0;
    
    // Find the order row
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 1).getValues();
    let orderRowNum = -1;
    for (let i = 0; i < ordersData.length; i++) {
      if (ordersData[i][0] === orderId) {
        orderRowNum = i + 2;
        break;
      }
    }
    
    if (orderRowNum === -1) {
      return { success: false, error: 'Order not found in sheet' };
    }
    
    const now = new Date();
    // Calculate first deduction date based on which pay period the order is received in
    const firstDeductionDate = getPaydayForReceivedDate(now);
    const isCashOrder = order.status === 'Pending - Cash';
    const isStorePaid = order.status === 'Store Paid';
    
    // Determine status:
    // - "Completed" if cash order or Store Paid (no deductions needed, just marking as received)
    // - "Store Paid" if $0 total (from Pending)
    // - "Active" otherwise (deductions will start)
    let newStatus;
    if (isCashOrder || isStorePaid) {
      newStatus = 'Completed';  // Cash/Store Paid orders become Completed when received
    } else if (newTotal === 0) {
      newStatus = 'Store Paid';
    } else {
      newStatus = 'Active';
    }
    
    // Update the original order
    ordersSheet.getRange(orderRowNum, 6).setValue(newTotal);                    // F: Total_Amount
    ordersSheet.getRange(orderRowNum, 8).setValue(newAmountPerPaycheck);        // H: Amount_Per_Paycheck
    if (!isCashOrder) {
      ordersSheet.getRange(orderRowNum, 9).setValue(new Date(firstDeductionDate)); // I: First_Deduction_Date
    }
    ordersSheet.getRange(orderRowNum, 12).setValue(isCashOrder ? 0 : newTotal); // L: Amount_Remaining (0 for cash - paid in full)
    ordersSheet.getRange(orderRowNum, 13).setValue(newStatus);                  // M: Status
    ordersSheet.getRange(orderRowNum, 17).setValue(now);                        // Q: Received_Date
    
    // For cash orders, also mark as fully paid
    if (isCashOrder) {
      ordersSheet.getRange(orderRowNum, 10).setValue(1);                        // J: Payments_Made
      ordersSheet.getRange(orderRowNum, 11).setValue(newTotal);                 // K: Amount_Paid
    }
    
    // Update quantity on partially received items (reduce to received qty only)
    for (const item of partialItemsToSplit) {
      for (let i = 0; i < allItemData.length; i++) {
        if (allItemData[i][lineIdCol - 1] === item.lineId) {
          const rowNum = i + 2;
          if (quantityCol > 0) itemsSheet.getRange(rowNum, quantityCol).setValue(item.receivedQty);
          if (lineTotalCol > 0) itemsSheet.getRange(rowNum, lineTotalCol).setValue(item.receivedAmount);
          break;
        }
      }
    }
    
    let newOrderId = null;
    const hasUnreceivedItems = itemsForNewOrder.length > 0 || partialItemsToSplit.length > 0;
    
    // If there are unreceived items or partial quantities, create a new order for them
    if (hasUnreceivedItems) {
      const totalUnreceivedAmount = parseFloat(unreceivedTotal.toFixed(2));
      
      // Generate new order ID
      newOrderId = generateOrderId();
      const newAmountPerPaycheck2 = totalUnreceivedAmount > 0 ? parseFloat((totalUnreceivedAmount / paymentPlan).toFixed(2)) : 0;
      
      // Create new order row
      const newOrderRow = [
        newOrderId,                          // A: Order_ID
        order.employeeId || '',              // B: Employee_ID
        order.employeeName,                  // C: Employee_Name
        order.location,                      // D: Location
        now,                                 // E: Order_Date
        totalUnreceivedAmount,               // F: Total_Amount
        paymentPlan,                         // G: Payment_Plan
        newAmountPerPaycheck2,               // H: Amount_Per_Paycheck
        null,                                // I: First_Deduction_Date (null until received)
        0,                                   // J: Payments_Made
        0,                                   // K: Amount_Paid
        totalUnreceivedAmount,               // L: Amount_Remaining
        'Pending',                           // M: Status
        `Split from ${orderId} - unreceived items`, // N: Notes
        'System',                            // O: Created_By
        now,                                 // P: Created_Date
        null                                 // Q: Received_Date
      ];
      
      ordersSheet.appendRow(newOrderRow);
      
      // Update fully unreceived items to point to the new order
      for (const item of itemsForNewOrder) {
        for (let i = 0; i < allItemData.length; i++) {
          if (allItemData[i][lineIdCol - 1] === item.lineId) {
            itemsSheet.getRange(i + 2, orderIdCol).setValue(newOrderId);
            break;
          }
        }
      }
      
      // Create new line items for partial quantities (the unreceived portion)
      for (const item of partialItemsToSplit) {
        const newLineId = 'LINE-' + Utilities.getUuid().substring(0, 8).toUpperCase();
        
        // Find original item row to copy data
        let origRow = null;
        for (let i = 0; i < allItemData.length; i++) {
          if (allItemData[i][lineIdCol - 1] === item.lineId) {
            origRow = allItemData[i];
            break;
          }
        }
        
        if (origRow) {
          // Create new row with unreceived quantity
          const newItemRow = [...origRow];
          newItemRow[lineIdCol - 1] = newLineId;                     // New Line_ID
          newItemRow[orderIdCol - 1] = newOrderId;                   // New Order_ID
          if (quantityCol > 0) newItemRow[quantityCol - 1] = item.unreceivedQty;
          if (lineTotalCol > 0) newItemRow[lineTotalCol - 1] = item.unreceivedAmount;
          
          // Reset received status for the new item
          const itemReceivedCol = itemHeaders.indexOf('Item_Received');
          const itemStatusCol = itemHeaders.indexOf('Item_Status');
          const receivedDateCol = itemHeaders.indexOf('Item_Received_Date');
          const receivedByCol = itemHeaders.indexOf('Item_Received_By');
          
          if (itemReceivedCol >= 0) newItemRow[itemReceivedCol] = false;
          if (itemStatusCol >= 0) newItemRow[itemStatusCol] = 'Pending';
          if (receivedDateCol >= 0) newItemRow[receivedDateCol] = '';
          if (receivedByCol >= 0) newItemRow[receivedByCol] = '';
          if (receivedQtyCol > 0) newItemRow[receivedQtyCol - 1] = '';
          
          itemsSheet.appendRow(newItemRow);
        }
      }
      
      const totalUnreceivedUnits = itemsForNewOrder.reduce((sum, i) => sum + (i.quantity || 1), 0) 
                                 + partialItemsToSplit.reduce((sum, i) => sum + i.unreceivedQty, 0);
      logActivity('CREATE', 'UNIFORM', `Split order created: ${newOrderId} from ${orderId} with ${totalUnreceivedUnits} unreceived units`);
    }
    
    logActivity('UPDATE', 'UNIFORM', `Order ${orderId} activated with ${itemsToReceive.length} items, total: $${newTotal}`);
    
    // CHUNK 19: Save action state for undo
    const afterState = {
      orders: [getOrderStateForUndo(orderId)]
    };
    const description = `Activated order ${orderId} (${order.employeeName}) - $${newTotal.toFixed(2)}`;
    const actionId = saveActionState('ACTIVATE_ORDER', description, [orderId], beforeState, afterState);
    
    return {
      success: true,
      message: `Order activated with ${itemsToReceive.length} received items`,
      orderId: orderId,
      newTotal: newTotal,
      status: newStatus,
      firstDeductionDate: firstDeductionDate,
      receivedCount: itemsToReceive.length,
      newOrderId: newOrderId,
      unreceivedCount: itemsForNewOrder.length + partialItemsToSplit.length,
      actionId: actionId // For undo capability
    };
    
  } catch (error) {
    console.error('Error activating order with received items:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Activates an order as "Store Paid" (cash payment)
 * No payroll deductions will be made
 * @param {string} orderId - The Order_ID to activate
 * @param {Array} receivedData - Array of {lineId, receivedQuantity} for items received
 * @returns {Object} Result
 */
function activateOrderAsCashPayment(orderId, receivedData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || !itemsSheet) {
      return { success: false, error: 'Required sheets not found' };
    }
    
    // Get order details
    const orderResult = getOrderDetails(orderId);
    if (!orderResult.success) {
      return orderResult;
    }
    
    const order = orderResult.order;
    
    // Only allow Pending orders to be converted to cash
    if (order.status !== 'Pending' && order.status !== 'Pending - Cash') {
      return { success: false, error: `Cannot mark as cash payment - order status is "${order.status}"` };
    }
    
    // Get item sheet headers for column lookups
    const itemHeaders = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
    const lineIdCol = itemHeaders.indexOf('Line_ID') + 1;
    const quantityCol = itemHeaders.indexOf('Quantity') + 1;
    const lineTotalCol = itemHeaders.indexOf('Line_Total') + 1;
    const receivedCol = itemHeaders.indexOf('Item_Received') + 1;
    const receivedDateCol = itemHeaders.indexOf('Item_Received_Date') + 1;
    const receivedByCol = itemHeaders.indexOf('Item_Received_By') + 1;
    const statusCol = itemHeaders.indexOf('Item_Status') + 1;
    const receivedQtyCol = itemHeaders.indexOf('Received_Quantity') + 1;
    
    // Create a map of received quantities from frontend
    const receivedMap = {};
    for (const item of receivedData) {
      receivedMap[item.lineId] = item.receivedQuantity;
    }
    
    // Get all item data
    const allItemData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();
    
    // Calculate the total for received items
    let receivedTotal = 0;
    const itemsToReceive = [];
    const itemsForNewOrder = [];
    
    for (const item of order.items) {
      if (item.itemStatus === 'Cancelled') continue;
      
      const orderedQty = item.quantity || 1;
      const receivedQty = receivedMap[item.lineId] !== undefined ? receivedMap[item.lineId] : 0;
      const unitPrice = item.lineTotal / orderedQty;
      const receivedAmount = receivedQty * unitPrice;
      const unreceivedQty = orderedQty - receivedQty;
      
      if (receivedQty > 0) {
        receivedTotal += receivedAmount;
        itemsToReceive.push({
          ...item,
          receivedQty: receivedQty,
          receivedAmount: receivedAmount
        });
      }
      
      if (unreceivedQty > 0) {
        itemsForNewOrder.push({
          ...item,
          unreceivedQty: unreceivedQty,
          unreceivedAmount: unreceivedQty * unitPrice
        });
      }
    }
    
    if (itemsToReceive.length === 0) {
      return { success: false, error: 'No items have been marked as received' };
    }
    
    const newTotal = parseFloat(receivedTotal.toFixed(2));
    const now = new Date();
    
    // Find the order row
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 1).getValues();
    let orderRowNum = -1;
    for (let i = 0; i < ordersData.length; i++) {
      if (ordersData[i][0] === orderId) {
        orderRowNum = i + 2;
        break;
      }
    }
    
    if (orderRowNum === -1) {
      return { success: false, error: 'Order not found in sheet' };
    }
    
    // Update the original order to "Store Paid" status (cash payment = no deductions)
    ordersSheet.getRange(orderRowNum, 6).setValue(newTotal);                // F: Total_Amount
    ordersSheet.getRange(orderRowNum, 8).setValue(0);                       // H: Amount_Per_Paycheck (no deductions)
    ordersSheet.getRange(orderRowNum, 10).setValue(1);                      // J: Payments_Made (paid in full via cash)
    ordersSheet.getRange(orderRowNum, 11).setValue(newTotal);               // K: Amount_Paid
    ordersSheet.getRange(orderRowNum, 12).setValue(0);                      // L: Amount_Remaining
    ordersSheet.getRange(orderRowNum, 13).setValue('Store Paid');           // M: Status
    ordersSheet.getRange(orderRowNum, 17).setValue(now);                    // Q: Received_Date
    
    // Update item statuses
    for (let i = 0; i < allItemData.length; i++) {
      const lineId = allItemData[i][lineIdCol - 1];
      const receivedQty = receivedMap[lineId];
      
      if (receivedQty !== undefined && receivedQty > 0) {
        const rowNum = i + 2;
        if (receivedCol > 0) itemsSheet.getRange(rowNum, receivedCol).setValue(true);
        if (receivedDateCol > 0) itemsSheet.getRange(rowNum, receivedDateCol).setValue(now);
        if (receivedByCol > 0) itemsSheet.getRange(rowNum, receivedByCol).setValue('Manager (Cash)');
        if (statusCol > 0) itemsSheet.getRange(rowNum, statusCol).setValue('Received');
        if (receivedQtyCol > 0) itemsSheet.getRange(rowNum, receivedQtyCol).setValue(receivedQty);
        
        // Update quantity if partial
        const orderedQty = allItemData[i][quantityCol - 1] || 1;
        if (receivedQty < orderedQty && quantityCol > 0) {
          itemsSheet.getRange(rowNum, quantityCol).setValue(receivedQty);
          const unitPrice = allItemData[i][lineTotalCol - 1] / orderedQty;
          if (lineTotalCol > 0) itemsSheet.getRange(rowNum, lineTotalCol).setValue(receivedQty * unitPrice);
        }
      }
    }
    
    // If there are unreceived items, create a new pending order for them
    let newOrderId = null;
    if (itemsForNewOrder.length > 0) {
      const totalUnreceivedAmount = itemsForNewOrder.reduce((sum, i) => sum + i.unreceivedAmount, 0);
      newOrderId = generateOrderId();
      
      const newOrderRow = [
        newOrderId,
        order.employeeId || '',
        order.employeeName,
        order.location,
        now,
        totalUnreceivedAmount,
        order.paymentPlan,
        parseFloat((totalUnreceivedAmount / order.paymentPlan).toFixed(2)),
        null,
        0,
        0,
        totalUnreceivedAmount,
        'Pending',
        order.notes ? `Split from ${orderId}. ${order.notes}` : `Split from ${orderId}`,
        '',
        '',
        null
      ];
      
      ordersSheet.appendRow(newOrderRow);
      
      // Create items for new order
      for (const item of itemsForNewOrder) {
        const newItemRow = [
          generateLineId(),
          newOrderId,
          item.itemId,
          item.itemName,
          item.size,
          item.unreceivedQty,
          item.unreceivedAmount,
          false,
          'Pending',
          null,
          null,
          null,
          null
        ];
        itemsSheet.appendRow(newItemRow);
      }
    }
    
    // Log the action
    logActivity('ORDER_CASH_PAYMENT', 'UNIFORM', `Order ${orderId} marked as cash payment ($${newTotal.toFixed(2)})`, orderId);
    
    return {
      success: true,
      message: `Order marked as Paid Cash - $${newTotal.toFixed(2)}`,
      orderId: orderId,
      total: newTotal,
      newOrderId: newOrderId,
      receivedCount: itemsToReceive.length
    };
    
  } catch (error) {
    console.error('Error activating order as cash payment:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Records a payment for an order
 * @param {string} orderId - The Order_ID
 * @returns {Object} Result
 */
function recordUniformPayment(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!sheet) {
      return { success: false, error: 'Orders sheet not found' };
    }
    
    const orders = getUniformOrders();
    const order = orders.find(o => o.orderId === orderId);
    
    if (!order) {
      return { success: false, error: 'Order not found' };
    }
    
    if (order.status === 'Completed' || order.status === 'Cancelled' || order.status === 'Store Paid' || order.status === 'Pending') {
      return { success: false, error: `Cannot record payment - order status is "${order.status}"` };
    }
    
    if (order.paymentsMade >= order.paymentPlan) {
      return { success: false, error: 'All payments already recorded' };
    }
    
    const row = order.rowIndex;
    const newPaymentsMade = order.paymentsMade + 1;
    const newAmountPaid = parseFloat((order.amountPaid + order.amountPerPaycheck).toFixed(2));
    let newAmountRemaining = parseFloat((order.totalAmount - newAmountPaid).toFixed(2));
    let newStatus = order.status;
    
    // Check if fully paid
    if (newPaymentsMade >= order.paymentPlan) {
      newStatus = 'Completed';
      newAmountRemaining = 0; // Handle rounding
    }
    
    // Update the order (columns shifted due to Location column)
    sheet.getRange(row, 10).setValue(newPaymentsMade);    // J: Payments_Made
    sheet.getRange(row, 11).setValue(newAmountPaid);      // K: Amount_Paid
    sheet.getRange(row, 12).setValue(newAmountRemaining); // L: Amount_Remaining
    sheet.getRange(row, 13).setValue(newStatus);          // M: Status
    
    return {
      success: true,
      message: `Payment ${newPaymentsMade} of ${order.paymentPlan} recorded`,
      paymentsMade: newPaymentsMade,
      amountPaid: newAmountPaid,
      amountRemaining: newAmountRemaining,
      status: newStatus
    };
    
  } catch (error) {
    console.error('Error recording payment:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Cancels an order
 * @param {string} orderId - The Order_ID
 * @returns {Object} Result
 */
function cancelUniformOrder(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!sheet) {
      return { success: false, error: 'Orders sheet not found' };
    }
    
    const orders = getUniformOrders();
    const order = orders.find(o => o.orderId === orderId);
    
    if (!order) {
      return { success: false, error: 'Order not found' };
    }
    
    sheet.getRange(order.rowIndex, 13).setValue('Cancelled');  // M: Status column
    
    return { success: true, message: 'Order cancelled successfully' };
    
  } catch (error) {
    console.error('Error cancelling order:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets payroll deductions due for a specific payday
 * @param {string} payday - The payday date (YYYY-MM-DD)
 * @returns {Object} Deductions data
 */
function getPayrollDeductions(payday) {
  try {
    // Use noon to avoid timezone issues
    const paydayDate = new Date(payday + 'T12:00:00');
    
    // Get BOTH Active AND Completed orders for historical viewing
    // Exclude: Pending, Cancelled, Store Paid (these never have deductions)
    const allOrders = getUniformOrders();
    const relevantOrders = allOrders.filter(o => 
      o.status === 'Active' || o.status === 'Completed'
    );
    
    const deductions = [];
    
    for (const order of relevantOrders) {
      if (!order.firstDeductionDate) continue;
      
      // Use noon to avoid timezone issues
      const firstDeduction = new Date(order.firstDeductionDate);
      if (typeof order.firstDeductionDate === 'string' && !order.firstDeductionDate.includes('T')) {
        firstDeduction.setTime(new Date(order.firstDeductionDate + 'T12:00:00').getTime());
      }
      firstDeduction.setHours(12, 0, 0, 0);
      
      // Calculate which payment number this payday represents
      const daysDiff = Math.round((paydayDate - firstDeduction) / (1000 * 60 * 60 * 24));
      
      // Check if this payday is a valid deduction date (every 14 days from first deduction)
      // Allow for small rounding errors (within 1 day)
      const periodsFromFirst = daysDiff / 14;
      const isValidPayday = daysDiff >= 0 && Math.abs(periodsFromFirst - Math.round(periodsFromFirst)) < 0.1;
      
      if (!isValidPayday) continue;
      
      const paymentNumber = Math.round(periodsFromFirst) + 1;
      
      // Check if this payment number is valid for this order's payment plan
      if (paymentNumber < 1 || paymentNumber > order.paymentPlan) continue;
      
      // Determine if this payment was already made or is still due
      const isPaid = paymentNumber <= order.paymentsMade;
      const isDue = paymentNumber > order.paymentsMade && order.status === 'Active';
      
      deductions.push({
        orderId: order.orderId,
        employeeId: order.employeeId,
        employeeName: order.employeeName,
        paymentNumber: paymentNumber,
        totalPayments: order.paymentPlan,
        amount: order.amountPerPaycheck,
        totalAmount: order.totalAmount,
        amountRemaining: order.amountRemaining,
        status: order.status,
        isPaid: isPaid,    // Was this payment already recorded?
        isDue: isDue       // Is this payment still due?
      });
    }
    
    // Sort by employee name
    deductions.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    
    const totalDeductions = deductions.reduce((sum, d) => sum + d.amount, 0);
    const paidDeductions = deductions.filter(d => d.isPaid);
    const dueDeductions = deductions.filter(d => d.isDue);
    
    return {
      success: true,
      payday: payday,
      deductions: deductions,
      count: deductions.length,
      total: parseFloat(totalDeductions.toFixed(2)),
      paidCount: paidDeductions.length,
      paidTotal: parseFloat(paidDeductions.reduce((sum, d) => sum + d.amount, 0).toFixed(2)),
      dueCount: dueDeductions.length,
      dueTotal: parseFloat(dueDeductions.reduce((sum, d) => sum + d.amount, 0).toFixed(2))
    };
    
  } catch (error) {
    console.error('Error getting payroll deductions:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Marks all deductions for a payday as paid
 * @param {string} payday - The payday date (YYYY-MM-DD)
 * @returns {Object} Result
 */
function markAllDeductionsPaid(payday) {
  try {
    const result = getPayrollDeductions(payday);
    
    if (!result.success) {
      return result;
    }
    
    let recorded = 0;
    const errors = [];
    
    for (const deduction of result.deductions) {
      const paymentResult = recordUniformPayment(deduction.orderId);
      if (paymentResult.success) {
        recorded++;
      } else {
        errors.push(`${deduction.employeeName}: ${paymentResult.error}`);
      }
    }
    
    return {
      success: true,
      message: `Recorded ${recorded} of ${result.deductions.length} payments`,
      recorded: recorded,
      errors: errors
    };
    
  } catch (error) {
    console.error('Error marking deductions paid:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets orders created in the last N days (for email summary)
 * @param {number} days - Number of days to look back
 * @returns {Array} Array of orders with items
 */
function getRecentOrders(days = 7) {
  try {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    cutoffDate.setHours(0, 0, 0, 0);
    
    const orders = getUniformOrders();
    const recentOrders = orders.filter(o => {
      const orderDate = new Date(o.orderDate);
      return orderDate >= cutoffDate;
    });
    
    // Get items for each order
    for (const order of recentOrders) {
      const details = getOrderDetails(order.orderId);
      if (details.success) {
        order.items = details.order.items;
      }
    }
    
    return recentOrders;
    
  } catch (error) {
    console.error('Error getting recent orders:', error);
    return [];
  }
}

/**
 * Sends weekly order summary email
 * This function should be set up with a weekly trigger (Saturday morning)
 */
function sendWeeklyOrderSummary() {
  try {
    const settings = getSettings();
    const recipients = settings.weeklyEmailRecipients;
    
    if (!recipients) {
      console.log('No email recipients configured');
      return { success: false, error: 'No email recipients configured' };
    }
    
    const orders = getRecentOrders(7);
    
    if (orders.length === 0) {
      console.log('No orders in the last week');
      return { success: true, message: 'No orders to report' };
    }
    
    // Calculate totals
    const totalOrders = orders.length;
    const totalAmount = orders.reduce((sum, o) => sum + o.totalAmount, 0);
    
    // Build email HTML
    const today = new Date();
    const weekAgo = new Date(today);
    weekAgo.setDate(weekAgo.getDate() - 7);
    
    const formatDate = (d) => new Date(d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    
    let html = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #E51636; color: white; padding: 20px; text-align: center;">
          <h1 style="margin: 0;">📦 Weekly Uniform Orders Summary</h1>
          <p style="margin: 10px 0 0 0;">Week of ${formatDate(weekAgo)} - ${formatDate(today)}</p>
        </div>
        
        <div style="background: #f5f5f5; padding: 15px; text-align: center;">
          <span style="font-size: 24px; font-weight: bold;">${totalOrders}</span> new order(s) totaling 
          <span style="font-size: 24px; font-weight: bold;">$${totalAmount.toFixed(2)}</span>
        </div>
    `;
    
    for (const order of orders) {
      const itemsList = (order.items || []).map(item => 
        `<li>${item.itemName} (${item.size}) x${item.quantity} - $${item.lineTotal.toFixed(2)}</li>`
      ).join('');
      
      const deductionInfo = order.totalAmount > 0 
        ? `${order.paymentPlan} payment(s) of $${order.amountPerPaycheck.toFixed(2)} starting ${formatDate(order.firstDeductionDate)}`
        : 'Store Paid';
      
      html += `
        <div style="border: 1px solid #ddd; margin: 15px; padding: 15px; background: white;">
          <div style="border-bottom: 1px solid #eee; padding-bottom: 10px; margin-bottom: 10px;">
            <strong style="color: #E51636;">Order #${order.orderId}</strong>
            <span style="float: right; color: #666;">${formatDate(order.orderDate)}</span>
          </div>
          <p style="margin: 5px 0;"><strong>Employee:</strong> ${order.employeeName}</p>
          <p style="margin: 5px 0;"><strong>Items:</strong></p>
          <ul style="margin: 5px 0;">${itemsList}</ul>
          <p style="margin: 10px 0 5px 0; padding-top: 10px; border-top: 1px solid #eee;">
            <strong>Total:</strong> $${order.totalAmount.toFixed(2)} | ${deductionInfo}
          </p>
        </div>
      `;
    }
    
    html += `
        <div style="background: #333; color: #999; padding: 15px; text-align: center; font-size: 12px;">
          This is an automated email from Payroll Review
        </div>
      </div>
    `;
    
    // Send email
    MailApp.sendEmail({
      to: recipients,
      subject: `📦 Weekly Uniform Orders: ${totalOrders} order(s), $${totalAmount.toFixed(2)}`,
      htmlBody: html
    });
    
    console.log(`Weekly summary sent to ${recipients}`);
    return { success: true, message: `Email sent to ${recipients}` };
    
  } catch (error) {
    console.error('Error sending weekly summary:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// DATA SAVING FUNCTIONS
// ============================================================================

/**
 * Saves OT data for a pay period
 * Also upserts employees and updates status based on 28-day rule
 * @param {Array} employees - Array of employee OT records (with employeeId if available)
 * @param {string} periodEnd - Period end date string
 * @param {boolean} overwrite - Whether to overwrite existing period data
 * @returns {Object} Result object
 */
function saveOTData(employees, periodEnd, overwrite) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    }
    
    const settings = getSettings();
    const importDate = new Date();
    const importedBy = Session.getActiveUser().getEmail() || 'Anonymous';
    
    // Check for existing period data
    const existingData = sheet.getDataRange().getValues();
    const periodDateStr = new Date(periodEnd).toDateString();
    
    let existingRows = [];
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] && new Date(existingData[i][0]).toDateString() === periodDateStr) {
        existingRows.push(i + 1); // 1-indexed row number
      }
    }
    
    if (existingRows.length > 0 && !overwrite) {
      return { 
        success: false, 
        error: 'duplicate',
        message: `Data for period ending ${periodEnd} already exists. Do you want to overwrite it?`,
        existingCount: existingRows.length
      };
    }
    
    // Delete existing rows if overwriting (delete from bottom to top to preserve row indices)
    if (existingRows.length > 0 && overwrite) {
      existingRows.sort((a, b) => b - a); // Sort descending
      for (const rowNum of existingRows) {
        sheet.deleteRow(rowNum);
      }
    }
    
    // ========== EMPLOYEE MANAGEMENT ==========
    // Upsert all employees in this batch
    const employeeDataForUpsert = employees.map(emp => ({
      employeeId: emp.employeeId || '',
      displayName: emp.employeeName,
      matchKey: emp.matchKey,
      location: emp.location === 'Multi' ? '' : emp.location // Don't set Multi as primary location
    }));
    
    const upsertResult = upsertEmployeesBatch(employeeDataForUpsert, periodEnd);
    console.log('Employee upsert result:', upsertResult);
    
    // Build a set of active match keys for status update
    const activeMatchKeys = new Set(employees.map(e => (e.matchKey || '').toLowerCase()));
    
    // ========== SAVE OT DATA ==========
    // Prepare new rows (now with 19 columns, including Employee_ID)
    const newRows = employees.map(emp => {
      const otCost = emp.totalOT * settings.hourlyWage * settings.otMultiplier;
      
      return [
        new Date(periodEnd),           // A: Period End
        emp.employeeName,              // B: Employee Name
        emp.matchKey,                  // C: Match Key (for deduplication)
        emp.location,                  // D: Location
        emp.chHours || 0,              // E: CH Hours
        emp.dbuHours || 0,             // F: DBU Hours
        emp.totalHours,                // G: Total Hours
        emp.regularHours || Math.min(emp.totalHours, 80),  // H: Regular Hours
        emp.week1Hours || 0,           // I: Week 1 Hours
        emp.week2Hours || 0,           // J: Week 2 Hours
        emp.week1OT || 0,              // K: Week 1 OT
        emp.week2OT || 0,              // L: Week 2 OT
        emp.totalOT,                   // M: Total OT
        parseFloat(otCost.toFixed(2)), // N: OT Cost
        emp.flag,                      // O: Flag
        emp.location === 'Multi',      // P: Is Multi-Location
        importDate,                    // Q: Import Date
        importedBy,                    // R: Imported By
        emp.employeeId || ''           // S: Employee_ID (new!)
      ];
    });
    
    // Append new rows
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
        .setValues(newRows);
    }
    
    // ========== UPDATE EMPLOYEE STATUSES ==========
    // Check for employees who haven't appeared in 28+ days
    updateEmployeeStatuses(periodEnd, activeMatchKeys);
    
    // ========== BACKFILL ==========
    // Try to backfill any missing Employee_IDs in older records
    backfillEmployeeIds();
    
    // Calculate totals for response
    const totalOT = employees.reduce((sum, e) => sum + e.totalOT, 0);
    const totalCost = totalOT * settings.hourlyWage * settings.otMultiplier;
    const multiCount = employees.filter(e => e.location === 'Multi').length;
    
    // Find high OT employees (Really High flag)
    const highOTEmployees = employees
      .filter(e => e.flag === 'Really High')
      .map(e => ({ name: e.employeeName, hours: e.totalOT }))
      .sort((a, b) => b.hours - a.hours);
    
    // Log the activity
    logActivity('CREATE', 'OT', 
      `OT data uploaded: ${employees.length} employees, ${totalOT.toFixed(1)} OT hrs, period ending ${periodEnd}`,
      periodEnd
    );
    
    // Send high OT notification if there are high OT employees
    if (highOTEmployees.length > 0) {
      try {
        sendHighOTNotification({
          count: highOTEmployees.length,
          employees: highOTEmployees,
          threshold: settings.reallyHighThreshold || 15
        });
      } catch (emailError) {
        console.error('Failed to send high OT notification email:', emailError);
        // Don't fail the request if email fails
      }
    }
    
    return {
      success: true,
      message: `Saved ${employees.length} employees for period ending ${periodEnd}`,
      stats: {
        employeeCount: employees.length,
        totalOT: parseFloat(totalOT.toFixed(2)),
        totalCost: parseFloat(totalCost.toFixed(2)),
        multiLocationCount: multiCount,
        wasOverwrite: existingRows.length > 0,
        employeesCreated: upsertResult.created || 0,
        employeesUpdated: upsertResult.updated || 0,
        highOTCount: highOTEmployees.length
      }
    };
    
  } catch (error) {
    console.error('Error saving OT data:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// DATA RETRIEVAL FUNCTIONS
// ============================================================================

/**
 * Gets all OT history data
 * @param {Object} filters - Optional filters (periodEnd, location, flag, employeeName)
 * @returns {Array} Array of OT records
 */
function getOTHistory(filters = {}) {
  try {
    console.log('getOTHistory called with filters:', JSON.stringify(filters));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!sheet) {
      console.log('OT_History sheet not found');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    console.log('Sheet last row:', lastRow);
    
    if (lastRow < 2) {
      console.log('No data rows found');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    console.log('Total rows including header:', data.length);
    
    const records = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // Skip empty rows
      
      // Convert dates to ISO strings for proper client/server serialization
      const periodEndDate = row[0] ? new Date(row[0]) : null;
      const importDateVal = row[16] ? new Date(row[16]) : null;
      
      const record = {
        periodEnd: periodEndDate ? periodEndDate.toISOString() : null,
        employeeName: row[1] || '',
        matchKey: row[2] || '',
        location: row[3] || '',
        chHours: parseFloat(row[4]) || 0,
        dbuHours: parseFloat(row[5]) || 0,
        totalHours: parseFloat(row[6]) || 0,
        regularHours: parseFloat(row[7]) || 0,
        week1Hours: parseFloat(row[8]) || 0,
        week2Hours: parseFloat(row[9]) || 0,
        week1OT: parseFloat(row[10]) || 0,
        week2OT: parseFloat(row[11]) || 0,
        totalOT: parseFloat(row[12]) || 0,
        otCost: parseFloat(row[13]) || 0,
        flag: row[14] || 'Normal',
        isMultiLocation: row[15] === true || row[15] === 'TRUE' || String(row[15]).toUpperCase() === 'TRUE',
        importDate: importDateVal ? importDateVal.toISOString() : null,
        importedBy: row[17] || '',
        employeeId: row[18] || ''  // Column S: Employee_ID
      };
      
      // Apply filters
      let include = true;
      
      if (filters.periodEnd) {
        // Normalize both dates to YYYY-MM-DD format for comparison
        const filterDate = new Date(filters.periodEnd);
        const recordDate = new Date(record.periodEnd);
        
        // Compare just the date parts (year, month, day)
        const filterKey = `${filterDate.getFullYear()}-${filterDate.getMonth()}-${filterDate.getDate()}`;
        const recordKey = `${recordDate.getFullYear()}-${recordDate.getMonth()}-${recordDate.getDate()}`;
        
        if (filterKey !== recordKey) include = false;
      }
      
      if (filters.location && filters.location !== 'All') {
        if (filters.location === 'Multi' && !record.isMultiLocation) include = false;
        else if (filters.location !== 'Multi' && record.location !== filters.location) include = false;
      }
      
      if (filters.flag && filters.flag !== 'All') {
        if (record.flag !== filters.flag) include = false;
      }
      
      if (filters.employeeName) {
        if (!record.employeeName.toLowerCase().includes(filters.employeeName.toLowerCase())) {
          include = false;
        }
      }
      
      if (include) {
        records.push(record);
      }
    }
    
    // Sort by period end (newest first), then by total OT (highest first)
    records.sort((a, b) => {
      const dateCompare = new Date(b.periodEnd) - new Date(a.periodEnd);
      if (dateCompare !== 0) return dateCompare;
      return b.totalOT - a.totalOT;
    });
    
    return records;
    
  } catch (error) {
    console.error('Error getting OT history:', error);
    return [];
  }
}

/**
 * Gets list of unique pay periods
 * @returns {Array} Array of period end date strings (ISO format for proper serialization)
 */
function getPayPeriods() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const periods = new Set();
    
    for (const row of data) {
      if (row[0]) {
        // Convert to timestamp for deduplication
        const d = new Date(row[0]);
        if (!isNaN(d.getTime())) {
          periods.add(d.getTime());
        }
      }
    }
    
    // Convert to sorted array (newest first) and return as ISO strings
    // This ensures proper serialization across client/server boundary
    return Array.from(periods)
      .sort((a, b) => b - a)
      .map(t => new Date(t).toISOString());
      
  } catch (error) {
    console.error('Error getting pay periods:', error);
    return [];
  }
}

/**
 * Gets data for a specific pay period
 * @param {string} periodEnd - Period end date
 * @returns {Object} Period data with employees and stats
 */
function getPeriodData(periodEnd) {
  try {
    console.log('getPeriodData called with:', periodEnd);
    const allEmployees = getOTHistory({ periodEnd: periodEnd });
    const settings = getSettings();
    
    // Filter out employees with 0 total hours - they shouldn't show up
    const employees = allEmployees.filter(e => (e.totalHours || 0) > 0);
    
    console.log('Found employees with hours:', employees.length, '(filtered from', allEmployees.length, ')');
    
    if (employees.length === 0) {
      return { success: false, error: 'No data found for this period' };
    }
    
    // Calculate period stats
    const totalOT = employees.reduce((sum, e) => sum + (e.totalOT || 0), 0);
    const totalCost = employees.reduce((sum, e) => sum + (e.otCost || 0), 0);
    const highOTCount = employees.filter(e => e.flag === 'Really High').length;
    const moderateOTCount = employees.filter(e => e.flag === 'Moderate').length;
    const multiCount = employees.filter(e => e.isMultiLocation).length;
    
    return {
      success: true,
      periodEnd: periodEnd,
      employees: employees,
      stats: {
        employeeCount: employees.length,
        totalOT: parseFloat(totalOT.toFixed(2)),
        totalCost: parseFloat(totalCost.toFixed(2)),
        highOTCount: highOTCount,
        moderateOTCount: moderateOTCount,
        multiLocationCount: multiCount,
        avgOT: employees.length > 0 ? parseFloat((totalOT / employees.length).toFixed(2)) : 0
      }
    };
    
  } catch (error) {
    console.error('Error getting period data:', error);
    return { success: false, error: error.message || 'Unknown error' };
  }
}

/**
 * Gets historical data for a specific employee
 * @param {string} employeeName - Employee name to search
 * @returns {Object} Employee history data
 */
function getEmployeeHistory(employeeName) {
  try {
    const allRecords = getOTHistory({ employeeName: employeeName });
    const settings = getSettings();
    
    if (allRecords.length === 0) {
      return { success: false, error: 'No records found for this employee' };
    }
    
    // Group by employee name (exact match)
    const exactMatches = allRecords.filter(r => 
      r.employeeName.toLowerCase() === employeeName.toLowerCase()
    );
    
    if (exactMatches.length === 0) {
      return { 
        success: true, 
        partialMatches: allRecords.slice(0, 10),
        message: 'No exact match found. Did you mean one of these?'
      };
    }
    
    // Calculate aggregates
    const totalOT = exactMatches.reduce((sum, r) => sum + r.totalOT, 0);
    const totalCost = exactMatches.reduce((sum, r) => sum + r.otCost, 0);
    const avgOT = totalOT / exactMatches.length;
    const multiLocationPeriods = exactMatches.filter(r => r.isMultiLocation).length;
    const highOTPeriods = exactMatches.filter(r => r.flag === 'Really High').length;
    
    // Check for consecutive high OT
    const sortedRecords = exactMatches.sort((a, b) => 
      new Date(b.periodEnd) - new Date(a.periodEnd)
    );
    
    let consecutiveHigh = 0;
    for (const record of sortedRecords) {
      if (record.flag === 'Really High') {
        consecutiveHigh++;
      } else {
        break;
      }
    }
    
    return {
      success: true,
      employeeName: exactMatches[0].employeeName,
      records: sortedRecords,
      stats: {
        periodsWorked: exactMatches.length,
        totalOT: parseFloat(totalOT.toFixed(2)),
        totalCost: parseFloat(totalCost.toFixed(2)),
        avgOTPerPeriod: parseFloat(avgOT.toFixed(2)),
        multiLocationPeriods: multiLocationPeriods,
        highOTPeriods: highOTPeriods,
        consecutiveHighOT: consecutiveHigh
      },
      alerts: consecutiveHigh >= settings.consecutiveAlertPeriods ? 
        [`⚠️ ${consecutiveHigh} consecutive periods of Really High OT`] : []
    };
    
  } catch (error) {
    console.error('Error getting employee history:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets list of all unique employee names
 * @returns {Array} Array of employee names
 */
function getAllEmployeeNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    // Get the last 4 pay periods
    const periods = getPayPeriods();
    const recentPeriods = periods.slice(0, 4); // Most recent 4
    const recentPeriodTimes = new Set(recentPeriods.map(p => new Date(p).getTime()));
    
    // Column layout: 1=Period End, 2=Name, 3=Match Key, 4=Location, 5=CH Hours, 6=DBU Hours, 7=Total Hours
    // Get columns 1, 2, and 7 (period, name, total hours)
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const activeNames = new Set();
    
    for (const row of data) {
      const periodEnd = row[0];
      const name = row[1];
      const totalHours = parseFloat(row[6]) || 0;  // Column 7 = index 6
      
      if (!name || !periodEnd) continue;
      
      // Check if this record is from a recent period AND has hours
      const periodTime = new Date(periodEnd).getTime();
      if (recentPeriodTimes.has(periodTime) && totalHours > 0) {
        activeNames.add(name);
      }
    }
    
    console.log('Active employees (last 4 periods with hours):', activeNames.size);
    return Array.from(activeNames).sort();
    
  } catch (error) {
    console.error('Error getting employee names:', error);
    return [];
  }
}

// ============================================================================
// ANALYTICS & TRENDS
// ============================================================================

/**
 * Gets trend data for charts
 * @param {string} timeRange - '3m', '6m', '1y', 'all'
 * @returns {Object} Trend data
 */
function getTrendData(timeRange) {
  try {
    console.log('getTrendData called with range:', timeRange);
    
    const allRecords = getOTHistory();
    const settings = getSettings();
    
    if (allRecords.length === 0) {
      return { success: false, error: 'No historical data available' };
    }
    
    console.log('Total records in history:', allRecords.length);
    
    // Filter by time range - create fresh Date for each calculation
    let cutoffDate = null;
    const now = new Date();
    
    switch (timeRange) {
      case '3m':
        cutoffDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
        break;
      case '6m':
        cutoffDate = new Date(now.getFullYear(), now.getMonth() - 6, now.getDate());
        break;
      case '1y':
        cutoffDate = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
        break;
      case 'all':
      default:
        cutoffDate = null; // All time
    }
    
    console.log('Cutoff date:', cutoffDate);
    
    const filteredRecords = cutoffDate 
      ? allRecords.filter(r => new Date(r.periodEnd) >= cutoffDate)
      : allRecords;
    
    console.log('Filtered records count:', filteredRecords.length);
    
    // Group by period
    const periodMap = new Map();
    for (const record of filteredRecords) {
      const periodKey = new Date(record.periodEnd).toISOString().split('T')[0];
      
      if (!periodMap.has(periodKey)) {
        periodMap.set(periodKey, {
          periodEnd: record.periodEnd,
          employees: [],
          totalOT: 0,
          totalCost: 0,
          multiCount: 0
        });
      }
      
      const period = periodMap.get(periodKey);
      period.employees.push(record);
      period.totalOT += record.totalOT;
      period.totalCost += record.otCost;
      if (record.isMultiLocation) period.multiCount++;
    }
    
    // Convert to array sorted by date
    const periodData = Array.from(periodMap.values())
      .sort((a, b) => new Date(a.periodEnd) - new Date(b.periodEnd));
    
    // Calculate monthly aggregates
    const monthlyMap = new Map();
    for (const record of filteredRecords) {
      const date = new Date(record.periodEnd);
      const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      
      if (!monthlyMap.has(monthKey)) {
        monthlyMap.set(monthKey, {
          month: monthKey,
          totalOT: 0,
          totalCost: 0,
          employeeCount: new Set()
        });
      }
      
      const month = monthlyMap.get(monthKey);
      month.totalOT += record.totalOT;
      month.totalCost += record.otCost;
      month.employeeCount.add(record.employeeName);
    }
    
    const monthlyData = Array.from(monthlyMap.entries())
      .map(([key, value]) => ({
        month: key,
        totalOT: parseFloat(value.totalOT.toFixed(2)),
        totalCost: parseFloat(value.totalCost.toFixed(2)),
        uniqueEmployees: value.employeeCount.size
      }))
      .sort((a, b) => a.month.localeCompare(b.month));
    
    // Top OT employees (aggregate)
    const employeeMap = new Map();
    for (const record of filteredRecords) {
      if (!employeeMap.has(record.employeeName)) {
        employeeMap.set(record.employeeName, {
          employeeName: record.employeeName,
          totalOT: 0,
          totalCost: 0,
          periods: 0,
          multiLocationPeriods: 0
        });
      }
      
      const emp = employeeMap.get(record.employeeName);
      emp.totalOT += record.totalOT;
      emp.totalCost += record.otCost;
      emp.periods++;
      if (record.isMultiLocation) emp.multiLocationPeriods++;
    }
    
    const topEmployees = Array.from(employeeMap.values())
      .map(e => ({
        ...e,
        totalOT: parseFloat(e.totalOT.toFixed(2)),
        totalCost: parseFloat(e.totalCost.toFixed(2)),
        avgOT: parseFloat((e.totalOT / e.periods).toFixed(2))
      }))
      .sort((a, b) => b.totalOT - a.totalOT)
      .slice(0, 15);
    
    // Location comparison - use actual hours data to determine location
    // This is more accurate than the location field which may be incorrect from bulk imports
    let chTotalOT = 0, dbuTotalOT = 0, multiTotalOT = 0;
    for (const record of filteredRecords) {
      const hasChHours = (record.chHours || 0) > 0;
      const hasDbuHours = (record.dbuHours || 0) > 0;
      
      if (hasChHours && hasDbuHours) {
        // Multi-location employee
        multiTotalOT += record.totalOT;
      } else if (hasDbuHours) {
        // DBU-only employee
        dbuTotalOT += record.totalOT;
      } else {
        // CH-only employee (or default)
        chTotalOT += record.totalOT;
      }
    }
    
    // Multi-location trend
    const multiTrend = periodData.map(p => ({
      periodEnd: p.periodEnd,
      count: p.multiCount,
      percentage: parseFloat(((p.multiCount / p.employees.length) * 100).toFixed(1))
    }));
    
    return {
      success: true,
      timeRange: timeRange,
      periodData: periodData.map(p => ({
        periodEnd: p.periodEnd,
        totalOT: parseFloat(p.totalOT.toFixed(2)),
        totalCost: parseFloat(p.totalCost.toFixed(2)),
        employeeCount: p.employees.length,
        multiCount: p.multiCount
      })),
      monthlyData: monthlyData,
      topEmployees: topEmployees,
      locationComparison: {
        [settings.location1Name]: parseFloat(chTotalOT.toFixed(2)),
        [settings.location2Name]: parseFloat(dbuTotalOT.toFixed(2)),
        'Multi-Location': parseFloat(multiTotalOT.toFixed(2))
      },
      multiLocationTrend: multiTrend,
      totals: {
        totalOT: parseFloat(filteredRecords.reduce((s, r) => s + r.totalOT, 0).toFixed(2)),
        totalCost: parseFloat(filteredRecords.reduce((s, r) => s + r.otCost, 0).toFixed(2)),
        periodsCount: periodData.length,
        uniqueEmployees: new Set(filteredRecords.map(r => r.employeeName)).size
      }
    };
    
  } catch (error) {
    console.error('Error getting trend data:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets multi-location transfer data for a specific month
 * Aggregates by employee for display (so each employee appears once with their total monthly hours)
 * But calculates transfer totals per-record for accurate payroll math
 * @param {string} month - Month in YYYY-MM format
 * @returns {Object} Transfer data for the month
 */
function getMonthlyTransferData(month) {
  try {
    console.log('getMonthlyTransferData called for month:', month);
    
    const allRecords = getOTHistory();
    
    // Filter to records in the specified month that are multi-location
    const monthRecords = allRecords.filter(r => {
      if (!r.periodEnd) return false;
      const d = new Date(r.periodEnd);
      const recordMonth = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      
      // Check if multi-location based on hours (more accurate than isMultiLocation flag)
      const hasChHours = (r.chHours || 0) > 0;
      const hasDbuHours = (r.dbuHours || 0) > 0;
      
      return recordMonth === month && hasChHours && hasDbuHours;
    });
    
    console.log('Found multi-location records for month:', monthRecords.length);
    
    // Calculate transfer totals PER RECORD (original logic - more accurate for payroll)
    // This determines transfer direction based on each pay period, not monthly aggregate
    let toCH = 0;
    let toDBU = 0;
    
    for (const r of monthRecords) {
      const chHours = r.chHours || 0;
      const dbuHours = r.dbuHours || 0;
      
      // Payout is to location with more hours for THIS pay period
      // Transfer is from the other location
      if (chHours > dbuHours) {
        toCH += dbuHours; // DBU hours transfer to CH
      } else {
        toDBU += chHours; // CH hours transfer to DBU
      }
    }
    
    // Aggregate records by employee name FOR DISPLAY ONLY
    const employeeMap = {};
    for (const r of monthRecords) {
      const name = r.employeeName;
      if (!employeeMap[name]) {
        employeeMap[name] = {
          employeeName: name,
          chHours: 0,
          dbuHours: 0,
          totalHours: 0,
          totalOT: 0,
          payPeriodCount: 0
        };
      }
      employeeMap[name].chHours += r.chHours || 0;
      employeeMap[name].dbuHours += r.dbuHours || 0;
      employeeMap[name].totalHours += r.totalHours || 0;
      employeeMap[name].totalOT += r.totalOT || 0;
      employeeMap[name].payPeriodCount++;
    }
    
    // Build employee list with rounded values
    const employees = Object.values(employeeMap).map(emp => {
      emp.chHours = parseFloat(emp.chHours.toFixed(2));
      emp.dbuHours = parseFloat(emp.dbuHours.toFixed(2));
      emp.totalHours = parseFloat(emp.totalHours.toFixed(2));
      emp.totalOT = parseFloat(emp.totalOT.toFixed(2));
      return emp;
    });
    
    // Sort by total hours descending
    employees.sort((a, b) => b.totalHours - a.totalHours);
    
    // Calculate NET transfer (larger minus smaller)
    // This shows what one restaurant owes the other after all transfers cancel out
    const netTransfer = Math.abs(toDBU - toCH);
    const netDirection = toDBU > toCH ? 'DBU' : (toCH > toDBU ? 'CH' : 'EVEN');
    
    return {
      success: true,
      month: month,
      employees: employees,
      toCH: parseFloat(toCH.toFixed(2)),
      toDBU: parseFloat(toDBU.toFixed(2)),
      netTransfer: parseFloat(netTransfer.toFixed(2)),
      netDirection: netDirection, // Which restaurant is OWED the net amount
      totalTransferred: parseFloat(netTransfer.toFixed(2)) // Now shows NET, not gross
    };
    
  } catch (error) {
    console.error('Error getting monthly transfer data:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// ALERTS
// ============================================================================

/**
 * Gets current alerts based on data patterns
 * @returns {Array} Array of alert objects
 */
function getAlerts() {
  try {
    const allRecords = getOTHistory();
    const settings = getSettings();
    const alerts = [];
    
    if (allRecords.length === 0) {
      return [];
    }
    
    // Get periods for comparison (these are ISO strings)
    const periods = getPayPeriods();
    if (periods.length < 2) {
      return [];
    }
    
    // Convert period strings to timestamps for comparison
    const currentPeriodTime = new Date(periods[0]).getTime();
    const previousPeriodTime = new Date(periods[1]).getTime();
    
    // Filter records by period (handle both string and Date periodEnd)
    const currentData = allRecords.filter(r => {
      const recordTime = new Date(r.periodEnd).getTime();
      return recordTime === currentPeriodTime;
    });
    const previousData = allRecords.filter(r => {
      const recordTime = new Date(r.periodEnd).getTime();
      return recordTime === previousPeriodTime;
    });
    
    console.log('Alert check - Current period records:', currentData.length);
    console.log('Alert check - Previous period records:', previousData.length);
    
    // Alert 1: Consecutive high OT employees (need 3+ consecutive periods with "Really High")
    const employeeMap = new Map();
    for (const record of allRecords) {
      if (!employeeMap.has(record.employeeName)) {
        employeeMap.set(record.employeeName, []);
      }
      employeeMap.get(record.employeeName).push(record);
    }
    
    const consecutiveThreshold = settings.consecutiveAlertPeriods || 3;
    
    for (const [name, records] of employeeMap) {
      // Sort by date descending (newest first)
      const sorted = records.sort((a, b) => new Date(b.periodEnd) - new Date(a.periodEnd));
      let consecutive = 0;
      
      for (const record of sorted) {
        if (record.flag === 'Really High') {
          consecutive++;
        } else {
          break;
        }
      }
      
      if (consecutive >= consecutiveThreshold) {
        alerts.push({
          type: 'consecutive_high',
          severity: 'high',
          title: 'Consecutive High OT',
          message: `${name} has had Really High OT for ${consecutive} consecutive periods`,
          employeeName: name,
          value: consecutive
        });
      }
    }
    
    // Alert 2: Period-over-period OT increase (triggers if OT increased by X% or more)
    const increaseThreshold = settings.monthlyIncreaseAlertPercent || 25;
    
    if (currentData.length > 0 && previousData.length > 0) {
      const currentTotal = currentData.reduce((s, r) => s + (r.totalOT || 0), 0);
      const previousTotal = previousData.reduce((s, r) => s + (r.totalOT || 0), 0);
      
      console.log('Alert check - Current OT total:', currentTotal, 'Previous OT total:', previousTotal);
      
      if (previousTotal > 0) {
        const percentChange = ((currentTotal - previousTotal) / previousTotal) * 100;
        
        if (percentChange >= increaseThreshold) {
          alerts.push({
            type: 'period_increase',
            severity: 'warning',
            title: 'OT Increased',
            message: `Total OT increased ${percentChange.toFixed(0)}% from previous period`,
            value: percentChange.toFixed(1)
          });
        }
      }
    }
    
    // Alert 3: New high OT employees (someone went from Normal/Moderate to Really High)
    if (currentData.length > 0 && previousData.length > 0) {
      const previousHighNames = new Set(
        previousData.filter(r => r.flag === 'Really High').map(r => r.employeeName)
      );
      
      const newHighOT = currentData.filter(r => 
        r.flag === 'Really High' && !previousHighNames.has(r.employeeName)
      );
      
      if (newHighOT.length > 0) {
        alerts.push({
          type: 'new_high_ot',
          severity: 'warning',
          title: 'New High OT',
          message: `${newHighOT.length} employee(s) have Really High OT who didn't before: ${newHighOT.slice(0, 3).map(r => r.employeeName).join(', ')}${newHighOT.length > 3 ? '...' : ''}`,
          employees: newHighOT.map(r => r.employeeName),
          value: newHighOT.length
        });
      }
    }
    
    // Alert 4: Multi-location increase (50% more employees working both locations)
    if (currentData.length > 0 && previousData.length > 0) {
      const currentMulti = currentData.filter(r => r.isMultiLocation).length;
      const previousMulti = previousData.filter(r => r.isMultiLocation).length;
      
      if (previousMulti > 0 && currentMulti > previousMulti * 1.5) {
        alerts.push({
          type: 'multi_increase',
          severity: 'info',
          title: 'Multi-Location Increase',
          message: `Multi-location employees increased from ${previousMulti} to ${currentMulti}`,
          value: currentMulti
        });
      }
    }
    
    // Alert 5: High OT count in current period (if more than 5 employees have Really High OT)
    const currentHighCount = currentData.filter(r => r.flag === 'Really High').length;
    if (currentHighCount >= 5) {
      alerts.push({
        type: 'high_ot_count',
        severity: 'warning',
        title: 'Many High OT Employees',
        message: `${currentHighCount} employees have Really High OT this period`,
        value: currentHighCount
      });
    }
    
    // Alert 6: New employees needing review (created via uniform orders)
    try {
      const needsReview = getEmployeesNeedingReview();
      if (needsReview && needsReview.length > 0) {
        alerts.push({
          type: 'employees_need_review',
          severity: 'info',
          title: 'New Employees to Review',
          message: `${needsReview.length} employee(s) were created via uniform orders and should be verified: ${needsReview.slice(0, 3).map(e => e.fullName).join(', ')}${needsReview.length > 3 ? '...' : ''}`,
          employees: needsReview.map(e => e.fullName),
          value: needsReview.length,
          link: 'settings' // Indicates clicking should go to settings
        });
      }
    } catch (e) {
      console.log('Could not check employees needing review:', e);
    }
    
    // Alert 7: Potential duplicate employees detected
    try {
      const duplicates = scanForDuplicateEmployees();
      if (duplicates && duplicates.length > 0) {
        alerts.push({
          type: 'potential_duplicates',
          severity: 'warning',
          title: 'Potential Duplicate Employees',
          message: `${duplicates.length} potential duplicate(s) found. Review in Settings to merge if needed.`,
          value: duplicates.length,
          link: 'settings'
        });
      }
    } catch (e) {
      console.log('Could not check for duplicates:', e);
    }
    
    console.log('Alert check - Total alerts generated:', alerts.length);
    
    return alerts;
    
  } catch (error) {
    console.error('Error getting alerts:', error);
    return [];
  }
}

// ============================================================================
// DATA MANAGEMENT
// ============================================================================

/**
 * Deletes all data for a specific pay period
 * @param {string} periodEnd - Period end date to delete
 * @returns {Object} Result object
 */
function deletePeriod(periodEnd) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!sheet) {
      return { success: false, error: 'No data sheet found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const periodDateStr = new Date(periodEnd).toDateString();
    
    // Find rows to delete (from bottom to top)
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && new Date(data[i][0]).toDateString() === periodDateStr) {
        rowsToDelete.push(i + 1);
      }
    }
    
    if (rowsToDelete.length === 0) {
      return { success: false, error: 'No data found for this period' };
    }
    
    // Delete rows from bottom to top
    rowsToDelete.sort((a, b) => b - a);
    for (const rowNum of rowsToDelete) {
      sheet.deleteRow(rowNum);
    }
    
    return { 
      success: true, 
      message: `Deleted ${rowsToDelete.length} records for period ending ${periodEnd}`
    };
    
  } catch (error) {
    console.error('Error deleting period:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Exports all data to CSV format
 * @returns {string} CSV content
 */
function exportToCSV() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return '';
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Convert to CSV
    const csv = data.map(row => 
      row.map(cell => {
        if (cell instanceof Date) {
          return cell.toISOString().split('T')[0];
        }
        if (typeof cell === 'string' && cell.includes(',')) {
          return `"${cell}"`;
        }
        return cell;
      }).join(',')
    ).join('\n');
    
    return csv;
    
  } catch (error) {
    console.error('Error exporting to CSV:', error);
    return '';
  }
}

// ============================================================================
// PDF EXPORT
// ============================================================================

/**
 * Generates OT-specific PDF reports (period, employee, trends)
 * @param {string} type - 'period', 'employee', 'trends'
 * @param {Object} params - Parameters for the report
 * @returns {Object} Result with PDF URL or error
 */
function generateOTPDFReport(type, params) {
  try {
    const settings = getSettings();
    let html = '';
    let title = '';
    
    switch (type) {
      case 'period':
        const periodData = getPeriodData(params.periodEnd);
        if (!periodData.success) {
          return { success: false, error: periodData.error };
        }
        title = `OT Report - Period Ending ${formatDateForDisplay(params.periodEnd)}`;
        html = generatePeriodReportHTML(periodData, settings);
        break;
        
      case 'employee':
        const empData = getEmployeeHistory(params.employeeName);
        if (!empData.success) {
          return { success: false, error: empData.error };
        }
        title = `OT Report - ${params.employeeName}`;
        html = generateEmployeeReportHTML(empData, settings);
        break;
        
      case 'trends':
        const trendData = getTrendData(params.timeRange || '6m');
        if (!trendData.success) {
          return { success: false, error: trendData.error };
        }
        title = `OT Trends Report - ${params.timeRange || '6 Months'}`;
        html = generateTrendsReportHTML(trendData, settings);
        break;
        
      default:
        return { success: false, error: 'Unknown report type' };
    }
    
    // Create PDF
    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName(title + '.pdf')
      .getAs('application/pdf');
    
    // Save to Drive and get URL
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();
    
    return {
      success: true,
      url: url,
      fileUrl: url, // Include both for compatibility
      title: title
    };
    
  } catch (error) {
    console.error('Error generating OT PDF:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Helper function to format date for display
 */
function formatDateForDisplay(date) {
  const d = new Date(date);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

/**
 * Converts decimal hours to HH:MM format
 * Example: 30.5 -> "30:30", 8.25 -> "8:15", 6.75 -> "6:45"
 * @param {number} decimalHours - Hours in decimal format
 * @returns {string} Hours in HH:MM format
 */
function formatHoursDisplay(decimalHours) {
  if (!decimalHours || decimalHours === 0) {
    return '0:00';
  }
  const hours = Math.floor(Math.abs(decimalHours));
  const minutes = Math.round((Math.abs(decimalHours) - hours) * 60);
  const sign = decimalHours < 0 ? '-' : '';
  return sign + hours + ':' + (minutes < 10 ? '0' : '') + minutes;
}

/**
 * Generates Time Transfer PDF report
 * @param {Array} data - Array of transfer records
 * @param {string} periodStr - Period string for title
 * @returns {Object} Result with PDF URL or error
 */
function generateTimeTransferPDF(data, periodStr) {
  try {
    // Get configured location names from settings
    const settings = getSettings();
    const loc1Name = settings.location1Name || 'Location 1';
    const loc2Name = settings.location2Name || 'Location 2';
    
    // Calculate totals
    let totalTransferred = 0;
    let toLoc1Hours = 0;
    let toLoc2Hours = 0;
    
    data.forEach(function(r) {
      totalTransferred += r.transferredHours || 0;
      // Check if payout location matches location 1 (partial match)
      const payoutLower = (r.payoutLocation || '').toLowerCase();
      const loc1Lower = loc1Name.toLowerCase();
      if (payoutLower.includes(loc1Lower.substring(0, Math.min(5, loc1Lower.length))) ||
          loc1Lower.includes(payoutLower.substring(0, Math.min(5, payoutLower.length)))) {
        toLoc1Hours += r.transferredHours || 0;
      } else {
        toLoc2Hours += r.transferredHours || 0;
      }
    });
    
    const tableRows = data.map(function(r) {
      return '<tr>' +
        '<td>' + r.employeeName + '</td>' +
        '<td>' + r.payoutLocation + '</td>' +
        '<td style="text-align: right;">' + formatHoursDisplay(r.chHours) + '</td>' +
        '<td style="text-align: right;">' + formatHoursDisplay(r.dbuHours) + '</td>' +
        '<td style="text-align: right; font-weight: bold;">' + formatHoursDisplay(r.transferredHours) + '</td>' +
        '<td>' + r.transferDirection + '</td>' +
      '</tr>';
    }).join('');
    
    const html = '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
        '<style>' +
          'body { font-family: Arial, sans-serif; padding: 40px; }' +
          'h1 { color: #E51636; margin-bottom: 5px; }' +
          '.subtitle { color: #666; margin-bottom: 30px; }' +
          '.stats { display: flex; gap: 40px; margin-bottom: 30px; padding: 20px; background: #f5f5f5; border-radius: 8px; }' +
          '.stat { text-align: center; }' +
          '.stat-value { font-size: 24px; font-weight: bold; color: #E51636; }' +
          '.stat-label { font-size: 12px; color: #666; text-transform: uppercase; }' +
          'table { width: 100%; border-collapse: collapse; margin-top: 20px; }' +
          'th { background: #E51636; color: white; padding: 12px 8px; text-align: left; font-size: 12px; }' +
          'td { padding: 10px 8px; border-bottom: 1px solid #eee; font-size: 11px; }' +
          'tr:nth-child(even) { background: #f9f9f9; }' +
          '.footer { margin-top: 30px; font-size: 10px; color: #999; text-align: center; }' +
          '.transfer-arrow { color: #E51636; font-weight: bold; }' +
        '</style>' +
      '</head>' +
      '<body>' +
        '<h1>Time Transfer Report</h1>' +
        '<p class="subtitle">Period: ' + periodStr + '</p>' +
        
        '<div class="stats">' +
          '<div class="stat">' +
            '<div class="stat-value">' + data.length + '</div>' +
            '<div class="stat-label">Multi-Location Employees</div>' +
          '</div>' +
          '<div class="stat">' +
            '<div class="stat-value">' + formatHoursDisplay(toLoc1Hours) + '</div>' +
            '<div class="stat-label">Hours → ' + loc1Name + '</div>' +
          '</div>' +
          '<div class="stat">' +
            '<div class="stat-value">' + formatHoursDisplay(toLoc2Hours) + '</div>' +
            '<div class="stat-label">Hours → ' + loc2Name + '</div>' +
          '</div>' +
          '<div class="stat">' +
            '<div class="stat-value">' + formatHoursDisplay(totalTransferred) + '</div>' +
            '<div class="stat-label">Total Transferred</div>' +
          '</div>' +
        '</div>' +
        
        '<table>' +
          '<thead>' +
            '<tr>' +
              '<th>Employee Name</th>' +
              '<th>Pay-out Location</th>' +
              '<th style="text-align: right;">' + loc1Name + ' Hours</th>' +
              '<th style="text-align: right;">' + loc2Name + ' Hours</th>' +
              '<th style="text-align: right;">Transferred</th>' +
              '<th>Direction</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' + tableRows + '</tbody>' +
        '</table>' +
        
        '<div class="footer">Generated on ' + new Date().toLocaleString() + '</div>' +
      '</body>' +
      '</html>';
    
    // Create PDF
    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName('Time_Transfer_' + periodStr.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf')
      .getAs('application/pdf');
    
    // Save to Drive and get URL
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      url: file.getUrl()
    };
    
  } catch (error) {
    console.error('Error generating time transfer PDF:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Generates HTML for period report PDF
 */
function generatePeriodReportHTML(data, settings) {
  const employees = data.employees;
  const stats = data.stats;
  
  let tableRows = employees.map(e => `
    <tr>
      <td>${e.employeeName}</td>
      <td>${e.location}</td>
      <td style="text-align: right;">${formatHoursDisplay(e.totalHours)}</td>
      <td style="text-align: right;">${formatHoursDisplay(e.totalOT)}</td>
      <td style="text-align: right;">$${e.otCost.toFixed(2)}</td>
      <td><span class="flag ${e.flag.toLowerCase().replace(' ', '-')}">${e.flag}</span></td>
    </tr>
  `).join('');
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 40px; }
        h1 { color: #E51636; margin-bottom: 5px; }
        .subtitle { color: #666; margin-bottom: 30px; }
        .stats { display: flex; gap: 20px; margin-bottom: 30px; }
        .stat { background: #f5f5f5; padding: 15px 25px; border-radius: 8px; }
        .stat-value { font-size: 24px; font-weight: bold; color: #333; }
        .stat-label { font-size: 12px; color: #666; text-transform: uppercase; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f5f5f5; font-weight: bold; font-size: 11px; text-transform: uppercase; }
        .flag { padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: bold; }
        .flag.normal { background: #ECFDF5; color: #059669; }
        .flag.moderate { background: #FFFBEB; color: #D97706; }
        .flag.really-high { background: #FEF2F2; color: #DC2626; }
        .footer { margin-top: 40px; font-size: 11px; color: #999; }
      </style>
    </head>
    <body>
      <h1>OT Report</h1>
      <p class="subtitle">Pay Period Ending ${formatDateForDisplay(data.periodEnd)}</p>
      
      <div class="stats">
        <div class="stat">
          <div class="stat-value">${stats.employeeCount}</div>
          <div class="stat-label">Employees</div>
        </div>
        <div class="stat">
          <div class="stat-value">${formatHoursDisplay(stats.totalOT)}</div>
          <div class="stat-label">Total OT Hours</div>
        </div>
        <div class="stat">
          <div class="stat-value">$${stats.totalCost.toFixed(2)}</div>
          <div class="stat-label">Total OT Cost</div>
        </div>
        <div class="stat">
          <div class="stat-value">${stats.highOTCount}</div>
          <div class="stat-label">High OT Alerts</div>
        </div>
      </div>
      
      <table>
        <thead>
          <tr>
            <th>Employee</th>
            <th>Location</th>
            <th style="text-align: right;">Total Hours</th>
            <th style="text-align: right;">OT Hours</th>
            <th style="text-align: right;">OT Cost</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
      
      <div class="footer">
        Generated on ${new Date().toLocaleString()} | Hourly Rate: $${settings.hourlyWage} | OT Multiplier: ${settings.otMultiplier}x
      </div>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for employee report PDF
 */
function generateEmployeeReportHTML(data, settings) {
  const stats = data.stats;
  
  let tableRows = data.records.map(r => `
    <tr>
      <td>${formatDateForDisplay(r.periodEnd)}</td>
      <td>${r.location}</td>
      <td style="text-align: right;">${formatHoursDisplay(r.totalHours)}</td>
      <td style="text-align: right;">${formatHoursDisplay(r.totalOT)}</td>
      <td style="text-align: right;">$${r.otCost.toFixed(2)}</td>
      <td><span class="flag ${r.flag.toLowerCase().replace(' ', '-')}">${r.flag}</span></td>
    </tr>
  `).join('');
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 40px; }
        h1 { color: #E51636; margin-bottom: 5px; }
        .subtitle { color: #666; margin-bottom: 30px; }
        .stats { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 30px; }
        .stat { background: #f5f5f5; padding: 12px 20px; border-radius: 8px; }
        .stat-value { font-size: 20px; font-weight: bold; color: #333; }
        .stat-label { font-size: 11px; color: #666; text-transform: uppercase; }
        .alert { background: #FEF2F2; border-left: 4px solid #DC2626; padding: 12px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f5f5f5; font-weight: bold; font-size: 11px; text-transform: uppercase; }
        .flag { padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: bold; }
        .flag.normal { background: #ECFDF5; color: #059669; }
        .flag.moderate { background: #FFFBEB; color: #D97706; }
        .flag.really-high { background: #FEF2F2; color: #DC2626; }
        .footer { margin-top: 40px; font-size: 11px; color: #999; }
      </style>
    </head>
    <body>
      <h1>${data.employeeName}</h1>
      <p class="subtitle">Employee OT History</p>
      
      ${data.alerts.length > 0 ? `<div class="alert">${data.alerts.join('<br>')}</div>` : ''}
      
      <div class="stats">
        <div class="stat">
          <div class="stat-value">${stats.periodsWorked}</div>
          <div class="stat-label">Periods Worked</div>
        </div>
        <div class="stat">
          <div class="stat-value">${formatHoursDisplay(stats.totalOT)}</div>
          <div class="stat-label">Total OT Hours</div>
        </div>
        <div class="stat">
          <div class="stat-value">$${stats.totalCost.toFixed(2)}</div>
          <div class="stat-label">Total OT Cost</div>
        </div>
        <div class="stat">
          <div class="stat-value">${formatHoursDisplay(stats.avgOTPerPeriod)}</div>
          <div class="stat-label">Avg OT/Period</div>
        </div>
        <div class="stat">
          <div class="stat-value">${stats.multiLocationPeriods}</div>
          <div class="stat-label">Multi-Location Periods</div>
        </div>
        <div class="stat">
          <div class="stat-value">${stats.highOTPeriods}</div>
          <div class="stat-label">High OT Periods</div>
        </div>
      </div>
      
      <table>
        <thead>
          <tr>
            <th>Period End</th>
            <th>Location</th>
            <th style="text-align: right;">Total Hours</th>
            <th style="text-align: right;">OT Hours</th>
            <th style="text-align: right;">OT Cost</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
      
      <div class="footer">
        Generated on ${new Date().toLocaleString()}
      </div>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for trends report PDF
 */
function generateTrendsReportHTML(data, settings) {
  let topEmployeesRows = data.topEmployees.slice(0, 10).map((e, i) => `
    <tr>
      <td>${i + 1}</td>
      <td>${e.employeeName}</td>
      <td style="text-align: right;">${e.periods}</td>
      <td style="text-align: right;">${formatHoursDisplay(e.totalOT)}</td>
      <td style="text-align: right;">$${e.totalCost.toFixed(2)}</td>
      <td style="text-align: right;">${formatHoursDisplay(e.avgOT)}</td>
    </tr>
  `).join('');
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 40px; }
        h1 { color: #E51636; margin-bottom: 5px; }
        h2 { color: #333; margin-top: 30px; font-size: 16px; }
        .subtitle { color: #666; margin-bottom: 30px; }
        .stats { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 30px; }
        .stat { background: #f5f5f5; padding: 12px 20px; border-radius: 8px; }
        .stat-value { font-size: 20px; font-weight: bold; color: #333; }
        .stat-label { font-size: 11px; color: #666; text-transform: uppercase; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f5f5f5; font-weight: bold; font-size: 11px; text-transform: uppercase; }
        .footer { margin-top: 40px; font-size: 11px; color: #999; }
      </style>
    </head>
    <body>
      <h1>OT Trends Report</h1>
      <p class="subtitle">Time Range: ${data.timeRange === '3m' ? '3 Months' : data.timeRange === '6m' ? '6 Months' : data.timeRange === '1y' ? '1 Year' : 'All Time'}</p>
      
      <div class="stats">
        <div class="stat">
          <div class="stat-value">${data.totals.periodsCount}</div>
          <div class="stat-label">Pay Periods</div>
        </div>
        <div class="stat">
          <div class="stat-value">${data.totals.uniqueEmployees}</div>
          <div class="stat-label">Unique Employees</div>
        </div>
        <div class="stat">
          <div class="stat-value">${formatHoursDisplay(data.totals.totalOT)}</div>
          <div class="stat-label">Total OT Hours</div>
        </div>
        <div class="stat">
          <div class="stat-value">$${data.totals.totalCost.toFixed(2)}</div>
          <div class="stat-label">Total OT Cost</div>
        </div>
      </div>
      
      <h2>Location Breakdown</h2>
      <div class="stats">
        ${Object.entries(data.locationComparison).map(([loc, hours]) => `
          <div class="stat">
            <div class="stat-value">${formatHoursDisplay(hours)}</div>
            <div class="stat-label">${loc} OT Hours</div>
          </div>
        `).join('')}
      </div>
      
      <h2>Top 10 OT Employees</h2>
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Employee</th>
            <th style="text-align: right;">Periods</th>
            <th style="text-align: right;">Total OT</th>
            <th style="text-align: right;">Total Cost</th>
            <th style="text-align: right;">Avg OT</th>
          </tr>
        </thead>
        <tbody>
          ${topEmployeesRows}
        </tbody>
      </table>
      
      <div class="footer">
        Generated on ${new Date().toLocaleString()}
      </div>
    </body>
    </html>
  `;
}

// =====================================================
// UNIFORM SUMMARY / REPORTS
// =====================================================

/**
 * Gets comprehensive Uniform Summary data for the reports page
 * @returns {Object} Summary statistics and lists
 */
function getUniformSummaryData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName('Uniform_Order_Items');
    
    // Default empty response
    const emptyResponse = {
      success: true,
      summary: {
        totalOrders: 0,
        activeOrders: 0,
        completedOrders: 0,
        totalRevenue: 0,
        totalCollected: 0,
        totalOutstanding: 0,
        uniqueEmployees: 0
      },
      byCategory: [],
      byLocation: [],
      topItems: [],
      recentOrders: [],
      upcomingDeductions: []
    };
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return emptyResponse;
    }
    
    // Read orders data (17 columns based on current structure)
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 17).getValues();
    
    // Read items data if available
    let itemsData = [];
    const lineItemTotals = {}; // Map orderId -> calculated total from line items
    
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 9).getValues();
      
      // Calculate correct totals from line items
      for (const row of itemsData) {
        const orderId = row[1];
        const lineTotal = parseFloat(row[7]) || 0;
        if (orderId) {
          lineItemTotals[orderId] = (lineItemTotals[orderId] || 0) + lineTotal;
        }
      }
    }
    
    // Initialize aggregations
    let totalOrders = 0;
    let activeOrders = 0;
    let completedOrders = 0;
    let totalRevenue = 0;
    let totalCollected = 0;
    
    const employeeSet = new Set();
    const locationMap = new Map();
    const categoryMap = new Map();
    const itemCountMap = new Map();
    const recentOrders = [];
    
    // Process orders
    ordersData.forEach((row, index) => {
      const orderId = row[0];
      if (!orderId) return;
      
      const employeeId = row[1] || '';
      const employeeName = row[2] || 'Unknown';
      const location = row[3] || 'Unknown';
      const orderDate = row[4] ? new Date(row[4]) : null;
      const paymentPlan = parseInt(row[6]) || 1;
      const paymentsMade = parseInt(row[9]) || 0;
      const status = row[12] || 'Active';
      
      // Use calculated total from line items (not stored value which may be wrong)
      const storedTotal = parseFloat(row[5]) || 0;
      const totalAmount = lineItemTotals[orderId] !== undefined ? lineItemTotals[orderId] : storedTotal;
      const amountPerCheck = totalAmount > 0 ? Math.round((totalAmount / paymentPlan) * 100) / 100 : 0;
      
      totalOrders++;
      totalRevenue += totalAmount;
      employeeSet.add(employeeId || employeeName);
      
      const collected = paymentsMade * amountPerCheck;
      totalCollected += Math.min(collected, totalAmount);
      
      if (status === 'Active') {
        activeOrders++;
      } else if (status === 'Completed') {
        completedOrders++;
      }
      
      // Track by location
      if (!locationMap.has(location)) {
        locationMap.set(location, { location: location, orders: 0, revenue: 0 });
      }
      locationMap.get(location).orders++;
      locationMap.get(location).revenue += totalAmount;
      
      // Collect recent orders (for display)
      if (orderDate) {
        recentOrders.push({
          orderId: orderId,
          employeeName: employeeName,
          location: location,
          orderDate: orderDate.toISOString(),
          totalAmount: totalAmount,
          status: status,
          paymentProgress: `${paymentsMade}/${paymentPlan}`
        });
      }
    });
    
    // Process items for category and item popularity
    itemsData.forEach(row => {
      const orderId = row[1];
      const category = row[2] || 'Uncategorized';
      const itemName = row[3] || 'Unknown Item';
      const quantity = parseInt(row[5]) || 1;
      const lineTotal = parseFloat(row[7]) || 0;
      
      // Track by category
      if (!categoryMap.has(category)) {
        categoryMap.set(category, { category: category, itemCount: 0, revenue: 0 });
      }
      categoryMap.get(category).itemCount += quantity;
      categoryMap.get(category).revenue += lineTotal;
      
      // Track item popularity
      if (!itemCountMap.has(itemName)) {
        itemCountMap.set(itemName, { itemName: itemName, category: category, quantity: 0 });
      }
      itemCountMap.get(itemName).quantity += quantity;
    });
    
    // Sort and format results
    const byLocation = Array.from(locationMap.values())
      .sort((a, b) => b.revenue - a.revenue)
      .map(loc => ({
        ...loc,
        revenue: Math.round(loc.revenue * 100) / 100
      }));
    
    const byCategory = Array.from(categoryMap.values())
      .sort((a, b) => b.revenue - a.revenue)
      .map(cat => ({
        ...cat,
        revenue: Math.round(cat.revenue * 100) / 100
      }));
    
    const topItems = Array.from(itemCountMap.values())
      .sort((a, b) => b.quantity - a.quantity)
      .slice(0, 10);
    
    // Sort recent orders (newest first), limit to 15
    const sortedRecentOrders = recentOrders
      .sort((a, b) => new Date(b.orderDate) - new Date(a.orderDate))
      .slice(0, 15)
      .map(order => ({
        ...order,
        orderDate: new Date(order.orderDate).toLocaleDateString('en-US', {
          month: 'short', day: 'numeric', year: 'numeric'
        }),
        totalAmount: Math.round(order.totalAmount * 100) / 100
      }));
    
    // Get upcoming deductions (next 2 paydays)
    const upcomingDeductions = getUpcomingUniformDeductions();
    
    return {
      success: true,
      summary: {
        totalOrders: totalOrders,
        activeOrders: activeOrders,
        completedOrders: completedOrders,
        totalRevenue: Math.round(totalRevenue * 100) / 100,
        totalCollected: Math.round(totalCollected * 100) / 100,
        totalOutstanding: Math.round((totalRevenue - totalCollected) * 100) / 100,
        uniqueEmployees: employeeSet.size
      },
      byCategory: byCategory,
      byLocation: byLocation,
      topItems: topItems,
      recentOrders: sortedRecentOrders,
      upcomingDeductions: upcomingDeductions
    };
    
  } catch (error) {
    console.error('Error getting uniform summary data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get upcoming uniform deductions for next 2 pay periods
 */
function getUpcomingUniformDeductions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return [];
    }
    
    // First, get correct totals from line items
    const lineItemTotals = {};
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 9).getValues();
      for (const row of itemsData) {
        const orderId = row[1];
        const lineTotal = parseFloat(row[7]) || 0;
        if (orderId) {
          lineItemTotals[orderId] = (lineItemTotals[orderId] || 0) + lineTotal;
        }
      }
    }
    
    const data = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 17).getValues();
    const activeOrders = data.filter(r => r[12] === 'Active');
    
    // Get next 2 paydays
    const paydays = getUpcomingPaydays(2);
    const results = [];
    
    paydays.forEach(paydayStr => {
      const paydayDate = new Date(paydayStr);
      let totalAmount = 0;
      const employees = new Set();
      
      activeOrders.forEach(order => {
        const orderId = order[0];
        const firstDeduction = order[8] ? new Date(order[8]) : null;
        const paymentPlan = parseInt(order[6]) || 1;
        const paymentsMade = parseInt(order[9]) || 0;
        const employeeName = order[2];
        
        // Use calculated total from line items
        const storedTotal = parseFloat(order[5]) || 0;
        const correctTotal = lineItemTotals[orderId] !== undefined ? lineItemTotals[orderId] : storedTotal;
        const amountPerPaycheck = correctTotal > 0 ? Math.round((correctTotal / paymentPlan) * 100) / 100 : 0;
        
        if (!firstDeduction || paymentsMade >= paymentPlan) return;
        
        const daysDiff = Math.round((paydayDate - firstDeduction) / (1000 * 60 * 60 * 24));
        const paymentNumber = Math.floor(daysDiff / 14) + 1;
        
        if (daysDiff >= 0 && paymentNumber > paymentsMade && paymentNumber <= paymentPlan) {
          totalAmount += amountPerPaycheck;
          employees.add(employeeName);
        }
      });
      
      if (employees.size > 0) {
        results.push({
          date: paydayDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
          dateISO: paydayStr,
          amount: Math.round(totalAmount * 100) / 100,
          employeeCount: employees.size
        });
      }
    });
    
    return results;
    
  } catch (error) {
    console.error('Error getting upcoming deductions:', error);
    return [];
  }
}

// =====================================================
// ACTIVITY LOG / AUDIT TRAIL
// =====================================================

/**
 * Initializes the Activity Log sheet if it doesn't exist
 */
function initializeActivityLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Activity_Log');
  
  if (!sheet) {
    sheet = ss.insertSheet('Activity_Log');
    const headers = ['Timestamp', 'User', 'Action', 'Category', 'Description', 'Related_ID'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(2, 200); // User
    sheet.setColumnWidth(3, 120); // Action
    sheet.setColumnWidth(4, 100); // Category
    sheet.setColumnWidth(5, 400); // Description
    sheet.setColumnWidth(6, 120); // Related_ID
  }
  
  return sheet;
}

/**
 * Logs an activity to the Activity_Log sheet
 * @param {string} action - The action type (e.g., 'CREATE', 'UPDATE', 'DELETE', 'STATUS_CHANGE')
 * @param {string} category - The category (e.g., 'PTO', 'UNIFORM', 'PAYROLL', 'OT', 'SETTINGS')
 * @param {string} description - Human-readable description of what happened
 * @param {string} relatedId - Optional ID of the related record
 */
function logActivity(action, category, description, relatedId = '') {
  try {
    const sheet = initializeActivityLog();
    
    // Try multiple methods to get the user email
    let user = 'Unknown';
    try {
      // First try session email (from authentication flow)
      const sessionEmail = getSessionEmail();
      if (sessionEmail && sessionEmail !== '') {
        user = sessionEmail;
      } else {
        // Then try getActiveUser (works when user triggers action)
        const activeUser = Session.getActiveUser();
        if (activeUser) {
          user = activeUser.getEmail() || '';
        }
        
        // If that's empty, try getEffectiveUser (works in more contexts)
        if (!user || user === '') {
          const effectiveUser = Session.getEffectiveUser();
          if (effectiveUser) {
            user = effectiveUser.getEmail() || 'System';
          }
        }
      }
      
      // Final fallback
      if (!user || user === '') {
        user = 'System';
      }
    } catch (userError) {
      user = 'System';
    }
    
    const timestamp = new Date();
    
    sheet.appendRow([
      timestamp,
      user,
      action,
      category,
      description,
      relatedId
    ]);
    
    // Keep only last 1000 entries to prevent sheet from growing too large
    const lastRow = sheet.getLastRow();
    if (lastRow > 1001) {
      sheet.deleteRows(2, lastRow - 1001);
    }
    
  } catch (error) {
    console.error('Error logging activity:', error);
    // Don't throw - logging should never break main functionality
  }
}

/**
 * Gets recent activity log entries
 * @param {number} limit - Maximum number of entries to return
 * @returns {Array} Array of activity entries
 */
function getActivityLog(limit = 50) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Activity_Log');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, 6).getValues();
    
    return data
      .map(row => ({
        timestamp: row[0] ? new Date(row[0]).toISOString() : null,
        timestampDisplay: row[0] ? new Date(row[0]).toLocaleString() : '',
        user: row[1],
        action: row[2],
        category: row[3],
        description: row[4],
        relatedId: row[5]
      }))
      .filter(entry => entry.timestamp)
      .reverse(); // Most recent first
      
  } catch (error) {
    console.error('Error getting activity log:', error);
    return [];
  }
}

// =====================================================
// EMAIL NOTIFICATIONS
// =====================================================

/**
 * Sends an email notification if notifications are enabled
 * @param {string} subject - Email subject
 * @param {string} body - Email body (HTML supported)
 * @param {string} notificationType - Type of notification for checking settings
 */
function sendNotificationEmail(subject, body, notificationType) {
  try {
    const settings = getSettings();
    
    // Check if notifications are enabled
    if (!settings.notificationsEnabled) {
      console.log('Notifications disabled, skipping email');
      return { success: false, reason: 'Notifications disabled' };
    }
    
    // Check if this notification type is enabled
    const typeSettings = {
      'pto': settings.notifyOnPTORequest,
      'highOT': settings.notifyOnHighOT,
      'payroll': settings.notifyOnPayrollDue,
      'uniform': settings.notifyOnUniformOrder
    };
    
    if (typeSettings[notificationType] === false) {
      console.log(`Notification type ${notificationType} disabled`);
      return { success: false, reason: 'Notification type disabled' };
    }
    
    // Get admin emails
    const adminEmails = settings.adminEmails;
    if (!adminEmails || adminEmails.trim() === '') {
      console.log('No admin emails configured');
      return { success: false, reason: 'No admin emails configured' };
    }
    
    // Parse email addresses (comma-separated)
    const emailList = adminEmails.split(',').map(e => e.trim()).filter(e => e);
    
    if (emailList.length === 0) {
      return { success: false, reason: 'No valid email addresses' };
    }
    
    // Send email
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #E51636; color: white; padding: 16px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0; font-size: 18px;">📋 Payroll Review Notification</h2>
        </div>
        <div style="background: #f9f9f9; padding: 24px; border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 8px 8px;">
          ${body}
          <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 24px 0;">
          <p style="color: #666; font-size: 12px; margin: 0;">
            This is an automated notification from Payroll Review. 
            <a href="${ScriptApp.getService().getUrl()}" style="color: #E51636;">Open App</a>
          </p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailList.join(','),
      subject: '[Payroll Review] ' + subject,
      htmlBody: htmlBody
    });
    
    console.log(`Email sent to ${emailList.length} recipient(s): ${subject}`);
    return { success: true, recipients: emailList.length };
    
  } catch (error) {
    console.error('Error sending notification email:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Sends PTO request notification
 */
function sendPTORequestNotification(ptoData) {
  const subject = `New PTO Request: ${ptoData.employeeName}`;
  const body = `
    <h3 style="color: #333; margin-top: 0;">New PTO Request Submitted</h3>
    <table style="width: 100%; border-collapse: collapse;">
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Employee:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${ptoData.employeeName}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Hours:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${ptoData.hours} hours</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Dates:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${ptoData.dateRange}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Payout Period:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${ptoData.payoutPeriod}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0;"><strong>PTO ID:</strong></td>
        <td style="padding: 8px 0;">${ptoData.ptoId}</td>
      </tr>
    </table>
  `;
  
  return sendNotificationEmail(subject, body, 'pto');
}

/**
 * Sends High OT alert notification
 */
function sendHighOTNotification(otData) {
  const subject = `High OT Alert: ${otData.count} Employee(s)`;
  
  let employeeList = '';
  if (otData.employees && otData.employees.length > 0) {
    employeeList = '<ul style="margin: 8px 0; padding-left: 20px;">';
    otData.employees.slice(0, 10).forEach(emp => {
      employeeList += `<li>${emp.name}: ${emp.hours} hrs OT</li>`;
    });
    employeeList += '</ul>';
  }
  
  const body = `
    <h3 style="color: #E51636; margin-top: 0;">⚠️ High Overtime Alert</h3>
    <p><strong>${otData.count} employee(s)</strong> have high overtime (≥${otData.threshold} hours) in the latest period.</p>
    ${employeeList}
    <p style="margin-top: 16px;">
      <a href="${ScriptApp.getService().getUrl()}#ot-trends" 
         style="background: #E51636; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        Review OT Trends
      </a>
    </p>
  `;
  
  return sendNotificationEmail(subject, body, 'highOT');
}

/**
 * Sends Payroll Due reminder notification
 */
function sendPayrollDueNotification(payrollData) {
  const subject = payrollData.daysUntil === 0 
    ? 'Payroll Due TODAY!' 
    : `Payroll Due in ${payrollData.daysUntil} Day(s)`;
  
  const urgencyColor = payrollData.daysUntil === 0 ? '#E51636' : '#D97706';
  
  const body = `
    <h3 style="color: ${urgencyColor}; margin-top: 0;">
      ${payrollData.daysUntil === 0 ? '🚨' : '⏰'} Payroll Reminder
    </h3>
    <p>Payroll is due <strong>${payrollData.daysUntil === 0 ? 'TODAY' : 'in ' + payrollData.daysUntil + ' day(s)'}</strong> (${payrollData.payrollDate}).</p>
    <table style="width: 100%; border-collapse: collapse; margin: 16px 0;">
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Uniform Deductions:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">$${payrollData.uniformTotal || '0.00'}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>PTO Payouts:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${payrollData.ptoCount || 0} request(s)</td>
      </tr>
    </table>
    <p style="margin-top: 16px;">
      <a href="${ScriptApp.getService().getUrl()}#payroll-processing" 
         style="background: ${urgencyColor}; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; display: inline-block;">
        Process Payroll Now
      </a>
    </p>
  `;
  
  return sendNotificationEmail(subject, body, 'payroll');
}

/**
 * Sends Uniform Order notification
 */
function sendUniformOrderNotification(orderData) {
  const subject = `New Uniform Order: ${orderData.employeeName}`;
  const body = `
    <h3 style="color: #333; margin-top: 0;">New Uniform Order Created</h3>
    <table style="width: 100%; border-collapse: collapse;">
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Employee:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${orderData.employeeName}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Location:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${orderData.location || 'N/A'}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Total:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">$${orderData.total}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;"><strong>Payment Plan:</strong></td>
        <td style="padding: 8px 0; border-bottom: 1px solid #eee;">${orderData.payments} payment(s)</td>
      </tr>
      <tr>
        <td style="padding: 8px 0;"><strong>Order ID:</strong></td>
        <td style="padding: 8px 0;">${orderData.orderId}</td>
      </tr>
    </table>
  `;
  
  return sendNotificationEmail(subject, body, 'uniform');
}

/**
 * Test email notification (for settings page)
 */
function sendTestNotification() {
  const subject = 'Test Notification';
  const body = `
    <h3 style="color: #333; margin-top: 0;">✅ Test Successful!</h3>
    <p>If you're seeing this email, your notification settings are configured correctly.</p>
    <p><strong>Timestamp:</strong> ${new Date().toLocaleString()}</p>
  `;
  
  // Force send even if notifications disabled (for testing)
  const settings = getSettings();
  const adminEmails = settings.adminEmails;
  
  if (!adminEmails || adminEmails.trim() === '') {
    return { success: false, error: 'No admin emails configured. Please enter email address(es) first.' };
  }
  
  try {
    const emailList = adminEmails.split(',').map(e => e.trim()).filter(e => e);
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #E51636; color: white; padding: 16px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0; font-size: 18px;">📋 Payroll Review Notification</h2>
        </div>
        <div style="background: #f9f9f9; padding: 24px; border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 8px 8px;">
          ${body}
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailList.join(','),
      subject: '[Payroll Review] ' + subject,
      htmlBody: htmlBody
    });
    
    return { success: true, message: `Test email sent to ${emailList.join(', ')}` };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// =====================================================
// WEEKLY SUMMARY EMAIL
// =====================================================

/**
 * Sends enhanced weekly summary email with:
 * - System Health Summary
 * - Payroll Preview (upcoming deductions)
 * - PTO requests and Uniform orders
 * - Action Items checklist
 * - Quick Stats
 * - Direct Links
 * 
 * This is triggered automatically every Saturday at 1PM
 */
function sendWeeklySummaryEmail() {
  try {
    const settings = getSettings();
    
    // Check if notifications are enabled
    if (!settings.notificationsEnabled) {
      console.log('Notifications disabled, skipping weekly summary');
      return { success: false, reason: 'Notifications disabled' };
    }
    
    // Get admin emails
    const adminEmails = settings.adminEmails;
    if (!adminEmails || adminEmails.trim() === '') {
      console.log('No admin emails configured');
      return { success: false, reason: 'No admin emails configured' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const appUrl = ScriptApp.getService().getUrl();
    
    // Format date range for display
    const dateFormatter = (date) => {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM d, yyyy');
    };
    const weekRangeStr = `${dateFormatter(weekAgo)} - ${dateFormatter(now)}`;
    
    // Get all the data we need
    const ptoRequests = getRecentPTORequests(weekAgo);
    const uniformOrders = getRecentUniformOrders(weekAgo);
    const healthCheck = runSystemHealthCheck();
    const payrollPreview = getPayrollPreview();
    const quickStats = getQuickStats();
    const actionItems = getActionItems();
    
    // ========== BUILD QUICK STATS SECTION ==========
    const quickStatsSection = `
      <div style="margin-bottom: 24px;">
        <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
          📈 QUICK STATS
        </h3>
        <div style="display: flex; flex-wrap: wrap; gap: 12px;">
          <div style="flex: 1; min-width: 140px; background: #FEF3C7; padding: 16px; border-radius: 8px; text-align: center;">
            <div style="font-size: 28px; font-weight: 700; color: #D97706;">${quickStats.pendingOrders}</div>
            <div style="font-size: 12px; color: #92400E;">Pending Orders</div>
          </div>
          <div style="flex: 1; min-width: 140px; background: #DBEAFE; padding: 16px; border-radius: 8px; text-align: center;">
            <div style="font-size: 28px; font-weight: 700; color: #2563EB;">${quickStats.activeDeductions}</div>
            <div style="font-size: 12px; color: #1E40AF;">Active Deductions</div>
          </div>
          <div style="flex: 1; min-width: 140px; background: ${quickStats.daysSinceOT > 7 ? '#FEE2E2' : '#DCFCE7'}; padding: 16px; border-radius: 8px; text-align: center;">
            <div style="font-size: 28px; font-weight: 700; color: ${quickStats.daysSinceOT > 7 ? '#DC2626' : '#059669'};">${quickStats.daysSinceOT !== null ? quickStats.daysSinceOT : '—'}</div>
            <div style="font-size: 12px; color: ${quickStats.daysSinceOT > 7 ? '#991B1B' : '#166534'};">Days Since OT Upload</div>
          </div>
        </div>
      </div>
    `;
    
    // ========== BUILD PAYROLL PREVIEW SECTION ==========
    const payrollPreviewSection = payrollPreview.employeeCount > 0 ? `
      <div style="margin-bottom: 24px;">
        <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
          💰 UPCOMING PAYROLL DEDUCTIONS
        </h3>
        <div style="background: #F0FDF4; border: 1px solid #86EFAC; border-radius: 8px; padding: 16px;">
          <div style="font-size: 16px; margin-bottom: 8px;">
            <strong>This pay period:</strong> ${payrollPreview.employeeCount} employee(s), <span style="color: #059669; font-weight: 700;">$${payrollPreview.totalAmount.toFixed(2)}</span> total
          </div>
          ${payrollPreview.finalPayments > 0 ? `
            <div style="font-size: 14px; color: #166534;">
              ✅ ${payrollPreview.finalPayments} order(s) completing final payments this cycle
            </div>
          ` : ''}
        </div>
      </div>
    ` : '';
    
    // ========== BUILD SYSTEM HEALTH SECTION ==========
    let healthSection = '';
    if (healthCheck && healthCheck.checks) {
      const warnings = healthCheck.checks.filter(c => c.status === 'warning');
      const errors = healthCheck.checks.filter(c => c.status === 'error');
      
      if (warnings.length > 0 || errors.length > 0) {
        healthSection = `
          <div style="margin-bottom: 24px;">
            <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
              🏥 SYSTEM HEALTH
            </h3>
            <div style="background: ${errors.length > 0 ? '#FEE2E2' : '#FEF3C7'}; border: 1px solid ${errors.length > 0 ? '#FECACA' : '#FCD34D'}; border-radius: 8px; padding: 16px;">
              ${errors.map(e => `
                <div style="margin-bottom: 8px; color: #DC2626;">
                  ❌ <strong>${e.title}:</strong> ${e.message}
                </div>
              `).join('')}
              ${warnings.map(w => `
                <div style="margin-bottom: 8px; color: #D97706;">
                  ⚠️ <strong>${w.title}:</strong> ${w.message}
                </div>
              `).join('')}
              <div style="margin-top: 12px;">
                <a href="${appUrl}#system-health" style="color: #E51636; font-size: 13px;">View Full Health Report →</a>
              </div>
            </div>
          </div>
        `;
      }
    }
    
    // ========== BUILD ACTION ITEMS SECTION ==========
    const actionItemsSection = actionItems.length > 0 ? `
      <div style="margin-bottom: 24px;">
        <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
          ✅ ACTION ITEMS
        </h3>
        <div style="background: #F9FAFB; border-radius: 8px; padding: 16px;">
          ${actionItems.map(item => `
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px; padding: 8px; background: white; border-radius: 6px; border-left: 4px solid ${item.priority === 'high' ? '#DC2626' : item.priority === 'medium' ? '#D97706' : '#9CA3AF'};">
              <span style="font-size: 16px;">${item.icon}</span>
              <span style="font-size: 14px; color: #374151;">${item.text}</span>
            </div>
          `).join('')}
        </div>
      </div>
    ` : '';
    
    // ========== BUILD PTO SECTION ==========
    let ptoSection = '';
    if (ptoRequests.length > 0) {
      ptoSection = `
        <div style="margin-bottom: 24px;">
          <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
            🏖️ PTO REQUESTS (${ptoRequests.length} new)
          </h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr style="background: #f5f5f5;">
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Employee</th>
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Dates</th>
              <th style="padding: 8px; text-align: center; border-bottom: 1px solid #ddd;">Hours</th>
              <th style="padding: 8px; text-align: center; border-bottom: 1px solid #ddd;">Status</th>
            </tr>
            ${ptoRequests.map(req => `
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${req.employeeName}</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${req.dateRange}</td>
                <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">${req.hours}</td>
                <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">
                  <span style="background: ${getStatusColor(req.status)}; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px;">
                    ${req.status}
                  </span>
                </td>
              </tr>
            `).join('')}
          </table>
        </div>
      `;
    }
    
    // ========== BUILD UNIFORM SECTION ==========
    let uniformSection = '';
    let uniformTotal = 0;
    if (uniformOrders.length > 0) {
      uniformOrders.forEach(order => uniformTotal += order.total);
      
      uniformSection = `
        <div style="margin-bottom: 24px;">
          <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
            👕 NEW UNIFORM ORDERS (${uniformOrders.length}) - $${uniformTotal.toFixed(2)} total
          </h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr style="background: #f5f5f5;">
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Employee</th>
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Items</th>
              <th style="padding: 8px; text-align: left; border-bottom: 1px solid #ddd;">Notes</th>
              <th style="padding: 8px; text-align: right; border-bottom: 1px solid #ddd;">Total</th>
              <th style="padding: 8px; text-align: center; border-bottom: 1px solid #ddd;">Payments</th>
            </tr>
            ${uniformOrders.map(order => `
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${order.employeeName}</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-size: 12px;">${order.itemSummary}</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-size: 12px; color: #666; max-width: 200px;">${order.notes || '-'}</td>
                <td style="padding: 8px; text-align: right; border-bottom: 1px solid #eee;">$${order.total.toFixed(2)}</td>
                <td style="padding: 8px; text-align: center; border-bottom: 1px solid #eee;">${order.payments}</td>
              </tr>
            `).join('')}
          </table>
        </div>
      `;
    }
    
    // ========== BUILD DIRECT LINKS SECTION ==========
    const linksSection = `
      <div style="margin-bottom: 16px;">
        <h3 style="color: #E51636; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 2px solid #E51636;">
          🔗 QUICK LINKS
        </h3>
        <div style="display: flex; flex-wrap: wrap; gap: 8px;">
          <a href="${appUrl}#uniforms-orders" style="background: #E51636; color: white; padding: 8px 16px; text-decoration: none; border-radius: 6px; font-size: 13px; font-weight: 500;">
            Pending Orders
          </a>
          <a href="${appUrl}#ot-upload" style="background: #2563EB; color: white; padding: 8px 16px; text-decoration: none; border-radius: 6px; font-size: 13px; font-weight: 500;">
            Upload OT
          </a>
          <a href="${appUrl}#system-health" style="background: #059669; color: white; padding: 8px 16px; text-decoration: none; border-radius: 6px; font-size: 13px; font-weight: 500;">
            System Health
          </a>
          <a href="${appUrl}#home" style="background: #6B7280; color: white; padding: 8px 16px; text-decoration: none; border-radius: 6px; font-size: 13px; font-weight: 500;">
            Dashboard
          </a>
        </div>
      </div>
    `;
    
    // ========== BUILD SUBJECT LINE ==========
    let subjectParts = [];
    if (healthCheck.errorCount > 0) subjectParts.push(`⚠️ ${healthCheck.errorCount} error(s)`);
    if (payrollPreview.employeeCount > 0) subjectParts.push(`$${payrollPreview.totalAmount.toFixed(0)} deductions`);
    if (ptoRequests.length > 0) subjectParts.push(`${ptoRequests.length} PTO`);
    if (uniformOrders.length > 0) subjectParts.push(`${uniformOrders.length} uniforms`);
    
    const subject = subjectParts.length > 0 
      ? `Weekly Summary: ${subjectParts.join(', ')}`
      : 'Weekly Summary: All Clear ✓';
    
    // ========== ASSEMBLE FULL HTML ==========
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #E51636 0%, #B91230 100%); color: white; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0; font-size: 20px;">📊 Weekly Payroll Summary</h2>
          <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px;">${weekRangeStr}</p>
        </div>
        <div style="background: #ffffff; padding: 24px; border: 1px solid #e0e0e0; border-top: none;">
          
          ${quickStatsSection}
          
          ${healthSection}
          
          ${payrollPreviewSection}
          
          ${actionItemsSection}
          
          ${ptoSection || `
            <div style="margin-bottom: 24px; padding: 16px; background: #f8f9fa; border-radius: 8px; text-align: center;">
              <p style="margin: 0; color: #666;">🏖️ No new PTO requests this week</p>
            </div>
          `}
          
          ${uniformSection || `
            <div style="margin-bottom: 24px; padding: 16px; background: #f8f9fa; border-radius: 8px; text-align: center;">
              <p style="margin: 0; color: #666;">👕 No new uniform orders this week</p>
            </div>
          `}
          
          ${linksSection}
          
        </div>
        <div style="padding: 12px 24px; background: #f5f5f5; border-radius: 0 0 8px 8px; border: 1px solid #e0e0e0; border-top: none;">
          <p style="color: #888; font-size: 11px; margin: 0; text-align: center;">
            Weekly summary sent every Saturday at 1 PM • 
            <a href="${appUrl}#settings" style="color: #E51636;">Manage notifications</a>
          </p>
        </div>
      </div>
    `;
    
    // Send email
    const emailList = adminEmails.split(',').map(e => e.trim()).filter(e => e);
    
    MailApp.sendEmail({
      to: emailList.join(','),
      subject: '[Payroll Review] ' + subject,
      htmlBody: htmlBody
    });
    
    console.log(`Enhanced weekly summary sent to ${emailList.length} recipient(s)`);
    logActivity('SEND_EMAIL', 'NOTIFICATION', `Weekly summary: PTO: ${ptoRequests.length}, Uniforms: ${uniformOrders.length}, Health errors: ${healthCheck.errorCount}`, null);
    
    return { 
      success: true, 
      ptoCount: ptoRequests.length, 
      uniformCount: uniformOrders.length,
      healthErrors: healthCheck.errorCount,
      healthWarnings: healthCheck.warningCount,
      recipients: emailList.length 
    };
    
  } catch (error) {
    console.error('Error sending weekly summary:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Helper function to get status color
 */
function getStatusColor(status) {
  const colors = {
    'Pending': '#D97706',
    'Approved': '#059669',
    'Denied': '#DC2626',
    'Paid': '#2563EB',
    'Completed': '#059669'
  };
  return colors[status] || '#6B7280';
}

/**
 * Get PTO requests from the past week
 */
function getRecentPTORequests(sinceDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO_Requests');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const submittedIdx = headers.indexOf('Submitted_Date');
    const nameIdx = headers.indexOf('Employee_Name');
    const startIdx = headers.indexOf('Start_Date');
    const endIdx = headers.indexOf('End_Date');
    const hoursIdx = headers.indexOf('Total_Hours');
    const statusIdx = headers.indexOf('Status');
    
    const requests = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const submittedDate = new Date(row[submittedIdx]);
      
      if (submittedDate >= sinceDate) {
        const startDate = new Date(row[startIdx]);
        const endDate = new Date(row[endIdx]);
        
        requests.push({
          employeeName: row[nameIdx],
          dateRange: formatDateRange(startDate, endDate),
          hours: row[hoursIdx],
          status: row[statusIdx] || 'Pending'
        });
      }
    }
    
    return requests;
    
  } catch (error) {
    console.error('Error getting recent PTO requests:', error);
    return [];
  }
}

/**
 * Get Uniform orders from the past week
 */
function getRecentUniformOrders(sinceDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Uniform_Orders');
    const itemsSheet = ss.getSheetByName('Uniform_Order_Items');
    
    if (!ordersSheet) return [];
    
    const ordersData = ordersSheet.getDataRange().getValues();
    if (ordersData.length <= 1) return [];
    
    const headers = ordersData[0];
    const orderIdIdx = headers.indexOf('Order_ID');
    const dateIdx = headers.indexOf('Order_Date');
    const nameIdx = headers.indexOf('Employee_Name');
    const paymentsIdx = headers.indexOf('Payment_Plan');
    const notesIdx = headers.indexOf('Notes');
    
    // Get all items for reference - including lineTotal for correct totals
    let allItems = [];
    let orderTotals = {}; // Calculate totals from line items
    
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 9).getValues();
      
      for (const row of itemsData) {
        const orderId = row[1];
        const lineTotal = parseFloat(row[7]) || 0; // Line_Total is column 8 (index 7)
        
        if (orderId) {
          // Accumulate totals by order
          orderTotals[orderId] = (orderTotals[orderId] || 0) + lineTotal;
          
          // Store item details
          allItems.push({
            orderId: orderId,
            itemName: row[3],
            size: row[4],
            quantity: row[5]
          });
        }
      }
    }
    
    const orders = [];
    
    for (let i = 1; i < ordersData.length; i++) {
      const row = ordersData[i];
      const orderDate = new Date(row[dateIdx]);
      const orderId = row[orderIdIdx];
      
      if (orderDate >= sinceDate && orderId) {
        // Get items for this order
        const orderItems = allItems.filter(item => item.orderId === orderId);
        
        // Build item summary
        let itemSummary = 'N/A';
        if (orderItems.length > 0) {
          itemSummary = orderItems.map(item => {
            const size = item.size && item.size !== 'One Size' ? ` (${item.size})` : '';
            return `${item.itemName}${size}`;
          }).join(', ');
        }
        
        // Use calculated total from line items (not the stored incorrect value)
        const calculatedTotal = orderTotals[orderId] || 0;
        
        orders.push({
          employeeName: row[nameIdx] || 'Unknown',
          itemSummary: itemSummary,
          total: calculatedTotal,
          payments: parseInt(row[paymentsIdx]) || 1,
          notes: notesIdx >= 0 ? (row[notesIdx] || '') : ''
        });
      }
    }
    
    return orders;
    
  } catch (error) {
    console.error('Error getting recent uniform orders:', error);
    return [];
  }
}

// ============================================================================
// CHUNK 11: ENHANCED WEEKLY SUMMARY EMAIL HELPERS
// ============================================================================

/**
 * Get payroll preview for upcoming pay date
 * @returns {Object} Preview data for upcoming deductions
 */
function getPayrollPreview() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return { employeeCount: 0, totalAmount: 0, finalPayments: 0, orders: [] };
    }
    
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
    
    let employeeCount = 0;
    let totalAmount = 0;
    let finalPayments = 0;
    const employeeSet = new Set();
    
    for (const row of ordersData) {
      const status = row[12];
      const amountPerPaycheck = parseFloat(row[7]) || 0;
      const employeeName = row[2];
      const paymentsMade = parseInt(row[9]) || 0;
      const paymentPlan = parseInt(row[6]) || 1;
      
      if (status === 'Active' && amountPerPaycheck > 0) {
        employeeSet.add(employeeName);
        totalAmount += amountPerPaycheck;
        
        // Check if this is the final payment
        if (paymentsMade + 1 >= paymentPlan) {
          finalPayments++;
        }
      }
    }
    
    return {
      employeeCount: employeeSet.size,
      totalAmount: parseFloat(totalAmount.toFixed(2)),
      finalPayments: finalPayments
    };
    
  } catch (error) {
    console.error('Error getting payroll preview:', error);
    return { employeeCount: 0, totalAmount: 0, finalPayments: 0 };
  }
}

/**
 * Get quick stats for the dashboard summary
 * @returns {Object} Quick stats object
 */
function getQuickStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Pending orders count
    let pendingOrders = 0;
    let activeDeductions = 0;
    let lastOTUpload = null;
    let daysSinceOT = null;
    
    // Count pending orders
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
      for (const row of ordersData) {
        const status = row[12];
        if (status === 'Pending' || status === 'Pending - Cash') {
          pendingOrders++;
        }
        if (status === 'Active') {
          activeDeductions++;
        }
      }
    }
    
    // Get last OT upload date
    const otSheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    if (otSheet && otSheet.getLastRow() > 1) {
      const periods = otSheet.getRange(2, 1, otSheet.getLastRow() - 1, 1).getValues();
      const dates = periods.map(p => p[0] ? new Date(p[0]).getTime() : 0).filter(d => d > 0);
      
      if (dates.length > 0) {
        lastOTUpload = new Date(Math.max(...dates));
        daysSinceOT = Math.floor((new Date() - lastOTUpload) / (1000 * 60 * 60 * 24));
      }
    }
    
    return {
      pendingOrders: pendingOrders,
      activeDeductions: activeDeductions,
      lastOTUpload: lastOTUpload,
      daysSinceOT: daysSinceOT
    };
    
  } catch (error) {
    console.error('Error getting quick stats:', error);
    return { pendingOrders: 0, activeDeductions: 0, lastOTUpload: null, daysSinceOT: null };
  }
}

/**
 * Get action items for the weekly summary
 * @returns {Array} Array of action item objects
 */
function getActionItems() {
  const actions = [];
  const stats = getQuickStats();
  
  // OT upload reminder
  if (stats.daysSinceOT === null) {
    actions.push({
      icon: '📊',
      text: 'Upload OT data - no data found',
      priority: 'high'
    });
  } else if (stats.daysSinceOT > 7) {
    actions.push({
      icon: '📊',
      text: `Upload OT data (last upload: ${stats.daysSinceOT} days ago)`,
      priority: 'high'
    });
  } else if (stats.daysSinceOT > 5) {
    actions.push({
      icon: '📊',
      text: 'Upload OT data by Monday EOD',
      priority: 'medium'
    });
  }
  
  // Pending orders
  if (stats.pendingOrders > 0) {
    actions.push({
      icon: '👕',
      text: `Process ${stats.pendingOrders} pending uniform order(s)`,
      priority: stats.pendingOrders > 5 ? 'high' : 'medium'
    });
  }
  
  // System health check reminder
  actions.push({
    icon: '🔍',
    text: 'Review System Health Dashboard',
    priority: 'low'
  });
  
  return actions;
}

/**
 * Helper to format date range
 */
function formatDateRange(start, end) {
  const options = { month: 'short', day: 'numeric' };
  const startStr = start.toLocaleDateString('en-US', options);
  const endStr = end.toLocaleDateString('en-US', options);
  
  if (startStr === endStr) {
    return startStr;
  }
  return `${startStr} - ${endStr}`;
}

/**
 * Sets up the weekly summary trigger for Saturday at 1PM
 * Call this once to initialize the trigger
 */
function setupWeeklySummaryTrigger() {
  // First, remove any existing weekly summary triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendWeeklySummaryEmail') {
      ScriptApp.deleteTrigger(trigger);
      console.log('Removed existing weekly summary trigger');
    }
  }
  
  // Create new trigger for Saturday at 1 PM
  ScriptApp.newTrigger('sendWeeklySummaryEmail')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(13) // 1 PM
    .create();
  
  console.log('Weekly summary trigger created for Saturday at 1 PM');
  logActivity('system', 'Weekly summary trigger configured', 'Saturday at 1 PM');
  
  return { 
    success: true, 
    message: 'Weekly summary email scheduled for Saturday at 1 PM' 
  };
}

/**
 * Removes the weekly summary trigger
 */
function removeWeeklySummaryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendWeeklySummaryEmail') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  
  console.log(`Removed ${removed} weekly summary trigger(s)`);
  return { success: true, removed: removed };
}

/**
 * Check if weekly summary trigger is set up
 */
function isWeeklySummaryTriggerActive() {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendWeeklySummaryEmail') {
      return { 
        active: true, 
        nextRun: 'Saturday at 1 PM' 
      };
    }
  }
  
  return { active: false };
}

/**
 * Manually send weekly summary (for testing)
 */
function sendWeeklySummaryNow() {
  return sendWeeklySummaryEmail();
}

// =====================================================
// COMMENTS SYSTEM
// =====================================================

/**
 * Initializes the Comments sheet if it doesn't exist
 */
function initializeCommentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Comments');
  
  if (!sheet) {
    sheet = ss.insertSheet('Comments');
    const headers = ['Comment_ID', 'Record_Type', 'Record_ID', 'Comment_Text', 'Created_By', 'Created_Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 120); // Comment_ID
    sheet.setColumnWidth(2, 100); // Record_Type
    sheet.setColumnWidth(3, 120); // Record_ID
    sheet.setColumnWidth(4, 400); // Comment_Text
    sheet.setColumnWidth(5, 200); // Created_By
    sheet.setColumnWidth(6, 160); // Created_Date
  }
  
  return sheet;
}

/**
 * Get comments for a specific record
 */
function getComments(recordType, recordId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Comments');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    
    const comments = data
      .filter(row => row[1] === recordType && row[2] === recordId)
      .map(row => ({
        id: row[0],
        recordType: row[1],
        recordId: row[2],
        text: row[3],
        createdBy: row[4],
        createdDate: row[5] ? new Date(row[5]).toISOString() : null
      }))
      .sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));
    
    return comments;
    
  } catch (error) {
    console.error('Error getting comments:', error);
    return [];
  }
}

/**
 * Add a new comment
 */
function addComment(recordType, recordId, commentText) {
  try {
    const sheet = initializeCommentsSheet();
    
    // Generate comment ID
    const lastRow = sheet.getLastRow();
    let commentNumber = 1;
    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue();
      if (lastId && typeof lastId === 'string' && lastId.startsWith('CMT-')) {
        commentNumber = parseInt(lastId.replace('CMT-', '')) + 1;
      }
    }
    const commentId = 'CMT-' + String(commentNumber).padStart(5, '0');
    
    // Get user
    let user = 'Unknown';
    try {
      const activeUser = Session.getActiveUser();
      if (activeUser) user = activeUser.getEmail() || 'Unknown';
      if (!user || user === '') {
        const effectiveUser = Session.getEffectiveUser();
        if (effectiveUser) user = effectiveUser.getEmail() || 'System';
      }
    } catch (e) {
      user = 'System';
    }
    
    // Add comment
    sheet.appendRow([
      commentId,
      recordType,
      recordId,
      commentText,
      user,
      new Date()
    ]);
    
    return { success: true, commentId: commentId };
    
  } catch (error) {
    console.error('Error adding comment:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Delete a comment
 */
function deleteComment(commentId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Comments');
    
    if (!sheet) {
      return { success: false, error: 'Comments sheet not found' };
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === commentId) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    
    return { success: false, error: 'Comment not found' };
    
  } catch (error) {
    console.error('Error deleting comment:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get comment count for a record
 */
function getCommentCount(recordType, recordId) {
  try {
    const comments = getComments(recordType, recordId);
    return comments.length;
  } catch (error) {
    return 0;
  }
}

// =====================================================
// DATA BACKUP SYSTEM
// =====================================================

/**
 * Create a full backup of all data
 */
function createBackup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const backupName = `Payroll_Backup_${timestamp}`;
    
    // Create a copy of the spreadsheet
    const backup = ss.copy(backupName);
    const backupId = backup.getId();
    const backupUrl = backup.getUrl();
    
    // Log the backup
    logActivity('CREATE', 'SETTINGS', `Backup created: ${backupName}`, backupId);
    
    // Store backup info in settings or a backup log
    storeBackupHistory(backupId, backupName, backupUrl);
    
    return {
      success: true,
      backupId: backupId,
      backupName: backupName,
      backupUrl: backupUrl,
      message: `Backup created successfully: ${backupName}`
    };
    
  } catch (error) {
    console.error('Error creating backup:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Store backup history
 */
function storeBackupHistory(backupId, backupName, backupUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Backup_History');
    
    if (!sheet) {
      sheet = ss.insertSheet('Backup_History');
      const headers = ['Backup_ID', 'Backup_Name', 'URL', 'Created_Date', 'Created_By'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    let user = 'Unknown';
    try {
      user = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'System';
    } catch (e) {}
    
    sheet.appendRow([backupId, backupName, backupUrl, new Date(), user]);
    
    // Keep only last 20 backups in history
    if (sheet.getLastRow() > 21) {
      sheet.deleteRow(2);
    }
    
  } catch (error) {
    console.error('Error storing backup history:', error);
  }
}

/**
 * Get backup history
 */
function getBackupHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Backup_History');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    
    return data.map(row => ({
      id: row[0],
      name: row[1],
      url: row[2],
      createdDate: row[3] ? new Date(row[3]).toISOString() : null,
      createdBy: row[4]
    })).reverse(); // Most recent first
    
  } catch (error) {
    console.error('Error getting backup history:', error);
    return [];
  }
}

/**
 * Export all data as JSON for download
 */
function exportDataAsJSON() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = {};
    
    // Sheets to export
    const sheetsToExport = [
      'OT_History', 'Employees', 'Uniform_Catalog', 'Uniform_Orders', 
      'Uniform_Order_Items', 'PTO', 'Payroll_Settings', 'Settings'
    ];
    
    sheetsToExport.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() >= 1) {
        const values = sheet.getDataRange().getValues();
        const headers = values[0];
        const rows = values.slice(1).map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });
        data[sheetName] = rows;
      }
    });
    
    data.exportDate = new Date().toISOString();
    data.exportedBy = Session.getActiveUser().getEmail() || 'Unknown';
    
    return {
      success: true,
      data: JSON.stringify(data, null, 2)
    };
    
  } catch (error) {
    console.error('Error exporting data:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// PDF EXPORT
// =====================================================

/**
 * Generate PDF report - unified function for all report types
 */
function generatePDFReport(reportType, options = {}) {
  try {
    // Handle OT-specific reports via dedicated function
    if (['period', 'employee', 'trends'].includes(reportType)) {
      return generateOTPDFReport(reportType, options);
    }
    
    let htmlContent = '';
    let reportTitle = '';
    
    switch (reportType) {
      case 'payroll':
        reportTitle = 'Payroll Report';
        htmlContent = generatePayrollReportHTML(options);
        break;
      case 'ot-summary':
        reportTitle = 'OT Summary Report';
        htmlContent = generateOTSummaryHTML(options);
        break;
      case 'uniform-deductions':
        reportTitle = 'Uniform Deductions Report';
        htmlContent = generateUniformDeductionsHTML(options);
        break;
      case 'pto-summary':
        reportTitle = 'PTO Summary Report';
        htmlContent = generatePTOReportHTML(options);
        break;
      default:
        return { success: false, error: 'Unknown report type: ' + reportType };
    }
    
    // Create HTML for PDF
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 40px; color: #333; }
          h1 { color: #E51636; border-bottom: 2px solid #E51636; padding-bottom: 10px; }
          h2 { color: #374151; margin-top: 30px; }
          table { width: 100%; border-collapse: collapse; margin: 20px 0; }
          th { background: #F3F4F6; text-align: left; padding: 10px; border: 1px solid #E5E7EB; }
          td { padding: 10px; border: 1px solid #E5E7EB; }
          .header-info { display: flex; justify-content: space-between; margin-bottom: 20px; }
          .stat-box { background: #F9FAFB; padding: 15px; border-radius: 8px; margin: 10px 0; }
          .stat-value { font-size: 24px; font-weight: bold; color: #E51636; }
          .stat-label { font-size: 12px; color: #6B7280; text-transform: uppercase; }
          .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #E5E7EB; font-size: 12px; color: #9CA3AF; }
        </style>
      </head>
      <body>
        <h1>📋 ${reportTitle}</h1>
        <div class="header-info">
          <div>Generated: ${new Date().toLocaleString()}</div>
        </div>
        ${htmlContent}
        <div class="footer">
          Generated by Payroll Review System
        </div>
      </body>
      </html>
    `;
    
    // Convert to PDF blob
    const blob = HtmlService.createHtmlOutput(fullHtml).getBlob().setName(`${reportTitle.replace(/\s+/g, '_')}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.pdf`).getAs('application/pdf');
    
    // Save to Drive and get URL
    const file = DriveApp.createFile(blob);
    const fileUrl = file.getUrl();
    
    // Log activity
    logActivity('EXPORT', 'REPORT', `PDF exported: ${reportTitle}`, file.getId());
    
    return {
      success: true,
      fileUrl: fileUrl,
      url: fileUrl, // Include both for compatibility
      fileName: file.getName()
    };
    
  } catch (error) {
    console.error('Error generating PDF:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Generate Payroll Report HTML
 */
function generatePayrollReportHTML(options) {
  const payrollDate = options.payrollDate || getPayrollSetting('next_payroll_date');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get data sections
  const uniformData = getUniformDeductionsForPayroll(ss, payrollDate);
  const ptoData = getPTOForPayrollReport(Utilities.formatDate(new Date(payrollDate), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  
  let html = `
    <h2>Payroll Date: ${new Date(payrollDate).toLocaleDateString()}</h2>
    
    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin: 20px 0;">
      <div class="stat-box">
        <div class="stat-value">$${uniformData.total.toFixed(2)}</div>
        <div class="stat-label">Uniform Deductions</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${uniformData.orders.length}</div>
        <div class="stat-label">Uniform Orders</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${Object.keys(ptoData).length}</div>
        <div class="stat-label">PTO Payouts</div>
      </div>
    </div>
  `;
  
  // Uniform deductions table
  if (uniformData.orders.length > 0) {
    html += `
      <h2>Uniform Deductions</h2>
      <table>
        <tr>
          <th>Employee</th>
          <th>Order ID</th>
          <th>Payment #</th>
          <th>Amount</th>
        </tr>
        ${uniformData.orders.map(o => `
          <tr>
            <td>${o.employeeName}</td>
            <td>${o.orderId}</td>
            <td>${o.paymentNumber} of ${o.totalPayments}</td>
            <td>$${o.deductionAmount.toFixed(2)}</td>
          </tr>
        `).join('')}
      </table>
    `;
  }
  
  // PTO payouts
  const ptoEntries = Object.values(ptoData);
  if (ptoEntries.length > 0) {
    html += `
      <h2>PTO Payouts</h2>
      <table>
        <tr>
          <th>Employee</th>
          <th>Hours</th>
          <th>PTO IDs</th>
        </tr>
        ${ptoEntries.map(p => `
          <tr>
            <td>${p.employeeName}</td>
            <td>${p.totalHours}</td>
            <td>${p.ptoRecords.map(r => r.ptoId).join(', ')}</td>
          </tr>
        `).join('')}
      </table>
    `;
  }
  
  return html;
}

/**
 * Generate OT Summary HTML
 */
function generateOTSummaryHTML(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return '<p>No OT data available.</p>';
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  
  // Get unique periods
  const periods = [...new Set(data.map(r => r[0] ? new Date(r[0]).toDateString() : null).filter(p => p))];
  periods.sort((a, b) => new Date(b) - new Date(a));
  
  // Calculate totals
  let totalOTHours = 0;
  let totalOTCost = 0;
  let employeeCount = new Set();
  
  data.forEach(row => {
    totalOTHours += parseFloat(row[12]) || 0;
    totalOTCost += parseFloat(row[13]) || 0;
    if (row[1]) employeeCount.add(row[1]);
  });
  
  return `
    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin: 20px 0;">
      <div class="stat-box">
        <div class="stat-value">${periods.length}</div>
        <div class="stat-label">Pay Periods</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${employeeCount.size}</div>
        <div class="stat-label">Employees</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${totalOTHours.toFixed(1)}</div>
        <div class="stat-label">Total OT Hours</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">$${totalOTCost.toFixed(2)}</div>
        <div class="stat-label">Total OT Cost</div>
      </div>
    </div>
    <p>Report includes data from ${periods.length} pay periods.</p>
  `;
}

/**
 * Generate Uniform Deductions HTML
 */
function generateUniformDeductionsHTML(options) {
  const orders = getUniformOrders({ status: 'Active' });
  
  let totalDeductions = 0;
  orders.forEach(o => {
    totalDeductions += o.amountRemaining || 0;
  });
  
  return `
    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin: 20px 0;">
      <div class="stat-box">
        <div class="stat-value">${orders.length}</div>
        <div class="stat-label">Active Orders</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">$${totalDeductions.toFixed(2)}</div>
        <div class="stat-label">Total Outstanding</div>
      </div>
    </div>
    
    <h2>Active Orders</h2>
    <table>
      <tr>
        <th>Order ID</th>
        <th>Employee</th>
        <th>Total</th>
        <th>Remaining</th>
        <th>Payments</th>
      </tr>
      ${orders.slice(0, 50).map(o => `
        <tr>
          <td>${o.orderId}</td>
          <td>${o.employeeName}</td>
          <td>$${o.totalAmount.toFixed(2)}</td>
          <td>$${o.amountRemaining.toFixed(2)}</td>
          <td>${o.paymentsMade} of ${o.paymentPlan}</td>
        </tr>
      `).join('')}
    </table>
  `;
}

/**
 * Generate PTO Report HTML
 */
function generatePTOReportHTML(options) {
  const data = getPTOSummaryData();
  
  return `
    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin: 20px 0;">
      <div class="stat-box">
        <div class="stat-value">${data.totalRequests}</div>
        <div class="stat-label">Total Requests</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${data.totalHours}</div>
        <div class="stat-label">Total Hours</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${data.pendingCount}</div>
        <div class="stat-label">Pending</div>
      </div>
      <div class="stat-box">
        <div class="stat-value">${data.paidCount}</div>
        <div class="stat-label">Paid</div>
      </div>
    </div>
    
    <h2>Top Employees by PTO Hours</h2>
    <table>
      <tr>
        <th>Employee</th>
        <th>Location</th>
        <th>Requests</th>
        <th>Total Hours</th>
      </tr>
      ${(data.topEmployees || []).slice(0, 10).map(e => `
        <tr>
          <td>${e.name}</td>
          <td>${e.location}</td>
          <td>${e.requestCount}</td>
          <td>${e.totalHours}</td>
        </tr>
      `).join('')}
    </table>
  `;
}

// =====================================================
// SYSTEM HEALTH CHECK
// =====================================================

/**
 * Runs comprehensive system health checks
 * @returns {Object} Health check results
 */
function runSystemHealthCheck() {
  try {
    console.log('Starting system health check...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const checks = [];
    let errorCount = 0;
    let warningCount = 0;
  
  // ========== CHECK 1: Required Sheets Exist ==========
  const requiredSheets = [
    { name: 'OT_History', description: 'Overtime tracking data' },
    { name: 'Employees', description: 'Employee roster' },
    { name: 'Uniform_Catalog', description: 'Uniform items and prices' },
    { name: 'Uniform_Orders', description: 'Uniform order records' },
    { name: 'Uniform_Order_Items', description: 'Individual order line items' },
    { name: 'PTO', description: 'PTO requests' },
    { name: 'Settings', description: 'Application settings' },
    { name: 'Payroll_Settings', description: 'Payroll configuration' },
    { name: 'System_Counters', description: 'ID generation counters' }
  ];
  
  const missingSheets = [];
  for (const req of requiredSheets) {
    if (!ss.getSheetByName(req.name)) {
      missingSheets.push(req);
    }
  }
  
  if (missingSheets.length === 0) {
    checks.push({
      id: 'sheets_exist',
      status: 'success',
      title: 'All Required Sheets Present',
      message: `All ${requiredSheets.length} required sheets exist.`,
      details: requiredSheets.map(s => s.name),
      howToFix: null
    });
  } else {
    errorCount++;
    checks.push({
      id: 'sheets_exist',
      status: 'error',
      title: 'Missing Required Sheets',
      message: `${missingSheets.length} required sheet(s) are missing.`,
      details: missingSheets.map(s => `${s.name} - ${s.description}`),
      howToFix: 'Go to Settings and click "Initialize All Sheets" to create missing sheets, or check if sheets were accidentally renamed or deleted.'
    });
  }
  
  // ========== CHECK 2: Employee Roster Has Data ==========
  const empSheet = ss.getSheetByName('Employees');
  if (empSheet) {
    const empCount = Math.max(0, empSheet.getLastRow() - 1);
    if (empCount > 0) {
      checks.push({
        id: 'employees_exist',
        status: 'success',
        title: 'Employee Roster Populated',
        message: `${empCount} employees in the system.`,
        details: null,
        howToFix: null
      });
    } else {
      warningCount++;
      checks.push({
        id: 'employees_exist',
        status: 'warning',
        title: 'No Employees Found',
        message: 'The employee roster is empty.',
        details: null,
        howToFix: 'Upload OT data to automatically populate the employee roster, or manually add employees to the Employees sheet.'
      });
    }
  }
  
  // ========== CHECK 3: Uniform Catalog Has Items ==========
  const catalogSheet = ss.getSheetByName('Uniform_Catalog');
  if (catalogSheet) {
    const catalogCount = Math.max(0, catalogSheet.getLastRow() - 1);
    if (catalogCount > 0) {
      checks.push({
        id: 'catalog_populated',
        status: 'success',
        title: 'Uniform Catalog Populated',
        message: `${catalogCount} items in the catalog.`,
        details: null,
        howToFix: null
      });
    } else {
      warningCount++;
      checks.push({
        id: 'catalog_populated',
        status: 'warning',
        title: 'Uniform Catalog Empty',
        message: 'No items in the uniform catalog.',
        details: null,
        howToFix: 'Go to Uniforms → Catalog and add uniform items with prices.'
      });
    }
  }
  
  // ========== CHECK 4: OT Data Freshness ==========
  const otSheet = ss.getSheetByName('OT_History');
  if (otSheet && otSheet.getLastRow() > 1) {
    const periods = otSheet.getRange(2, 1, otSheet.getLastRow() - 1, 1).getValues();
    const dates = periods.map(p => p[0] ? new Date(p[0]).getTime() : 0).filter(d => d > 0);
    
    if (dates.length > 0) {
      const latestUpload = new Date(Math.max(...dates));
      const daysSinceUpload = Math.floor((new Date() - latestUpload) / (1000 * 60 * 60 * 24));
      
      if (daysSinceUpload <= 7) {
        checks.push({
          id: 'ot_freshness',
          status: 'success',
          title: 'OT Data Up to Date',
          message: `Last OT data uploaded ${daysSinceUpload} day(s) ago.`,
          details: [`Latest period end: ${latestUpload.toLocaleDateString()}`],
          howToFix: null
        });
      } else if (daysSinceUpload <= 14) {
        warningCount++;
        checks.push({
          id: 'ot_freshness',
          status: 'warning',
          title: 'OT Data May Be Stale',
          message: `Last OT data uploaded ${daysSinceUpload} days ago.`,
          details: [`Latest period end: ${latestUpload.toLocaleDateString()}`],
          howToFix: 'Upload new OT data from HotSchedules. Go to Overtime → Upload Data.'
        });
      } else {
        errorCount++;
        checks.push({
          id: 'ot_freshness',
          status: 'error',
          title: 'OT Data Outdated',
          message: `No OT data uploaded in ${daysSinceUpload} days!`,
          details: [`Latest period end: ${latestUpload.toLocaleDateString()}`],
          howToFix: 'Upload new OT data immediately. Go to Overtime → Upload Data.'
        });
      }
    }
  } else {
    warningCount++;
    checks.push({
      id: 'ot_freshness',
      status: 'warning',
      title: 'No OT Data',
      message: 'No overtime data has been uploaded yet.',
      details: null,
      howToFix: 'Upload OT data from HotSchedules. Go to Overtime → Upload Data.'
    });
  }
  
  // ========== CHECK 5: Stale Pending Orders ==========
  const ordersSheet = ss.getSheetByName('Uniform_Orders');
  if (ordersSheet && ordersSheet.getLastRow() > 1) {
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
    const now = new Date();
    const staleOrders30 = [];
    const staleOrders60 = [];
    const staleOrders90 = []; // Critical - needs immediate attention
    
    for (const row of ordersData) {
      const orderId = row[0];
      const employeeName = row[2];
      const orderDate = row[4] ? new Date(row[4]) : null;
      const status = row[12];
      
      if (status === 'Pending' && orderDate) {
        const daysOld = Math.floor((now - orderDate) / (1000 * 60 * 60 * 24));
        if (daysOld > 90) {
          staleOrders90.push({ orderId, employeeName, daysOld });
        } else if (daysOld > 60) {
          staleOrders60.push({ orderId, employeeName, daysOld });
        } else if (daysOld > 30) {
          staleOrders30.push({ orderId, employeeName, daysOld });
        }
      }
    }
    
    if (staleOrders90.length > 0) {
      errorCount++;
      checks.push({
        id: 'stale_orders_90',
        status: 'error',
        title: 'Critical: Orders Pending 90+ Days',
        message: `${staleOrders90.length} order(s) pending for over 90 days! Requires immediate action.`,
        details: staleOrders90.map(o => `${o.orderId} - ${o.employeeName} (${o.daysOld} days)`),
        howToFix: 'These orders need immediate attention. Either activate them if items were received, or cancel if the order is no longer valid. Contact Jeff if unsure.'
      });
    }
    
    if (staleOrders60.length > 0) {
      errorCount++;
      checks.push({
        id: 'stale_orders_60',
        status: 'error',
        title: 'Very Old Pending Orders',
        message: `${staleOrders60.length} order(s) pending for over 60 days!`,
        details: staleOrders60.map(o => `${o.orderId} - ${o.employeeName} (${o.daysOld} days)`),
        howToFix: 'Review these orders. Either mark items as received and activate, or cancel if no longer needed. Go to Uniforms → Orders.'
      });
    }
    
    if (staleOrders30.length > 0) {
      warningCount++;
      checks.push({
        id: 'stale_orders_30',
        status: 'warning',
        title: 'Old Pending Orders',
        message: `${staleOrders30.length} order(s) pending for over 30 days.`,
        details: staleOrders30.map(o => `${o.orderId} - ${o.employeeName} (${o.daysOld} days)`),
        howToFix: 'Review these orders to ensure uniforms have been received and orders are activated.'
      });
    }
    
    if (staleOrders30.length === 0 && staleOrders60.length === 0 && staleOrders90.length === 0) {
      checks.push({
        id: 'stale_orders',
        status: 'success',
        title: 'No Stale Orders',
        message: 'All pending orders are less than 30 days old.',
        details: null,
        howToFix: null
      });
    }
  }
  
  // ========== CHECK 6: System Counters ==========
  const countersSheet = ss.getSheetByName('System_Counters');
  if (countersSheet && countersSheet.getLastRow() >= 2) {
    const counterData = countersSheet.getRange(2, 1, 2, 2).getValues();
    const hasOrderCounter = counterData.some(r => r[0] === 'Order_ID_Counter');
    const hasLineCounter = counterData.some(r => r[0] === 'Line_ID_Counter');
    
    if (hasOrderCounter && hasLineCounter) {
      checks.push({
        id: 'counters_valid',
        status: 'success',
        title: 'System Counters Active',
        message: 'ID generation counters are properly configured.',
        details: counterData.map(r => `${r[0]}: ${r[1]}`),
        howToFix: null
      });
    } else {
      warningCount++;
      checks.push({
        id: 'counters_valid',
        status: 'warning',
        title: 'System Counters Incomplete',
        message: 'Some ID counters may be missing.',
        details: null,
        howToFix: 'Run "Initialize System Counters" from the script editor, or create the first order to auto-initialize.'
      });
    }
  } else {
    warningCount++;
    checks.push({
      id: 'counters_valid',
      status: 'warning',
      title: 'System Counters Not Initialized',
      message: 'The System_Counters sheet is empty or missing.',
      details: null,
      howToFix: 'Create a new uniform order to automatically initialize the counters, or run initializeSystemCounters() from the script editor.'
    });
  }
  
  // ========== CHECK 7: Active Orders Have Valid Data ==========
  if (ordersSheet && ordersSheet.getLastRow() > 1) {
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
    const invalidOrders = [];
    
    for (const row of ordersData) {
      const orderId = row[0];
      const employeeName = row[2];
      const totalAmount = parseFloat(row[5]) || 0;
      const status = row[12];
      
      if (status === 'Active') {
        if (!employeeName || employeeName.trim() === '') {
          invalidOrders.push({ orderId, issue: 'Missing employee name' });
        } else if (totalAmount <= 0) {
          invalidOrders.push({ orderId, issue: 'Zero or negative amount' });
        }
      }
    }
    
    if (invalidOrders.length === 0) {
      checks.push({
        id: 'valid_orders',
        status: 'success',
        title: 'Active Orders Valid',
        message: 'All active orders have valid employee and amount data.',
        details: null,
        howToFix: null
      });
    } else {
      errorCount++;
      checks.push({
        id: 'valid_orders',
        status: 'error',
        title: 'Invalid Active Orders',
        message: `${invalidOrders.length} active order(s) have data issues.`,
        details: invalidOrders.map(o => `${o.orderId}: ${o.issue}`),
        howToFix: 'Review and fix these orders in the Uniform_Orders sheet directly, or contact Jeff.'
      });
    }
  }
  
  // ========== CHECK 8: Auto-Archive Triggers ==========
  try {
    const triggerStatus = getArchiveTriggerStatus();
    if (triggerStatus.annualArchiveActive) {
      checks.push({
        id: 'archive_triggers',
        status: 'success',
        title: 'Auto-Archive Configured',
        message: 'Annual archive trigger is active. Old data will be archived automatically.',
        details: ['Annual archive runs on January 1st'],
        howToFix: null
      });
    } else {
      warningCount++;
      checks.push({
        id: 'archive_triggers',
        status: 'warning',
        title: 'Auto-Archive Not Configured',
        message: 'No automatic archive trigger is set up.',
        details: null,
        howToFix: 'Go to Settings and click "Enable Auto-Archive" to automatically clean up old data each year.'
      });
    }
  } catch (e) {
    // If we can't check triggers, just skip this check
    console.log('Could not check archive triggers:', e);
  }
  
  // ========== CHECK 9: Deduction Conflict Detection (Chunk 14) ==========
  try {
    const conflicts = detectDeductionConflicts();
    
    if (conflicts.totalConflicts === 0) {
      checks.push({
        id: 'deduction_conflicts',
        status: 'success',
        title: 'No Deduction Conflicts',
        message: 'All deduction data is consistent with no integrity issues.',
        details: null,
        howToFix: null
      });
    } else {
      // Build combined details list
      const conflictDetails = [];
      
      if (conflicts.duplicates.length > 0) {
        conflictDetails.push(`📋 ${conflicts.duplicates.length} potential duplicate(s)`);
        conflicts.duplicates.slice(0, 3).forEach(d => conflictDetails.push(`   • ${d.description}`));
      }
      
      if (conflicts.zombieOrders.length > 0) {
        conflictDetails.push(`🧟 ${conflicts.zombieOrders.length} zombie order(s)`);
        conflicts.zombieOrders.slice(0, 3).forEach(z => conflictDetails.push(`   • ${z.description}`));
      }
      
      if (conflicts.overcharges.length > 0) {
        conflictDetails.push(`💰 ${conflicts.overcharges.length} over-charge(s)`);
        conflicts.overcharges.slice(0, 3).forEach(o => conflictDetails.push(`   • ${o.description}`));
      }
      
      if (conflicts.orphans.length > 0) {
        conflictDetails.push(`👻 ${conflicts.orphans.length} orphaned record issue(s)`);
        conflicts.orphans.slice(0, 3).forEach(o => conflictDetails.push(`   • ${o.description}`));
      }
      
      // Determine severity based on conflict types
      const hasHighPriority = conflicts.overcharges.length > 0 || conflicts.duplicates.length > 0;
      
      if (hasHighPriority) {
        errorCount++;
        checks.push({
          id: 'deduction_conflicts',
          status: 'error',
          title: 'Deduction Conflicts Detected',
          message: `${conflicts.totalConflicts} data integrity issue(s) found that need attention.`,
          details: conflictDetails,
          howToFix: 'Review the Uniform Orders page. Cancel duplicates, update zombie order statuses, and process any needed refunds for over-charges.'
        });
      } else {
        warningCount++;
        checks.push({
          id: 'deduction_conflicts',
          status: 'warning',
          title: 'Minor Deduction Issues',
          message: `${conflicts.totalConflicts} minor data issue(s) found.`,
          details: conflictDetails,
          howToFix: 'Review and clean up these records when convenient. These are not urgent but should be addressed.'
        });
      }
    }
  } catch (e) {
    console.log('Could not run conflict detection:', e);
  }
  
  // ========== Determine Overall Status ==========
  let overallStatus = 'healthy';
  if (warningCount > 0) overallStatus = 'warnings';
  if (errorCount > 0) overallStatus = 'errors';
  
  console.log('Health check complete. Status: ' + overallStatus);
  
  return {
    success: true,
    status: overallStatus,
    errorCount: errorCount,
    warningCount: warningCount,
    successCount: checks.filter(c => c.status === 'success').length,
    checks: checks,
    lastRun: new Date().toISOString()
  };
  } catch (error) {
    console.error('Health check failed:', error);
    return {
      success: false,
      status: 'errors',
      errorCount: 1,
      warningCount: 0,
      successCount: 0,
      checks: [{
        id: 'system_error',
        status: 'error',
        title: 'Health Check Failed',
        message: 'An error occurred while running health checks.',
        details: [error.message || String(error)],
        howToFix: 'Contact Jeff. The error details have been logged.'
      }],
      lastRun: new Date().toISOString()
    };
  }
}

// ============================================================
// CHUNK 7: AUTO-ARCHIVE & SELF-CLEANING FUNCTIONS
// ============================================================

/**
 * Archive completed orders from prior years
 * Moves orders with status "Completed" from years before current year
 * to an "Archive_Orders_[YEAR]" sheet
 * @param {boolean} dryRun - If true, only returns what would be archived without moving
 * @returns {Object} Result with counts and details
 */
function archiveCompletedOrders(dryRun = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet) {
      return { success: false, error: 'Uniform_Orders sheet not found' };
    }
    
    const currentYear = new Date().getFullYear();
    const ordersData = ordersSheet.getDataRange().getValues();
    const headers = ordersData[0];
    
    // Find orders to archive (Completed status from prior years)
    const ordersToArchive = [];
    const rowsToDelete = []; // Store row indices to delete (1-based)
    
    for (let i = 1; i < ordersData.length; i++) {
      const row = ordersData[i];
      const orderId = row[0];
      const orderDate = row[4];
      const status = row[12];
      
      if (status === 'Completed' && orderDate) {
        const orderYear = new Date(orderDate).getFullYear();
        if (orderYear < currentYear) {
          ordersToArchive.push({
            rowIndex: i + 1, // 1-based for sheet operations
            year: orderYear,
            data: row,
            orderId: orderId
          });
          rowsToDelete.push(i + 1);
        }
      }
    }
    
    if (ordersToArchive.length === 0) {
      return { 
        success: true, 
        message: 'No completed orders from prior years to archive',
        archivedCount: 0 
      };
    }
    
    if (dryRun) {
      return {
        success: true,
        dryRun: true,
        wouldArchive: ordersToArchive.length,
        orders: ordersToArchive.map(o => ({ orderId: o.orderId, year: o.year }))
      };
    }
    
    // Group orders by year
    const ordersByYear = {};
    for (const order of ordersToArchive) {
      if (!ordersByYear[order.year]) {
        ordersByYear[order.year] = [];
      }
      ordersByYear[order.year].push(order);
    }
    
    // Archive to year-specific sheets
    let totalArchived = 0;
    for (const [year, orders] of Object.entries(ordersByYear)) {
      const archiveSheetName = `Archive_Orders_${year}`;
      let archiveSheet = ss.getSheetByName(archiveSheetName);
      
      // Create archive sheet if it doesn't exist
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveSheetName);
        archiveSheet.appendRow(headers);
        archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        // Protect the archive sheet (read-only for most users)
        const protection = archiveSheet.protect().setDescription('Archived orders - read only');
        protection.setWarningOnly(true); // Shows warning but doesn't block
      }
      
      // Append orders to archive
      for (const order of orders) {
        archiveSheet.appendRow(order.data);
        totalArchived++;
      }
    }
    
    // Also archive related order items
    if (itemsSheet) {
      const itemsData = itemsSheet.getDataRange().getValues();
      const itemsHeaders = itemsData[0];
      const archivedOrderIds = ordersToArchive.map(o => o.orderId);
      
      const itemsToArchive = [];
      const itemRowsToDelete = [];
      
      for (let i = 1; i < itemsData.length; i++) {
        const orderId = itemsData[i][1]; // Order_ID is column B
        if (archivedOrderIds.includes(orderId)) {
          const orderYear = ordersToArchive.find(o => o.orderId === orderId)?.year;
          if (orderYear) {
            itemsToArchive.push({ data: itemsData[i], year: orderYear });
            itemRowsToDelete.push(i + 1);
          }
        }
      }
      
      // Archive items by year
      for (const item of itemsToArchive) {
        const archiveItemsSheetName = `Archive_Order_Items_${item.year}`;
        let archiveItemsSheet = ss.getSheetByName(archiveItemsSheetName);
        
        if (!archiveItemsSheet) {
          archiveItemsSheet = ss.insertSheet(archiveItemsSheetName);
          archiveItemsSheet.appendRow(itemsHeaders);
          archiveItemsSheet.getRange(1, 1, 1, itemsHeaders.length).setFontWeight('bold');
          const protection = archiveItemsSheet.protect().setDescription('Archived order items - read only');
          protection.setWarningOnly(true);
        }
        
        archiveItemsSheet.appendRow(item.data);
      }
      
      // Delete archived items from main sheet (in reverse order to preserve indices)
      itemRowsToDelete.sort((a, b) => b - a);
      for (const rowIndex of itemRowsToDelete) {
        itemsSheet.deleteRow(rowIndex);
      }
    }
    
    // Delete archived orders from main sheet (in reverse order)
    rowsToDelete.sort((a, b) => b - a);
    for (const rowIndex of rowsToDelete) {
      ordersSheet.deleteRow(rowIndex);
    }
    
    // Log the archive action
    logActivity('ARCHIVE', 'UNIFORM', `Archived ${totalArchived} completed orders from prior years`);
    
    return {
      success: true,
      message: `Successfully archived ${totalArchived} orders`,
      archivedCount: totalArchived,
      byYear: Object.fromEntries(Object.entries(ordersByYear).map(([y, o]) => [y, o.length]))
    };
    
  } catch (error) {
    console.error('Error archiving orders:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Archive old OT data
 * Keeps current year data in main sheet, moves older data to archive
 * @param {boolean} dryRun - If true, only returns what would be archived
 * @returns {Object} Result with counts and details
 */
function archiveOTData(dryRun = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const otSheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    
    if (!otSheet || otSheet.getLastRow() < 2) {
      return { success: true, message: 'No OT data to archive', archivedCount: 0 };
    }
    
    const currentYear = new Date().getFullYear();
    const data = otSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find rows to archive (from years before current year)
    const rowsToArchive = [];
    const rowIndicesToDelete = [];
    
    for (let i = 1; i < data.length; i++) {
      const periodEnd = data[i][0]; // Period_End is column A
      if (periodEnd) {
        const periodYear = new Date(periodEnd).getFullYear();
        if (periodYear < currentYear) {
          rowsToArchive.push({
            rowIndex: i + 1,
            year: periodYear,
            data: data[i]
          });
          rowIndicesToDelete.push(i + 1);
        }
      }
    }
    
    if (rowsToArchive.length === 0) {
      return { success: true, message: 'No OT data from prior years to archive', archivedCount: 0 };
    }
    
    if (dryRun) {
      const byYear = {};
      for (const row of rowsToArchive) {
        byYear[row.year] = (byYear[row.year] || 0) + 1;
      }
      return {
        success: true,
        dryRun: true,
        wouldArchive: rowsToArchive.length,
        byYear: byYear
      };
    }
    
    // Group by year
    const dataByYear = {};
    for (const row of rowsToArchive) {
      if (!dataByYear[row.year]) {
        dataByYear[row.year] = [];
      }
      dataByYear[row.year].push(row.data);
    }
    
    // Archive to year-specific sheets
    let totalArchived = 0;
    for (const [year, rows] of Object.entries(dataByYear)) {
      const archiveSheetName = `Archive_OT_${year}`;
      let archiveSheet = ss.getSheetByName(archiveSheetName);
      
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveSheetName);
        archiveSheet.appendRow(headers);
        archiveSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        const protection = archiveSheet.protect().setDescription('Archived OT data - read only');
        protection.setWarningOnly(true);
      }
      
      for (const rowData of rows) {
        archiveSheet.appendRow(rowData);
        totalArchived++;
      }
    }
    
    // Delete archived rows from main sheet (in reverse order)
    rowIndicesToDelete.sort((a, b) => b - a);
    for (const rowIndex of rowIndicesToDelete) {
      otSheet.deleteRow(rowIndex);
    }
    
    logActivity('ARCHIVE', 'OT', `Archived ${totalArchived} OT records from prior years`);
    
    return {
      success: true,
      message: `Successfully archived ${totalArchived} OT records`,
      archivedCount: totalArchived,
      byYear: Object.fromEntries(Object.entries(dataByYear).map(([y, r]) => [y, r.length]))
    };
    
  } catch (error) {
    console.error('Error archiving OT data:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get active employees only (filters out inactive)
 * Uses the existing Status column (index 4) in Employees sheet
 * Status can be 'Active', 'Inactive', or blank (treated as Active)
 * @returns {Array} Array of active employee objects
 */
function getActiveEmployees() {
  // Use the existing getEmployees function with status filter
  // This ensures consistency with the rest of the codebase
  const allEmployees = getEmployees();
  
  // Filter to only active employees
  // Status column defaults to 'Active' if blank (see getEmployees line 1234)
  return allEmployees.filter(emp => {
    const status = (emp.status || 'Active').toLowerCase();
    return status === 'active' || status === '';
  }).map(emp => ({
    id: emp.employeeId,
    name: emp.fullName,
    location: emp.primaryLocation
  }));
}

/**
 * Set employee active status
 * Uses the existing Status column (column E, index 4) in Employees sheet
 * @param {string} employeeId - Employee ID
 * @param {boolean} isActive - Whether employee should be active
 * @returns {Object} Result
 */
function setEmployeeActiveStatus(employeeId, isActive) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (!empSheet) {
      return { success: false, error: 'Employees sheet not found' };
    }
    
    const data = empSheet.getDataRange().getValues();
    
    // Employee_ID is column A (index 0), Status is column E (index 4)
    const ID_COL = 0;
    const STATUS_COL = 4; // Column E = Status
    
    // Find employee row
    let employeeRow = -1;
    let employeeName = '';
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][ID_COL]) === String(employeeId)) {
        employeeRow = i + 1; // 1-based for sheet operations
        employeeName = data[i][1] || ''; // Full_Name is column B
        break;
      }
    }
    
    if (employeeRow === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Update status column (column E = column 5 in 1-based)
    const newStatus = isActive ? 'Active' : 'Inactive';
    empSheet.getRange(employeeRow, STATUS_COL + 1).setValue(newStatus);
    
    logActivity('UPDATE', 'EMPLOYEE', `Set employee ${employeeName || employeeId} status to ${newStatus}`, employeeId);
    
    return { 
      success: true, 
      message: `Employee ${employeeName || employeeId} ${isActive ? 'activated' : 'deactivated'} successfully` 
    };
    
  } catch (error) {
    console.error('Error setting employee status:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Run all archive operations (for annual cleanup)
 * @param {boolean} dryRun - If true, only preview what would happen
 * @returns {Object} Combined results
 */
function runAnnualArchive(dryRun = false) {
  console.log('Running annual archive...' + (dryRun ? ' (DRY RUN)' : ''));
  
  const ordersResult = archiveCompletedOrders(dryRun);
  const otResult = archiveOTData(dryRun);
  
  return {
    success: ordersResult.success && otResult.success,
    dryRun: dryRun,
    orders: ordersResult,
    otData: otResult,
    timestamp: new Date().toISOString()
  };
}

/**
 * Setup automatic archive triggers
 * Creates time-based triggers for:
 * - Annual archive on January 1st
 * - Daily health check (optional)
 */
function setupAutoArchiveTriggers() {
  // Remove existing archive triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'runScheduledAnnualArchive' || funcName === 'runScheduledHealthCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create annual archive trigger (January 1st at 2 AM)
  ScriptApp.newTrigger('runScheduledAnnualArchive')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  
  console.log('Auto-archive triggers configured');
  
  return { success: true, message: 'Archive triggers set up successfully' };
}

/**
 * Scheduled function for annual archive (called by trigger)
 */
function runScheduledAnnualArchive() {
  // Only run in January
  const now = new Date();
  if (now.getMonth() !== 0) { // 0 = January
    console.log('Not January, skipping annual archive');
    return;
  }
  
  console.log('Running scheduled annual archive');
  const result = runAnnualArchive(false);
  
  // Send summary email
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      const subject = 'OT Tracker Pro - Annual Archive Complete';
      const body = `
Annual archive completed on ${new Date().toLocaleDateString()}

Orders Archived: ${result.orders.archivedCount || 0}
OT Records Archived: ${result.otData.archivedCount || 0}

${result.success ? '✅ All archives completed successfully.' : '⚠️ Some errors occurred. Please check the system.'}

This is an automated message from OT Tracker Pro.
      `.trim();
      
      MailApp.sendEmail(email, subject, body);
    }
  } catch (e) {
    console.error('Failed to send archive summary email:', e);
  }
  
  return result;
}

/**
 * Check if archive triggers are set up
 * @returns {Object} Trigger status
 */
function getArchiveTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  let annualArchive = false;
  
  for (const trigger of triggers) {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'runScheduledAnnualArchive') {
      annualArchive = true;
    }
  }
  
  return {
    annualArchiveActive: annualArchive
  };
}

/**
 * Remove all archive triggers
 */
function removeArchiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  for (const trigger of triggers) {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'runScheduledAnnualArchive' || funcName === 'runScheduledHealthCheck') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  
  return { success: true, removed: removed };
}

// ============================================================
// CHUNK 9: PRE-PAYROLL VALIDATION SCANNER
// ============================================================

/**
 * Run pre-payroll validation checks
 * @returns {Object} Validation results with errors, warnings, and passes
 */
function runPayrollValidation() {
  try {
    console.log('Running payroll validation...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const checks = [];
    let errorCount = 0;
    let warningCount = 0;
    
    // ===== CHECK 1: Orders activated but missing payment schedules =====
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
      const missingSchedule = [];
      
      for (const row of ordersData) {
        const orderId = row[0];
        const employeeName = row[2];
        const status = row[12];
        const amountPerPaycheck = parseFloat(row[7]) || 0;
        const firstDeductionDate = row[8];
        
        if (status === 'Active' && (amountPerPaycheck <= 0 || !firstDeductionDate)) {
          missingSchedule.push({ orderId, employeeName });
        }
      }
      
      if (missingSchedule.length > 0) {
        errorCount++;
        checks.push({
          id: 'missing_schedule',
          status: 'error',
          title: 'Orders Missing Payment Schedule',
          message: `${missingSchedule.length} active order(s) have no payment schedule set.`,
          details: missingSchedule.map(o => `${o.orderId} - ${o.employeeName}`),
          action: 'uniforms-orders'
        });
      } else {
        checks.push({
          id: 'missing_schedule',
          status: 'success',
          title: 'Payment Schedules Complete',
          message: 'All active orders have payment schedules configured.',
          details: null,
          action: null
        });
      }
    }
    
    // ===== CHECK 2: Inactive employees with active deductions =====
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    if (ordersSheet && empSheet && ordersSheet.getLastRow() > 1 && empSheet.getLastRow() > 1) {
      const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
      const empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 5).getValues();
      
      // Build map of inactive employees
      const inactiveEmps = new Set();
      for (const emp of empData) {
        const empId = emp[0];
        const status = (emp[4] || 'Active').toLowerCase();
        if (status === 'inactive') {
          inactiveEmps.add(String(empId));
        }
      }
      
      const inactiveWithDeductions = [];
      for (const row of ordersData) {
        const orderId = row[0];
        const empId = row[1];
        const employeeName = row[2];
        const status = row[12];
        
        if (status === 'Active' && inactiveEmps.has(String(empId))) {
          inactiveWithDeductions.push({ orderId, employeeName });
        }
      }
      
      if (inactiveWithDeductions.length > 0) {
        warningCount++;
        checks.push({
          id: 'inactive_deductions',
          status: 'warning',
          title: 'Inactive Employees with Deductions',
          message: `${inactiveWithDeductions.length} inactive employee(s) have active deductions.`,
          details: inactiveWithDeductions.map(o => `${o.orderId} - ${o.employeeName}`),
          action: 'uniforms-orders'
        });
      } else {
        checks.push({
          id: 'inactive_deductions',
          status: 'success',
          title: 'No Inactive Employee Deductions',
          message: 'No inactive employees have pending deductions.',
          details: null,
          action: null
        });
      }
    }
    
    // ===== CHECK 3: OT data freshness =====
    const otSheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
    if (otSheet && otSheet.getLastRow() > 1) {
      const periods = otSheet.getRange(2, 1, otSheet.getLastRow() - 1, 1).getValues();
      const dates = [];
      for (const p of periods) {
        if (p[0]) {
          try {
            const d = new Date(p[0]);
            const t = d.getTime();
            if (!isNaN(t) && t > 0) {
              dates.push(t);
            }
          } catch (e) {
            // Skip invalid date
          }
        }
      }
      
      if (dates.length > 0) {
        const latestTimestamp = Math.max(...dates);
        const latestUpload = new Date(latestTimestamp);
        
        // Validate latestUpload is a valid date before using it
        if (!isNaN(latestUpload.getTime())) {
          const daysSinceUpload = Math.floor((new Date() - latestUpload) / (1000 * 60 * 60 * 24));
          const uploadDateStr = latestUpload.toLocaleDateString();
          
          if (daysSinceUpload <= 7) {
            checks.push({
              id: 'ot_freshness',
              status: 'success',
              title: 'OT Data Current',
              message: `OT data uploaded ${daysSinceUpload} day(s) ago.`,
              details: null,
              action: null
            });
          } else if (daysSinceUpload <= 14) {
            warningCount++;
            checks.push({
              id: 'ot_freshness',
              status: 'warning',
              title: 'OT Data May Be Stale',
              message: `OT data is ${daysSinceUpload} days old. Consider uploading fresh data.`,
              details: [`Last upload: ${uploadDateStr}`],
              action: 'ot-upload'
            });
          } else {
            errorCount++;
            checks.push({
              id: 'ot_freshness',
              status: 'error',
              title: 'OT Data Outdated',
              message: `OT data is ${daysSinceUpload} days old! Upload current data before processing.`,
              details: [`Last upload: ${uploadDateStr}`],
              action: 'ot-upload'
            });
          }
        }
      }
    } else {
      warningCount++;
      checks.push({
        id: 'ot_freshness',
        status: 'warning',
        title: 'No OT Data',
        message: 'No overtime data found. Upload OT data if needed.',
        details: null,
        action: 'ot-upload'
      });
    }
    
    // ===== CHECK 4: Orders with $0 total but active =====
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
      const zeroOrders = [];
      
      for (const row of ordersData) {
        const orderId = row[0];
        const employeeName = row[2];
        const totalAmount = parseFloat(row[5]) || 0;
        const status = row[12];
        
        if (status === 'Active' && totalAmount <= 0) {
          zeroOrders.push({ orderId, employeeName });
        }
      }
      
      if (zeroOrders.length > 0) {
        errorCount++;
        checks.push({
          id: 'zero_orders',
          status: 'error',
          title: 'Zero-Value Active Orders',
          message: `${zeroOrders.length} active order(s) have $0.00 total.`,
          details: zeroOrders.map(o => `${o.orderId} - ${o.employeeName}`),
          action: 'uniforms-orders'
        });
      } else {
        checks.push({
          id: 'zero_orders',
          status: 'success',
          title: 'Order Values Valid',
          message: 'All active orders have valid amounts.',
          details: null,
          action: null
        });
      }
    }
    
    // ===== CHECK 5: Deduction totals reconciliation =====
    // Get next payroll date to check what's due THIS period vs future periods
    let nextPayrollDateStr = null;
    try {
      const nextPayrollDate = getPayrollSetting('next_payroll_date');
      if (nextPayrollDate) {
        const npd = new Date(nextPayrollDate);
        if (!isNaN(npd.getTime())) {
          nextPayrollDateStr = npd.toISOString().split('T')[0];
        }
      }
    } catch (e) {
      console.log('Could not parse next payroll date for deduction check');
    }
    
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
      let totalExpected = 0;
      let totalRemaining = 0;
      let activeCount = 0;
      let dueThisPeriodCount = 0;
      let dueThisPeriodAmount = 0;
      
      for (const row of ordersData) {
        const status = row[12];
        const amountRemaining = parseFloat(row[11]) || 0;
        const totalAmount = parseFloat(row[5]) || 0;
        const amountPaid = parseFloat(row[10]) || 0;
        const nextPaymentDate = row[8]; // Column I = Next Payment Date
        
        if (status === 'Active') {
          activeCount++;
          totalRemaining += amountRemaining;
          totalExpected += (totalAmount - amountPaid);
          
          // Check if this order has a payment due on the next payroll date
          if (nextPaymentDate && nextPayrollDateStr) {
            try {
              let nextPaymentStr = null;
              if (nextPaymentDate instanceof Date && !isNaN(nextPaymentDate.getTime())) {
                nextPaymentStr = nextPaymentDate.toISOString().split('T')[0];
              } else if (typeof nextPaymentDate === 'string') {
                const parsed = new Date(nextPaymentDate);
                if (!isNaN(parsed.getTime())) {
                  nextPaymentStr = parsed.toISOString().split('T')[0];
                }
              }
              
              if (nextPaymentStr === nextPayrollDateStr) {
                dueThisPeriodCount++;
                const paymentAmount = parseFloat(row[7]) || 0; // Amount_Per_Paycheck (column H, index 7)
                dueThisPeriodAmount += Math.min(paymentAmount, amountRemaining);
              }
            } catch (e) {
              // Skip this row if date parsing fails
              console.log('Skipping row due to date parse error');
            }
          }
        }
      }
      
      const discrepancy = Math.abs(totalRemaining - totalExpected);
      if (discrepancy > 0.01) {
        warningCount++;
        checks.push({
          id: 'deduction_reconcile',
          status: 'warning',
          title: 'Deduction Discrepancy',
          message: `$${discrepancy.toFixed(2)} discrepancy in deduction totals.`,
          details: [
            `Expected remaining: $${totalExpected.toFixed(2)}`,
            `Recorded remaining: $${totalRemaining.toFixed(2)}`
          ],
          action: 'uniforms-orders'
        });
      } else if (dueThisPeriodCount > 0) {
        // Show what's due this period
        checks.push({
          id: 'deduction_reconcile',
          status: 'success',
          title: 'Deductions Ready',
          message: `${dueThisPeriodCount} deduction(s) due this period ($${dueThisPeriodAmount.toFixed(2)}).`,
          details: activeCount > dueThisPeriodCount ? [`${activeCount - dueThisPeriodCount} more orders scheduled for future periods.`] : null,
          action: null
        });
      } else if (activeCount > 0) {
        // No deductions due this period, but there are future ones
        checks.push({
          id: 'deduction_reconcile',
          status: 'success',
          title: 'No Deductions This Period',
          message: `${activeCount} active order(s) scheduled for future pay periods.`,
          details: null,
          action: null
        });
      } else {
        // No active orders at all
        checks.push({
          id: 'deduction_reconcile',
          status: 'success',
          title: 'No Active Deductions',
          message: 'No uniform deduction orders currently active.',
          details: null,
          action: null
        });
      }
    }
    
    // ===== CHECK 6: Deduction Conflict Detection (Chunk 14) =====
    const conflicts = detectDeductionConflicts();
    if (conflicts.totalConflicts > 0) {
      if (conflicts.duplicates.length > 0) {
        errorCount++;
        checks.push({
          id: 'conflict_duplicates',
          status: 'error',
          title: 'Duplicate Deductions Detected',
          message: `${conflicts.duplicates.length} potential duplicate deduction(s) found.`,
          details: conflicts.duplicates.map(d => d.description),
          action: 'uniforms-orders',
          howToFix: 'Review these orders and cancel any duplicates. Check if the same item was ordered twice accidentally.'
        });
      }
      
      if (conflicts.zombieOrders.length > 0) {
        errorCount++;
        checks.push({
          id: 'conflict_zombies',
          status: 'error',
          title: 'Zombie Orders Detected',
          message: `${conflicts.zombieOrders.length} order(s) have status/balance mismatches.`,
          details: conflicts.zombieOrders.map(z => z.description),
          action: 'uniforms-orders',
          howToFix: 'Update order status to match payment status, or adjust remaining balance.'
        });
      }
      
      if (conflicts.overcharges.length > 0) {
        errorCount++;
        checks.push({
          id: 'conflict_overcharge',
          status: 'error',
          title: 'Over-Charged Orders',
          message: `${conflicts.overcharges.length} order(s) have been charged more than the order total.`,
          details: conflicts.overcharges.map(o => o.description),
          action: 'uniforms-orders',
          howToFix: 'Review payment history and issue refund or adjustment as needed.'
        });
      }
      
      if (conflicts.orphans.length > 0) {
        warningCount++;
        checks.push({
          id: 'conflict_orphans',
          status: 'warning',
          title: 'Orphaned Data',
          message: `${conflicts.orphans.length} data integrity issue(s) found.`,
          details: conflicts.orphans.map(o => o.description),
          action: 'uniforms-orders',
          howToFix: 'Review and clean up orphaned records in the spreadsheet.'
        });
      }
    } else {
      checks.push({
        id: 'conflict_check',
        status: 'success',
        title: 'No Deduction Conflicts',
        message: 'All deduction data is consistent.',
        details: null,
        action: null
      });
    }
    
    // Determine overall status
    let overallStatus = 'ready';
    if (warningCount > 0) overallStatus = 'warnings';
    if (errorCount > 0) overallStatus = 'errors';
    
    return {
      success: true,
      status: overallStatus,
      errorCount: errorCount,
      warningCount: warningCount,
      passCount: checks.filter(c => c.status === 'success').length,
      checks: checks,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Payroll validation failed:', error);
    return {
      success: false,
      status: 'errors',
      errorCount: 1,
      warningCount: 0,
      passCount: 0,
      checks: [{
        id: 'system_error',
        status: 'error',
        title: 'Validation Failed',
        message: error.message || 'An error occurred during validation.',
        details: null,
        action: null
      }],
      timestamp: new Date().toISOString()
    };
  }
}

// ============================================================
// CHUNK 15: SMART PAYROLL PERIOD CALENDAR
// ============================================================

/**
 * Get comprehensive payroll calendar data
 * @param {number} monthsAhead - How many months of paydays to return (default 3)
 * @returns {Object} Calendar data including paydays, current period, and deductions
 */
function getPayrollCalendarData(monthsAhead = 3) {
  try {
    const settings = getSettings();
    const referenceDate = new Date(settings.paydayReference || '2024-11-29');
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // ========== Calculate Pay Periods (Sunday-Saturday) ==========
    const paydays = [];
    let currentPayday = new Date(referenceDate);
    
    // Find the next payday after today
    while (currentPayday <= today) {
      currentPayday.setDate(currentPayday.getDate() + 14);
    }
    const nextPayday = new Date(currentPayday);
    
    // Calculate CURRENT pay period we're in (Sunday to Saturday)
    // Find the most recent Sunday (or today if it's Sunday)
    const currentPeriodStart = new Date(today);
    const dayOfWeek = currentPeriodStart.getDay(); // 0 = Sunday
    currentPeriodStart.setDate(currentPeriodStart.getDate() - dayOfWeek);
    
    // Period ends on the following Saturday
    const currentPeriodEnd = new Date(currentPeriodStart);
    currentPeriodEnd.setDate(currentPeriodEnd.getDate() + 6);
    
    // Days until next payroll
    const daysUntilPayday = Math.ceil((nextPayday - today) / (1000 * 60 * 60 * 24));
    
    // Generate upcoming paydays (next 3 months worth)
    const endDate = new Date(today);
    endDate.setMonth(endDate.getMonth() + monthsAhead);
    
    let paydayDate = new Date(nextPayday);
    while (paydayDate <= endDate) {
      // Get preview for this payday
      const preview = getPayrollDatePreview(paydayDate);
      
      paydays.push({
        date: paydayDate.toISOString().split('T')[0],
        dateFormatted: formatDateShort(paydayDate),
        dayOfWeek: paydayDate.toLocaleDateString('en-US', { weekday: 'short' }),
        employeeCount: preview.employeeCount,
        totalAmount: preview.totalAmount,
        completingOrders: preview.completingOrders,
        isPast: paydayDate < today,
        isNext: paydays.length === 0
      });
      
      paydayDate.setDate(paydayDate.getDate() + 14);
    }
    
    // ========== Get OT Upload Due Dates (Mondays of PAYROLL weeks only) ==========
    // OT is due on the Monday of the same week as payday (Friday)
    // If payday is Friday, Monday of that week = payday - 4 days
    const otDueDates = [];
    
    for (const payday of paydays) {
      // Parse date carefully to avoid timezone issues
      const parts = payday.date.split('-');
      const paydayDateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]), 12, 0, 0);
      
      // Monday of the same week = Friday - 4 days
      const mondayOfPayweek = new Date(paydayDateObj);
      mondayOfPayweek.setDate(paydayDateObj.getDate() - 4);
      
      otDueDates.push({
        date: mondayOfPayweek.toISOString().split('T')[0],
        dateFormatted: formatDateShort(mondayOfPayweek)
      });
    }
    
    // ========== Current Cycle Stats ==========
    const nextPaydayPreview = paydays.length > 0 ? paydays[0] : null;
    
    return {
      success: true,
      currentPeriod: {
        start: currentPeriodStart.toISOString().split('T')[0],
        end: currentPeriodEnd.toISOString().split('T')[0],
        startFormatted: formatDateShort(currentPeriodStart),
        endFormatted: formatDateShort(currentPeriodEnd)
      },
      nextPayday: {
        date: nextPayday.toISOString().split('T')[0],
        dateFormatted: formatDateShort(nextPayday),
        daysUntil: daysUntilPayday,
        employeeCount: nextPaydayPreview ? nextPaydayPreview.employeeCount : 0,
        totalAmount: nextPaydayPreview ? nextPaydayPreview.totalAmount : 0,
        completingOrders: nextPaydayPreview ? nextPaydayPreview.completingOrders : 0
      },
      upcomingPaydays: paydays.slice(0, 6), // Next 6 paydays
      otDueDates: otDueDates,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Error getting payroll calendar data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get deduction preview for a specific payroll date
 * @param {Date|string} paydayDate - The payday to get preview for
 * @returns {Object} Preview of deductions for that date
 */
function getPayrollDatePreview(paydayDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return { employeeCount: 0, totalAmount: 0, completingOrders: 0, orders: [] };
    }
    
    const targetDate = new Date(paydayDate);
    targetDate.setHours(0, 0, 0, 0);
    
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
    
    const affectedOrders = [];
    let totalAmount = 0;
    let completingOrders = 0;
    const employeeSet = new Set();
    
    for (const row of ordersData) {
      const orderId = row[0];
      const employeeName = row[2];
      const amountPerPaycheck = parseFloat(row[7]) || 0;
      const firstDeductionDate = row[8] ? new Date(row[8]) : null;
      const paymentsMade = parseInt(row[9]) || 0;
      const paymentPlan = parseInt(row[6]) || 1;
      const amountRemaining = parseFloat(row[11]) || 0;
      const status = row[12];
      
      if (status !== 'Active' || amountPerPaycheck <= 0 || !firstDeductionDate) {
        continue;
      }
      
      // Check if this payday falls on a scheduled deduction
      firstDeductionDate.setHours(0, 0, 0, 0);
      
      // Calculate days between first deduction and target date
      const daysDiff = Math.round((targetDate - firstDeductionDate) / (1000 * 60 * 60 * 24));
      
      // If daysDiff is non-negative and divisible by 14, it's a scheduled deduction
      if (daysDiff >= 0 && daysDiff % 14 === 0) {
        // Calculate which payment number this would be
        const paymentNumber = Math.floor(daysDiff / 14) + 1;
        
        // Only include if we haven't exceeded the payment plan
        if (paymentNumber <= paymentPlan && paymentsMade < paymentPlan) {
          employeeSet.add(employeeName);
          
          // Calculate actual deduction amount (might be less for final payment)
          const deductionAmount = Math.min(amountPerPaycheck, amountRemaining);
          totalAmount += deductionAmount;
          
          // Check if this is the final payment
          const isFinalPayment = paymentNumber === paymentPlan || deductionAmount >= amountRemaining - 0.01;
          if (isFinalPayment) {
            completingOrders++;
          }
          
          affectedOrders.push({
            orderId,
            employeeName,
            deductionAmount,
            paymentNumber,
            totalPayments: paymentPlan,
            isFinalPayment,
            amountRemaining
          });
        }
      }
    }
    
    return {
      employeeCount: employeeSet.size,
      totalAmount: parseFloat(totalAmount.toFixed(2)),
      completingOrders,
      orders: affectedOrders
    };
    
  } catch (error) {
    console.error('Error getting payroll date preview:', error);
    return { employeeCount: 0, totalAmount: 0, completingOrders: 0, orders: [] };
  }
}

/**
 * Get historical payroll data for past cycles
 * @param {number} cyclesBack - How many past cycles to retrieve
 * @returns {Array} Array of past payroll cycle data
 */
function getPastPayrollCycles(cyclesBack = 6) {
  try {
    const settings = getSettings();
    const referenceDate = new Date(settings.paydayReference || '2024-11-29');
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Find current payday
    let currentPayday = new Date(referenceDate);
    while (currentPayday <= today) {
      currentPayday.setDate(currentPayday.getDate() + 14);
    }
    
    // Go back to find past paydays
    const pastCycles = [];
    let payday = new Date(currentPayday);
    payday.setDate(payday.getDate() - 14); // Start with most recent past payday
    
    for (let i = 0; i < cyclesBack; i++) {
      if (payday < referenceDate) break;
      
      const preview = getPayrollDatePreview(payday);
      pastCycles.push({
        date: payday.toISOString().split('T')[0],
        dateFormatted: formatDateShort(payday),
        employeeCount: preview.employeeCount,
        totalAmount: preview.totalAmount,
        completingOrders: preview.completingOrders
      });
      
      payday.setDate(payday.getDate() - 14);
    }
    
    return pastCycles;
    
  } catch (error) {
    console.error('Error getting past payroll cycles:', error);
    return [];
  }
}

/**
 * Helper function to format date in short format
 */
function formatDateShort(date) {
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

// ============================================================
// CHUNK 14: DEDUCTION CONFLICT DETECTOR
// ============================================================

/**
 * Detect deduction conflicts and data integrity issues
 * @returns {Object} Conflicts found with details and suggested fixes
 */
function detectDeductionConflicts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    const conflicts = {
      duplicates: [],      // Same employee + item + date
      orphans: [],         // Data integrity issues
      overcharges: [],     // Amount paid > total
      zombieOrders: [],    // Status/balance mismatches
      totalConflicts: 0
    };
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return conflicts;
    }
    
    // Get all orders data
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 13).getValues();
    
    // ========== 1. DUPLICATE DETECTION ==========
    // Group by Employee + Order Date to find potential duplicates
    const orderGroups = {};
    
    for (const row of ordersData) {
      const orderId = row[0];
      const employeeName = row[2];
      let orderDate = '';
      if (row[4]) {
        try {
          const d = new Date(row[4]);
          if (!isNaN(d.getTime())) {
            orderDate = d.toDateString();
          }
        } catch (e) {
          // Skip invalid date
        }
      }
      const status = row[12];
      const totalAmount = parseFloat(row[5]) || 0;
      
      // Only consider Active or Pending orders for duplicates
      if (status !== 'Active' && status !== 'Pending' && status !== 'Pending - Cash') continue;
      
      const key = `${employeeName}|${orderDate}`;
      if (!orderGroups[key]) {
        orderGroups[key] = [];
      }
      orderGroups[key].push({ orderId, employeeName, orderDate, totalAmount, status });
    }
    
    // Find groups with multiple orders (potential duplicates)
    for (const key in orderGroups) {
      const group = orderGroups[key];
      if (group.length > 1) {
        // Check if they have similar totals (within 10% or same items)
        const totalSum = group.reduce((sum, o) => sum + o.totalAmount, 0);
        const avgTotal = totalSum / group.length;
        
        // Flag as potential duplicate
        conflicts.duplicates.push({
          type: 'duplicate',
          employeeName: group[0].employeeName,
          orderDate: group[0].orderDate,
          orderIds: group.map(o => o.orderId),
          description: `${group[0].employeeName} has ${group.length} orders on ${group[0].orderDate}: ${group.map(o => o.orderId).join(', ')}`
        });
      }
    }
    
    // ========== 2. ZOMBIE ORDER DETECTION ==========
    // Status = "Completed" but Amount_Remaining > 0
    // OR Status = "Active" but Amount_Remaining = 0
    
    for (const row of ordersData) {
      const orderId = row[0];
      const employeeName = row[2];
      const totalAmount = parseFloat(row[5]) || 0;
      const amountPaid = parseFloat(row[10]) || 0;
      const amountRemaining = parseFloat(row[11]) || 0;
      const status = row[12];
      
      // Zombie Type 1: Completed but still has remaining balance
      if (status === 'Completed' && amountRemaining > 0.01) {
        conflicts.zombieOrders.push({
          type: 'zombie_completed_with_balance',
          orderId,
          employeeName,
          status,
          amountRemaining,
          description: `${orderId} (${employeeName}): Status "Completed" but $${amountRemaining.toFixed(2)} still remaining`
        });
      }
      
      // Zombie Type 2: Active but nothing remaining
      if (status === 'Active' && amountRemaining <= 0) {
        conflicts.zombieOrders.push({
          type: 'zombie_active_no_balance',
          orderId,
          employeeName,
          status,
          amountPaid,
          totalAmount,
          description: `${orderId} (${employeeName}): Status "Active" but fully paid ($${amountPaid.toFixed(2)}/$${totalAmount.toFixed(2)})`
        });
      }
    }
    
    // ========== 3. OVER-CHARGE DETECTION ==========
    // Amount_Paid > Total_Amount
    
    for (const row of ordersData) {
      const orderId = row[0];
      const employeeName = row[2];
      const totalAmount = parseFloat(row[5]) || 0;
      const amountPaid = parseFloat(row[10]) || 0;
      
      if (totalAmount > 0 && amountPaid > totalAmount + 0.01) {
        const overchargeAmount = amountPaid - totalAmount;
        conflicts.overcharges.push({
          type: 'overcharge',
          orderId,
          employeeName,
          totalAmount,
          amountPaid,
          overchargeAmount,
          description: `${orderId} (${employeeName}): Over-charged by $${overchargeAmount.toFixed(2)} (Paid $${amountPaid.toFixed(2)} on $${totalAmount.toFixed(2)} order)`
        });
      }
    }
    
    // ========== 4. ORPHAN DETECTION ==========
    // Check for items without valid orders (if items sheet exists)
    
    if (itemsSheet && itemsSheet.getLastRow() > 1) {
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 3).getValues();
      const validOrderIds = new Set(ordersData.map(row => row[0]).filter(id => id));
      
      const orphanedItems = [];
      for (const row of itemsData) {
        const lineId = row[0];
        const orderId = row[1];
        
        if (orderId && !validOrderIds.has(orderId)) {
          orphanedItems.push(orderId);
        }
      }
      
      // Get unique orphaned order IDs
      const uniqueOrphanOrderIds = [...new Set(orphanedItems)];
      if (uniqueOrphanOrderIds.length > 0) {
        conflicts.orphans.push({
          type: 'orphaned_items',
          count: uniqueOrphanOrderIds.length,
          description: `${uniqueOrphanOrderIds.length} item(s) reference non-existent orders: ${uniqueOrphanOrderIds.slice(0, 5).join(', ')}${uniqueOrphanOrderIds.length > 5 ? '...' : ''}`
        });
      }
    }
    
    // Calculate total conflicts
    conflicts.totalConflicts = 
      conflicts.duplicates.length + 
      conflicts.zombieOrders.length + 
      conflicts.overcharges.length + 
      conflicts.orphans.length;
    
    // Log if conflicts found
    if (conflicts.totalConflicts > 0) {
      console.log(`Deduction conflict check found ${conflicts.totalConflicts} issue(s)`);
      logActivity('CONFLICT_CHECK', 'SYSTEM', 
        `Found ${conflicts.totalConflicts} deduction conflict(s): ${conflicts.duplicates.length} duplicates, ${conflicts.zombieOrders.length} zombies, ${conflicts.overcharges.length} overcharges, ${conflicts.orphans.length} orphans`,
        null
      );
    }
    
    return conflicts;
    
  } catch (error) {
    console.error('Error detecting deduction conflicts:', error);
    return {
      duplicates: [],
      orphans: [],
      overcharges: [],
      zombieOrders: [],
      totalConflicts: 0,
      error: error.message
    };
  }
}

/**
 * Run conflict detection and return formatted results
 * Can be called standalone or as part of health check
 * @returns {Object} Formatted conflict report
 */
function runConflictCheck() {
  const conflicts = detectDeductionConflicts();
  
  const report = {
    timestamp: new Date().toISOString(),
    hasConflicts: conflicts.totalConflicts > 0,
    summary: {
      total: conflicts.totalConflicts,
      duplicates: conflicts.duplicates.length,
      zombieOrders: conflicts.zombieOrders.length,
      overcharges: conflicts.overcharges.length,
      orphans: conflicts.orphans.length
    },
    details: conflicts,
    recommendations: []
  };
  
  // Add recommendations
  if (conflicts.duplicates.length > 0) {
    report.recommendations.push({
      priority: 'high',
      issue: 'Duplicate Deductions',
      action: 'Review orders placed on the same date by the same employee. Cancel any duplicate orders.',
      affectedOrders: conflicts.duplicates.flatMap(d => d.orderIds)
    });
  }
  
  if (conflicts.zombieOrders.length > 0) {
    report.recommendations.push({
      priority: 'high',
      issue: 'Zombie Orders',
      action: 'Update order status to match payment status. Mark fully-paid orders as "Completed".',
      affectedOrders: conflicts.zombieOrders.map(z => z.orderId)
    });
  }
  
  if (conflicts.overcharges.length > 0) {
    report.recommendations.push({
      priority: 'critical',
      issue: 'Over-Charged Orders',
      action: 'Review payment history immediately. Process refunds for over-charged amounts.',
      affectedOrders: conflicts.overcharges.map(o => o.orderId)
    });
  }
  
  if (conflicts.orphans.length > 0) {
    report.recommendations.push({
      priority: 'low',
      issue: 'Orphaned Records',
      action: 'Clean up orphaned item records that reference deleted orders.',
      affectedOrders: []
    });
  }
  
  return report;
}

// ============================================================
// CHUNK 10: SMART DEFAULTS & HISTORICAL PATTERNS
// ============================================================

/**
 * Get employee's size history from past orders
 * @param {string} employeeId - Employee ID
 * @returns {Object} Size history by item type
 */
function getEmployeeSizeHistory(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    if (!ordersSheet || !itemsSheet) {
      return { success: true, sizes: {}, hasHistory: false };
    }
    
    // Get all orders for this employee
    const ordersData = ordersSheet.getLastRow() > 1 
      ? ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 5).getValues()
      : [];
    
    const employeeOrderIds = [];
    for (const row of ordersData) {
      if (String(row[1]) === String(employeeId)) {
        employeeOrderIds.push(row[0]); // Order_ID
      }
    }
    
    if (employeeOrderIds.length === 0) {
      return { success: true, sizes: {}, hasHistory: false };
    }
    
    // Get items from those orders
    const itemsData = itemsSheet.getLastRow() > 1
      ? itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 6).getValues()
      : [];
    
    // Build size history: { "Polo Shirt": "Medium", "Pants": "32x30" }
    const sizeHistory = {};
    const sizeCounts = {}; // Track most common size per item
    
    for (const row of itemsData) {
      const orderId = row[1];
      const itemName = row[3];
      const size = row[4];
      
      if (employeeOrderIds.includes(orderId) && itemName && size) {
        if (!sizeCounts[itemName]) {
          sizeCounts[itemName] = {};
        }
        sizeCounts[itemName][size] = (sizeCounts[itemName][size] || 0) + 1;
      }
    }
    
    // Get most common size for each item
    for (const [itemName, sizes] of Object.entries(sizeCounts)) {
      let maxCount = 0;
      let mostCommon = null;
      for (const [size, count] of Object.entries(sizes)) {
        if (count > maxCount) {
          maxCount = count;
          mostCommon = size;
        }
      }
      if (mostCommon) {
        sizeHistory[itemName] = mostCommon;
      }
    }
    
    return {
      success: true,
      sizes: sizeHistory,
      hasHistory: Object.keys(sizeHistory).length > 0,
      orderCount: employeeOrderIds.length
    };
    
  } catch (error) {
    console.error('Error getting size history:', error);
    return { success: false, sizes: {}, hasHistory: false, error: error.message };
  }
}

/**
 * Get average order value for a location
 * @param {string} location - Location name (optional, if blank gets overall average)
 * @returns {Object} Average order statistics
 */
function getLocationOrderAverage(location) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return { success: true, average: 0, count: 0, hasData: false };
    }
    
    const ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 6).getValues();
    
    let totalAmount = 0;
    let orderCount = 0;
    
    for (const row of ordersData) {
      const orderLocation = row[3];
      const amount = parseFloat(row[5]) || 0;
      
      // Filter by location if specified
      if (!location || orderLocation === location) {
        if (amount > 0) {
          totalAmount += amount;
          orderCount++;
        }
      }
    }
    
    const average = orderCount > 0 ? totalAmount / orderCount : 0;
    
    return {
      success: true,
      average: Math.round(average * 100) / 100,
      count: orderCount,
      hasData: orderCount > 0,
      location: location || 'All Locations'
    };
    
  } catch (error) {
    console.error('Error getting order average:', error);
    return { success: false, average: 0, count: 0, hasData: false, error: error.message };
  }
}

// ============================================================
// CHUNK 20: MONTHLY AUTO-BACKUP SYSTEM
// ============================================================

/**
 * Main backup function - creates complete system backup
 * @param {boolean} isManual - True if triggered manually (affects notification text)
 * @returns {Object} Backup result with status and details
 */
function monthlyBackup(isManual = false) {
  try {
    console.log('Starting monthly backup...');
    const startTime = new Date();
    
    // Create or get the main backup folder
    const mainFolder = getOrCreateBackupFolder();
    
    // Create a dated subfolder for this backup
    const dateStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM');
    const timestamp = Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
    const backupFolderName = 'Backup_' + dateStr + '_' + timestamp;
    const backupFolder = mainFolder.createFolder(backupFolderName);
    
    // Sheets to backup
    const sheetsToBackup = [
      'Uniform_Orders',
      'Uniform_Order_Items',
      'Uniform_Catalog',
      'Employees',
      'OT_History',
      'PTO',
      'Activity_Log',
      'Settings',
      'Payroll_Settings',
      'System_Counters'
    ];
    
    const backupResults = [];
    let totalRows = 0;
    
    // Export each sheet to CSV
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    for (const sheetName of sheetsToBackup) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 0) {
        const result = exportSheetToCSV(sheet, backupFolder);
        backupResults.push({
          sheet: sheetName,
          success: result.success,
          rows: result.rows || 0
        });
        totalRows += result.rows || 0;
      } else {
        backupResults.push({
          sheet: sheetName,
          success: false,
          rows: 0,
          reason: sheet ? 'Empty sheet' : 'Sheet not found'
        });
      }
    }
    
    // Generate manifest file
    const manifest = generateBackupManifest(backupResults, startTime, ss);
    backupFolder.createFile('Backup_Manifest.txt', manifest);
    
    // Generate readme file
    const readme = generateBackupReadme();
    backupFolder.createFile('README.txt', readme);
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    const successCount = backupResults.filter(r => r.success).length;
    
    // Log the activity
    logActivity('BACKUP', 'SYSTEM', 
      `${isManual ? 'Manual' : 'Scheduled'} backup completed: ${successCount}/${sheetsToBackup.length} sheets, ${totalRows} rows`,
      backupFolderName
    );
    
    // Store last backup info in settings
    updateBackupStatus(startTime, successCount, sheetsToBackup.length, backupFolder.getUrl());
    
    // Send notification email
    sendBackupNotification(true, {
      folderName: backupFolderName,
      folderUrl: backupFolder.getUrl(),
      filesCount: successCount,
      totalSheets: sheetsToBackup.length,
      rowsCount: totalRows,
      duration: duration,
      isManual: isManual,
      results: backupResults
    });
    
    return {
      success: true,
      message: `Backup completed: ${successCount} files backed up`,
      folderUrl: backupFolder.getUrl(),
      folderName: backupFolderName,
      filesCount: successCount,
      totalRows: totalRows,
      duration: duration
    };
    
  } catch (error) {
    console.error('Backup failed:', error);
    
    // Log the failure
    logActivity('BACKUP_FAILED', 'SYSTEM', `Backup failed: ${error.message}`, null);
    
    // Send failure notification
    sendBackupNotification(false, { error: error.message });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get or create the main Payroll_System_Backups folder in Drive
 */
function getOrCreateBackupFolder() {
  const folderName = 'Payroll_System_Backups';
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // Create the folder
  const folder = DriveApp.createFolder(folderName);
  folder.setDescription('Automated backups from the Payroll System. Do not delete.');
  return folder;
}

/**
 * Export a sheet to CSV and save to Drive folder
 */
function exportSheetToCSV(sheet, folder) {
  try {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      return { success: false, rows: 0 };
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    // Convert to CSV
    const csv = data.map(row => 
      row.map(cell => {
        // Handle different data types
        if (cell === null || cell === undefined) {
          return '';
        }
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
        // Escape quotes and wrap in quotes if contains comma, newline, or quote
        const str = String(cell);
        if (str.includes(',') || str.includes('\n') || str.includes('"')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');
    
    // Create the file
    folder.createFile(sheetName + '.csv', csv, MimeType.CSV);
    
    return { success: true, rows: lastRow };
    
  } catch (error) {
    console.error('Error exporting sheet ' + sheet.getName() + ':', error);
    return { success: false, rows: 0, error: error.message };
  }
}

/**
 * Generate the backup manifest file content
 */
function generateBackupManifest(results, timestamp, spreadsheet) {
  const successCount = results.filter(r => r.success).length;
  const totalRows = results.reduce((sum, r) => sum + (r.rows || 0), 0);
  
  // Get some stats
  const ordersSheet = spreadsheet.getSheetByName('Uniform_Orders');
  const empSheet = spreadsheet.getSheetByName('Employees');
  const totalOrders = ordersSheet ? Math.max(0, ordersSheet.getLastRow() - 1) : 0;
  const totalEmployees = empSheet ? Math.max(0, empSheet.getLastRow() - 1) : 0;
  
  let manifest = '========================================\n';
  manifest += '       PAYROLL SYSTEM BACKUP MANIFEST\n';
  manifest += '========================================\n\n';
  manifest += 'Backup Date: ' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') + '\n';
  manifest += 'Timezone: ' + Session.getScriptTimeZone() + '\n';
  manifest += 'System Version: 1.0\n\n';
  manifest += '--- STATISTICS ---\n';
  manifest += 'Files Included: ' + successCount + '/' + results.length + '\n';
  manifest += 'Total Rows: ' + totalRows + '\n';
  manifest += 'Total Orders: ' + totalOrders + '\n';
  manifest += 'Total Employees: ' + totalEmployees + '\n';
  manifest += 'Backup Status: ' + (successCount === results.length ? 'Complete' : 'Partial') + '\n\n';
  manifest += '--- FILES ---\n';
  
  for (const result of results) {
    const status = result.success ? '✓' : '✗';
    const rows = result.success ? ` (${result.rows} rows)` : ` - ${result.reason || 'Failed'}`;
    manifest += `${status} ${result.sheet}.csv${rows}\n`;
  }
  
  manifest += '\n--- END MANIFEST ---\n';
  
  return manifest;
}

/**
 * Generate the backup readme file content
 */
function generateBackupReadme() {
  let readme = '========================================\n';
  readme += '    HOW TO RESTORE FROM THIS BACKUP\n';
  readme += '========================================\n\n';
  readme += 'This folder contains a complete backup of the Payroll System data.\n\n';
  readme += '--- RESTORATION STEPS ---\n\n';
  readme += '1. Open the Google Spreadsheet containing the Payroll System\n\n';
  readme += '2. For each CSV file in this backup folder:\n';
  readme += '   a. Open the corresponding sheet (e.g., Uniform_Orders)\n';
  readme += '   b. Select all data (Ctrl+A) and delete it\n';
  readme += '   c. Go to File > Import\n';
  readme += '   d. Upload the CSV file from this backup\n';
  readme += '   e. Choose "Replace data at selected cell"\n';
  readme += '   f. Click Import\n\n';
  readme += '3. Repeat for all sheets that need restoration\n\n';
  readme += '4. Verify the data looks correct\n\n';
  readme += '5. Test the system to ensure it works properly\n\n';
  readme += '--- FILE DESCRIPTIONS ---\n\n';
  readme += 'Uniform_Orders.csv      - All uniform orders (pending, active, completed)\n';
  readme += 'Uniform_Order_Items.csv - Individual line items for each order\n';
  readme += 'Uniform_Catalog.csv     - Available uniform items and prices\n';
  readme += 'Employees.csv           - Employee roster with IDs and locations\n';
  readme += 'OT_History.csv          - Historical overtime data by pay period\n';
  readme += 'PTO.csv                 - PTO requests and balances\n';
  readme += 'Activity_Log.csv        - System activity audit trail\n';
  readme += 'Settings.csv            - Application settings\n';
  readme += 'Payroll_Settings.csv    - Payroll configuration\n';
  readme += 'System_Counters.csv     - ID generation counters\n\n';
  readme += '--- CONTACT ---\n\n';
  readme += 'If you need help restoring data, contact Jeff (Operator).\n\n';
  readme += '--- END README ---\n';
  
  return readme;
}

/**
 * Send backup notification email
 */
function sendBackupNotification(success, details) {
  try {
    const settings = getSettings();
    const recipients = settings.weeklyEmailRecipients || Session.getActiveUser().getEmail();
    
    if (!recipients) {
      console.log('No email recipients configured for backup notifications');
      return;
    }
    
    let subject, body;
    
    if (success) {
      subject = '✅ Payroll System Backup Completed';
      body = '<html><body style="font-family: Arial, sans-serif; padding: 20px;">';
      body += '<h2 style="color: #10B981;">✅ Backup Completed Successfully</h2>';
      body += '<p>A ' + (details.isManual ? 'manual' : 'scheduled') + ' backup of the Payroll System has been completed.</p>';
      body += '<table style="border-collapse: collapse; margin: 20px 0;">';
      body += '<tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>Backup Folder:</strong></td><td style="padding: 8px; border-bottom: 1px solid #eee;">' + details.folderName + '</td></tr>';
      body += '<tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>Files Backed Up:</strong></td><td style="padding: 8px; border-bottom: 1px solid #eee;">' + details.filesCount + ' of ' + details.totalSheets + '</td></tr>';
      body += '<tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>Total Rows:</strong></td><td style="padding: 8px; border-bottom: 1px solid #eee;">' + details.rowsCount.toLocaleString() + '</td></tr>';
      body += '<tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>Duration:</strong></td><td style="padding: 8px; border-bottom: 1px solid #eee;">' + details.duration + ' seconds</td></tr>';
      body += '</table>';
      body += '<p><a href="' + details.folderUrl + '" style="background: #10B981; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">View Backup Folder</a></p>';
      body += '<hr style="margin: 30px 0; border: none; border-top: 1px solid #eee;">';
      body += '<p style="color: #666; font-size: 12px;">This is an automated message from the Payroll System.</p>';
      body += '</body></html>';
    } else {
      subject = '❌ Payroll System Backup FAILED';
      body = '<html><body style="font-family: Arial, sans-serif; padding: 20px;">';
      body += '<h2 style="color: #EF4444;">❌ Backup Failed</h2>';
      body += '<p>The scheduled backup of the Payroll System has failed.</p>';
      body += '<p><strong>Error:</strong> ' + (details.error || 'Unknown error') + '</p>';
      body += '<p>Please check the system and run a manual backup as soon as possible.</p>';
      body += '<hr style="margin: 30px 0; border: none; border-top: 1px solid #eee;">';
      body += '<p style="color: #666; font-size: 12px;">This is an automated message from the Payroll System.</p>';
      body += '</body></html>';
    }
    
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: body
    });
    
  } catch (error) {
    console.error('Error sending backup notification:', error);
  }
}

/**
 * Update the backup status in settings
 */
function updateBackupStatus(timestamp, filesCount, totalFiles, folderUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      return;
    }
    
    // Find or create backup status rows
    const data = settingsSheet.getDataRange().getValues();
    const keys = ['last_backup_date', 'last_backup_status', 'last_backup_url'];
    const values = [
      Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      filesCount === totalFiles ? 'Complete' : 'Partial (' + filesCount + '/' + totalFiles + ')',
      folderUrl
    ];
    
    for (let i = 0; i < keys.length; i++) {
      let rowIndex = -1;
      for (let j = 0; j < data.length; j++) {
        if (data[j][0] === keys[i]) {
          rowIndex = j + 1;
          break;
        }
      }
      
      if (rowIndex > 0) {
        settingsSheet.getRange(rowIndex, 2).setValue(values[i]);
      } else {
        settingsSheet.appendRow([keys[i], values[i]]);
      }
    }
    
  } catch (error) {
    console.error('Error updating backup status:', error);
  }
}

/**
 * Get the last backup status
 */
function getBackupStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet || settingsSheet.getLastRow() < 2) {
      return { hasBackup: false };
    }
    
    const data = settingsSheet.getDataRange().getValues();
    const result = { hasBackup: false };
    
    for (const row of data) {
      if (row[0] === 'last_backup_date') result.lastDate = row[1];
      if (row[0] === 'last_backup_status') result.status = row[1];
      if (row[0] === 'last_backup_url') result.folderUrl = row[1];
    }
    
    result.hasBackup = !!result.lastDate;
    
    // Check if auto-backup trigger exists
    result.autoBackupEnabled = getBackupTriggerStatus().active;
    
    return result;
    
  } catch (error) {
    console.error('Error getting backup status:', error);
    return { hasBackup: false, error: error.message };
  }
}

/**
 * Run a manual backup (called from UI)
 */
function runManualBackup() {
  return monthlyBackup(true);
}

/**
 * Set up the monthly backup trigger
 */
function setupBackupTrigger() {
  try {
    // Remove existing backup triggers first
    removeBackupTriggers();
    
    // Create a weekly trigger that runs on Sundays at 2am
    // The function will check if it's the first Sunday of the month
    ScriptApp.newTrigger('runScheduledBackup')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(2)
      .create();
    
    logActivity('BACKUP_TRIGGER', 'SYSTEM', 'Auto-backup trigger enabled (first Sunday of each month)', null);
    
    return { success: true, message: 'Auto-backup enabled' };
    
  } catch (error) {
    console.error('Error setting up backup trigger:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Scheduled backup function (called by trigger)
 * Only runs on the first Sunday of the month
 */
function runScheduledBackup() {
  const today = new Date();
  const dayOfMonth = today.getDate();
  
  // Only run on the first Sunday (day 1-7 of the month)
  if (dayOfMonth <= 7) {
    console.log('Running scheduled monthly backup...');
    monthlyBackup(false);
  } else {
    console.log('Skipping backup - not first Sunday of month (day ' + dayOfMonth + ')');
  }
}

/**
 * Remove existing backup triggers
 */
function removeBackupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runScheduledBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Get backup trigger status
 */
function getBackupTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'runScheduledBackup') {
      return { active: true };
    }
  }
  return { active: false };
}

/**
 * Toggle auto-backup on/off
 */
function toggleAutoBackup(enable) {
  if (enable) {
    return setupBackupTrigger();
  } else {
    removeBackupTriggers();
    logActivity('BACKUP_TRIGGER', 'SYSTEM', 'Auto-backup trigger disabled', null);
    return { success: true, message: 'Auto-backup disabled' };
  }
}

// ============================================================
// CHUNK 19: ROLLBACK/UNDO SYSTEM
// ============================================================

const ACTION_HISTORY_SHEET = 'Action_History';
const UNDO_WINDOW_HOURS = 12;
const MAX_UNDO_ACTIONS = 10;

/**
 * Get or create the Action_History sheet
 */
function getOrCreateActionHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ACTION_HISTORY_SHEET);
  
  if (!sheet) {
    sheet = ss.insertSheet(ACTION_HISTORY_SHEET);
    const headers = [
      'Action_ID',       // A: Unique identifier
      'Timestamp',       // B: When action occurred
      'User',            // C: Who performed action
      'Action_Type',     // D: ACTIVATE_ORDER, BULK_ACTIVATE
      'Description',     // E: Human-readable description
      'Affected_IDs',    // F: JSON array of affected order IDs
      'Before_State',    // G: JSON of state before action
      'After_State',     // H: JSON of state after action
      'Expires_At',      // I: When undo expires
      'Is_Undone',       // J: Boolean - has this been undone
      'Undone_At',       // K: When it was undone
      'Undone_By'        // L: Who undid it
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * Save action state for potential undo
 * @param {string} actionType - Type of action (ACTIVATE_ORDER, BULK_ACTIVATE)
 * @param {string} description - Human-readable description
 * @param {Array} affectedIds - Array of order IDs affected
 * @param {Object} beforeState - State before the action
 * @param {Object} afterState - State after the action
 * @returns {string} Action ID for undo reference
 */
function saveActionState(actionType, description, affectedIds, beforeState, afterState) {
  try {
    const sheet = getOrCreateActionHistorySheet();
    
    // Generate action ID
    const actionId = 'ACT-' + Date.now();
    const now = new Date();
    const expiresAt = new Date(now.getTime() + (UNDO_WINDOW_HOURS * 60 * 60 * 1000));
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    
    // Add new action
    const row = [
      actionId,
      now,
      user,
      actionType,
      description,
      JSON.stringify(affectedIds),
      JSON.stringify(beforeState),
      JSON.stringify(afterState),
      expiresAt,
      false,  // Is_Undone
      null,   // Undone_At
      null    // Undone_By
    ];
    
    sheet.appendRow(row);
    
    // Auto-purge: keep only last MAX_UNDO_ACTIONS
    const lastRow = sheet.getLastRow();
    if (lastRow > MAX_UNDO_ACTIONS + 1) { // +1 for header
      const rowsToDelete = lastRow - MAX_UNDO_ACTIONS - 1;
      sheet.deleteRows(2, rowsToDelete); // Delete oldest rows (after header)
    }
    
    console.log('Saved action state:', actionId, actionType);
    return actionId;
    
  } catch (error) {
    console.error('Error saving action state:', error);
    return null;
  }
}

/**
 * Get the order state for undo purposes
 * @param {string} orderId - The order ID
 * @returns {Object} Order state snapshot
 */
function getOrderStateForUndo(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
  const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
  
  // Find order row
  const ordersData = ordersSheet.getDataRange().getValues();
  const orderHeaders = ordersData[0];
  let orderRow = null;
  let orderRowNum = -1;
  
  for (let i = 1; i < ordersData.length; i++) {
    if (ordersData[i][0] === orderId) {
      orderRow = ordersData[i];
      orderRowNum = i + 1;
      break;
    }
  }
  
  if (!orderRow) return null;
  
  // Build order state object
  const orderState = {
    rowNum: orderRowNum,
    orderId: orderId,
    status: orderRow[12],                    // M: Status
    totalAmount: orderRow[5],                // F: Total_Amount
    amountPerPaycheck: orderRow[7],          // H: Amount_Per_Paycheck
    firstDeductionDate: orderRow[8],         // I: First_Deduction_Date
    paymentsMade: orderRow[9],               // J: Payments_Made
    amountPaid: orderRow[10],                // K: Amount_Paid
    amountRemaining: orderRow[11],           // L: Amount_Remaining
    receivedDate: orderRow[16],              // Q: Received_Date
    items: []
  };
  
  // Get items for this order
  const itemsData = itemsSheet.getDataRange().getValues();
  const itemHeaders = itemsData[0];
  
  for (let i = 1; i < itemsData.length; i++) {
    if (itemsData[i][1] === orderId) { // Order_ID column
      orderState.items.push({
        rowNum: i + 1,
        lineId: itemsData[i][0],
        itemReceived: itemsData[i][9],       // Item_Received
        itemReceivedDate: itemsData[i][10],  // Item_Received_Date
        itemStatus: itemsData[i][11]         // Item_Status
      });
    }
  }
  
  return orderState;
}

/**
 * Check if an action can be undone
 * @param {string} actionId - The action ID to check
 * @returns {Object} { canUndo: boolean, reason: string }
 */
function canUndoAction(actionId) {
  try {
    const sheet = getOrCreateActionHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === actionId) {
        const expiresAt = new Date(data[i][8]);
        const isUndone = data[i][9];
        
        if (isUndone) {
          return { canUndo: false, reason: 'This action has already been undone' };
        }
        
        if (new Date() > expiresAt) {
          return { canUndo: false, reason: 'Undo window has expired (12 hours)' };
        }
        
        return { canUndo: true, reason: null };
      }
    }
    
    return { canUndo: false, reason: 'Action not found' };
    
  } catch (error) {
    return { canUndo: false, reason: error.message };
  }
}

/**
 * Undo an order activation action
 * @param {string} actionId - The action ID to undo
 * @returns {Object} Result
 */
function undoAction(actionId) {
  try {
    // Check if can undo
    const canUndoResult = canUndoAction(actionId);
    if (!canUndoResult.canUndo) {
      return { success: false, error: canUndoResult.reason };
    }
    
    const sheet = getOrCreateActionHistorySheet();
    const data = sheet.getDataRange().getValues();
    
    let actionRow = null;
    let actionRowNum = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === actionId) {
        actionRow = data[i];
        actionRowNum = i + 1;
        break;
      }
    }
    
    if (!actionRow) {
      return { success: false, error: 'Action not found' };
    }
    
    const actionType = actionRow[3];
    const beforeState = JSON.parse(actionRow[6]);
    const affectedIds = JSON.parse(actionRow[5]);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDER_ITEMS);
    
    // Restore each order to its before state
    for (const orderState of beforeState.orders) {
      const rowNum = orderState.rowNum;
      
      // Restore order fields
      ordersSheet.getRange(rowNum, 6).setValue(orderState.totalAmount);         // F: Total_Amount
      ordersSheet.getRange(rowNum, 8).setValue(orderState.amountPerPaycheck);   // H: Amount_Per_Paycheck
      ordersSheet.getRange(rowNum, 9).setValue(orderState.firstDeductionDate);  // I: First_Deduction_Date
      ordersSheet.getRange(rowNum, 10).setValue(orderState.paymentsMade);       // J: Payments_Made
      ordersSheet.getRange(rowNum, 11).setValue(orderState.amountPaid);         // K: Amount_Paid
      ordersSheet.getRange(rowNum, 12).setValue(orderState.amountRemaining);    // L: Amount_Remaining
      ordersSheet.getRange(rowNum, 13).setValue(orderState.status);             // M: Status
      ordersSheet.getRange(rowNum, 17).setValue(orderState.receivedDate);       // Q: Received_Date
      
      // Restore item states
      for (const itemState of orderState.items) {
        itemsSheet.getRange(itemState.rowNum, 10).setValue(itemState.itemReceived);      // Item_Received
        itemsSheet.getRange(itemState.rowNum, 11).setValue(itemState.itemReceivedDate);  // Item_Received_Date
        itemsSheet.getRange(itemState.rowNum, 12).setValue(itemState.itemStatus);        // Item_Status
      }
    }
    
    // Mark action as undone
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    sheet.getRange(actionRowNum, 10).setValue(true);  // Is_Undone
    sheet.getRange(actionRowNum, 11).setValue(now);   // Undone_At
    sheet.getRange(actionRowNum, 12).setValue(user);  // Undone_By
    
    // Log the undo action
    const description = actionRow[4];
    logActivity('UNDO', 'UNIFORM', `Undone: ${description}`, actionId);
    
    return {
      success: true,
      message: `Successfully undone: ${description}`,
      restoredOrders: affectedIds.length
    };
    
  } catch (error) {
    console.error('Error undoing action:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get recent undoable actions for display
 * @returns {Array} List of recent actions with undo status
 */
function getRecentUndoableActions() {
  try {
    const sheet = getOrCreateActionHistorySheet();
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    const actions = [];
    for (let i = data.length - 1; i >= 1; i--) { // Reverse order (newest first)
      const row = data[i];
      const expiresAt = new Date(row[8]);
      const isUndone = row[9];
      
      actions.push({
        actionId: row[0],
        timestamp: row[1],
        user: row[2],
        actionType: row[3],
        description: row[4],
        affectedIds: JSON.parse(row[5] || '[]'),
        canUndo: !isUndone && now < expiresAt,
        isUndone: isUndone,
        expiresAt: expiresAt,
        timeRemaining: !isUndone && now < expiresAt ? Math.round((expiresAt - now) / 1000 / 60) : 0 // minutes
      });
    }
    
    return { success: true, actions: actions };
    
  } catch (error) {
    return { success: false, error: error.message, actions: [] };
  }
}

/**
 * Initialize Action_History sheet (run manually from script editor)
 */
function initializeActionHistory() {
  const sheet = getOrCreateActionHistorySheet();
  console.log('Action_History sheet initialized');
  return { success: true, message: 'Action_History sheet ready' };
}

// ============================================================
// CHUNK 17: YEAR-END CLOSING WIZARD
// ============================================================

/**
 * Check if Year-End Wizard should be shown (Jan 1 - Jan 31)
 */
function shouldShowYearEndWizard() {
  const now = new Date();
  const month = now.getMonth(); // 0-11
  
  // Show only during January (month 0)
  return month === 0;
}

/**
 * Get Year-End Wizard status and progress
 */
function getYearEndWizardStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    // Default status
    let status = {
      currentStep: 1,
      completedSteps: [],
      lastUpdated: null,
      year: new Date().getFullYear()
    };
    
    // Try to load saved progress from Settings
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'YearEndWizard_Progress') {
          try {
            status = JSON.parse(data[i][1]);
          } catch (e) {
            console.log('Could not parse wizard progress');
          }
          break;
        }
      }
    }
    
    return { success: true, status: status, showWizard: shouldShowYearEndWizard() };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Save Year-End Wizard progress
 */
function saveYearEndWizardProgress(progress) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      return { success: false, error: 'Settings sheet not found' };
    }
    
    progress.lastUpdated = new Date().toISOString();
    
    // Find or create the progress row
    const data = settingsSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'YearEndWizard_Progress') {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow > 0) {
      settingsSheet.getRange(foundRow, 2).setValue(JSON.stringify(progress));
    } else {
      settingsSheet.appendRow(['YearEndWizard_Progress', JSON.stringify(progress)]);
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Step 1: Get issues to review before year-end
 */
function getYearEndReviewIssues() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentYear = new Date().getFullYear();
    const issues = {
      pendingOrders: [],
      staleEmployees: [],
      inactiveWithDeductions: []
    };
    
    // 1. Find pending orders from current year
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const ordersData = ordersSheet.getDataRange().getValues();
      const headers = ordersData[0];
      const statusIdx = headers.indexOf('Status');
      const dateIdx = headers.indexOf('Order_Date');
      const nameIdx = headers.indexOf('Employee_Name');
      const amountIdx = headers.indexOf('Total_Amount');
      
      for (let i = 1; i < ordersData.length; i++) {
        const status = ordersData[i][statusIdx];
        const orderDate = new Date(ordersData[i][dateIdx]);
        
        if ((status === 'Pending' || status === 'Pending - Cash') && orderDate.getFullYear() === currentYear) {
          issues.pendingOrders.push({
            orderId: ordersData[i][0],
            employeeName: ordersData[i][nameIdx],
            amount: ordersData[i][amountIdx],
            orderDate: orderDate.toLocaleDateString(),
            status: status,
            rowIndex: i + 1
          });
        }
      }
    }
    
    // 2. Find employees not in recent OT data (3+ months)
    const otSheet = ss.getSheetByName('OT_History');
    const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
    
    if (otSheet && empSheet && otSheet.getLastRow() > 1 && empSheet.getLastRow() > 1) {
      const threeMonthsAgo = new Date();
      threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
      
      // Get all employee names from OT data in last 3 months
      const otData = otSheet.getDataRange().getValues();
      const otHeaders = otData[0];
      const otDateIdx = otHeaders.indexOf('Period_End') >= 0 ? otHeaders.indexOf('Period_End') : 0;
      const otNameIdx = otHeaders.indexOf('Employee_Name') >= 0 ? otHeaders.indexOf('Employee_Name') : 1;
      
      const recentOTEmployees = new Set();
      for (let i = 1; i < otData.length; i++) {
        const periodEnd = new Date(otData[i][otDateIdx]);
        if (periodEnd >= threeMonthsAgo) {
          const name = String(otData[i][otNameIdx]).toLowerCase().trim();
          if (name) recentOTEmployees.add(name);
        }
      }
      
      // Compare against active employees
      const empData = empSheet.getDataRange().getValues();
      const empHeaders = empData[0];
      const empNameIdx = empHeaders.indexOf('Display_Name') >= 0 ? empHeaders.indexOf('Display_Name') : 1;
      const empStatusIdx = empHeaders.indexOf('Status') >= 0 ? empHeaders.indexOf('Status') : -1;
      const empIdIdx = 0;
      
      for (let i = 1; i < empData.length; i++) {
        const empStatus = empStatusIdx >= 0 ? empData[i][empStatusIdx] : 'Active';
        if (empStatus === 'Active') {
          const empName = String(empData[i][empNameIdx]).toLowerCase().trim();
          if (empName && !recentOTEmployees.has(empName)) {
            issues.staleEmployees.push({
              employeeId: empData[i][empIdIdx],
              employeeName: empData[i][empNameIdx],
              rowIndex: i + 1
            });
          }
        }
      }
    }
    
    // 3. Find inactive employees with active deductions
    if (ordersSheet && empSheet && ordersSheet.getLastRow() > 1 && empSheet.getLastRow() > 1) {
      const empData = empSheet.getDataRange().getValues();
      const empHeaders = empData[0];
      const empNameIdx = empHeaders.indexOf('Display_Name') >= 0 ? empHeaders.indexOf('Display_Name') : 1;
      const empStatusIdx = empHeaders.indexOf('Status') >= 0 ? empHeaders.indexOf('Status') : -1;
      
      const inactiveEmployees = new Set();
      for (let i = 1; i < empData.length; i++) {
        const status = empStatusIdx >= 0 ? empData[i][empStatusIdx] : 'Active';
        if (status === 'Inactive') {
          inactiveEmployees.add(String(empData[i][empNameIdx]).toLowerCase().trim());
        }
      }
      
      const ordersData = ordersSheet.getDataRange().getValues();
      const headers = ordersData[0];
      const statusIdx = headers.indexOf('Status');
      const nameIdx = headers.indexOf('Employee_Name');
      
      for (let i = 1; i < ordersData.length; i++) {
        const status = ordersData[i][statusIdx];
        const empName = String(ordersData[i][nameIdx]).toLowerCase().trim();
        
        if (status === 'Active' && inactiveEmployees.has(empName)) {
          issues.inactiveWithDeductions.push({
            orderId: ordersData[i][0],
            employeeName: ordersData[i][nameIdx],
            amount: ordersData[i][headers.indexOf('Amount_Remaining')],
            rowIndex: i + 1
          });
        }
      }
    }
    
    return { success: true, issues: issues };
  } catch (error) {
    console.error('Error getting year-end review issues:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Step 2: Generate Annual Reports
 */
function generateAnnualReports(year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetYear = year || new Date().getFullYear();
    
    const reports = {
      uniformSummary: { totalOrders: 0, totalAmount: 0, byLocation: {}, byEmployee: [] },
      otSummary: { totalHours: 0, totalCost: 0, byLocation: {} },
      year: targetYear
    };
    
    // Uniform Orders Summary
    const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const data = ordersSheet.getDataRange().getValues();
      const headers = data[0];
      const dateIdx = headers.indexOf('Order_Date');
      const amountIdx = headers.indexOf('Total_Amount');
      const locationIdx = headers.indexOf('Location');
      const nameIdx = headers.indexOf('Employee_Name');
      const statusIdx = headers.indexOf('Status');
      
      const employeeSpending = {};
      
      for (let i = 1; i < data.length; i++) {
        const orderDate = new Date(data[i][dateIdx]);
        const status = data[i][statusIdx];
        
        if (orderDate.getFullYear() === targetYear && status !== 'Cancelled') {
          const amount = parseFloat(data[i][amountIdx]) || 0;
          const location = data[i][locationIdx] || 'Unknown';
          const empName = data[i][nameIdx] || 'Unknown';
          
          reports.uniformSummary.totalOrders++;
          reports.uniformSummary.totalAmount += amount;
          
          // By location
          if (!reports.uniformSummary.byLocation[location]) {
            reports.uniformSummary.byLocation[location] = { orders: 0, amount: 0 };
          }
          reports.uniformSummary.byLocation[location].orders++;
          reports.uniformSummary.byLocation[location].amount += amount;
          
          // By employee
          if (!employeeSpending[empName]) {
            employeeSpending[empName] = { name: empName, orders: 0, amount: 0 };
          }
          employeeSpending[empName].orders++;
          employeeSpending[empName].amount += amount;
        }
      }
      
      reports.uniformSummary.byEmployee = Object.values(employeeSpending).sort((a, b) => b.amount - a.amount);
    }
    
    // OT Summary
    // Use hardcoded indices that match DashboardModule.gs (known to work correctly)
    // Column A (0) = Period_End, Column D (3) = Location, Column M (12) = Total_OT, Column N (13) = OT_Cost
    const otSheet = ss.getSheetByName('OT_History');
    if (otSheet && otSheet.getLastRow() > 1) {
      const data = otSheet.getRange(2, 1, otSheet.getLastRow() - 1, 14).getValues(); // Skip header row
      
      for (let i = 0; i < data.length; i++) {
        const periodEnd = new Date(data[i][0]); // Column A = Period_End
        
        if (periodEnd.getFullYear() === targetYear) {
          const hours = parseFloat(data[i][12]) || 0;  // Column M = Total_OT (index 12)
          const cost = parseFloat(data[i][13]) || 0;   // Column N = OT_Cost (index 13)
          const location = data[i][3] || 'Unknown';    // Column D = Location (index 3)
          
          reports.otSummary.totalHours += hours;
          reports.otSummary.totalCost += cost;
          
          if (!reports.otSummary.byLocation[location]) {
            reports.otSummary.byLocation[location] = { hours: 0, cost: 0 };
          }
          reports.otSummary.byLocation[location].hours += hours;
          reports.otSummary.byLocation[location].cost += cost;
        }
      }
    }
    
    return { success: true, reports: reports };
  } catch (error) {
    console.error('Error generating annual reports:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Export annual reports to Google Drive
 */
function exportAnnualReportsToDrive(year) {
  try {
    const targetYear = year || new Date().getFullYear();
    const reportsResult = generateAnnualReports(targetYear);
    
    if (!reportsResult.success) {
      return reportsResult;
    }
    
    const reports = reportsResult.reports;
    
    // Get or create Year-End folder
    const backupFolderName = 'Payroll_System_Backups';
    const yearEndFolderName = `Year_End_${targetYear}`;
    
    let parentFolder;
    const parentFolders = DriveApp.getFoldersByName(backupFolderName);
    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
    } else {
      parentFolder = DriveApp.createFolder(backupFolderName);
    }
    
    let yearEndFolder;
    const yearEndFolders = parentFolder.getFoldersByName(yearEndFolderName);
    if (yearEndFolders.hasNext()) {
      yearEndFolder = yearEndFolders.next();
    } else {
      yearEndFolder = parentFolder.createFolder(yearEndFolderName);
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const files = [];
    
    // Create Employee Spending Report CSV
    let employeeCSV = 'Employee Name,Orders,Total Spent\n';
    reports.uniformSummary.byEmployee.forEach(emp => {
      employeeCSV += `"${emp.name}",${emp.orders},$${emp.amount.toFixed(2)}\n`;
    });
    const empFile = yearEndFolder.createFile(`Employee_Spending_Report_${targetYear}_${timestamp}.csv`, employeeCSV, 'text/csv');
    files.push({ name: empFile.getName(), type: 'Employee Spending Report' });
    
    // Create Summary Report
    let summaryText = `YEAR-END SUMMARY REPORT - ${targetYear}\n`;
    summaryText += `Generated: ${new Date().toLocaleString()}\n\n`;
    summaryText += `=== UNIFORM ORDERS ===\n`;
    summaryText += `Total Orders: ${reports.uniformSummary.totalOrders}\n`;
    summaryText += `Total Amount: $${reports.uniformSummary.totalAmount.toFixed(2)}\n\n`;
    summaryText += `By Location:\n`;
    for (const [loc, data] of Object.entries(reports.uniformSummary.byLocation)) {
      summaryText += `  ${loc}: ${data.orders} orders, $${data.amount.toFixed(2)}\n`;
    }
    summaryText += `\n=== OVERTIME ===\n`;
    summaryText += `Total OT Hours: ${reports.otSummary.totalHours.toFixed(1)}\n`;
    summaryText += `Total OT Cost: $${reports.otSummary.totalCost.toFixed(2)}\n\n`;
    summaryText += `By Location:\n`;
    for (const [loc, data] of Object.entries(reports.otSummary.byLocation)) {
      summaryText += `  ${loc}: ${data.hours.toFixed(1)} hours, $${data.cost.toFixed(2)}\n`;
    }
    
    const summaryFile = yearEndFolder.createFile(`Year_End_Summary_${targetYear}_${timestamp}.txt`, summaryText, 'text/plain');
    files.push({ name: summaryFile.getName(), type: 'Year-End Summary' });
    
    return {
      success: true,
      folderName: yearEndFolderName,
      folderUrl: yearEndFolder.getUrl(),
      files: files
    };
  } catch (error) {
    console.error('Error exporting reports to Drive:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Step 3: Run Archive (wrapper around existing function)
 */
function runYearEndArchive() {
  try {
    // Use existing archive function
    const result = runAnnualArchive(false); // Not a dry run
    return result;
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Step 4: Create Year-End Backup
 */
function createYearEndBackup() {
  try {
    const year = new Date().getFullYear();
    
    // Get or create backup folder
    const backupFolderName = 'Payroll_System_Backups';
    const yearEndFolderName = `Year_End_${year}`;
    
    let parentFolder;
    const parentFolders = DriveApp.getFoldersByName(backupFolderName);
    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
    } else {
      parentFolder = DriveApp.createFolder(backupFolderName);
    }
    
    let yearEndFolder;
    const yearEndFolders = parentFolder.getFoldersByName(yearEndFolderName);
    if (yearEndFolders.hasNext()) {
      yearEndFolder = yearEndFolders.next();
    } else {
      yearEndFolder = parentFolder.createFolder(yearEndFolderName);
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const files = [];
    
    // Backup all sheets
    const sheets = ss.getSheets();
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (sheet.getLastRow() > 0) {
        const data = sheet.getDataRange().getValues();
        let csv = '';
        for (const row of data) {
          csv += row.map(cell => {
            const val = String(cell);
            if (val.includes(',') || val.includes('"') || val.includes('\n')) {
              return '"' + val.replace(/"/g, '""') + '"';
            }
            return val;
          }).join(',') + '\n';
        }
        
        const fileName = `${sheetName}_Backup_${timestamp}.csv`;
        yearEndFolder.createFile(fileName, csv, 'text/csv');
        files.push(fileName);
      }
    }
    
    // Create manifest
    let manifest = `YEAR-END BACKUP MANIFEST - ${year}\n`;
    manifest += `Created: ${new Date().toLocaleString()}\n`;
    manifest += `Total Files: ${files.length}\n\n`;
    manifest += `Files Included:\n`;
    files.forEach(f => manifest += `  - ${f}\n`);
    
    yearEndFolder.createFile(`Backup_Manifest_${timestamp}.txt`, manifest, 'text/plain');
    
    logActivity('YEAR_END_BACKUP', 'SYSTEM', `Year-end backup created with ${files.length} files`, null);
    
    return {
      success: true,
      folderName: yearEndFolderName,
      folderUrl: yearEndFolder.getUrl(),
      fileCount: files.length,
      files: files
    };
  } catch (error) {
    console.error('Error creating year-end backup:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Step 5: Send Year-End Summary Email
 */
function sendYearEndSummaryEmail(summaryData) {
  try {
    const settings = getSettings();
    const adminEmails = settings.adminEmails;
    
    if (!adminEmails || adminEmails.trim() === '') {
      return { success: false, error: 'No admin emails configured' };
    }
    
    const year = summaryData.year || new Date().getFullYear();
    const now = new Date();
    
    const html = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #E51636 0%, #C41230 100%); color: white; padding: 24px; text-align: center;">
          <h1 style="margin: 0;">🎉 ${year} Payroll Year Complete!</h1>
        </div>
        
        <div style="padding: 24px; background: #f9fafb;">
          <p>The year-end closing process has been completed successfully.</p>
          
          <h3 style="color: #1F2937; border-bottom: 2px solid #E51636; padding-bottom: 8px;">📊 Uniform Orders Summary</h3>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb;">Total Orders</td><td style="text-align: right; font-weight: bold;">${summaryData.uniformOrders || 0}</td></tr>
            <tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb;">Total Amount</td><td style="text-align: right; font-weight: bold;">$${(summaryData.uniformAmount || 0).toFixed(2)}</td></tr>
          </table>
          
          <h3 style="color: #1F2937; border-bottom: 2px solid #E51636; padding-bottom: 8px;">⏰ Overtime Summary</h3>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb;">Total OT Hours</td><td style="text-align: right; font-weight: bold;">${(summaryData.otHours || 0).toFixed(1)} hrs</td></tr>
            <tr><td style="padding: 8px; border-bottom: 1px solid #e5e7eb;">Total OT Cost</td><td style="text-align: right; font-weight: bold;">$${(summaryData.otCost || 0).toFixed(2)}</td></tr>
          </table>
          
          <h3 style="color: #1F2937; border-bottom: 2px solid #E51636; padding-bottom: 8px;">📁 Archive & Backup</h3>
          <ul style="color: #4B5563;">
            <li>Orders archived: ${summaryData.ordersArchived || 0}</li>
            <li>OT records archived: ${summaryData.otArchived || 0}</li>
            <li>Backup location: <a href="${summaryData.backupUrl || '#'}">${summaryData.backupFolder || 'Payroll_System_Backups'}</a></li>
          </ul>
          
          <div style="background: #DCFCE7; border-left: 4px solid #22C55E; padding: 16px; margin-top: 24px;">
            <strong>✅ System Ready for ${year + 1}</strong><br>
            All data has been archived and backed up. The system is ready for the new year.
          </div>
        </div>
        
        <div style="background: #1F2937; color: #9CA3AF; padding: 16px; text-align: center; font-size: 12px;">
          Generated on ${now.toLocaleString()}<br>
          Payroll Management System
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: adminEmails,
      subject: `🎉 ${year} Payroll Year-End Complete`,
      htmlBody: html
    });
    
    logActivity('YEAR_END_EMAIL', 'SYSTEM', `Year-end summary email sent for ${year}`, null);
    
    return { success: true, message: 'Year-end summary email sent' };
  } catch (error) {
    console.error('Error sending year-end email:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Mark employee as inactive
 */
function markEmployeeInactive(employeeId) {
  try {
    const result = setEmployeeActiveStatus(employeeId, false);
    if (result.success) {
      logActivity('EMPLOYEE_INACTIVE', 'EMPLOYEE', `Employee ${employeeId} marked inactive via Year-End Wizard`, employeeId);
    }
    return result;
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Mark multiple employees as inactive (bulk operation)
 */
function markEmployeesInactive(employeeIds) {
  try {
    if (!employeeIds || !Array.isArray(employeeIds) || employeeIds.length === 0) {
      return { success: false, error: 'No employees provided' };
    }
    
    let successCount = 0;
    const errors = [];
    
    for (const employeeId of employeeIds) {
      try {
        const result = setEmployeeActiveStatus(employeeId, false);
        if (result.success) {
          successCount++;
        } else {
          errors.push(`${employeeId}: ${result.error}`);
        }
      } catch (e) {
        errors.push(`${employeeId}: ${e.message}`);
      }
    }
    
    if (successCount > 0) {
      logActivity('BULK_EMPLOYEE_INACTIVE', 'EMPLOYEE', 
        `${successCount} employees marked inactive via Year-End Wizard`, 
        employeeIds.join(','));
    }
    
    return { 
      success: true, 
      count: successCount, 
      errors: errors.length > 0 ? errors : null 
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================================
// CHUNK 28: NEW USER ONBOARDING SYSTEM
// ============================================================

/**
 * Check if current user has completed onboarding
 * Returns user's onboarding status and preferences
 */
function checkOnboardingStatus() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    // Default status
    let status = {
      hasCompletedTour: false,
      newUserModeEnabled: false,
      tourCompletedDate: null,
      checklistProgress: []
    };
    
    if (!settingsSheet) {
      return { success: true, status: status, isNewUser: true };
    }
    
    // Look for user's onboarding record
    const data = settingsSheet.getDataRange().getValues();
    const onboardingKey = `Onboarding_${userEmail}`;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === onboardingKey) {
        try {
          status = JSON.parse(data[i][1]);
        } catch (e) {
          console.log('Could not parse onboarding status');
        }
        break;
      }
    }
    
    return { 
      success: true, 
      status: status, 
      isNewUser: !status.hasCompletedTour,
      userEmail: userEmail
    };
  } catch (error) {
    console.error('Error checking onboarding status:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Save onboarding progress/completion
 */
function saveOnboardingStatus(statusData) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      return { success: false, error: 'Settings sheet not found' };
    }
    
    const onboardingKey = `Onboarding_${userEmail}`;
    
    // Find or create the record
    const data = settingsSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === onboardingKey) {
        foundRow = i + 1;
        break;
      }
    }
    
    const statusJson = JSON.stringify(statusData);
    
    if (foundRow > 0) {
      settingsSheet.getRange(foundRow, 2).setValue(statusJson);
    } else {
      settingsSheet.appendRow([onboardingKey, statusJson]);
    }
    
    return { success: true };
  } catch (error) {
    console.error('Error saving onboarding status:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Mark the tour as complete
 */
function markTourComplete() {
  try {
    const result = checkOnboardingStatus();
    if (!result.success) return result;
    
    const status = result.status;
    status.hasCompletedTour = true;
    status.tourCompletedDate = new Date().toISOString();
    
    const saveResult = saveOnboardingStatus(status);
    if (!saveResult.success) return saveResult;
    
    logActivity('ONBOARDING_COMPLETE', 'SYSTEM', 'User completed onboarding tour', null);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Toggle New User Mode
 */
function setNewUserMode(enabled) {
  try {
    const result = checkOnboardingStatus();
    if (!result.success) return result;
    
    const status = result.status;
    status.newUserModeEnabled = enabled;
    
    return saveOnboardingStatus(status);
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Get onboarding checklist items with status
 */
function getOnboardingChecklist() {
  try {
    const result = checkOnboardingStatus();
    const progress = result.success ? (result.status.checklistProgress || []) : [];
    
    const checklist = [
      { id: 'tour', label: 'Complete the welcome tour', completed: result.status?.hasCompletedTour || false },
      { id: 'health', label: 'Review the System Health Dashboard', completed: progress.includes('health') },
      { id: 'uniform', label: 'View a pending uniform order', completed: progress.includes('uniform') },
      { id: 'calendar', label: 'Check the Payroll Calendar', completed: progress.includes('calendar') },
      { id: 'help', label: 'Visit the Help & Documentation page', completed: progress.includes('help') }
    ];
    
    return { success: true, checklist: checklist };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Mark a checklist item as complete
 */
function markChecklistItem(itemId) {
  try {
    const result = checkOnboardingStatus();
    if (!result.success) return result;
    
    const status = result.status;
    if (!status.checklistProgress) {
      status.checklistProgress = [];
    }
    
    if (!status.checklistProgress.includes(itemId)) {
      status.checklistProgress.push(itemId);
    }
    
    return saveOnboardingStatus(status);
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Reset onboarding for current user (for testing/replay)
 */
function resetOnboarding() {
  try {
    const status = {
      hasCompletedTour: false,
      newUserModeEnabled: false,
      tourCompletedDate: null,
      checklistProgress: []
    };
    
    return saveOnboardingStatus(status);
  } catch (error) {
    return { success: false, error: error.message };
  }
}

