/**
 * Reports Module - Handles payroll processing, reporting, and settings
 * Part of Chunk 7: Unified Reporting & Payroll Integration
 */

// =====================================================
// CONSTANTS
// =====================================================

const PAYROLL_SETTINGS_SHEET = 'Payroll_Settings';

// The payday reference date is the operator-editable `paydayReference` setting.
// All payday math goes through PaydayModule.gs (getPaydayReferenceDate_,
// generatePaydaySeries_, getPaydayForDate_, getPayPeriodForPayday_).

// =====================================================
// PAYROLL SETTINGS - INITIALIZATION
// =====================================================

/**
 * Initializes the Payroll_Settings tab if it doesn't exist
 * Creates default settings for bi-weekly payroll
 */
function initializePayrollSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PAYROLL_SETTINGS_SHEET);
  
  // Don't overwrite existing settings
  if (sheet) {
    console.log('Payroll_Settings tab already exists');
    return { success: true, message: 'Payroll_Settings already exists' };
  }
  
  // Create new sheet
  sheet = ss.insertSheet(PAYROLL_SETTINGS_SHEET);
  
  // Add headers
  const headers = ['Setting_Name', 'Setting_Value', 'Description', 'Last_Updated'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Calculate next Friday from today using reference date
  const nextPayday = calculateNextPayday();
  const today = new Date();
  
  // Default settings
  const settings = [
    ['payroll_frequency', '14', 'Days between paychecks', today],
    ['next_payroll_date', formatDateISO(nextPayday), 'Next Friday payroll date', today],
    ['payroll_day_of_week', 'Friday', 'Day of week employees get paid', today],
    ['last_payroll_processed', '', 'Most recent payroll date that was marked complete', today],
    ['default_ot_rate', '16.50', 'Default hourly rate for OT cost calculations', today],
    ['admin_access_passcode', '05894', 'Shared passcode for admin dashboard access', today]
  ];
  
  // Add settings
  sheet.getRange(2, 1, settings.length, 4).setValues(settings);
  
  // Format
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#E8EAED');
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 200);  // Setting_Name
  sheet.setColumnWidth(2, 150);  // Setting_Value
  sheet.setColumnWidth(3, 400);  // Description
  sheet.setColumnWidth(4, 150);  // Last_Updated
  
  console.log('Payroll_Settings tab created with next payroll date:', formatDateISO(nextPayday));
  
  return {
    success: true,
    message: 'Payroll_Settings tab created',
    nextPayrollDate: formatDateISO(nextPayday)
  };
}

/**
 * Calculates the next payday based on reference date (Friday, Nov 28, 2025)
 * Paydays are bi-weekly Fridays
 */
function calculateNextPayday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // First payday strictly after today, from the canonical series (PaydayModule.gs).
  const series = generatePaydaySeries_(1, 4);
  for (let i = 0; i < series.length; i++) {
    const candidate = new Date(series[i].getTime());
    candidate.setHours(0, 0, 0, 0);
    if (candidate > today) return series[i];
  }

  // Fallback (should never hit): two weeks past the most recent payday
  const fallback = generatePaydaySeries_(1, 0)[0];
  fallback.setDate(fallback.getDate() + 14);
  return fallback;
}

/**
 * Formats a date as ISO string (YYYY-MM-DD)
 * Returns empty string if date is invalid
 */
function formatDateISO(date) {
  if (!date) return '';
  try {
    const d = new Date(date);
    if (isNaN(d.getTime())) return '';
    // Use LOCAL date components, not UTC (toISOString uses UTC which can shift dates)
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  } catch (e) {
    console.log('formatDateISO error:', e);
    return '';
  }
}

// =====================================================
// PAYROLL SETTINGS - GET/UPDATE
// =====================================================

/**
 * Retrieves a specific setting value from Payroll_Settings
 * @param {string} settingName - Name of the setting to retrieve
 * @returns {any} The setting value (converted to appropriate type)
 */
function getPayrollSetting(settingName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PAYROLL_SETTINGS_SHEET);
    
    if (!sheet) {
      console.log('Payroll_Settings sheet not found, initializing...');
      initializePayrollSettings();
      return getPayrollSetting(settingName); // Retry after init
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === settingName) {
        const value = data[i][1];
        
        // Type conversion based on setting name
        if (settingName === 'payroll_frequency') {
          return parseInt(value) || 14;
        }
        if (settingName === 'next_payroll_date' || settingName === 'last_payroll_processed') {
          if (!value) return null;
          // Handle both Date objects and string values from spreadsheet
          if (value instanceof Date) {
            return value;
          }
          // If it's a string, parse it
          const parsed = new Date(value + 'T12:00:00');
          if (!isNaN(parsed.getTime())) {
            return parsed;
          }
          // Try parsing as-is
          const directParse = new Date(value);
          if (!isNaN(directParse.getTime())) {
            return directParse;
          }
          return null;
        }
        if (settingName === 'default_ot_rate') {
          return parseFloat(value) || 16.50;
        }
        
        return value;
      }
    }
    
    console.log('Setting not found:', settingName);
    return null;
    
  } catch (error) {
    console.error('Error getting payroll setting:', error);
    return null;
  }
}

/**
 * Updates a specific setting value in Payroll_Settings
 * @param {string} settingName - Name of the setting to update
 * @param {any} newValue - New value to set
 * @returns {Object} Success/error response
 */
function updatePayrollSetting(settingName, newValue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PAYROLL_SETTINGS_SHEET);
    
    if (!sheet) {
      return { success: false, error: 'Payroll_Settings sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === settingName) {
        const oldValue = data[i][1];
        
        // Convert value to string for storage
        let valueToStore = newValue;
        if (newValue instanceof Date) {
          valueToStore = formatDateISO(newValue);
        } else if (typeof newValue === 'number') {
          valueToStore = newValue.toString();
        }
        
        // Update Setting_Value (column B)
        sheet.getRange(i + 1, 2).setValue(valueToStore);
        
        // Update Last_Updated (column D)
        sheet.getRange(i + 1, 4).setValue(new Date());
        
        console.log(`Updated ${settingName}: ${oldValue} -> ${valueToStore}`);
        
        return {
          success: true,
          settingName: settingName,
          oldValue: oldValue,
          newValue: valueToStore
        };
      }
    }
    
    return { success: false, error: 'Setting not found: ' + settingName };
    
  } catch (error) {
    console.error('Error updating payroll setting:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// VALIDATION DISMISSAL PERSISTENCE
// =====================================================

/**
 * Dismisses a validation check by storing its ID and a signature of the
 * current details in Payroll_Settings.  The signature ensures the dismissal
 * auto-invalidates when the underlying issue changes (new items appear, etc.).
 *
 * @param {string} checkId - The validation check ID to dismiss
 * @param {string} detailsSignature - A stable fingerprint of the check's details
 * @returns {Object} success/error
 */
function dismissValidationCheckServer(checkId, detailsSignature) {
  try {
    var raw = getPayrollSetting('dismissed_validations');
    var dismissed = [];
    if (raw) {
      try { dismissed = JSON.parse(raw); } catch (e) { dismissed = []; }
    }

    dismissed = dismissed.filter(function(d) { return d.id !== checkId; });
    dismissed.push({
      id: checkId,
      sig: detailsSignature || '',
      ts: new Date().toISOString()
    });

    var json = JSON.stringify(dismissed);
    var result = updatePayrollSetting('dismissed_validations', json);

    // If the setting row doesn't exist yet, append it
    if (!result.success && result.error && result.error.indexOf('not found') !== -1) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(PAYROLL_SETTINGS_SHEET);
      if (sheet) {
        sheet.appendRow(['dismissed_validations', json, 'JSON array of dismissed validation check IDs', new Date()]);
      }
    }

    return { success: true };
  } catch (error) {
    console.error('Error dismissing validation check:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Retrieves stored validation dismissals from Payroll_Settings.
 * @returns {Array} Array of {id, sig, ts} objects
 */
function getDismissedValidationChecks() {
  try {
    var raw = getPayrollSetting('dismissed_validations');
    if (!raw) return [];
    return JSON.parse(raw);
  } catch (e) {
    return [];
  }
}

/**
 * Clears all validation dismissals (useful for a fresh start).
 * @returns {Object} success/error
 */
function clearDismissedValidationChecks() {
  return updatePayrollSetting('dismissed_validations', '[]');
}

// =====================================================
// PAYROLL DATE FUNCTIONS
// =====================================================

/**
 * Returns the next scheduled payroll date
 * @returns {Object} Object with date, dateString, and displayString
 */
function getNextPayrollDate() {
  try {
    let nextDate = getPayrollSetting('next_payroll_date');
    
    // If no date set or date is in the past, calculate next one
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Validate nextDate is a proper Date object
    if (!nextDate || !(nextDate instanceof Date) || isNaN(nextDate.getTime()) || nextDate < today) {
      console.log('Next payroll date is missing, invalid, or past, calculating...');
      nextDate = calculateNextPayday();
      
      // Update the setting
      updatePayrollSetting('next_payroll_date', nextDate);
    }
    
    return {
      date: nextDate,
      dateString: formatDateISO(nextDate),
      displayString: nextDate.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      })
    };
    
  } catch (error) {
    console.error('Error getting next payroll date:', error);
    // Fallback to calculation
    const fallback = calculateNextPayday();
    return {
      date: fallback,
      dateString: formatDateISO(fallback),
      displayString: fallback.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      })
    };
  }
}

/**
 * Returns payroll dates for dropdowns - includes historical and future dates
 * Uses the canonical payday series (PaydayModule.gs) as the calculation base
 * All paydays are bi-weekly Fridays
 * @param {number} futureCount - How many future dates to return (default 6)
 * @param {number} historyCount - How many past dates to return (default 26 = ~1 year)
 * @returns {Array} Array of date objects with display strings, sorted chronologically
 */
function getUpcomingPayrollDatesFromSettings(futureCount = 6, historyCount = 26) {
  try {
    const today = new Date();
    today.setHours(12, 0, 0, 0); // Use noon to avoid timezone edge cases

    const dates = [];

    // Canonical series (PaydayModule.gs) — already includes the current payday + history + future.
    generatePaydaySeries_(historyCount, futureCount).forEach(function(payDate) {
      if (isNaN(payDate.getTime())) return;
      addPaydayToList(dates, payDate, today);
    });

    // Sort by date (oldest first)
    dates.sort((a, b) => new Date(a.date + 'T12:00:00') - new Date(b.date + 'T12:00:00'));

    return dates;

  } catch (error) {
    console.error('Error getting upcoming payroll dates:', error);
    return [];
  }
}

/**
 * Helper function to add a payday to the dates list with formatted display strings
 */
function addPaydayToList(dates, payDate, today) {
  const dateStr = formatDateISO(payDate);
  
  // Check if this is the most recent payday (current period)
  const dayDiff = Math.floor((payDate - today) / (1000 * 60 * 60 * 24));
  const isCurrent = dayDiff >= -13 && dayDiff <= 0; // Within the past 2 weeks
  const isNext = dayDiff > 0 && dayDiff <= 14; // Within the next 2 weeks
  
  // English format with indicator
  let englishDisplay = payDate.toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
  
  if (isCurrent) {
    englishDisplay += ' (Current Period)';
  } else if (isNext) {
    englishDisplay += ' (Next Payday)';
  }
  
  // Spanish format
  const spanishDisplay = payDate.toLocaleDateString('es-ES', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
  const spanishCapitalized = spanishDisplay.charAt(0).toUpperCase() + spanishDisplay.slice(1);
  
  dates.push({
    date: dateStr,
    displayEnglish: englishDisplay,
    displaySpanish: spanishCapitalized,
    isCurrent: isCurrent,
    isNext: isNext
  });
}

/**
 * Advances the payroll date to the next period after marking complete
 * @returns {Object} Success/error response with old and new dates
 */
function advancePayrollDate() {
  try {
    const currentNext = getPayrollSetting('next_payroll_date');
    
    if (!currentNext) {
      return { success: false, error: 'No current payroll date found' };
    }
    
    // Calculate new next date (14 days later)
    const newNextDate = new Date(currentNext);
    newNextDate.setDate(newNextDate.getDate() + 14);
    
    // Update last_payroll_processed to current date
    updatePayrollSetting('last_payroll_processed', currentNext);
    
    // Update next_payroll_date to new date
    updatePayrollSetting('next_payroll_date', newNextDate);
    
    console.log(`Payroll date advanced: ${formatDateISO(currentNext)} -> ${formatDateISO(newNextDate)}`);
    
    return {
      success: true,
      previousPayrollDate: formatDateISO(currentNext),
      newPayrollDate: formatDateISO(newNextDate),
      message: 'Payroll date advanced by 14 days'
    };
    
  } catch (error) {
    console.error('Error advancing payroll date:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Calculates the pay period dates for a given payroll date
 * Pay period ends on Saturday (6 days before Friday payday)
 * Pay period is 14 days
 * 
 * @param {string|Date} payrollDate - The Friday payment date
 * @returns {Object} Object with periodStart and periodEnd dates
 */
function calculatePayPeriodDates(payrollDate) {
  // Canonical pay-period math lives in PaydayModule.gs (getPayPeriodForPayday_).
  const period = getPayPeriodForPayday_(payrollDate);
  return {
    periodStart: period.periodStart,
    periodStartString: formatDateISO(period.periodStart),
    periodEnd: period.periodEnd,
    periodEndString: formatDateISO(period.periodEnd)
  };
}

// =====================================================
// PAYROLL REPORT GENERATION
// =====================================================

/**
 * Generates the complete payroll processing report for a specific payroll date
 * This is the main function that gathers all data from OT, Uniforms, PTO, and Transfers
 * 
 * @param {string|Date} payrollDate - The Friday payment date (e.g., "2025-12-13")
 * @returns {Object} Complete report object with all four sections
 */
function generatePayrollProcessingReport(payrollDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Validate and parse payroll date
    let payday;
    if (typeof payrollDate === 'string') {
      payday = new Date(payrollDate + 'T12:00:00');
    } else {
      payday = new Date(payrollDate);
    }
    
    if (isNaN(payday.getTime())) {
      return { success: false, error: 'Invalid payroll date' };
    }
    
    // Calculate pay period dates
    const payPeriod = calculatePayPeriodDates(payday);
    
    console.log(`Generating report for payday: ${formatDateISO(payday)}`);
    console.log(`Pay period: ${payPeriod.periodStartString} to ${payPeriod.periodEndString}`);
    
    // STEP 1: Get Overtime Data
    const overtimeData = getOvertimeForPayroll(ss, payPeriod.periodEnd);
    
    // STEP 2: Get Uniform Deductions
    const uniformData = getUniformDeductionsForPayroll(ss, payday);
    
    // STEP 3: Get PTO Payouts
    const ptoData = getPTOForPayroll(ss, payday);
    
    // STEP 4: Get Time Transfers (multi-location employees)
    const transferData = getTimeTransfersForPayroll(ss, payPeriod.periodEnd);
    
    // Calculate unique employees with activity
    const employeeIds = new Set();
    overtimeData.employees.forEach(e => employeeIds.add(e.employeeId));
    uniformData.orders.forEach(o => {
      const normalizedName = normalizeEmployeeName(o.employeeName || '');
      const key = normalizedName || (o.employeeName || '').trim() || (o.employeeId || '').toString();
      if (key) employeeIds.add(key);
    });
    ptoData.requests.forEach(r => employeeIds.add(r.employeeId));
    transferData.employees.forEach(e => employeeIds.add(e.employeeId));
    
    // Build complete report object
    const report = {
      success: true,
      payrollDate: formatDateISO(payday),
      payrollDateDisplay: payday.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      }),
      payPeriodStart: payPeriod.periodStartString,
      payPeriodEnd: payPeriod.periodEndString,
      generatedDate: new Date().toISOString(),
      generatedBy: Session.getActiveUser().getEmail() || 'Unknown',
      
      overtime: overtimeData,
      uniformDeductions: uniformData,
      ptoPayout: ptoData,
      timeTransfers: transferData,
      
      summary: {
        employeesWithActivity: employeeIds.size,
        totalOTHours: overtimeData.totalOTHours,
        totalUniformDeductions: uniformData.totalDeductions,
        totalPTOHours: ptoData.totalHours
      }
    };
    
    console.log('Report generated successfully');
    return report;
    
  } catch (error) {
    console.error('Error generating payroll report:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets overtime data for the payroll report
 * Queries OT_History for the specific period end date
 */
function getOvertimeForPayroll(ss, periodEndDate) {
  try {
    const sheet = ss.getSheetByName('OT_History');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { employeeCount: 0, totalOTHours: 0, employees: [] };
    }
    
    const periodEndStr = formatDateISO(periodEndDate);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
    
    // Filter to this period and employees with OT > 0
    const employees = [];
    
    data.forEach(row => {
      const rowPeriodEnd = row[0] ? formatDateISO(new Date(row[0])) : null;
      
      if (rowPeriodEnd !== periodEndStr) return;
      
      const totalOT = parseFloat(row[12]) || 0; // Column M = Total OT
      if (totalOT <= 0) return; // Skip employees with no OT
      
      const employeeName = row[1] || '';
      const matchKey = row[2] || employeeName.toLowerCase();
      const location = row[3] || 'Unknown';
      const chHours = parseFloat(row[4]) || 0;
      const dbuHours = parseFloat(row[5]) || 0;
      const totalHours = parseFloat(row[6]) || 0;
      const regularHours = parseFloat(row[7]) || 0;
      const week1OT = parseFloat(row[10]) || 0;
      const week2OT = parseFloat(row[11]) || 0;
      const isMultiLocation = row[15] === true || row[15] === 'TRUE';
      
      employees.push({
        employeeId: matchKey,
        employeeName: employeeName,
        location: location,
        regularHours: regularHours,
        regularMinutes: Math.floor(regularHours * 60),
        otHours: totalOT,
        otMinutes: Math.floor(totalOT * 60),
        totalHours: totalHours,
        totalMinutes: Math.floor(totalHours * 60),
        week1OT: week1OT,
        week2OT: week2OT,
        chHours: chHours,
        dbuHours: dbuHours,
        isMultiLocation: isMultiLocation
      });
    });
    
    // Sort by location, then by name
    employees.sort((a, b) => {
      if (a.location !== b.location) return a.location.localeCompare(b.location);
      return a.employeeName.localeCompare(b.employeeName);
    });
    
    const totalOTHours = employees.reduce((sum, e) => sum + e.otHours, 0);
    
    return {
      employeeCount: employees.length,
      totalOTHours: Math.floor(totalOTHours * 100) / 100,
      employees: employees
    };
    
  } catch (error) {
    console.error('Error getting overtime for payroll:', error);
    return { employeeCount: 0, totalOTHours: 0, employees: [] };
  }
}

// =====================================================
// UNIFORM DEDUCTION CORE MATH (single source of truth)
// -----------------------------------------------------
// Every uniform-deduction calculation in the system (payroll report, deductions
// view, dashboard metrics) flows through these two helpers so the dollar amount
// and the "which payday does this check land on" decision are computed in exactly
// one place. Do not re-implement this math elsewhere.
// =====================================================

/**
 * Normalizes a stored First_Deduction_Date (Date or string) to a 'YYYY-MM-DD' string.
 * Spreadsheet date cells come back as a Date; we read UTC components because dates are
 * written either as local-noon (parseLocalDate_) or legacy midnight-UTC, and both resolve
 * to the correct calendar day under UTC components for US timezones.
 * @param {Date|string} value
 * @returns {string|null}
 */
function normalizeFirstDeductionStr_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    const year = value.getUTCFullYear();
    const month = String(value.getUTCMonth() + 1).padStart(2, '0');
    const day = String(value.getUTCDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  const parsed = new Date(String(value).indexOf('T') !== -1 ? value : value + 'T12:00:00');
  if (isNaN(parsed.getTime())) return null;
  return formatDateISO(parsed);
}

/**
 * Resolves an order's payment context. Honors Needs_Review (stored schedule preserved
 * until a manager resolves it); otherwise uses the live (non-cancelled) line-item total.
 * @param {Object} params - { firstDeductionRaw, paymentSchedule, checksCompleted,
 *                            storedTotal, storedPerCheck, liveTotal, needsReview }
 * @returns {Object|null} { firstDeductionStr, paymentSchedule, checksCompleted,
 *                          totalAmount, amountPerCheck } or null if no valid deduction date
 */
function buildUniformOrderContext_(params) {
  const firstDeductionStr = normalizeFirstDeductionStr_(params.firstDeductionRaw);
  if (!firstDeductionStr) return null;

  const paymentSchedule = parseInt(params.paymentSchedule) || 1;
  const checksCompleted = parseInt(params.checksCompleted) || 0;
  const storedTotal = parseFloat(params.storedTotal) || 0;
  const storedPerCheck = parseFloat(params.storedPerCheck) || 0;
  const liveTotal = (params.liveTotal !== undefined && params.liveTotal !== null)
    ? parseFloat(params.liveTotal) || 0
    : storedTotal;

  let totalAmount;
  let amountPerCheck;
  if (params.needsReview) {
    totalAmount = storedTotal;
    amountPerCheck = storedPerCheck > 0
      ? storedPerCheck
      : (storedTotal > 0 ? Math.round((storedTotal / paymentSchedule) * 100) / 100 : 0);
  } else {
    totalAmount = liveTotal;
    amountPerCheck = totalAmount > 0 ? Math.round((totalAmount / paymentSchedule) * 100) / 100 : 0;
  }

  return {
    firstDeductionStr: firstDeductionStr,
    paymentSchedule: paymentSchedule,
    checksCompleted: checksCompleted,
    totalAmount: totalAmount,
    amountPerCheck: amountPerCheck
  };
}

/**
 * Determines whether a given payday lands on one of an order's deduction dates, and if so
 * which check it is and how much to deduct (final check absorbs the rounding remainder).
 * @param {Object} ctx - output of buildUniformOrderContext_
 * @param {string} paydayStr - target payday 'YYYY-MM-DD'
 * @returns {Object} { lands, checkNumber, checksRemaining, isFinalPayment, deductionAmount,
 *                     alreadyRecorded, amountRemaining, deductionDateStrs }
 */
function computeUniformDeduction_(ctx, paydayStr) {
  const result = {
    lands: false, checkNumber: 0, checksRemaining: 0, isFinalPayment: false,
    deductionAmount: 0, alreadyRecorded: false, amountRemaining: 0, deductionDateStrs: []
  };
  if (!ctx || !ctx.firstDeductionStr) return result;

  // Build the full schedule of deduction date strings (local-noon math, no UTC drift).
  const baseDate = new Date(ctx.firstDeductionStr + 'T12:00:00');
  for (let i = 0; i < ctx.paymentSchedule; i++) {
    const d = new Date(baseDate);
    d.setDate(d.getDate() + (i * 14));
    result.deductionDateStrs.push(formatDateISO(d));
  }

  const checkIndex = result.deductionDateStrs.indexOf(paydayStr);
  if (checkIndex === -1) return result;

  result.lands = true;
  const checkNumber = checkIndex + 1;
  result.checkNumber = checkNumber;
  result.checksRemaining = ctx.paymentSchedule - checkNumber;
  result.isFinalPayment = checkNumber === ctx.paymentSchedule;
  result.alreadyRecorded = checkIndex < ctx.checksCompleted;

  // Use schedule position (not recorded count) so unrecorded prior checks don't inflate this one.
  let deductionAmount = ctx.amountPerCheck;
  if (result.isFinalPayment) {
    const paidSoFar = (checkNumber - 1) * ctx.amountPerCheck;
    deductionAmount = Math.round((ctx.totalAmount - paidSoFar) * 100) / 100;
  }
  result.deductionAmount = deductionAmount;

  const priorPayments = (checkNumber - 1) * ctx.amountPerCheck;
  result.amountRemaining = Math.max(0, Math.round((ctx.totalAmount - priorPayments - deductionAmount) * 100) / 100);

  return result;
}

/**
 * Gets uniform deductions for the payroll report
 * Calculates which orders have a deduction on this specific payday
 */
function getUniformDeductionsForPayroll(ss, payday) {
  try {
    const sheet = ss.getSheetByName('Uniform_Orders');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { employeeCount: 0, totalDeductions: 0, orders: [] };
    }
    
    // Accept either a Date or a 'YYYY-MM-DD' string. Parse strings at local noon so
    // formatDateISO doesn't shift the day backward in negative-UTC-offset timezones.
    const paydayDate = (payday instanceof Date)
      ? payday
      : new Date(String(payday).indexOf('T') !== -1 ? payday : payday + 'T12:00:00');
    const paydayStr = formatDateISO(paydayDate);

    // Read items sheet once — build both totals map and items-by-order map
    // Skip cancelled items (Item_Status === 'Cancelled') so they don't inflate totals
    const lineItemTotals = {};
    const itemsByOrderId = {};
    const cancelledByOrderId = {};
    const itemsSheet = ss.getSheetByName('Uniform_Order_Items');
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const iHeaders = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
      const iStatusCol = iHeaders.indexOf('Item_Status');
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();
      for (const row of itemsData) {
        const orderId = row[1];
        if (!orderId) continue;
        const isCancelled = iStatusCol >= 0 && row[iStatusCol] === 'Cancelled';
        const lineTotal = parseFloat(row[7]) || 0;
        if (isCancelled) {
          cancelledByOrderId[orderId] = (cancelledByOrderId[orderId] || 0) + lineTotal;
          continue;
        }
        lineItemTotals[orderId] = (lineItemTotals[orderId] || 0) + lineTotal;
        if (!itemsByOrderId[orderId]) itemsByOrderId[orderId] = [];
        itemsByOrderId[orderId].push({
          description: row[3] || '',
          size: row[4] || '',
          quantity: parseInt(row[5]) || 1,
          unitPrice: parseFloat(row[6]) || 0,
          lineTotal: lineTotal,
          isReplacement: row[8] === true
        });
      }
    }

    // Read orders using header-based column lookup (Needs_Review/Review_Reason are appended)
    const oHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const oNeedsCol = oHeaders.indexOf('Needs_Review');
    const oReasonCol = oHeaders.indexOf('Review_Reason');
    const oPerCheckCol = oHeaders.indexOf('Amount_Per_Paycheck');
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    const orders = [];
    let activeCount = 0;
    let noFirstDeductionCount = 0;
    
    data.forEach((row, index) => {
      const orderId = row[0];
      if (!orderId) return;

      const status = row[12]; // Status column
      if (status !== 'Active') return; // Only active orders
      activeCount++;

      const firstDeductionDate = row[8]; // First_Deduction_Date
      if (!firstDeductionDate) {
        noFirstDeductionCount++;
        return;
      }

      // Flagged orders (Needs_Review) keep their stored schedule until manager resolves —
      // per-check amount stays as-is so payments continue on the original plan.
      // Unflagged orders recompute from live (non-cancelled) line items.
      const needsReview = oNeedsCol >= 0 ? (row[oNeedsCol] === true || row[oNeedsCol] === 'TRUE') : false;
      const reviewReason = oReasonCol >= 0 ? (row[oReasonCol] || '') : '';

      // Single source of truth for the date + dollar math.
      const ctx = buildUniformOrderContext_({
        firstDeductionRaw: firstDeductionDate,
        paymentSchedule: row[6],
        checksCompleted: row[9],
        storedTotal: row[5],
        storedPerCheck: oPerCheckCol >= 0 ? row[oPerCheckCol] : 0,
        liveTotal: lineItemTotals[orderId],
        needsReview: needsReview
      });
      if (!ctx) return;

      const calc = computeUniformDeduction_(ctx, paydayStr);
      if (!calc.lands) return; // No deduction on this date

      const items = itemsByOrderId[orderId] || [];

      orders.push({
        employeeId: row[1],
        employeeName: row[2],
        location: row[3] || 'Unknown',
        orderId: orderId,
        orderDate: row[4] ? formatDateISO(new Date(row[4])) : '',
        totalOrderCost: ctx.totalAmount,
        paymentSchedule: ctx.paymentSchedule,
        amountPerCheck: ctx.amountPerCheck,
        checkNumber: calc.checkNumber,
        checksCompleted: ctx.checksCompleted,
        checksRemaining: calc.checksRemaining,
        deductionAmount: calc.alreadyRecorded ? 0 : calc.deductionAmount,
        amountRemaining: calc.amountRemaining,
        isFinalPayment: calc.isFinalPayment,
        alreadyRecorded: calc.alreadyRecorded,
        needsReview: needsReview,
        reviewReason: reviewReason,
        cancelledAmount: Math.round((cancelledByOrderId[orderId] || 0) * 100) / 100,
        items: items,
        rowIndex: index + 2
      });
    });
    
    // Sort by location, then by name
    orders.sort((a, b) => {
      if (a.location !== b.location) return a.location.localeCompare(b.location);
      return a.employeeName.localeCompare(b.employeeName);
    });
    
    const totalDeductions = orders.reduce((sum, o) => sum + (o.alreadyRecorded ? 0 : o.deductionAmount), 0);
    const uniqueEmployees = new Set(
      orders.filter(o => !o.alreadyRecorded).map(o => {
        const normalizedName = normalizeEmployeeName(o.employeeName || '');
        return normalizedName || (o.employeeName || '').trim() || (o.employeeId || '').toString();
      }).filter(Boolean)
    );
    
    // Group by employee for payroll processing view (keep per-order breakdown)
    const groupedMap = {};
    orders.forEach(order => {
      const normalizedName = normalizeEmployeeName(order.employeeName || '');
      const key = normalizedName || (order.employeeName || '').trim() || (order.employeeId || '').toString();
      if (!groupedMap[key]) {
        groupedMap[key] = {
          employeeKey: key,
          employeeId: order.employeeId,
          employeeName: order.employeeName,
          locationSet: new Set(),
          totalDeduction: 0,
          orders: []
        };
      }
      
      const group = groupedMap[key];
      group.locationSet.add(order.location || 'Unknown');
      group.totalDeduction += order.deductionAmount || 0;
      group.orders.push(order);
    });
    
    const groupedOrders = Object.values(groupedMap).map(group => {
      const locations = Array.from(group.locationSet);
      return {
        employeeKey: group.employeeKey,
        employeeId: group.employeeId,
        employeeName: group.employeeName,
        locations: locations,
        locationSummary: locations.length === 1 ? locations[0] : 'Multiple Locations',
        totalDeduction: Math.round(group.totalDeduction * 100) / 100,
        orderCount: group.orders.length,
        orders: group.orders
      };
    });
    
    // Sort by name for payroll view
    groupedOrders.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    
    return {
      employeeCount: groupedOrders.length,
      totalDeductions: Math.round(totalDeductions * 100) / 100,
      totalOrderCount: orders.length,
      orders: groupedOrders,
      rawOrders: orders
    };
    
  } catch (error) {
    console.error('Error getting uniform deductions for payroll:', error);
    return { employeeCount: 0, totalDeductions: 0, orders: [] };
  }
}

/**
 * Deduction reconciliation health check (on-demand, surfaced in System Health).
 * Scans every Active uniform order and flags where the stored Total_Amount,
 * Amount_Per_Paycheck, or Amount_Remaining have drifted from the live (non-cancelled)
 * line items — i.e. data corruption or manual sheet edits. Needs_Review orders are
 * reported as informational only (they intentionally hold their stored schedule).
 * @returns {Object} { success, checkedCount, discrepancyCount, discrepancies, summary }
 */
function runDeductionReconciliationCheck(token) {
  requireValidSession_(token);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Uniform_Orders');
    if (!ordersSheet || ordersSheet.getLastRow() < 2) {
      return { success: true, checkedCount: 0, discrepancyCount: 0, discrepancies: [], summary: 'No uniform orders to check.' };
    }

    // Live (non-cancelled) line-item totals per order
    const lineItemTotals = {};
    const itemsSheet = ss.getSheetByName('Uniform_Order_Items');
    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const iHeaders = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
      const iStatusCol = iHeaders.indexOf('Item_Status');
      const itemsData = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).getValues();
      for (const row of itemsData) {
        const orderId = row[1];
        if (!orderId) continue;
        if (iStatusCol >= 0 && row[iStatusCol] === 'Cancelled') continue;
        lineItemTotals[orderId] = (lineItemTotals[orderId] || 0) + (parseFloat(row[7]) || 0);
      }
    }

    const oHeaders = ordersSheet.getRange(1, 1, 1, ordersSheet.getLastColumn()).getValues()[0];
    const needsCol = oHeaders.indexOf('Needs_Review');
    const orderDiscountCol = oHeaders.indexOf('Order_Discount');
    const data = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, ordersSheet.getLastColumn()).getValues();

    const discrepancies = [];
    let checkedCount = 0;
    const round2 = v => Math.round(v * 100) / 100;

    data.forEach((row, index) => {
      const orderId = row[0];
      if (!orderId) return;
      if (row[12] !== 'Active') return; // Status column
      checkedCount++;

      const employeeName = row[2] || '';
      const storedTotal = round2(parseFloat(row[5]) || 0);
      const paymentPlan = parseInt(row[6]) || 1;
      const storedPerCheck = round2(parseFloat(row[7]) || 0);
      const paymentsMade = parseInt(row[9]) || 0;
      const amountPaid = round2(parseFloat(row[10]) || 0);
      const storedRemaining = round2(parseFloat(row[11]) || 0);
      const needsReview = needsCol >= 0 ? (row[needsCol] === true || row[needsCol] === 'TRUE') : false;
      const oDiscount = orderDiscountCol >= 0 ? (parseFloat(row[orderDiscountCol]) || 0) : 0;
      const rowNum = index + 2;

      const flag = (field, stored, expected, severity) => {
        discrepancies.push({ orderId, employeeName, rowIndex: rowNum, field, stored, expected, severity, needsReview });
      };

      // Payments-made sanity (applies to every order)
      if (paymentsMade < 0 || paymentsMade > paymentPlan) {
        flag('Payments_Made', paymentsMade, '0–' + paymentPlan, 'error');
      }

      // Needs_Review orders intentionally hold their stored schedule — informational only
      if (needsReview) {
        flag('Needs_Review', 'flagged', 'manager to resolve', 'info');
        return;
      }

      const liveTotal = lineItemTotals[orderId] !== undefined
        ? round2(Math.max(0, lineItemTotals[orderId] - oDiscount))
        : storedTotal;
      const expectedPerCheck = liveTotal > 0 ? round2(liveTotal / paymentPlan) : 0;
      const expectedRemaining = round2(liveTotal - amountPaid);

      if (Math.abs(liveTotal - storedTotal) >= 0.01) {
        flag('Total_Amount', storedTotal, liveTotal, 'warning');
      }
      if (Math.abs(expectedPerCheck - storedPerCheck) >= 0.01) {
        flag('Amount_Per_Paycheck', storedPerCheck, expectedPerCheck, 'warning');
      }
      if (Math.abs(expectedRemaining - storedRemaining) >= 0.01) {
        flag('Amount_Remaining', storedRemaining, expectedRemaining, 'warning');
      }
    });

    return {
      success: true,
      checkedCount: checkedCount,
      discrepancyCount: discrepancies.length,
      discrepancies: discrepancies,
      summary: discrepancies.length === 0
        ? `All ${checkedCount} active orders are consistent with their line items.`
        : `${discrepancies.length} discrepancy(ies) found across ${checkedCount} active orders.`
    };

  } catch (error) {
    console.error('Error in runDeductionReconciliationCheck:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets order items for a specific order
 */
function getOrderItems(ss, orderId) {
  try {
    const sheet = ss.getSheetByName('Uniform_Order_Items');
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    
    return data
      .filter(row => row[1] === orderId)
      .map(row => ({
        description: row[3] || '',
        size: row[4] || '',
        quantity: parseInt(row[5]) || 1,
        unitPrice: parseFloat(row[6]) || 0,
        lineTotal: parseFloat(row[7]) || 0,
        isReplacement: row[8] === true
      }));
      
  } catch (error) {
    console.error('Error getting order items:', error);
    return [];
  }
}

/**
 * Gets PTO payouts for the payroll report
 * Queries PTO where Payout_Period matches payday and not yet paid
 */
function getPTOForPayroll(ss, payday) {
  try {
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { employeeCount: 0, totalHours: 0, requests: [] };
    }
    
    const paydayStr = formatDateISO(payday);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    
    const requests = [];
    
    data.forEach((row, index) => {
      const ptoId = row[0];
      if (!ptoId) return;
      
      const payoutPeriod = row[7] ? formatDateISO(new Date(row[7])) : null;
      if (payoutPeriod !== paydayStr) return;
      
      const paidOut = row[10] === true || row[10] === 'TRUE';
      
      const status = row[11];
      if (status === 'Cancelled' || status === 'Denied') return;
      
      const startDate = row[5] ? new Date(row[5]) : null;
      const endDate = row[6] ? new Date(row[6]) : null;
      
      // Calculate duration
      let durationDays = 1;
      if (startDate && endDate) {
        durationDays = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
      }
      
      requests.push({
        ptoId: ptoId,
        employeeId: row[1],
        employeeName: row[2],
        location: row[3] || 'Unknown',
        hoursRequested: parseFloat(row[4]) || 0,
        ptoStartDate: startDate ? formatDateISO(startDate) : '',
        ptoEndDate: endDate ? formatDateISO(endDate) : '',
        durationDays: durationDays,
        submissionDate: row[9] ? new Date(row[9]).toISOString() : '',
        hotSchedulesConfirmed: row[8] === true || row[8] === 'TRUE',
        notes: row[12] || '',
        alreadyRecorded: paidOut,
        rowIndex: index + 2
      });
    });
    
    // Sort by location, then by name
    requests.sort((a, b) => {
      if (a.location !== b.location) return a.location.localeCompare(b.location);
      return a.employeeName.localeCompare(b.employeeName);
    });
    
    const totalHours = requests.filter(r => !r.alreadyRecorded).reduce((sum, r) => sum + r.hoursRequested, 0);
    const uniqueEmployees = new Set(requests.filter(r => !r.alreadyRecorded).map(r => r.employeeId));
    
    return {
      employeeCount: uniqueEmployees.size,
      totalHours: totalHours,
      requests: requests
    };
    
  } catch (error) {
    console.error('Error getting PTO for payroll:', error);
    return { employeeCount: 0, totalHours: 0, requests: [] };
  }
}

/**
 * Gets time transfers for multi-location employees
 * These are employees who worked at multiple locations
 * Uses configured location names from settings
 */
function getTimeTransfersForPayroll(ss, periodEndDate) {
  try {
    const sheet = ss.getSheetByName('OT_History');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { employeeCount: 0, employees: [] };
    }
    
    const settings = getSettings();
    const loc1Name = settings.location1Name || 'Location 1';
    const loc2Name = settings.location2Name || 'Location 2';
    
    const periodEndStr = formatDateISO(periodEndDate);
    const colCount = Math.min(sheet.getLastColumn(), 23);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();
    
    const floor2 = v => Math.floor(v * 100) / 100;
    
    const employees = [];
    
    data.forEach(row => {
      const rowPeriodEnd = row[0] ? formatDateISO(new Date(row[0])) : null;
      if (rowPeriodEnd !== periodEndStr) return;
      
      const isMultiLocation = row[15] === true || row[15] === 'TRUE';
      if (!isMultiLocation) return;
      
      const employeeName = row[1] || '';
      const matchKey = row[2] || employeeName.toLowerCase();
      
      const loc1Hours = parseFloat(row[4]) || 0;
      const loc2Hours = parseFloat(row[5]) || 0;
      const totalHours = parseFloat(row[6]) || 0;
      const week1Hours = parseFloat(row[8]) || 0;
      const week2Hours = parseFloat(row[9]) || 0;
      const week1OT = parseFloat(row[10]) || 0;
      const week2OT = parseFloat(row[11]) || 0;
      const totalOT = parseFloat(row[12]) || 0;
      
      // Per-week-per-location data (columns T-W, indices 19-22)
      const hasWeeklyBreakdown = colCount >= 23 && ((parseFloat(row[19]) || 0) > 0 || (parseFloat(row[20]) || 0) > 0 || (parseFloat(row[21]) || 0) > 0 || (parseFloat(row[22]) || 0) > 0);
      
      let w1Loc1, w1Loc2, w2Loc1, w2Loc2;
      
      if (hasWeeklyBreakdown) {
        w1Loc1 = parseFloat(row[19]) || 0;
        w1Loc2 = parseFloat(row[20]) || 0;
        w2Loc1 = parseFloat(row[21]) || 0;
        w2Loc2 = parseFloat(row[22]) || 0;
      } else {
        // Fallback: split period totals proportionally across weeks using combined week hours
        const totalCombined = loc1Hours + loc2Hours;
        const loc1Ratio = totalCombined > 0 ? loc1Hours / totalCombined : 0;
        const loc2Ratio = totalCombined > 0 ? loc2Hours / totalCombined : 0;
        w1Loc1 = floor2(week1Hours * loc1Ratio);
        w1Loc2 = floor2(week1Hours * loc2Ratio);
        w2Loc1 = floor2(week2Hours * loc1Ratio);
        w2Loc2 = floor2(week2Hours * loc2Ratio);
      }
      
      // Per-week proportional OT allocation
      const w1Total = w1Loc1 + w1Loc2;
      const w2Total = w2Loc1 + w2Loc2;
      const w1Loc1Ratio = w1Total > 0 ? w1Loc1 / w1Total : 0;
      const w1Loc2Ratio = w1Total > 0 ? w1Loc2 / w1Total : 0;
      const w2Loc1Ratio = w2Total > 0 ? w2Loc1 / w2Total : 0;
      const w2Loc2Ratio = w2Total > 0 ? w2Loc2 / w2Total : 0;
      
      const w1Loc1OT = floor2(week1OT * w1Loc1Ratio);
      const w1Loc2OT = floor2(week1OT * w1Loc2Ratio);
      const w2Loc1OT = floor2(week2OT * w2Loc1Ratio);
      const w2Loc2OT = floor2(week2OT * w2Loc2Ratio);
      
      const loc1OTTotal = w1Loc1OT + w2Loc1OT;
      const loc2OTTotal = w1Loc2OT + w2Loc2OT;
      
      const w1Loc1Reg = floor2(Math.max(0, w1Loc1 - w1Loc1OT));
      const w1Loc2Reg = floor2(Math.max(0, w1Loc2 - w1Loc2OT));
      const w2Loc1Reg = floor2(Math.max(0, w2Loc1 - w2Loc1OT));
      const w2Loc2Reg = floor2(Math.max(0, w2Loc2 - w2Loc2OT));
      
      const loc1RegTotal = w1Loc1Reg + w2Loc1Reg;
      const loc2RegTotal = w1Loc2Reg + w2Loc2Reg;
      
      // Transfer direction: paid from location with more hours
      let transferFrom, transferTo, transferHours, actualPaidFromLocation;
      if (loc1Hours >= loc2Hours) {
        actualPaidFromLocation = loc1Name;
        transferFrom = loc2Name;
        transferTo = loc1Name;
        transferHours = loc2Hours;
      } else {
        actualPaidFromLocation = loc2Name;
        transferFrom = loc1Name;
        transferTo = loc2Name;
        transferHours = loc1Hours;
      }
      
      const grandTotalHours = loc1Hours + loc2Hours;
      const grandTotalMinutes = Math.floor(grandTotalHours * 60);
      
      employees.push({
        employeeId: matchKey,
        employeeName: employeeName,
        homeLocation: actualPaidFromLocation,
        hasWeeklyBreakdown: hasWeeklyBreakdown,
        locations: [
          {
            name: loc1Name,
            regularHours: Math.max(0, loc1RegTotal),
            regularMinutes: Math.floor(Math.max(0, loc1RegTotal) * 60),
            otHours: Math.max(0, loc1OTTotal),
            otMinutes: Math.floor(Math.max(0, loc1OTTotal) * 60),
            totalHours: loc1Hours,
            totalMinutes: Math.floor(loc1Hours * 60),
            week1Hours: w1Loc1,
            week1Regular: w1Loc1Reg,
            week1OT: w1Loc1OT,
            week2Hours: w2Loc1,
            week2Regular: w2Loc1Reg,
            week2OT: w2Loc1OT
          },
          {
            name: loc2Name,
            regularHours: Math.max(0, loc2RegTotal),
            regularMinutes: Math.floor(Math.max(0, loc2RegTotal) * 60),
            otHours: Math.max(0, loc2OTTotal),
            otMinutes: Math.floor(Math.max(0, loc2OTTotal) * 60),
            totalHours: loc2Hours,
            totalMinutes: Math.floor(loc2Hours * 60),
            week1Hours: w1Loc2,
            week1Regular: w1Loc2Reg,
            week1OT: w1Loc2OT,
            week2Hours: w2Loc2,
            week2Regular: w2Loc2Reg,
            week2OT: w2Loc2OT
          }
        ],
        paidFromLocation: actualPaidFromLocation,
        grandTotalHours: grandTotalHours,
        grandTotalMinutes: grandTotalMinutes,
        week1Hours: week1Hours,
        week2Hours: week2Hours,
        week1OT: week1OT,
        week2OT: week2OT,
        totalOT: totalOT,
        transferFrom: transferFrom,
        transferTo: transferTo,
        transferHours: transferHours,
        transferMinutes: Math.floor(transferHours * 60)
      });
    });
    
    employees.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    
    return {
      employeeCount: employees.length,
      employees: employees
    };
    
  } catch (error) {
    console.error('Error getting time transfers for payroll:', error);
    return { employeeCount: 0, employees: [] };
  }
}

/**
 * On-demand reconciliation for any pay period — callable from the OT Reconciliation view
 */
function getReconciliationForPeriod(periodEndStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const periodEndDate = new Date(periodEndStr);
    if (isNaN(periodEndDate.getTime())) {
      return { success: false, error: 'Invalid period date' };
    }
    const result = getTimeTransfersForPayroll(ss, periodEndDate);
    return { success: true, data: result };
  } catch (error) {
    console.error('Error in getReconciliationForPeriod:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// PAYROLL COMPLETION FUNCTIONS
// =====================================================

/**
 * Marks payroll as complete - updates uniforms, PTO, and advances date
 * 
 * @param {string|Date} payrollDate - The payroll date being marked complete
 * @param {Array} skippedOrderIds - Order IDs to exclude from marking complete
 * @returns {Object} Summary of what was processed
 */
function markPayrollComplete(token, payrollDate, skippedOrderIds = []) {
  requireValidSession_(token);
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    return { success: false, error: 'Payroll is already being processed. Please wait a moment and refresh before trying again.' };
  }
  try {
    const payday = typeof payrollDate === 'string'
      ? new Date(payrollDate + 'T12:00:00') 
      : new Date(payrollDate);
    
    const paydayStr = formatDateISO(payday);
    
    // Validate this is the expected payroll date
    const nextPayroll = getPayrollSetting('next_payroll_date');
    const nextPayrollStr = nextPayroll ? formatDateISO(nextPayroll) : null;
    
    if (nextPayrollStr && paydayStr !== nextPayrollStr) {
      console.log(`Warning: Marking ${paydayStr} complete, but next scheduled is ${nextPayrollStr}`);
      // Don't block - allow marking any date complete
    }
    
    const results = {
      success: true,
      payrollDate: paydayStr,
      uniformPayments: { marked: 0, skipped: 0, totalAmount: 0, errors: [] },
      ptoPayouts: { marked: 0, totalHours: 0, errors: [] },
      newPayrollDate: null
    };
    
    // STEP 1: Mark uniform payments complete
    const uniformResult = markUniformPaymentsComplete(payday, skippedOrderIds);
    results.uniformPayments = uniformResult;
    
    // STEP 2: Mark PTO as paid
    const ptoResult = markPTOPaymentsComplete(payday);
    results.ptoPayouts = ptoResult;
    
    // STEP 3: Advance payroll date
    const advanceResult = advancePayrollDate();
    if (advanceResult.success) {
      results.newPayrollDate = advanceResult.newPayrollDate;
    }
    
    results.message = `Payroll marked complete. ${results.uniformPayments.marked} uniform payments, ${results.ptoPayouts.marked} PTO payouts. Next payroll: ${results.newPayrollDate}`;
    
    console.log(results.message);
    
    // Log the activity
    logActivity('COMPLETE', 'PAYROLL', 
      `Payroll completed for ${paydayStr}: ${results.uniformPayments.marked} uniform payments ($${results.uniformPayments.totalAmount}), ${results.ptoPayouts.marked} PTO payouts (${results.ptoPayouts.totalHours} hrs)`,
      paydayStr
    );
    
    return results;
    
  } catch (error) {
    console.error('Error marking payroll complete:', error);
    return { success: false, error: error.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Marks uniform payments complete for a specific payday
 */
function markUniformPaymentsComplete(payday, skippedOrderIds = []) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const report = getUniformDeductionsForPayroll(ss, payday);
    const sheet = ss.getSheetByName('Uniform_Orders');
    if (!sheet) return { marked: 0, skipped: 0, totalAmount: 0, errors: ['Orders sheet not found'] };

    const skippedSet = new Set(skippedOrderIds);
    let marked = 0;
    let skipped = 0;
    let totalAmount = 0;
    const errors = [];

    const ordersToProcess = Array.isArray(report.rawOrders) && report.rawOrders.length > 0
      ? report.rawOrders
      : report.orders.flatMap(entry => entry.orders || []).filter(o => o && o.orderId);

    if (sheet.getLastRow() < 2) {
      return { marked: 0, skipped: 0, totalAmount: 0, errors: null };
    }

    // Resolve columns by header so a future column insert/reorder can't make us
    // write payment data to the wrong cells. Fail safe if a required column is gone.
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const col = {};
    headers.forEach(function(h, i) { col[h] = i; });
    const required = ['Order_ID', 'Total_Amount', 'Payment_Plan', 'Amount_Per_Paycheck', 'Payments_Made', 'Amount_Paid', 'Amount_Remaining', 'Status'];
    const missing = required.filter(function(h) { return col[h] === undefined; });
    if (missing.length) {
      return { marked: 0, skipped: 0, totalAmount: 0, errors: ['Uniform_Orders missing columns: ' + missing.join(', ')] };
    }

    // Read orders sheet once for current payment state
    const ordersData = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    const orderRowMap = {};
    ordersData.forEach(function(row, idx) {
      if (row[col['Order_ID']]) {
        orderRowMap[row[col['Order_ID']]] = {
          rowIndex: idx + 2,
          paymentsMade: parseInt(row[col['Payments_Made']]) || 0,
          amountPaid: parseFloat(row[col['Amount_Paid']]) || 0,
          totalAmount: parseFloat(row[col['Total_Amount']]) || 0,
          paymentPlan: parseInt(row[col['Payment_Plan']]) || 1,
          amountPerCheck: parseFloat(row[col['Amount_Per_Paycheck']]) || 0,
          status: row[col['Status']]
        };
      }
    });

    ordersToProcess.forEach(order => {
      if (skippedSet.has(order.orderId) || order.alreadyRecorded) {
        skipped++;
        return;
      }

      try {
        const current = orderRowMap[order.orderId];
        if (!current) { errors.push(order.orderId + ': Order not found'); return; }
        if (current.status === 'Completed' || current.status === 'Cancelled' || current.status === 'Store Paid' || current.status === 'Pending') {
          errors.push(order.orderId + ': Status is "' + current.status + '"'); return;
        }
        if (current.paymentsMade >= current.paymentPlan) {
          errors.push(order.orderId + ': All payments already recorded'); return;
        }

        const newPaymentsMade = current.paymentsMade + 1;
        const newAmountPaid = parseFloat((current.amountPaid + current.amountPerCheck).toFixed(2));
        let newRemaining = parseFloat((current.totalAmount - newAmountPaid).toFixed(2));
        let newStatus = current.status;
        if (newPaymentsMade >= current.paymentPlan) { newStatus = 'Completed'; newRemaining = 0; }

        sheet.getRange(current.rowIndex, col['Payments_Made'] + 1).setValue(newPaymentsMade);
        sheet.getRange(current.rowIndex, col['Amount_Paid'] + 1).setValue(newAmountPaid);
        sheet.getRange(current.rowIndex, col['Amount_Remaining'] + 1).setValue(newRemaining);
        sheet.getRange(current.rowIndex, col['Status'] + 1).setValue(newStatus);

        marked++;
        totalAmount += order.deductionAmount;
      } catch (e) {
        errors.push(order.orderId + ': ' + e.message);
      }
    });

    return {
      marked: marked,
      skipped: skipped,
      totalAmount: Math.round(totalAmount * 100) / 100,
      errors: errors.length > 0 ? errors : null
    };

  } catch (error) {
    console.error('Error marking uniform payments complete:', error);
    return { marked: 0, skipped: 0, totalAmount: 0, errors: [error.message] };
  }
}

/**
 * Marks PTO payments complete for a specific payday
 */
function markPTOPaymentsComplete(payday) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const report = getPTOForPayroll(ss, payday);
    
    // Only process PTO that hasn't already been recorded
    const pendingRequests = report.requests.filter(r => !r.alreadyRecorded);
    
    if (pendingRequests.length === 0) {
      return { marked: 0, totalHours: 0, errors: null };
    }
    
    const ptoIds = pendingRequests.map(r => r.ptoId);
    const paydayStr = formatDateISO(payday);
    
    // Call existing batch function from PTOModule.gs
    const result = batchMarkPTOPaid_(ptoIds, paydayStr);
    
    return {
      marked: result.updatedCount || 0,
      totalHours: result.totalHours || 0,
      errors: result.errors
    };
    
  } catch (error) {
    console.error('Error marking PTO payments complete:', error);
    return { marked: 0, totalHours: 0, errors: [error.message] };
  }
}

/**
 * Skips a uniform payment, pushing remaining payments back by 14 days
 * 
 * @param {string} orderId - The order ID to skip
 * @param {string|Date} currentPayrollDate - The payroll date being skipped
 * @returns {Object} Result with new schedule
 */
function skipUniformPayment(token, orderId, currentPayrollDate) {
  requireValidSession_(token);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Uniform_Orders');
    
    if (!sheet) {
      return { success: false, error: 'Uniform_Orders sheet not found' };
    }
    
    // Find the order
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
    let orderRow = -1;
    let order = null;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === orderId) {
        orderRow = i + 2;
        order = data[i];
        break;
      }
    }
    
    if (!order) {
      return { success: false, error: 'Order not found' };
    }
    
    if (order[12] !== 'Active') {
      return { success: false, error: 'Order is not active' };
    }
    
    const oldFirstDeduction = order[8] ? new Date(order[8]) : null;
    if (!oldFirstDeduction) {
      return { success: false, error: 'No first deduction date set' };
    }
    
    // Calculate new first deduction date (+14 days)
    const newFirstDeduction = new Date(oldFirstDeduction);
    newFirstDeduction.setDate(newFirstDeduction.getDate() + 14);
    
    // Update the order
    sheet.getRange(orderRow, 9).setValue(newFirstDeduction); // First_Deduction_Date
    
    // Add note about skip
    const existingNotes = order[13] || '';
    const today = new Date().toLocaleDateString('en-US');
    const newNote = existingNotes 
      ? `${existingNotes}; Payment skipped on ${today}`
      : `Payment skipped on ${today}`;
    sheet.getRange(orderRow, 14).setValue(newNote); // Notes column
    
    // Calculate new schedule
    const paymentSchedule = parseInt(order[6]) || 1;
    const checksCompleted = parseInt(order[9]) || 0;
    const newSchedule = [];
    
    for (let i = checksCompleted; i < paymentSchedule; i++) {
      const deductionDate = new Date(newFirstDeduction);
      deductionDate.setDate(deductionDate.getDate() + ((i - checksCompleted) * 14));
      newSchedule.push(formatDateISO(deductionDate));
    }
    
    console.log(`Skipped payment for ${orderId}. New first deduction: ${formatDateISO(newFirstDeduction)}`);
    
    return {
      success: true,
      orderId: orderId,
      employeeName: order[2],
      oldFirstDeduction: formatDateISO(oldFirstDeduction),
      newFirstDeduction: formatDateISO(newFirstDeduction),
      newSchedule: newSchedule,
      message: `Payment skipped. Next deduction: ${formatDateISO(newFirstDeduction)}`
    };
    
  } catch (error) {
    console.error('Error skipping uniform payment:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// EXPORT FUNCTIONS
// =====================================================

/**
 * Generates HTML content for the payroll report (for print/PDF)
 * Opens in a new window for printing
 * 
 * @param {Object} reportData - The complete report object
 * @returns {string} HTML content
 */
function generatePayrollReportHTML(reportData) {
  if (!reportData || !reportData.success) {
    return '<html><body><h1>Error generating report</h1></body></html>';
  }
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Payroll Report - ${reportData.payrollDate}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { 
      font-family: 'Segoe UI', Arial, sans-serif; 
      font-size: 11px; 
      line-height: 1.4;
      color: #333;
      padding: 20px;
      max-width: 1000px;
      margin: 0 auto;
    }
    .header { 
      text-align: center; 
      border-bottom: 3px solid #E51636; 
      padding-bottom: 15px; 
      margin-bottom: 20px;
    }
    .header h1 { 
      color: #E51636; 
      font-size: 24px; 
      margin-bottom: 5px;
    }
    .header .dates { 
      color: #666; 
      font-size: 12px;
    }
    .section { 
      margin-bottom: 25px; 
      border: 1px solid #ddd;
      border-radius: 8px;
      overflow: hidden;
    }
    .section-header { 
      background: #E51636; 
      color: white; 
      padding: 10px 15px;
      font-size: 14px;
      font-weight: bold;
    }
    .section-summary {
      background: #f9f9f9;
      padding: 8px 15px;
      border-bottom: 1px solid #ddd;
      font-size: 11px;
      color: #666;
    }
    .section-content { padding: 10px 15px; }
    .employee-card { 
      padding: 10px 0; 
      border-bottom: 1px solid #eee;
    }
    .employee-card:last-child { border-bottom: none; }
    .employee-name { 
      font-weight: bold; 
      font-size: 12px;
      color: #333;
    }
    .employee-location { 
      color: #666; 
      font-size: 10px;
    }
    .employee-details { 
      margin-top: 5px; 
      padding-left: 15px;
      font-size: 10px;
      color: #555;
    }
    .detail-row { margin: 2px 0; }
    .badge { 
      display: inline-block;
      padding: 2px 6px;
      border-radius: 3px;
      font-size: 9px;
      font-weight: bold;
    }
    .badge-final { background: #22C55E; color: white; }
    .badge-confirmed { background: #22C55E; color: white; }
    .badge-not-confirmed { background: #F59E0B; color: white; }
    .items-list { 
      margin-top: 5px;
      padding-left: 15px;
      font-size: 9px;
      color: #777;
    }
    .summary-section {
      background: #f5f5f5;
      padding: 15px;
      border-radius: 8px;
      margin-top: 20px;
    }
    .summary-section h3 { 
      color: #E51636;
      margin-bottom: 10px;
    }
    .summary-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 10px;
    }
    .summary-item {
      text-align: center;
      padding: 10px;
      background: white;
      border-radius: 5px;
    }
    .summary-value { 
      font-size: 18px; 
      font-weight: bold;
      color: #E51636;
    }
    .summary-label { 
      font-size: 10px; 
      color: #666;
    }
    .footer { 
      margin-top: 30px; 
      text-align: center; 
      color: #999;
      font-size: 9px;
    }
    @media print {
      body { padding: 10px; }
      .section { page-break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>PAYROLL PROCESSING REPORT</h1>
    <div class="dates">
      <strong>Payment Date:</strong> ${reportData.payrollDateDisplay}<br>
      <strong>Pay Period:</strong> ${formatDisplayDate(reportData.payPeriodStart)} - ${formatDisplayDate(reportData.payPeriodEnd)}<br>
      <strong>Generated:</strong> ${new Date(reportData.generatedDate).toLocaleString()}
    </div>
  </div>

  <!-- OVERTIME SECTION -->
  <div class="section">
    <div class="section-header">📊 OVERTIME</div>
    <div class="section-summary">
      ${reportData.overtime.employeeCount} employees with OT • ${reportData.overtime.totalOTHours} total hours
    </div>
    <div class="section-content">
      ${reportData.overtime.employees.length === 0 ? '<p style="color: #999;">No overtime this period</p>' : 
        reportData.overtime.employees.map(emp => `
          <div class="employee-card">
            <div class="employee-name">${emp.employeeName}</div>
            <div class="employee-location">${emp.location}</div>
            <div class="employee-details">
              <div class="detail-row">Regular: ${emp.regularHours} hrs (${emp.regularMinutes} min)</div>
              <div class="detail-row">OT: ${emp.otHours} hrs (${emp.otMinutes} min) — Week 1: ${emp.week1OT}, Week 2: ${emp.week2OT}</div>
              <div class="detail-row">Total: ${emp.totalHours} hrs (${emp.totalMinutes} min)</div>
            </div>
          </div>
        `).join('')
      }
    </div>
  </div>

  <!-- UNIFORM DEDUCTIONS SECTION -->
  <div class="section">
    <div class="section-header">👕 UNIFORM DEDUCTIONS</div>
    <div class="section-summary">
      ${reportData.uniformDeductions.employeeCount} employees • $${reportData.uniformDeductions.totalDeductions.toFixed(2)} total deductions
    </div>
    <div class="section-content">
      ${reportData.uniformDeductions.orders.length === 0 ? '<p style="color: #999;">No uniform deductions this period</p>' :
        reportData.uniformDeductions.orders.map(emp => `
          <div class="employee-card">
            <div class="employee-name">${emp.employeeName} <span class="badge" style="background: #f0f0f0; color: #333;">$${emp.totalDeduction.toFixed(2)}</span></div>
            <div class="employee-location">${emp.locationSummary} • ${emp.orderCount} order${emp.orderCount !== 1 ? 's' : ''}</div>
            <div class="employee-details">
              ${emp.orders.map(order => `
                <div class="detail-row" style="margin-bottom: 6px;">
                  <strong>${order.orderId}</strong> • ${order.orderDate} • ${order.location}<br>
                  Check ${order.checkNumber} of ${order.paymentSchedule} • <strong>Deduct: $${order.deductionAmount.toFixed(2)}</strong> ${order.isFinalPayment ? '<span class="badge badge-final">FINAL</span>' : ''}<br>
                  Remaining after this: $${order.amountRemaining.toFixed(2)}<br>
                  Items: ${order.items.map(i => `${i.description}${i.size ? ' ('+i.size+')' : ''}`).join(', ')}
                </div>
              `).join('')}
            </div>
          </div>
        `).join('')
      }
    </div>
  </div>

  <!-- PTO PAYOUTS SECTION -->
  <div class="section">
    <div class="section-header">🏖️ PTO PAYOUTS</div>
    <div class="section-summary">
      ${reportData.ptoPayout.employeeCount} employees • ${reportData.ptoPayout.totalHours} total hours
    </div>
    <div class="section-content">
      ${reportData.ptoPayout.requests.length === 0 ? '<p style="color: #999;">No PTO payouts this period</p>' :
        reportData.ptoPayout.requests.map(pto => `
          <div class="employee-card">
            <div class="employee-name">${pto.employeeName} <span class="badge ${pto.hotSchedulesConfirmed ? 'badge-confirmed' : 'badge-not-confirmed'}">${pto.hotSchedulesConfirmed ? '✓ HS' : '✗ HS'}</span></div>
            <div class="employee-location">${pto.location} • ${pto.ptoId}</div>
            <div class="employee-details">
              <div class="detail-row">Hours: ${pto.hoursRequested} • Dates: ${formatDisplayDate(pto.ptoStartDate)}${pto.ptoStartDate !== pto.ptoEndDate ? ' - ' + formatDisplayDate(pto.ptoEndDate) : ''} (${pto.durationDays} day${pto.durationDays > 1 ? 's' : ''})</div>
              ${pto.notes ? `<div class="detail-row">Notes: ${pto.notes}</div>` : ''}
            </div>
          </div>
        `).join('')
      }
    </div>
  </div>

  <!-- TIME TRANSFERS SECTION -->
  <div class="section">
    <div class="section-header">🔄 TIME TRANSFERS</div>
    <div class="section-summary">
      ${reportData.timeTransfers.employeeCount} employees worked multiple locations
    </div>
    <div class="section-content">
      ${reportData.timeTransfers.employees.length === 0 ? '<p style="color: #999;">No multi-location employees this period</p>' :
        reportData.timeTransfers.employees.map(emp => {
          const totalMins = emp.grandTotalMinutes || 0;
          const totalHrs = Math.floor(totalMins / 60);
          const remainingMins = totalMins % 60;
          const totalTimeFormatted = totalHrs + ' hrs ' + remainingMins + ' min';
          const totalRegular = emp.locations.reduce((sum, loc) => sum + (loc.regularHours || 0), 0);
          const totalOT = emp.locations.reduce((sum, loc) => sum + (loc.otHours || 0), 0);
          return `
          <div class="employee-card">
            <div class="employee-name" style="display: flex; justify-content: space-between; align-items: center;">
              <span>${emp.employeeName}</span>
              <span style="font-size: 0.85em; color: #666; font-weight: 600;">Total: ${totalTimeFormatted} — Reg: ${(Math.floor(totalRegular * 100) / 100).toFixed(2)}, OT: ${(Math.floor(totalOT * 100) / 100).toFixed(2)}</span>
            </div>
            <div class="employee-location">Paid from: ${emp.paidFromLocation}</div>
            <div class="employee-details">
              ${emp.locations.map(loc => `
                <div class="detail-row"><strong>${loc.name}:</strong> ${loc.totalHours} hrs (${loc.totalMinutes} min) — Reg: ${loc.regularHours}, OT: ${loc.otHours}</div>
              `).join('')}
              <div class="detail-row" style="margin-top: 5px; color: #E51636;">
                <strong>Transfer:</strong> ${emp.transferHours} hrs (${emp.transferMinutes} min) from ${emp.transferFrom} → ${emp.transferTo}
              </div>
            </div>
          </div>
        `}).join('')
      }
    </div>
  </div>

  <!-- SUMMARY -->
  <div class="summary-section">
    <h3>SUMMARY</h3>
    <div class="summary-grid">
      <div class="summary-item">
        <div class="summary-value">${reportData.overtime.employeeCount}</div>
        <div class="summary-label">Employees with OT</div>
        <div class="summary-label">${reportData.overtime.totalOTHours} hrs</div>
      </div>
      <div class="summary-item">
        <div class="summary-value">${reportData.uniformDeductions.employeeCount}</div>
        <div class="summary-label">Uniform Deductions</div>
        <div class="summary-label">$${reportData.uniformDeductions.totalDeductions.toFixed(2)}</div>
      </div>
      <div class="summary-item">
        <div class="summary-value">${reportData.ptoPayout.employeeCount}</div>
        <div class="summary-label">PTO Payouts</div>
        <div class="summary-label">${reportData.ptoPayout.totalHours} hrs</div>
      </div>
      <div class="summary-item">
        <div class="summary-value">${reportData.timeTransfers.employeeCount}</div>
        <div class="summary-label">Time Transfers</div>
        <div class="summary-label">Multi-location</div>
      </div>
    </div>
  </div>

  <div class="footer">
    Generated by Payroll System • ${new Date().toLocaleString()}
  </div>
</body>
</html>
  `;
  
  return html;
}

/**
 * Helper to format date for display
 */
function formatDisplayDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T12:00:00');
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

/**
 * Generates CSV content for the payroll report
 * Simple summary format - one row per employee with activity
 * 
 * @param {Object} reportData - The complete report object
 * @returns {string} CSV content
 */
function generatePayrollReportCSV(reportData) {
  if (!reportData || !reportData.success) {
    return 'Error,Unable to generate report';
  }
  
  // Collect all unique employees with their data.
  // All three sources are keyed by a canonical name match-key (generateMatchKey) so the
  // same person is never split across rows because OT, uniforms, and PTO stored slightly
  // different IDs for them.
  const employeeMap = new Map();
  const keyOf = (name, fallback) =>
    generateMatchKey(name || '') || (name || '').toString().trim().toLowerCase() || String(fallback || '');

  const getOrCreate = (key, name, location) => {
    if (!employeeMap.has(key)) {
      employeeMap.set(key, {
        name: name,
        location: location,
        otHours: 0,
        uniformDeduction: 0,
        ptoHours: 0,
        notes: []
      });
    }
    return employeeMap.get(key);
  };

  // Process OT employees
  reportData.overtime.employees.forEach(emp => {
    const e = getOrCreate(keyOf(emp.employeeName, emp.employeeId), emp.employeeName, emp.location);
    e.otHours = emp.otHours;
    if (emp.isMultiLocation && !e.notes.includes('Multi-location')) e.notes.push('Multi-location');
  });

  // Process uniform deductions
  reportData.uniformDeductions.orders.forEach(emp => {
    const location = emp.locations && emp.locations.length === 1 ? emp.locations[0] : (emp.locationSummary || 'Multiple Locations');
    const e = getOrCreate(keyOf(emp.employeeName, emp.employeeKey || emp.employeeId), emp.employeeName, location);
    e.uniformDeduction += emp.totalDeduction || 0;
    if ((emp.orders || []).some(o => o.isFinalPayment)) {
      e.notes.push('Final uniform payment');
    }
  });

  // Process PTO
  reportData.ptoPayout.requests.forEach(pto => {
    const e = getOrCreate(keyOf(pto.employeeName, pto.employeeId), pto.employeeName, pto.location);
    e.ptoHours += pto.hoursRequested;
  });

  // Process transfers
  reportData.timeTransfers.employees.forEach(emp => {
    const key = keyOf(emp.employeeName, emp.employeeId);
    if (employeeMap.has(key)) {
      const e = employeeMap.get(key);
      if (!e.notes.includes('Multi-location')) {
        e.notes.push(`Transfer from ${emp.transferFrom}`);
      }
    }
  });
  
  // Build CSV
  const header = 'Employee_Name,Location,OT_Hours,Uniform_Deduction,PTO_Hours,Notes';
  const rows = [];
  
  // Sort by location, then name
  const sorted = Array.from(employeeMap.entries()).sort((a, b) => {
    if (a[1].location !== b[1].location) return a[1].location.localeCompare(b[1].location);
    return a[1].name.localeCompare(b[1].name);
  });
  
  sorted.forEach(([id, emp]) => {
    const notes = emp.notes.join('; ').replace(/"/g, '""');
    rows.push(`"${emp.name}","${emp.location}",${emp.otHours},${emp.uniformDeduction.toFixed(2)},${emp.ptoHours},"${notes}"`);
  });
  
  return header + '\n' + rows.join('\n');
}

