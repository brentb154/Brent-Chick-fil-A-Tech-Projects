/**
 * Reports Module - Handles payroll processing, reporting, and settings
 * Part of Chunk 7: Unified Reporting & Payroll Integration
 */

// =====================================================
// CONSTANTS
// =====================================================

const PAYROLL_SETTINGS_SHEET = 'Payroll_Settings';

// Reference payday: Friday, November 28, 2025 (also aligns with Nov 29, 2024)
// Using noon to avoid any timezone edge cases
const REFERENCE_PAYDAY = new Date(2025, 10, 28, 12, 0, 0); // Month is 0-indexed, Nov 28 2025 is Friday

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
    ['default_ot_rate', '16.50', 'Default hourly rate for OT cost calculations', today]
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
  
  // Days since reference payday
  const daysSinceRef = Math.floor((today - REFERENCE_PAYDAY) / (1000 * 60 * 60 * 24));
  
  // Number of complete 14-day periods since reference
  const periodsSinceRef = Math.floor(daysSinceRef / 14);
  
  // Next payday is the next multiple of 14 days from reference
  let nextPayday = new Date(REFERENCE_PAYDAY);
  nextPayday.setDate(REFERENCE_PAYDAY.getDate() + ((periodsSinceRef + 1) * 14));
  
  // If next payday is today or in the past, get the one after
  if (nextPayday <= today) {
    nextPayday.setDate(nextPayday.getDate() + 14);
  }
  
  return nextPayday;
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
 * Uses REFERENCE_PAYDAY (Friday, November 28, 2025) as the calculation base
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
    
    // Clone REFERENCE_PAYDAY to avoid mutating the constant
    let mostRecentPayday = new Date(REFERENCE_PAYDAY.getTime());
    
    // Move forward until we pass today
    while (mostRecentPayday < today) {
      mostRecentPayday.setDate(mostRecentPayday.getDate() + 14);
    }
    
    // If mostRecentPayday is in the future, go back one period
    if (mostRecentPayday > today) {
      mostRecentPayday.setDate(mostRecentPayday.getDate() - 14);
    }
    
    // Verify mostRecentPayday is a Friday (day 5)
    if (mostRecentPayday.getDay() !== 5) {
      console.error('WARNING: mostRecentPayday is not a Friday! Day of week: ' + mostRecentPayday.getDay() + ', Date: ' + mostRecentPayday);
    }
    
    // Generate historical dates (going backwards from most recent)
    for (let i = historyCount - 1; i >= 0; i--) {
      const payDate = new Date(mostRecentPayday.getTime());
      payDate.setDate(mostRecentPayday.getDate() - (i * 14));
      
      // Skip if date is invalid
      if (isNaN(payDate.getTime())) continue;
      
      addPaydayToList(dates, payDate, today);
    }
    
    // Generate future dates (starting from next payday after most recent)
    for (let i = 1; i <= futureCount; i++) {
      const payDate = new Date(mostRecentPayday.getTime());
      payDate.setDate(mostRecentPayday.getDate() + (i * 14));
      
      // Skip if date is invalid
      if (isNaN(payDate.getTime())) continue;
      
      addPaydayToList(dates, payDate, today);
    }
    
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
  const payday = new Date(payrollDate);
  if (typeof payrollDate === 'string') {
    payday.setTime(new Date(payrollDate + 'T12:00:00').getTime());
  }
  
  // Period end is Saturday, 6 days before Friday payday
  // Actually, if payday is Friday, period ends the Saturday BEFORE that
  // Friday Dec 13 payday -> Period ends Saturday Dec 7
  const periodEnd = new Date(payday);
  periodEnd.setDate(periodEnd.getDate() - 6); // Go back to Saturday
  
  // Period start is 13 days before period end (14 day period)
  const periodStart = new Date(periodEnd);
  periodStart.setDate(periodStart.getDate() - 13);
  
  return {
    periodStart: periodStart,
    periodStartString: formatDateISO(periodStart),
    periodEnd: periodEnd,
    periodEndString: formatDateISO(periodEnd)
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
    uniformData.orders.forEach(o => employeeIds.add(o.employeeId));
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
        regularMinutes: Math.round(regularHours * 60),
        otHours: totalOT,
        otMinutes: Math.round(totalOT * 60),
        totalHours: totalHours,
        totalMinutes: Math.round(totalHours * 60),
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
      totalOTHours: Math.round(totalOTHours * 100) / 100,
      employees: employees
    };
    
  } catch (error) {
    console.error('Error getting overtime for payroll:', error);
    return { employeeCount: 0, totalOTHours: 0, employees: [] };
  }
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
    
    const paydayStr = formatDateISO(payday);
    console.log(`getUniformDeductionsForPayroll called with payday: ${payday}, formatted as: ${paydayStr}`);
    
    // First, get all line item totals to calculate correct order totals
    const lineItemTotals = {};
    const itemsSheet = ss.getSheetByName('Uniform_Order_Items');
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
    
    // Read orders - columns adjusted for current structure
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
    
    console.log(`Total rows in Uniform_Orders: ${data.length}`);
    
    const orders = [];
    let activeCount = 0;
    let noFirstDeductionCount = 0;
    
    data.forEach((row, index) => {
      const orderId = row[0];
      if (!orderId) return;
      
      const status = row[12]; // Status column
      const firstDeductionDate = row[8]; // First_Deduction_Date
      
      // Debug: Log all orders regardless of status (show raw date for troubleshooting)
      if (index < 10) {
        console.log(`Row ${index}: OrderID=${orderId}, Status=${status}, FirstDeduction=${firstDeductionDate}`);
      }
      
      if (status !== 'Active') return; // Only active orders
      activeCount++;
      
      if (!firstDeductionDate) {
        noFirstDeductionCount++;
        console.log(`Order ${orderId} is Active but has NO First_Deduction_Date!`);
        return;
      }
      
      // Get the first deduction date as an ISO string for consistent comparison
      // IMPORTANT: Dates from Google Sheets are stored as midnight UTC, but displayed in local time.
      // When the spreadsheet stores "2026-01-09", it comes back as midnight UTC Jan 9 = 6PM CST Jan 8
      // We need to use UTC components to get the actual date that was stored.
      let firstDeductionStr;
      if (firstDeductionDate instanceof Date) {
        // Use UTC components to avoid timezone shift issues with spreadsheet dates
        const year = firstDeductionDate.getUTCFullYear();
        const month = String(firstDeductionDate.getUTCMonth() + 1).padStart(2, '0');
        const day = String(firstDeductionDate.getUTCDate()).padStart(2, '0');
        firstDeductionStr = `${year}-${month}-${day}`;
      } else {
        // It's a string - parse and format to ensure consistency
        const parsed = new Date(firstDeductionDate + 'T12:00:00'); // Parse as local noon
        if (isNaN(parsed.getTime())) return;
        firstDeductionStr = formatDateISO(parsed);
      }
      
      const paymentSchedule = parseInt(row[6]) || 1;
      const checksCompleted = parseInt(row[9]) || 0;
      
      // Use calculated total from line items (not stored value which may be wrong)
      const storedTotal = parseFloat(row[5]) || 0;
      const totalAmount = lineItemTotals[orderId] !== undefined ? lineItemTotals[orderId] : storedTotal;
      
      // Recalculate amount per check based on correct total
      const amountPerCheck = totalAmount > 0 ? Math.round((totalAmount / paymentSchedule) * 100) / 100 : 0;
      
      // Calculate all deduction date STRINGS for this order (avoids timezone issues)
      const deductionDateStrs = [];
      // Parse firstDeductionStr as local noon to avoid any timezone drift when adding days
      const baseDate = new Date(firstDeductionStr + 'T12:00:00');
      for (let i = 0; i < paymentSchedule; i++) {
        const deductionDate = new Date(baseDate);
        deductionDate.setDate(deductionDate.getDate() + (i * 14));
        deductionDateStrs.push(formatDateISO(deductionDate));
      }
      
      // Debug log for troubleshooting (now showing corrected date string)
      console.log(`Order ${orderId}: firstDeduction=${firstDeductionStr}, payday=${paydayStr}, deductionDates=${deductionDateStrs.join(',')}, match=${deductionDateStrs.includes(paydayStr)}`);
      
      // Check if this payday matches any of the remaining deduction dates
      const checkIndex = deductionDateStrs.findIndex(dStr => dStr === paydayStr);
      
      if (checkIndex === -1) return; // No deduction on this date
      if (checkIndex < checksCompleted) return; // This check already paid
      
      const checkNumber = checkIndex + 1;
      const checksRemaining = paymentSchedule - checkNumber;
      const isFinalPayment = checkNumber === paymentSchedule;
      
      // Calculate deduction amount (last check might be adjusted for rounding)
      let deductionAmount = amountPerCheck;
      if (isFinalPayment) {
        const paidSoFar = checksCompleted * amountPerCheck;
        deductionAmount = Math.round((totalAmount - paidSoFar) * 100) / 100;
      }
      
      const amountRemaining = Math.round((totalAmount - (checksCompleted * amountPerCheck) - deductionAmount) * 100) / 100;
      
      // Get order items
      const items = getOrderItems(ss, orderId);
      
      orders.push({
        employeeId: row[1],
        employeeName: row[2],
        location: row[3] || 'Unknown',
        orderId: orderId,
        orderDate: row[4] ? formatDateISO(new Date(row[4])) : '',
        totalOrderCost: totalAmount,
        paymentSchedule: paymentSchedule,
        amountPerCheck: amountPerCheck,
        checkNumber: checkNumber,
        checksCompleted: checksCompleted,
        checksRemaining: checksRemaining,
        deductionAmount: deductionAmount,
        amountRemaining: Math.max(0, amountRemaining),
        isFinalPayment: isFinalPayment,
        items: items,
        rowIndex: index + 2
      });
    });
    
    // Sort by location, then by name
    orders.sort((a, b) => {
      if (a.location !== b.location) return a.location.localeCompare(b.location);
      return a.employeeName.localeCompare(b.employeeName);
    });
    
    const totalDeductions = orders.reduce((sum, o) => sum + o.deductionAmount, 0);
    const uniqueEmployees = new Set(orders.map(o => o.employeeId));
    
    console.log(`Summary: ${activeCount} active orders, ${noFirstDeductionCount} missing FirstDeduction, ${orders.length} matched payday ${paydayStr}`);
    
    return {
      employeeCount: uniqueEmployees.size,
      totalDeductions: Math.round(totalDeductions * 100) / 100,
      orders: orders
    };
    
  } catch (error) {
    console.error('Error getting uniform deductions for payroll:', error);
    return { employeeCount: 0, totalDeductions: 0, orders: [] };
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
      if (paidOut) return; // Already paid
      
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
        rowIndex: index + 2
      });
    });
    
    // Sort by location, then by name
    requests.sort((a, b) => {
      if (a.location !== b.location) return a.location.localeCompare(b.location);
      return a.employeeName.localeCompare(b.employeeName);
    });
    
    const totalHours = requests.reduce((sum, r) => sum + r.hoursRequested, 0);
    const uniqueEmployees = new Set(requests.map(r => r.employeeId));
    
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
    
    // Get configured location names from settings
    const settings = getSettings();
    const loc1Name = settings.location1Name || 'Location 1';
    const loc2Name = settings.location2Name || 'Location 2';
    
    const periodEndStr = formatDateISO(periodEndDate);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
    
    const employees = [];
    
    data.forEach(row => {
      const rowPeriodEnd = row[0] ? formatDateISO(new Date(row[0])) : null;
      if (rowPeriodEnd !== periodEndStr) return;
      
      const isMultiLocation = row[15] === true || row[15] === 'TRUE';
      if (!isMultiLocation) return;
      
      const employeeName = row[1] || '';
      const matchKey = row[2] || employeeName.toLowerCase();
      const paidFromLocation = row[3] || 'Unknown'; // This is the location they're paid from
      
      // Column 4 = Location 1 hours, Column 5 = Location 2 hours
      const loc1Hours = parseFloat(row[4]) || 0;
      const loc2Hours = parseFloat(row[5]) || 0;
      const totalHours = parseFloat(row[6]) || 0;
      const week1OT = parseFloat(row[10]) || 0;
      const week2OT = parseFloat(row[11]) || 0;
      const totalOT = parseFloat(row[12]) || 0;
      
      // Determine transfer direction based on where employee worked MORE hours
      // They should be paid from the location with more hours, and the other hours transfer there
      let transferFrom = '';
      let transferTo = '';
      let transferHours = 0;
      let actualPaidFromLocation = '';
      
      if (loc1Hours >= loc2Hours) {
        // Worked more at Location 1, so paid from Location 1
        // Transfer Location 2 hours TO Location 1
        actualPaidFromLocation = loc1Name;
        transferFrom = loc2Name;
        transferTo = loc1Name;
        transferHours = loc2Hours;
      } else {
        // Worked more at Location 2, so paid from Location 2
        // Transfer Location 1 hours TO Location 2
        actualPaidFromLocation = loc2Name;
        transferFrom = loc1Name;
        transferTo = loc2Name;
        transferHours = loc1Hours;
      }
      
      // Calculate total hours across all locations
      const grandTotalHours = loc1Hours + loc2Hours;
      const grandTotalMinutes = Math.round(grandTotalHours * 60);
      
      // Calculate OT per location (approximate split based on hours ratio)
      const totalLocationHours = loc1Hours + loc2Hours;
      const loc1Ratio = totalLocationHours > 0 ? loc1Hours / totalLocationHours : 0;
      const loc2Ratio = totalLocationHours > 0 ? loc2Hours / totalLocationHours : 0;
      
      // Regular hours = total - OT
      const totalRegular = totalHours - totalOT;
      const loc1Regular = Math.round(Math.min(loc1Hours, totalRegular * loc1Ratio) * 100) / 100;
      const loc2Regular = Math.round(Math.min(loc2Hours, totalRegular * loc2Ratio) * 100) / 100;
      const loc1OT = Math.round((loc1Hours - loc1Regular) * 100) / 100;
      const loc2OT = Math.round((loc2Hours - loc2Regular) * 100) / 100;
      
      employees.push({
        employeeId: matchKey,
        employeeName: employeeName,
        homeLocation: actualPaidFromLocation,
        locations: [
          {
            name: loc1Name,
            regularHours: Math.max(0, loc1Regular),
            regularMinutes: Math.round(Math.max(0, loc1Regular) * 60),
            otHours: Math.max(0, loc1OT),
            otMinutes: Math.round(Math.max(0, loc1OT) * 60),
            totalHours: loc1Hours,
            totalMinutes: Math.round(loc1Hours * 60)
          },
          {
            name: loc2Name,
            regularHours: Math.max(0, loc2Regular),
            regularMinutes: Math.round(Math.max(0, loc2Regular) * 60),
            otHours: Math.max(0, loc2OT),
            otMinutes: Math.round(Math.max(0, loc2OT) * 60),
            totalHours: loc2Hours,
            totalMinutes: Math.round(loc2Hours * 60)
          }
        ],
        paidFromLocation: actualPaidFromLocation,
        grandTotalHours: grandTotalHours,
        grandTotalMinutes: grandTotalMinutes,
        transferFrom: transferFrom,
        transferTo: transferTo,
        transferHours: transferHours,
        transferMinutes: Math.round(transferHours * 60)
      });
    });
    
    // Sort by name
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
function markPayrollComplete(payrollDate, skippedOrderIds = []) {
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
  }
}

/**
 * Marks uniform payments complete for a specific payday
 */
function markUniformPaymentsComplete(payday, skippedOrderIds = []) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const report = getUniformDeductionsForPayroll(ss, payday);
    
    let marked = 0;
    let skipped = 0;
    let totalAmount = 0;
    const errors = [];
    
    report.orders.forEach(order => {
      if (skippedOrderIds.includes(order.orderId)) {
        skipped++;
        return;
      }
      
      try {
        // Call existing function from Code.gs
        const result = recordUniformPayment(order.orderId);
        if (result.success) {
          marked++;
          totalAmount += order.deductionAmount;
        } else {
          errors.push(`${order.orderId}: ${result.error}`);
        }
      } catch (e) {
        errors.push(`${order.orderId}: ${e.message}`);
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
    
    if (report.requests.length === 0) {
      return { marked: 0, totalHours: 0, errors: null };
    }
    
    const ptoIds = report.requests.map(r => r.ptoId);
    const paydayStr = formatDateISO(payday);
    
    // Call existing batch function from PTOModule.gs
    const result = batchMarkPTOPaid(ptoIds, paydayStr);
    
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
function skipUniformPayment(orderId, currentPayrollDate) {
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
        reportData.uniformDeductions.orders.map(order => `
          <div class="employee-card">
            <div class="employee-name">${order.employeeName} ${order.isFinalPayment ? '<span class="badge badge-final">FINAL</span>' : ''}</div>
            <div class="employee-location">${order.location} • ${order.orderId}</div>
            <div class="employee-details">
              <div class="detail-row">Order Date: ${order.orderDate} • Total: $${order.totalOrderCost.toFixed(2)}</div>
              <div class="detail-row">Payment: Check ${order.checkNumber} of ${order.paymentSchedule} • <strong>Deduct: $${order.deductionAmount.toFixed(2)}</strong></div>
              <div class="detail-row">Remaining after this: $${order.amountRemaining.toFixed(2)}</div>
              <div class="items-list">Items: ${order.items.map(i => `${i.description}${i.size ? ' ('+i.size+')' : ''}`).join(', ')}</div>
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
              <span style="font-size: 0.85em; color: #666; font-weight: 600;">Total: ${totalTimeFormatted} — Reg: ${totalRegular.toFixed(2)}, OT: ${totalOT.toFixed(2)}</span>
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
        <div class="summary-value">${reportData.uniformDeductions.orders.length}</div>
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
    Generated by Payroll Review • ${new Date().toLocaleString()}
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
  
  // Collect all unique employees with their data
  const employeeMap = new Map();
  
  // Process OT employees
  reportData.overtime.employees.forEach(emp => {
    if (!employeeMap.has(emp.employeeId)) {
      employeeMap.set(emp.employeeId, {
        name: emp.employeeName,
        location: emp.location,
        otHours: 0,
        uniformDeduction: 0,
        ptoHours: 0,
        notes: []
      });
    }
    const e = employeeMap.get(emp.employeeId);
    e.otHours = emp.otHours;
    if (emp.isMultiLocation) e.notes.push('Multi-location');
  });
  
  // Process uniform deductions
  reportData.uniformDeductions.orders.forEach(order => {
    if (!employeeMap.has(order.employeeId)) {
      employeeMap.set(order.employeeId, {
        name: order.employeeName,
        location: order.location,
        otHours: 0,
        uniformDeduction: 0,
        ptoHours: 0,
        notes: []
      });
    }
    const e = employeeMap.get(order.employeeId);
    e.uniformDeduction += order.deductionAmount;
    if (order.isFinalPayment) e.notes.push('Final uniform payment');
  });
  
  // Process PTO
  reportData.ptoPayout.requests.forEach(pto => {
    if (!employeeMap.has(pto.employeeId)) {
      employeeMap.set(pto.employeeId, {
        name: pto.employeeName,
        location: pto.location,
        otHours: 0,
        uniformDeduction: 0,
        ptoHours: 0,
        notes: []
      });
    }
    const e = employeeMap.get(pto.employeeId);
    e.ptoHours += pto.hoursRequested;
  });
  
  // Process transfers
  reportData.timeTransfers.employees.forEach(emp => {
    if (employeeMap.has(emp.employeeId)) {
      const e = employeeMap.get(emp.employeeId);
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

