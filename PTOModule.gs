/**
 * PTO Module - Handles all PTO (Paid Time Off) functionality
 * Part of Chunk 6: PTO System with Employee Request Form
 */

// =====================================================
// INITIALIZATION
// =====================================================

/**
 * Initializes the PTO tab if it doesn't exist
 * Creates headers and formats the sheet
 */
function initializePTOTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PTO');
  
  if (!sheet) {
    sheet = ss.insertSheet('PTO');
    
    // Set headers (13 columns)
    const headers = [
      'PTO_ID',
      'Employee_ID', 
      'Employee_Name',
      'Location',
      'Hours_Requested',
      'PTO_Start_Date',
      'PTO_End_Date',
      'Payout_Period',
      'HotSchedules_Confirmed',
      'Submission_Date',
      'Paid_Out',
      'Status',
      'Notes'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E8EAED');
    headerRange.setHorizontalAlignment('center');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // PTO_ID
    sheet.setColumnWidth(2, 100);  // Employee_ID
    sheet.setColumnWidth(3, 150);  // Employee_Name
    sheet.setColumnWidth(4, 130);  // Location
    sheet.setColumnWidth(5, 80);   // Hours_Requested
    sheet.setColumnWidth(6, 110);  // PTO_Start_Date
    sheet.setColumnWidth(7, 110);  // PTO_End_Date
    sheet.setColumnWidth(8, 110);  // Payout_Period
    sheet.setColumnWidth(9, 140);  // HotSchedules_Confirmed
    sheet.setColumnWidth(10, 140); // Submission_Date
    sheet.setColumnWidth(11, 80);  // Paid_Out
    sheet.setColumnWidth(12, 90);  // Status
    sheet.setColumnWidth(13, 200); // Notes
    
    console.log('PTO tab created successfully');
    return { success: true, message: 'PTO tab created' };
  }
  
  console.log('PTO tab already exists');
  return { success: true, message: 'PTO tab already exists' };
}

// =====================================================
// EMPLOYEE-FACING FUNCTIONS (Public Access)
// =====================================================

/**
 * Gets list of active employees for the PTO form dropdown
 * Returns: Array of {id, name, location} objects sorted alphabetically
 */
function getActiveEmployeesForPTO() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // First, try to get employees from OT_History (most accurate - shows who's actually working)
    const historySheet = ss.getSheetByName('OT_History');
    
    if (historySheet && historySheet.getLastRow() >= 2) {
      // Get the last 4 pay periods
      const periods = getPayPeriods();
      
      if (periods && periods.length > 0) {
        const recentPeriods = periods.slice(0, 4);
        const recentPeriodTimes = new Set(recentPeriods.map(p => new Date(p).getTime()));
        
        // Get data from OT_History: Period End, Name, Match Key, Location, Total Hours (cols 1,2,3,4,7)
        const lastRow = historySheet.getLastRow();
        const data = historySheet.getRange(2, 1, lastRow - 1, 7).getValues();
        
        // Build map of active employees with their location
        const activeEmployees = new Map();
        
        for (const row of data) {
          const periodEnd = row[0];
          const displayName = row[1];
          const matchKey = row[2] || (displayName ? displayName.toLowerCase() : null);
          const location = row[3] || 'Unknown';
          const totalHours = parseFloat(row[6]) || 0;
          
          if (!displayName || !periodEnd) continue;
          
          // Check if this record is from a recent period AND has hours
          const periodTime = new Date(periodEnd).getTime();
          if (recentPeriodTimes.has(periodTime) && totalHours > 0) {
            const key = matchKey || displayName.toLowerCase();
            if (!activeEmployees.has(key)) {
              activeEmployees.set(key, {
                id: key,
                name: displayName,
                location: location
              });
            }
          }
        }
        
        if (activeEmployees.size > 0) {
          const employees = Array.from(activeEmployees.values())
            .sort((a, b) => a.name.localeCompare(b.name));
          console.log('Active employees from OT_History:', employees.length);
          return employees;
        }
      }
    }
    
    // Fallback: Get employees from Employees sheet
    console.log('Falling back to Employees sheet');
    const empSheet = ss.getSheetByName('Employees');
    
    if (empSheet && empSheet.getLastRow() >= 2) {
      const lastRow = empSheet.getLastRow();
      // Columns: A=Employee_ID(match_key), B=Display_Name, C=Match_Key, D=Primary_Location
      const data = empSheet.getRange(2, 1, lastRow - 1, 4).getValues();
      
      const employees = data
        .filter(row => row[0] && row[1]) // Has ID and name
        .map(row => ({
          id: row[0],
          name: row[1],
          location: row[3] || 'Unknown'
        }))
        .sort((a, b) => a.name.localeCompare(b.name));
      
      console.log('Employees from Employees sheet:', employees.length);
      return employees;
    }
    
    console.log('No employee data found in any sheet');
    return [];
    
  } catch (error) {
    console.error('Error getting active employees for PTO:', error);
    return [];
  }
}

/**
 * Gets upcoming payroll dates for the payout period dropdown
 * Paydays are on FRIDAYS, bi-weekly
 * Reference date: Friday, November 28, 2025
 * Excludes dates less than 3 days away
 * 
 * @param {string} minDateStr - Optional. ISO date string (YYYY-MM-DD). Only return paydays AFTER this date.
 *                              Used for PTO to ensure payout is after time off ends.
 */
function getUpcomingPayrollDates(minDateStr) {
  try {
    const dates = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const minDaysAway = 3; // Minimum 3 days until payday to include it
    
    // If a minimum date is provided (PTO end date), use it as the baseline
    let effectiveMinDate = today;
    if (minDateStr) {
      const providedMin = new Date(minDateStr + 'T00:00:00');
      if (!isNaN(providedMin.getTime()) && providedMin > today) {
        effectiveMinDate = providedMin;
      }
    }
    
    // Reference payday: Friday, November 28, 2025
    const referencePayday = new Date(2025, 10, 28); // Month is 0-indexed (10 = November)
    
    // Calculate days since reference from the effective minimum date
    const daysSinceRef = Math.floor((effectiveMinDate - referencePayday) / (1000 * 60 * 60 * 24));
    
    // Find next payday (bi-weekly from reference)
    // Number of complete 14-day periods since reference
    const periodsSinceRef = Math.floor(daysSinceRef / 14);
    
    // Next payday is the next multiple of 14 days from reference
    let nextPayday = new Date(referencePayday);
    nextPayday.setDate(referencePayday.getDate() + ((periodsSinceRef + 1) * 14));
    
    // Make sure nextPayday is after the effectiveMinDate
    while (nextPayday <= effectiveMinDate) {
      nextPayday.setDate(nextPayday.getDate() + 14);
    }
    
    // Also ensure it's at least minDaysAway from today
    const daysToFirst = Math.floor((nextPayday - today) / (1000 * 60 * 60 * 24));
    if (daysToFirst < minDaysAway) {
      nextPayday.setDate(nextPayday.getDate() + 14);
    }
    
    // Generate next 6 paydays (bi-weekly)
    for (let i = 0; i < 6; i++) {
      const payDate = new Date(nextPayday);
      payDate.setDate(payDate.getDate() + (i * 14));
      
      const dateStr = payDate.toISOString().split('T')[0];
      
      // English format
      const englishDisplay = payDate.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      });
      
      // Spanish format
      const spanishDisplay = payDate.toLocaleDateString('es-ES', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      });
      // Capitalize first letter
      const spanishCapitalized = spanishDisplay.charAt(0).toUpperCase() + spanishDisplay.slice(1);
      
      dates.push({
        date: dateStr,
        displayEnglish: englishDisplay,
        displaySpanish: spanishCapitalized
      });
    }
    
    return dates;
    
  } catch (error) {
    console.error('Error getting upcoming payroll dates:', error);
    return [];
  }
}

/**
 * Submits a PTO request from the employee form
 * This is the ONLY function employees can call - it only creates records
 * 
 * @param {Object} requestData - The PTO request data
 * @returns {Object} Success/error response with PTO_ID
 */
function submitPTORequest(requestData) {
  try {
    // Validate inputs
    if (!requestData.employeeId) {
      return { success: false, error: 'Please select your name / Por favor selecciona tu nombre' };
    }
    
    if (!requestData.hoursRequested || requestData.hoursRequested <= 0) {
      return { success: false, error: 'Hours must be greater than zero / Las horas deben ser mayores que cero' };
    }
    
    if (requestData.hoursRequested > 80) {
      return { success: false, error: 'Hours cannot exceed 80 / Las horas no pueden exceder 80' };
    }
    
    if (!requestData.ptoStartDate) {
      return { success: false, error: 'Please select a start date / Por favor selecciona una fecha de inicio' };
    }
    
    if (!requestData.ptoEndDate) {
      return { success: false, error: 'Please select an end date / Por favor selecciona una fecha de fin' };
    }
    
    // Validate dates
    const startDate = new Date(requestData.ptoStartDate + 'T12:00:00');
    const endDate = new Date(requestData.ptoEndDate + 'T12:00:00');
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    if (isNaN(startDate.getTime())) {
      return { success: false, error: 'Invalid start date / Fecha de inicio inválida' };
    }
    
    if (isNaN(endDate.getTime())) {
      return { success: false, error: 'Invalid end date / Fecha de fin inválida' };
    }
    
    if (endDate < startDate) {
      return { success: false, error: 'End date must be after start date / La fecha de fin debe ser después de la fecha de inicio' };
    }
    
    if (!requestData.payoutPeriod) {
      return { success: false, error: 'Please select a pay period / Por favor selecciona un período de pago' };
    }
    
    // Get employee details - try multiple matching strategies
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName('Employees');
    
    let employeeName = null;
    let location = 'Unknown';
    let employeeId = requestData.employeeId;
    
    // Strategy 1: Look up in Employees sheet
    if (empSheet && empSheet.getLastRow() >= 2) {
      const empData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, 5).getValues();
      // Columns: A=Employee_ID(match_key), B=Display_Name, C=Match_Key, D=Location
      
      // Try to find by match key (case-insensitive)
      const searchKey = requestData.employeeId.toLowerCase().trim();
      const employee = empData.find(row => {
        const rowId = (row[0] || '').toString().toLowerCase().trim();
        const rowMatchKey = (row[2] || '').toString().toLowerCase().trim();
        const rowName = (row[1] || '').toString().toLowerCase().trim();
        return rowId === searchKey || rowMatchKey === searchKey || rowName === searchKey;
      });
      
      if (employee) {
        employeeName = employee[1] || requestData.employeeId;
        location = employee[3] || 'Unknown';
        employeeId = employee[0] || requestData.employeeId;
      }
    }
    
    // Strategy 2: If not found in Employees, look up in OT_History
    if (!employeeName) {
      const historySheet = ss.getSheetByName('OT_History');
      if (historySheet && historySheet.getLastRow() >= 2) {
        const histData = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 4).getValues();
        // Columns: A=Period End, B=Name, C=Match Key, D=Location
        
        const searchKey = requestData.employeeId.toLowerCase().trim();
        const histRecord = histData.find(row => {
          const rowMatchKey = (row[2] || '').toString().toLowerCase().trim();
          const rowName = (row[1] || '').toString().toLowerCase().trim();
          return rowMatchKey === searchKey || rowName === searchKey;
        });
        
        if (histRecord) {
          employeeName = histRecord[1];
          location = histRecord[3] || 'Unknown';
        }
      }
    }
    
    // If still not found, use the ID as the name (last resort)
    if (!employeeName) {
      // Capitalize the employee ID as a display name
      employeeName = requestData.employeeId
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
        .join(' ');
      console.log('Employee not found in sheets, using ID as name:', employeeName);
    }
    
    // Initialize PTO tab if needed
    initializePTOTab();
    
    // Generate PTO_ID
    const ptoSheet = ss.getSheetByName('PTO');
    const lastRow = ptoSheet.getLastRow();
    let ptoNumber = 1;
    
    if (lastRow > 1) {
      const lastId = ptoSheet.getRange(lastRow, 1).getValue();
      if (lastId && typeof lastId === 'string' && lastId.startsWith('PTO-')) {
        ptoNumber = parseInt(lastId.replace('PTO-', '')) + 1;
      } else {
        ptoNumber = lastRow;
      }
    }
    
    const ptoId = 'PTO-' + String(ptoNumber).padStart(5, '0');
    
    // Build PTO record (13 columns)
    const ptoRecord = [
      ptoId,                                          // A: PTO_ID
      requestData.employeeId,                         // B: Employee_ID
      employeeName,                                   // C: Employee_Name
      location,                                       // D: Location
      parseFloat(requestData.hoursRequested),         // E: Hours_Requested
      startDate,                                      // F: PTO_Start_Date
      endDate,                                        // G: PTO_End_Date
      new Date(requestData.payoutPeriod + 'T12:00:00'), // H: Payout_Period
      requestData.hotSchedulesConfirmed === true,     // I: HotSchedules_Confirmed
      new Date(),                                     // J: Submission_Date
      false,                                          // K: Paid_Out
      'Pending',                                      // L: Status
      requestData.notes || ''                         // M: Notes
    ];
    
    // Append to PTO sheet
    ptoSheet.appendRow(ptoRecord);
    
    // Format the dates in the response
    const startDateStr = startDate.toLocaleDateString('en-US', { 
      month: 'short', day: 'numeric', year: 'numeric' 
    });
    const endDateStr = endDate.toLocaleDateString('en-US', { 
      month: 'short', day: 'numeric', year: 'numeric' 
    });
    const payoutDateStr = new Date(requestData.payoutPeriod + 'T12:00:00').toLocaleDateString('en-US', {
      weekday: 'long', month: 'short', day: 'numeric', year: 'numeric'
    });
    
    console.log(`PTO request submitted: ${ptoId} for ${employeeName}`);
    
    // Log the activity
    logActivity('CREATE', 'PTO', 
      `PTO request submitted: ${employeeName} - ${parseFloat(requestData.hoursRequested)} hrs (${startDateStr} - ${endDateStr})`,
      ptoId
    );
    
    // Note: Immediate email notifications disabled - using weekly summary instead (Saturday 1PM)
    // PTO requests are included in the weekly summary email
    const dateRange = startDateStr === endDateStr ? startDateStr : `${startDateStr} - ${endDateStr}`;
    
    return {
      success: true,
      ptoId: ptoId,
      employeeName: employeeName,
      hours: parseFloat(requestData.hoursRequested),
      startDate: startDateStr,
      endDate: endDateStr,
      dateRange: dateRange,
      payoutPeriod: payoutDateStr,
      payoutPeriodSpanish: new Date(requestData.payoutPeriod + 'T12:00:00').toLocaleDateString('es-ES', {
        weekday: 'long', month: 'short', day: 'numeric', year: 'numeric'
      })
    };
    
  } catch (error) {
    console.error('Error submitting PTO request:', error);
    return { 
      success: false, 
      error: 'Error saving request. Please try again. / Error al guardar la solicitud. Por favor, inténtalo de nuevo.' 
    };
  }
}

/**
 * Cancel a pending PTO request (for employee undo)
 * Only works for Pending requests
 * @param {string} ptoId - The PTO_ID to cancel
 * @returns {Object} Result
 */
function cancelPTORequest(ptoId) {
  console.log('cancelPTORequest called with ptoId:', ptoId);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ptoSheet = ss.getSheetByName('PTO');
    
    console.log('PTO Sheet found:', !!ptoSheet);
    
    if (!ptoSheet || ptoSheet.getLastRow() < 2) {
      console.log('PTO sheet not found or empty');
      return { success: false, error: 'PTO records not found' };
    }
    
    // Find the PTO request
    const ptoData = ptoSheet.getDataRange().getValues();
    const headers = ptoData[0];
    console.log('PTO Headers:', headers);
    
    const statusColIdx = headers.indexOf('Status');
    const ptoIdColIdx = headers.indexOf('PTO_ID');
    
    console.log('Status column index:', statusColIdx, 'PTO_ID column index:', ptoIdColIdx);
    
    if (ptoIdColIdx < 0 || statusColIdx < 0) {
      console.log('Missing columns - statusColIdx:', statusColIdx, 'ptoIdColIdx:', ptoIdColIdx);
      return { success: false, error: 'Required columns not found in PTO sheet. Headers: ' + headers.join(', ') };
    }
    
    let ptoRow = null;
    let ptoRowNum = -1;
    
    console.log('Searching for PTO ID:', ptoId, 'Type:', typeof ptoId);
    
    for (let i = 1; i < ptoData.length; i++) {
      const rowPtoId = ptoData[i][ptoIdColIdx];
      // Compare as strings to handle potential type mismatches
      if (String(rowPtoId).trim() === String(ptoId).trim()) {
        ptoRow = ptoData[i];
        ptoRowNum = i + 1;
        console.log('Found PTO at row', ptoRowNum, 'Data:', ptoRow);
        break;
      }
    }
    
    if (!ptoRow) {
      // List all PTO IDs for debugging
      const allPtoIds = ptoData.slice(1).map(row => row[ptoIdColIdx]);
      console.log('Available PTO IDs:', allPtoIds);
      return { success: false, error: 'PTO request not found. Looking for: ' + ptoId };
    }
    
    const currentStatus = ptoRow[statusColIdx];
    
    // Only allow cancellation of pending requests
    if (currentStatus !== 'Pending') {
      return { success: false, error: `Cannot cancel - request status is "${currentStatus}". Only pending requests can be cancelled.` };
    }
    
    console.log(`Cancelling PTO request ${ptoId} at row ${ptoRowNum}, status column ${statusColIdx + 1}`);
    
    // Update status to Cancelled
    ptoSheet.getRange(ptoRowNum, statusColIdx + 1).setValue('Cancelled');
    
    // Force the write to complete
    SpreadsheetApp.flush();
    
    // Log the activity
    logActivity('CANCEL', 'PTO', `PTO request ${ptoId} cancelled by employee (undo)`, ptoId);
    
    console.log(`PTO request ${ptoId} cancelled successfully`);
    
    return {
      success: true,
      message: `PTO request ${ptoId} has been cancelled`,
      ptoId: ptoId
    };
    
  } catch (error) {
    console.error('Error cancelling PTO request:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// ADMIN FUNCTIONS (Restricted Access)
// =====================================================

/**
 * Gets PTO records with optional filtering
 * Admin function - should check authorization in production
 * 
 * @param {Object} filters - Optional filters (status, location, employeeId, payoutPeriod)
 * @returns {Array} Array of PTO records
 */
function getPTORecords(filters = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    const today = new Date();
    
    let records = data
      .map((row, index) => {
        if (!row[0]) return null; // Skip empty rows
        
        const startDate = row[5] ? new Date(row[5]) : null;
        const endDate = row[6] ? new Date(row[6]) : null;
        const payoutDate = row[7] ? new Date(row[7]) : null;
        const submissionDate = row[9] ? new Date(row[9]) : null;
        
        // Calculate duration in days
        let durationDays = 1;
        if (startDate && endDate) {
          durationDays = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        }
        
        // Calculate days until payout
        let daysUntilPayout = null;
        if (payoutDate) {
          daysUntilPayout = Math.ceil((payoutDate - today) / (1000 * 60 * 60 * 24));
        }
        
        return {
          ptoId: row[0],
          employeeId: row[1],
          employeeName: row[2],
          location: row[3],
          hoursRequested: parseFloat(row[4]) || 0,
          ptoStartDate: startDate ? startDate.toISOString().split('T')[0] : null,
          ptoEndDate: endDate ? endDate.toISOString().split('T')[0] : null,
          payoutPeriod: payoutDate ? payoutDate.toISOString().split('T')[0] : null,
          hotSchedulesConfirmed: row[8] === true,
          submissionDate: submissionDate ? submissionDate.toISOString() : null,
          paidOut: row[10] === true,
          status: row[11] || 'Pending',
          notes: row[12] || '',
          durationDays: durationDays,
          daysUntilPayout: daysUntilPayout,
          rowIndex: index + 2
        };
      })
      .filter(r => r !== null);
    
    // Apply filters
    if (filters.status) {
      records = records.filter(r => r.status === filters.status);
    }
    if (filters.statuses && Array.isArray(filters.statuses)) {
      records = records.filter(r => filters.statuses.includes(r.status));
    }
    if (filters.location) {
      records = records.filter(r => r.location === filters.location);
    }
    if (filters.employeeId) {
      records = records.filter(r => r.employeeId === filters.employeeId);
    }
    if (filters.payoutPeriod) {
      records = records.filter(r => r.payoutPeriod === filters.payoutPeriod);
    }
    if (filters.unpaidOnly) {
      records = records.filter(r => r.paidOut === false);
    }
    
    // Sort by submission date (newest first), then by employee name
    records.sort((a, b) => {
      const dateA = a.submissionDate ? new Date(a.submissionDate) : new Date(0);
      const dateB = b.submissionDate ? new Date(b.submissionDate) : new Date(0);
      if (dateB - dateA !== 0) return dateB - dateA;
      return a.employeeName.localeCompare(b.employeeName);
    });
    
    return records;
    
  } catch (error) {
    console.error('Error getting PTO records:', error);
    return [];
  }
}

/**
 * Marks a single PTO record as paid
 * 
 * @param {string} ptoId - The PTO_ID to mark as paid
 * @param {string} notes - Optional note to add
 * @returns {Object} Success/error response
 */
function markPTOPaid(ptoId, notes = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet) {
      return { success: false, error: 'PTO sheet not found' };
    }
    
    const records = getPTORecords();
    const record = records.find(r => r.ptoId === ptoId);
    
    if (!record) {
      return { success: false, error: 'PTO record not found' };
    }
    
    if (record.paidOut) {
      return { success: false, error: 'PTO already marked as paid' };
    }
    
    const row = record.rowIndex;
    const today = new Date().toLocaleDateString('en-US');
    
    // Update Paid_Out (column K = 11)
    sheet.getRange(row, 11).setValue(true);
    
    // Update Status (column L = 12)
    sheet.getRange(row, 12).setValue('Paid');
    
    // Update Notes (column M = 13)
    const existingNotes = sheet.getRange(row, 13).getValue() || '';
    const newNote = notes ? `Paid on ${today}. ${notes}` : `Paid on ${today}`;
    const updatedNotes = existingNotes ? `${existingNotes}; ${newNote}` : newNote;
    sheet.getRange(row, 13).setValue(updatedNotes);
    
    console.log(`PTO ${ptoId} marked as paid`);
    
    // Log the activity
    logActivity('STATUS_CHANGE', 'PTO', 
      `PTO marked as paid: ${record.employeeName} - ${record.hoursRequested} hrs`,
      ptoId
    );
    
    return {
      success: true,
      ptoId: ptoId,
      employeeName: record.employeeName,
      hours: record.hoursRequested
    };
    
  } catch (error) {
    console.error('Error marking PTO as paid:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Batch marks multiple PTO records as paid
 * 
 * @param {Array} ptoIds - Array of PTO_IDs to mark as paid
 * @param {string} payrollDate - The payroll date (for notes)
 * @returns {Object} Success/error response with counts
 */
function batchMarkPTOPaid(ptoIds, payrollDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet) {
      return { success: false, error: 'PTO sheet not found' };
    }
    
    const records = getPTORecords();
    let updatedCount = 0;
    let totalHours = 0;
    const errors = [];
    
    ptoIds.forEach(ptoId => {
      const record = records.find(r => r.ptoId === ptoId);
      
      if (!record) {
        errors.push(`${ptoId}: Not found`);
        return;
      }
      
      if (record.paidOut) {
        errors.push(`${ptoId}: Already paid`);
        return;
      }
      
      const row = record.rowIndex;
      
      // Update Paid_Out
      sheet.getRange(row, 11).setValue(true);
      
      // Update Status
      sheet.getRange(row, 12).setValue('Paid');
      
      // Update Notes
      const existingNotes = sheet.getRange(row, 13).getValue() || '';
      const newNote = `Paid in payroll ${payrollDate}`;
      const updatedNotes = existingNotes ? `${existingNotes}; ${newNote}` : newNote;
      sheet.getRange(row, 13).setValue(updatedNotes);
      
      updatedCount++;
      totalHours += record.hoursRequested;
    });
    
    console.log(`Batch marked ${updatedCount} PTO records as paid`);
    
    return {
      success: true,
      updatedCount: updatedCount,
      totalHours: totalHours,
      payrollDate: payrollDate,
      errors: errors.length > 0 ? errors : null
    };
    
  } catch (error) {
    console.error('Error batch marking PTO as paid:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Gets PTO records for a specific payroll date (for payroll report)
 * Groups by employee and sums hours
 * 
 * @param {string} payrollDate - The payroll date (YYYY-MM-DD)
 * @returns {Array} Array of employees with their PTO for this payroll
 */
function getPTOForPayrollReport(payrollDate) {
  try {
    // Get pending PTO for this payout period
    const records = getPTORecords({
      payoutPeriod: payrollDate,
      unpaidOnly: true
    });
    
    // Group by employee
    const byEmployee = {};
    
    records.forEach(record => {
      if (!byEmployee[record.employeeId]) {
        byEmployee[record.employeeId] = {
          employeeId: record.employeeId,
          employeeName: record.employeeName,
          location: record.location,
          totalHours: 0,
          ptoRecords: []
        };
      }
      
      byEmployee[record.employeeId].totalHours += record.hoursRequested;
      byEmployee[record.employeeId].ptoRecords.push({
        ptoId: record.ptoId,
        hours: record.hoursRequested,
        dates: record.ptoStartDate === record.ptoEndDate 
          ? record.ptoStartDate 
          : `${record.ptoStartDate} - ${record.ptoEndDate}`,
        hotSchedulesConfirmed: record.hotSchedulesConfirmed
      });
    });
    
    // Convert to array and sort by name
    const result = Object.values(byEmployee).sort((a, b) => 
      a.employeeName.localeCompare(b.employeeName)
    );
    
    return result;
    
  } catch (error) {
    console.error('Error getting PTO for payroll report:', error);
    return [];
  }
}

/**
 * Gets PTO summary statistics
 * 
 * @param {Object} filters - Optional filters
 * @returns {Object} Summary statistics
 */
function getPTOSummaryStats(filters = {}) {
  try {
    const allRecords = getPTORecords(filters);
    
    // Basic counts
    const totalRecords = allRecords.length;
    const pendingRecords = allRecords.filter(r => r.status === 'Pending');
    const paidRecords = allRecords.filter(r => r.status === 'Paid');
    
    const totalHoursPending = pendingRecords.reduce((sum, r) => sum + r.hoursRequested, 0);
    const totalHoursPaid = paidRecords.reduce((sum, r) => sum + r.hoursRequested, 0);
    
    // By location
    const byLocation = {};
    allRecords.forEach(record => {
      const loc = record.location || 'Unknown';
      if (!byLocation[loc]) {
        byLocation[loc] = { pending: 0, pendingHours: 0, paid: 0, paidHours: 0 };
      }
      if (record.status === 'Pending') {
        byLocation[loc].pending++;
        byLocation[loc].pendingHours += record.hoursRequested;
      } else if (record.status === 'Paid') {
        byLocation[loc].paid++;
        byLocation[loc].paidHours += record.hoursRequested;
      }
    });
    
    // Upcoming payouts (next 2 periods)
    const upcomingPayouts = {};
    const payrollDates = getUpcomingPayrollDates().slice(0, 2);
    
    payrollDates.forEach(pd => {
      const recordsForDate = pendingRecords.filter(r => r.payoutPeriod === pd.date);
      upcomingPayouts[pd.date] = {
        count: recordsForDate.length,
        hours: recordsForDate.reduce((sum, r) => sum + r.hoursRequested, 0)
      };
    });
    
    // Recent submissions (last 7 days)
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const recentSubmissions = allRecords.filter(r => {
      if (!r.submissionDate) return false;
      return new Date(r.submissionDate) >= sevenDaysAgo;
    }).length;
    
    return {
      totalRecords: totalRecords,
      totalHoursPending: totalHoursPending,
      totalHoursPaid: totalHoursPaid,
      pendingCount: pendingRecords.length,
      paidCount: paidRecords.length,
      byLocation: byLocation,
      upcomingPayouts: upcomingPayouts,
      recentSubmissions: recentSubmissions
    };
    
  } catch (error) {
    console.error('Error getting PTO summary stats:', error);
    return {
      totalRecords: 0,
      totalHoursPending: 0,
      totalHoursPaid: 0,
      pendingCount: 0,
      paidCount: 0,
      byLocation: {},
      upcomingPayouts: {},
      recentSubmissions: 0
    };
  }
}

/**
 * Updates a PTO record (admin only)
 * 
 * @param {string} ptoId - The PTO_ID to update
 * @param {Object} updates - Fields to update
 * @returns {Object} Success/error response
 */
function updatePTORecord(ptoId, updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet) {
      return { success: false, error: 'PTO sheet not found' };
    }
    
    const records = getPTORecords();
    const record = records.find(r => r.ptoId === ptoId);
    
    if (!record) {
      return { success: false, error: 'PTO record not found' };
    }
    
    const row = record.rowIndex;
    
    // Update allowed fields
    if (updates.hoursRequested !== undefined) {
      sheet.getRange(row, 5).setValue(parseFloat(updates.hoursRequested));
    }
    if (updates.ptoStartDate !== undefined) {
      sheet.getRange(row, 6).setValue(new Date(updates.ptoStartDate + 'T12:00:00'));
    }
    if (updates.ptoEndDate !== undefined) {
      sheet.getRange(row, 7).setValue(new Date(updates.ptoEndDate + 'T12:00:00'));
    }
    if (updates.payoutPeriod !== undefined) {
      sheet.getRange(row, 8).setValue(new Date(updates.payoutPeriod + 'T12:00:00'));
    }
    if (updates.status !== undefined) {
      sheet.getRange(row, 12).setValue(updates.status);
      // Also update Paid_Out if status is Paid
      if (updates.status === 'Paid') {
        sheet.getRange(row, 11).setValue(true);
      }
    }
    if (updates.notes !== undefined) {
      sheet.getRange(row, 13).setValue(updates.notes);
    }
    
    console.log(`PTO ${ptoId} updated`);
    
    return { success: true, ptoId: ptoId };
    
  } catch (error) {
    console.error('Error updating PTO record:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Deletes a PTO record (admin only)
 * 
 * @param {string} ptoId - The PTO_ID to delete
 * @returns {Object} Success/error response
 */
function deletePTORecord(ptoId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet) {
      return { success: false, error: 'PTO sheet not found' };
    }
    
    const records = getPTORecords();
    const record = records.find(r => r.ptoId === ptoId);
    
    if (!record) {
      return { success: false, error: 'PTO record not found' };
    }
    
    // Delete the row
    sheet.deleteRow(record.rowIndex);
    
    console.log(`PTO ${ptoId} deleted`);
    
    return { success: true, ptoId: ptoId };
    
  } catch (error) {
    console.error('Error deleting PTO record:', error);
    return { success: false, error: error.message };
  }
}

// =====================================================
// PTO SUMMARY / REPORTS
// =====================================================

/**
 * Gets comprehensive PTO summary data for the reports page
 * @returns {Object} Summary statistics and lists
 */
function getPTOSummaryData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PTO');
    
    // Default empty response
    const emptyResponse = {
      success: true,
      summary: {
        totalRequests: 0,
        totalHours: 0,
        pendingRequests: 0,
        pendingHours: 0,
        paidRequests: 0,
        paidHours: 0,
        averageHoursPerRequest: 0,
        uniqueEmployees: 0
      },
      byLocation: [],
      topEmployees: [],
      recentActivity: [],
      byMonth: []
    };
    
    if (!sheet || sheet.getLastRow() < 2) {
      return emptyResponse;
    }
    
    // Read all PTO data
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    
    // Initialize aggregations
    let totalRequests = 0;
    let totalHours = 0;
    let pendingRequests = 0;
    let pendingHours = 0;
    let paidRequests = 0;
    let paidHours = 0;
    
    const employeeMap = new Map(); // Track hours per employee
    const locationMap = new Map(); // Track by location
    const monthlyData = {}; // Track by month
    const recentRecords = [];
    const uniqueEmployees = new Set();
    
    // Process each record
    data.forEach((row, index) => {
      const ptoId = row[0];
      if (!ptoId) return; // Skip empty rows
      
      const employeeId = row[1] || '';
      const employeeName = row[2] || 'Unknown';
      const location = row[3] || 'Unknown';
      const hoursRequested = parseFloat(row[4]) || 0;
      const ptoStartDate = row[5] ? new Date(row[5]) : null;
      const submissionDate = row[9] ? new Date(row[9]) : null;
      const paidOut = row[10] === true || row[10] === 'TRUE';
      const status = row[11] || 'Pending';
      
      // Skip cancelled/denied
      if (status === 'Cancelled' || status === 'Denied') return;
      
      totalRequests++;
      totalHours += hoursRequested;
      uniqueEmployees.add(employeeId || employeeName);
      
      if (paidOut || status === 'Paid') {
        paidRequests++;
        paidHours += hoursRequested;
      } else {
        pendingRequests++;
        pendingHours += hoursRequested;
      }
      
      // Track by employee
      const empKey = employeeId || employeeName.toLowerCase();
      if (!employeeMap.has(empKey)) {
        employeeMap.set(empKey, {
          employeeId: employeeId,
          employeeName: employeeName,
          location: location,
          totalHours: 0,
          requestCount: 0
        });
      }
      const empData = employeeMap.get(empKey);
      empData.totalHours += hoursRequested;
      empData.requestCount++;
      
      // Track by location
      if (!locationMap.has(location)) {
        locationMap.set(location, { location: location, hours: 0, requests: 0 });
      }
      locationMap.get(location).hours += hoursRequested;
      locationMap.get(location).requests++;
      
      // Track by month (last 6 months)
      if (submissionDate) {
        const monthKey = submissionDate.toISOString().slice(0, 7); // YYYY-MM
        if (!monthlyData[monthKey]) {
          monthlyData[monthKey] = { month: monthKey, hours: 0, requests: 0 };
        }
        monthlyData[monthKey].hours += hoursRequested;
        monthlyData[monthKey].requests++;
      }
      
      // Collect for recent activity (with row data)
      if (submissionDate) {
        recentRecords.push({
          ptoId: ptoId,
          employeeName: employeeName,
          location: location,
          hours: hoursRequested,
          status: paidOut ? 'Paid' : status,
          submissionDate: submissionDate.toISOString(),
          ptoStartDate: ptoStartDate ? ptoStartDate.toISOString().split('T')[0] : null
        });
      }
    });
    
    // Sort and limit top employees
    const topEmployees = Array.from(employeeMap.values())
      .sort((a, b) => b.totalHours - a.totalHours)
      .slice(0, 10)
      .map((emp, index) => ({
        rank: index + 1,
        ...emp,
        totalHours: Math.round(emp.totalHours * 10) / 10
      }));
    
    // Sort by location
    const byLocation = Array.from(locationMap.values())
      .sort((a, b) => b.hours - a.hours)
      .map(loc => ({
        ...loc,
        hours: Math.round(loc.hours * 10) / 10
      }));
    
    // Sort recent activity (newest first), limit to 15
    const recentActivity = recentRecords
      .sort((a, b) => new Date(b.submissionDate) - new Date(a.submissionDate))
      .slice(0, 15)
      .map(rec => ({
        ...rec,
        submissionDate: new Date(rec.submissionDate).toLocaleDateString('en-US', {
          month: 'short', day: 'numeric', year: 'numeric'
        })
      }));
    
    // Get last 6 months of data
    const sortedMonths = Object.values(monthlyData)
      .sort((a, b) => a.month.localeCompare(b.month))
      .slice(-6)
      .map(m => ({
        month: formatMonthDisplay(m.month),
        hours: Math.round(m.hours * 10) / 10,
        requests: m.requests
      }));
    
    return {
      success: true,
      summary: {
        totalRequests: totalRequests,
        totalHours: Math.round(totalHours * 10) / 10,
        pendingRequests: pendingRequests,
        pendingHours: Math.round(pendingHours * 10) / 10,
        paidRequests: paidRequests,
        paidHours: Math.round(paidHours * 10) / 10,
        averageHoursPerRequest: totalRequests > 0 ? Math.round((totalHours / totalRequests) * 10) / 10 : 0,
        uniqueEmployees: uniqueEmployees.size
      },
      byLocation: byLocation,
      topEmployees: topEmployees,
      recentActivity: recentActivity,
      byMonth: sortedMonths
    };
    
  } catch (error) {
    console.error('Error getting PTO summary data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Helper to format month string (YYYY-MM) to display format
 */
function formatMonthDisplay(monthStr) {
  const [year, month] = monthStr.split('-');
  const date = new Date(parseInt(year), parseInt(month) - 1, 1);
  return date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
}

