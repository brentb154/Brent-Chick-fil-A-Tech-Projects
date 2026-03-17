// ============================================
// PHASE 19: PROBATION AND ACTION TRACKING
// ============================================
// Functions for tracking 30-day probation periods
// and required actions based on point thresholds
// ============================================

// Sheet names for tracking
const PROBATION_SHEET_NAME = 'Probation_Tracking';
const ACTION_SHEET_NAME = 'Action_Tracking';

// Probation constants
const PROBATION_THRESHOLD = 9;
const DEFAULT_PROBATION_DAYS = 30;
const MAX_EXTENSION_DAYS = 90;

// Required actions by threshold
const THRESHOLD_ACTIONS = {
  6: [
    { name: 'Director meeting required', description: 'Hold meeting with director to discuss performance' },
    { name: 'Remove day from schedule', description: 'Remove one day from work schedule' }
  ],
  9: [
    { name: 'Final written warning', description: 'Deliver formal written warning document' },
    { name: '30-day probation started', description: 'Begin 30-day probation monitoring period' },
    { name: 'Director meeting required', description: 'Hold meeting with director to discuss probation' }
  ],
  12: [
    { name: '3-day suspension', description: 'Suspend employee for 3 days without pay' },
    { name: 'Reduced hours', description: 'Reduce scheduled hours until improvement shown' }
  ]
};

// ============================================
// SHEET INITIALIZATION
// ============================================

/**
 * Ensures the Probation_Tracking and Action_Tracking sheets exist.
 * Creates them with proper headers if missing.
 */
function ensureProbationSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Check/create Probation_Tracking sheet
  let probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);
  if (!probationSheet) {
    probationSheet = ss.insertSheet(PROBATION_SHEET_NAME);
    probationSheet.getRange('A1:M1').setValues([[
      'Probation_ID',
      'Employee_ID',
      'Employee_Name',
      'Start_Date',
      'Original_End_Date',
      'Current_End_Date',
      'Status',
      'Qualifying_Infraction_ID',
      'Extended_By',
      'Extension_Reason',
      'Ended_Early_By',
      'Early_End_Reason',
      'Completion_Notes'
    ]]);
    probationSheet.getRange('A1:M1').setFontWeight('bold');
    probationSheet.setFrozenRows(1);
    console.log('Created Probation_Tracking sheet');
  }

  // Check/create Action_Tracking sheet
  let actionSheet = ss.getSheetByName(ACTION_SHEET_NAME);
  if (!actionSheet) {
    actionSheet = ss.insertSheet(ACTION_SHEET_NAME);
    actionSheet.getRange('A1:I1').setValues([[
      'Action_ID',
      'Employee_ID',
      'Employee_Name',
      'Action_Name',
      'Threshold_Level',
      'Completed_Date',
      'Completed_By',
      'Notes',
      'Timestamp'
    ]]);
    actionSheet.getRange('A1:I1').setFontWeight('bold');
    actionSheet.setFrozenRows(1);
    console.log('Created Action_Tracking sheet');
  }

  return { probationSheet, actionSheet };
}

// ============================================
// PROBATION STATUS FUNCTIONS
// ============================================

/**
 * Checks the probation status for an employee.
 *
 * @param {string} employeeId - Employee ID to check
 * @returns {Object} Probation status information
 */
function checkProbationStatus(employeeId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);

    if (!probationSheet) {
      return {
        is_on_probation: false,
        message: 'Probation tracking not initialized'
      };
    }

    const data = probationSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indices
    const colIndex = {
      probation_id: headers.indexOf('Probation_ID'),
      employee_id: headers.indexOf('Employee_ID'),
      employee_name: headers.indexOf('Employee_Name'),
      start_date: headers.indexOf('Start_Date'),
      original_end_date: headers.indexOf('Original_End_Date'),
      current_end_date: headers.indexOf('Current_End_Date'),
      status: headers.indexOf('Status'),
      qualifying_infraction_id: headers.indexOf('Qualifying_Infraction_ID'),
      extended_by: headers.indexOf('Extended_By'),
      extension_reason: headers.indexOf('Extension_Reason'),
      ended_early_by: headers.indexOf('Ended_Early_By'),
      early_end_reason: headers.indexOf('Early_End_Reason')
    };

    // Find active probation for this employee
    let activeProbation = null;
    let probationHistory = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[colIndex.employee_id] === employeeId) {
        const probationRecord = {
          probation_id: row[colIndex.probation_id],
          start_date: row[colIndex.start_date],
          original_end_date: row[colIndex.original_end_date],
          current_end_date: row[colIndex.current_end_date],
          status: row[colIndex.status],
          qualifying_infraction_id: row[colIndex.qualifying_infraction_id],
          extended_by: row[colIndex.extended_by],
          extension_reason: row[colIndex.extension_reason],
          ended_early_by: row[colIndex.ended_early_by],
          early_end_reason: row[colIndex.early_end_reason],
          row_number: i + 1
        };

        if (row[colIndex.status] === 'Active') {
          activeProbation = probationRecord;
        }
        probationHistory.push(probationRecord);
      }
    }

    // Get current point total directly (avoid circular call to getEmployeeDetailData)
    let currentPoints = 0;
    try {
      const pointData = calculatePoints(employeeId);
      currentPoints = pointData.total_points || 0;
    } catch (e) {
      console.log('Could not calculate points for probation check:', e.toString());
    }

    if (activeProbation) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const endDate = new Date(activeProbation.current_end_date);
      endDate.setHours(0, 0, 0, 0);

      const startDate = new Date(activeProbation.start_date);
      startDate.setHours(0, 0, 0, 0);

      const daysRemaining = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));
      const totalDays = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24));
      const daysElapsed = Math.ceil((today - startDate) / (1000 * 60 * 60 * 24));
      const progressPercent = Math.min(100, Math.round((daysElapsed / totalDays) * 100));

      return {
        is_on_probation: true,
        probation_id: activeProbation.probation_id,
        probation_start_date: activeProbation.start_date,
        probation_end_date: activeProbation.current_end_date,
        original_end_date: activeProbation.original_end_date,
        days_remaining: Math.max(0, daysRemaining),
        days_elapsed: daysElapsed,
        total_days: totalDays,
        progress_percent: progressPercent,
        qualifying_infraction_id: activeProbation.qualifying_infraction_id,
        can_end_early: currentPoints < PROBATION_THRESHOLD,
        current_points: currentPoints,
        was_extended: activeProbation.extended_by ? true : false,
        extended_by: activeProbation.extended_by,
        extension_reason: activeProbation.extension_reason,
        row_number: activeProbation.row_number,
        probation_history: probationHistory
      };
    }

    // Not on probation - return history if any
    return {
      is_on_probation: false,
      current_points: currentPoints,
      probation_history: probationHistory,
      last_probation: probationHistory.length > 0 ? probationHistory[probationHistory.length - 1] : null
    };

  } catch (error) {
    console.error('Error checking probation status:', error.toString());
    return {
      is_on_probation: false,
      error: error.message
    };
  }
}

/**
 * Creates a new probation record when employee crosses 9-point threshold.
 * Called automatically when an infraction pushes employee to/past 9 points.
 *
 * @param {string} employeeId - Employee ID
 * @param {string} employeeName - Employee full name
 * @param {string} qualifyingInfractionId - The infraction that caused threshold crossing
 * @returns {Object} Result with probation details
 */
function createProbationRecord(employeeId, employeeName, qualifyingInfractionId) {
  try {
    ensureProbationSheets();

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);

    // Check if already on active probation
    const currentStatus = checkProbationStatus(employeeId);
    if (currentStatus.is_on_probation) {
      return {
        success: false,
        error: 'Employee is already on active probation',
        existing_probation: currentStatus
      };
    }

    // Generate probation ID
    const probationId = 'PROB-' + Date.now().toString(36).toUpperCase();

    // Calculate dates
    const startDate = new Date();
    const endDate = new Date();
    endDate.setDate(endDate.getDate() + DEFAULT_PROBATION_DAYS);

    // Create probation record
    const newRow = [
      probationId,
      employeeId,
      employeeName,
      startDate,
      endDate,
      endDate,  // Current_End_Date = Original_End_Date initially
      'Active',
      qualifyingInfractionId,
      '',  // Extended_By
      '',  // Extension_Reason
      '',  // Ended_Early_By
      '',  // Early_End_Reason
      ''   // Completion_Notes
    ];

    probationSheet.appendRow(newRow);

    // Send notification email
    sendProbationStartEmail(employeeId, employeeName, startDate, endDate);

    // Log the action
    logToAuditTrail('Probation Started', employeeId, {
      probation_id: probationId,
      start_date: startDate.toISOString(),
      end_date: endDate.toISOString(),
      qualifying_infraction: qualifyingInfractionId
    });

    console.log(`Probation created for ${employeeName}: ${probationId}`);

    return {
      success: true,
      probation_id: probationId,
      start_date: startDate,
      end_date: endDate,
      message: `30-day probation started for ${employeeName}`
    };

  } catch (error) {
    console.error('Error creating probation record:', error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Extends an employee's probation period.
 *
 * @param {string} employeeId - Employee ID
 * @param {number} extensionDays - Number of days to extend
 * @param {string} extendedByDirector - Director making the extension
 * @param {string} reason - Reason for extension
 * @returns {Object} Result with new end date
 */
function extendProbation(employeeId, extensionDays, extendedByDirector, reason) {
  try {
    // Validate session
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Only Directors and Operators can extend probation' };
    }

    // Get current user's email if not provided
    if (!extendedByDirector) {
      extendedByDirector = Session.getActiveUser().getEmail() || session.email || 'Unknown';
    }

    // Validate inputs
    extensionDays = parseInt(extensionDays);
    if (isNaN(extensionDays) || extensionDays < 1 || extensionDays > MAX_EXTENSION_DAYS) {
      return { success: false, error: `Extension must be between 1 and ${MAX_EXTENSION_DAYS} days` };
    }

    if (!reason || reason.trim() === '') {
      return { success: false, error: 'A reason for extension is required' };
    }

    // Check current probation status
    const status = checkProbationStatus(employeeId);
    if (!status.is_on_probation) {
      return { success: false, error: 'Employee is not currently on probation' };
    }

    // Update probation record
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);
    const rowNumber = status.row_number;

    // Calculate new end date
    const currentEndDate = new Date(status.probation_end_date);
    const newEndDate = new Date(currentEndDate);
    newEndDate.setDate(newEndDate.getDate() + extensionDays);

    // Update the row
    probationSheet.getRange(rowNumber, 6).setValue(newEndDate);  // Current_End_Date
    probationSheet.getRange(rowNumber, 9).setValue(extendedByDirector);  // Extended_By
    probationSheet.getRange(rowNumber, 10).setValue(reason);  // Extension_Reason

    // Get employee name for notification
    const employeeData = getEmployeeDetailData(employeeId);
    const employeeName = employeeData.success ? employeeData.employee.full_name : employeeId;

    // Send notification
    sendProbationExtendedEmail(employeeId, employeeName, extensionDays, newEndDate, extendedByDirector, reason);

    // Log the action
    logToAuditTrail('Probation Extended', employeeId, {
      probation_id: status.probation_id,
      extension_days: extensionDays,
      new_end_date: newEndDate.toISOString(),
      extended_by: extendedByDirector,
      reason: reason
    });

    return {
      success: true,
      message: `Probation extended by ${extensionDays} days`,
      new_end_date: newEndDate,
      previous_end_date: currentEndDate
    };

  } catch (error) {
    console.error('Error extending probation:', error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Ends probation early before the scheduled end date.
 *
 * @param {string} employeeId - Employee ID
 * @param {string} endedByDirector - Director ending the probation
 * @param {string} reason - Reason for early ending
 * @returns {Object} Result
 */
function endProbationEarly(employeeId, endedByDirector, reason) {
  try {
    // Validate session
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Only Directors and Operators can end probation early' };
    }

    // Get current user's email if not provided
    if (!endedByDirector) {
      endedByDirector = Session.getActiveUser().getEmail() || session.email || 'Unknown';
    }

    // Validate inputs
    if (!reason || reason.trim() === '') {
      return { success: false, error: 'A reason for ending probation early is required' };
    }

    // Check current probation status
    const status = checkProbationStatus(employeeId);
    if (!status.is_on_probation) {
      return { success: false, error: 'Employee is not currently on probation' };
    }

    // Update probation record
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);
    const rowNumber = status.row_number;

    // Update the row
    probationSheet.getRange(rowNumber, 7).setValue('Ended Early');  // Status
    probationSheet.getRange(rowNumber, 11).setValue(endedByDirector);  // Ended_Early_By
    probationSheet.getRange(rowNumber, 12).setValue(reason);  // Early_End_Reason

    // Get employee name for notification
    const employeeData = getEmployeeDetailData(employeeId);
    const employeeName = employeeData.success ? employeeData.employee.full_name : employeeId;

    // Send notification
    sendProbationEndedEarlyEmail(employeeId, employeeName, endedByDirector, reason);

    // Log the action
    logToAuditTrail('Probation Ended Early', employeeId, {
      probation_id: status.probation_id,
      ended_by: endedByDirector,
      reason: reason,
      days_remaining: status.days_remaining
    });

    return {
      success: true,
      message: 'Probation ended early',
      ended_by: endedByDirector
    };

  } catch (error) {
    console.error('Error ending probation early:', error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Completes a probation that has reached its end date.
 * Called by daily trigger or manually.
 *
 * @param {string} employeeId - Employee ID
 * @param {string} notes - Optional completion notes
 * @returns {Object} Result
 */
function completeProbation(employeeId, notes) {
  try {
    const status = checkProbationStatus(employeeId);
    if (!status.is_on_probation) {
      return { success: false, error: 'Employee is not on probation' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);
    const rowNumber = status.row_number;

    // Update status to Completed
    probationSheet.getRange(rowNumber, 7).setValue('Completed');  // Status
    probationSheet.getRange(rowNumber, 13).setValue(notes || 'Probation period completed');  // Completion_Notes

    // Get employee name
    const employeeData = getEmployeeDetailData(employeeId);
    const employeeName = employeeData.success ? employeeData.employee.full_name : employeeId;

    // Send notification
    sendProbationCompletedEmail(employeeId, employeeName);

    // Log
    logToAuditTrail('Probation Completed', employeeId, {
      probation_id: status.probation_id,
      notes: notes
    });

    return {
      success: true,
      message: 'Probation completed'
    };

  } catch (error) {
    console.error('Error completing probation:', error.toString());
    return { success: false, error: error.message };
  }
}

// ============================================
// ACTION TRACKING FUNCTIONS
// ============================================

/**
 * Gets required actions for an employee based on their point thresholds.
 *
 * @param {string} employeeId - Employee ID
 * @returns {Object} List of required and completed actions
 */
function getRequiredActions(employeeId) {
  try {
    // Get employee's current points directly (avoid circular call to getEmployeeDetailData)
    const pointData = calculatePoints(employeeId);
    const currentPoints = pointData.total_points || 0;

    // Get employee name from active employees list
    const employees = getActiveEmployees();
    const employee = employees.find(emp => emp.employee_id === employeeId);
    const employeeName = employee ? employee.full_name : 'Unknown';

    // Get completed actions for this employee
    const completedActions = getCompletedActionsForEmployee(employeeId);

    // Build list of required actions based on thresholds reached
    const requiredActions = [];

    for (const [threshold, actions] of Object.entries(THRESHOLD_ACTIONS)) {
      const thresholdNum = parseInt(threshold);

      if (currentPoints >= thresholdNum) {
        for (const action of actions) {
          // Check if this action was already completed
          const completed = completedActions.find(
            ca => ca.action_name === action.name && ca.threshold_level === thresholdNum
          );

          requiredActions.push({
            action_name: action.name,
            description: action.description,
            required_by_threshold: thresholdNum,
            completed: completed ? true : false,
            completed_date: completed ? completed.completed_date : null,
            completed_by: completed ? completed.completed_by : null,
            notes: completed ? completed.notes : null,
            action_id: completed ? completed.action_id : null
          });
        }
      }
    }

    // Sort by threshold level, then by completion status
    requiredActions.sort((a, b) => {
      if (a.required_by_threshold !== b.required_by_threshold) {
        return a.required_by_threshold - b.required_by_threshold;
      }
      return a.completed === b.completed ? 0 : (a.completed ? 1 : -1);
    });

    return {
      success: true,
      employee_id: employeeId,
      employee_name: employeeName,
      current_points: currentPoints,
      actions: requiredActions,
      total_required: requiredActions.length,
      total_completed: requiredActions.filter(a => a.completed).length,
      total_pending: requiredActions.filter(a => !a.completed).length
    };

  } catch (error) {
    console.error('Error getting required actions:', error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Gets all completed actions for an employee from Action_Tracking sheet.
 *
 * @param {string} employeeId - Employee ID
 * @returns {Array} List of completed action records
 */
function getCompletedActionsForEmployee(employeeId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const actionSheet = ss.getSheetByName(ACTION_SHEET_NAME);

    if (!actionSheet) {
      return [];
    }

    const data = actionSheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const colIndex = {
      action_id: headers.indexOf('Action_ID'),
      employee_id: headers.indexOf('Employee_ID'),
      action_name: headers.indexOf('Action_Name'),
      threshold_level: headers.indexOf('Threshold_Level'),
      completed_date: headers.indexOf('Completed_Date'),
      completed_by: headers.indexOf('Completed_By'),
      notes: headers.indexOf('Notes')
    };

    const completedActions = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[colIndex.employee_id] === employeeId) {
        completedActions.push({
          action_id: row[colIndex.action_id],
          action_name: row[colIndex.action_name],
          threshold_level: row[colIndex.threshold_level],
          completed_date: row[colIndex.completed_date],
          completed_by: row[colIndex.completed_by],
          notes: row[colIndex.notes]
        });
      }
    }

    return completedActions;

  } catch (error) {
    console.error('Error getting completed actions:', error.toString());
    return [];
  }
}

/**
 * Logs that a required action has been completed.
 *
 * @param {string} employeeId - Employee ID
 * @param {string} actionName - Name of the action completed
 * @param {string} completedByDirector - Director who completed/logged the action
 * @param {Date} completionDate - Date the action was completed
 * @param {string} notes - Optional notes about completion
 * @returns {Object} Result
 */
function logActionCompleted(employeeId, actionName, completedByDirector, completionDate, notes) {
  try {
    // Validate session
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Only Directors and Operators can log action completion' };
    }

    // Get current user's email if not provided
    if (!completedByDirector) {
      completedByDirector = Session.getActiveUser().getEmail() || session.email || 'Unknown';
    }

    // Validate completion date
    const compDate = new Date(completionDate);
    const today = new Date();
    today.setHours(23, 59, 59, 999);

    if (compDate > today) {
      return { success: false, error: 'Completion date cannot be in the future' };
    }

    // Get employee name
    const employeeData = getEmployeeDetailData(employeeId);
    if (!employeeData.success) {
      return { success: false, error: 'Employee not found' };
    }
    const employeeName = employeeData.employee.full_name;

    // Determine threshold level for this action
    let thresholdLevel = 0;
    for (const [threshold, actions] of Object.entries(THRESHOLD_ACTIONS)) {
      if (actions.some(a => a.name === actionName)) {
        thresholdLevel = parseInt(threshold);
        break;
      }
    }

    ensureProbationSheets();

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const actionSheet = ss.getSheetByName(ACTION_SHEET_NAME);

    // Check if action already logged for this employee at this threshold
    const existingActions = getCompletedActionsForEmployee(employeeId);
    const existing = existingActions.find(
      a => a.action_name === actionName && a.threshold_level === thresholdLevel
    );

    if (existing) {
      // Update existing record
      const data = actionSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === existing.action_id) {
          actionSheet.getRange(i + 1, 6).setValue(compDate);  // Completed_Date
          actionSheet.getRange(i + 1, 7).setValue(completedByDirector);  // Completed_By
          actionSheet.getRange(i + 1, 8).setValue(notes || '');  // Notes
          actionSheet.getRange(i + 1, 9).setValue(new Date());  // Timestamp
          break;
        }
      }

      return {
        success: true,
        message: 'Action record updated',
        action_id: existing.action_id,
        updated: true
      };
    }

    // Create new action record
    const actionId = 'ACT-' + Date.now().toString(36).toUpperCase();

    const newRow = [
      actionId,
      employeeId,
      employeeName,
      actionName,
      thresholdLevel,
      compDate,
      completedByDirector,
      notes || '',
      new Date()
    ];

    actionSheet.appendRow(newRow);

    // Log to audit trail
    logToAuditTrail('Action Completed', employeeId, {
      action_id: actionId,
      action_name: actionName,
      threshold_level: thresholdLevel,
      completed_by: completedByDirector
    });

    return {
      success: true,
      message: 'Action logged successfully',
      action_id: actionId,
      updated: false
    };

  } catch (error) {
    console.error('Error logging action completed:', error.toString());
    return { success: false, error: error.message };
  }
}

// ============================================
// AUTOMATIC PROBATION TRIGGER
// ============================================

/**
 * Checks if an infraction crosses the 9-point threshold and creates probation.
 * Should be called after adding an infraction.
 *
 * @param {string} employeeId - Employee ID
 * @param {string} employeeName - Employee name
 * @param {number} pointsBefore - Points before this infraction
 * @param {number} pointsAfter - Points after this infraction
 * @param {string} infractionId - ID of the infraction just added
 * @returns {Object} Result indicating if probation was started
 */
function checkAndCreateProbation(employeeId, employeeName, pointsBefore, pointsAfter, infractionId) {
  try {
    // Check if threshold was crossed (from below 9 to 9 or above)
    if (pointsBefore < PROBATION_THRESHOLD && pointsAfter >= PROBATION_THRESHOLD) {
      // Check if already on probation
      const status = checkProbationStatus(employeeId);
      if (status.is_on_probation) {
        return {
          probation_started: false,
          reason: 'Already on probation'
        };
      }

      // Create new probation
      const result = createProbationRecord(employeeId, employeeName, infractionId);

      return {
        probation_started: result.success,
        probation_id: result.probation_id,
        start_date: result.start_date,
        end_date: result.end_date,
        message: result.message || result.error
      };
    }

    return {
      probation_started: false,
      reason: 'Threshold not crossed'
    };

  } catch (error) {
    console.error('Error in checkAndCreateProbation:', error.toString());
    return {
      probation_started: false,
      error: error.message
    };
  }
}

/**
 * Daily trigger function to check and complete expired probations.
 * Should be set up as a daily time-driven trigger.
 */
function dailyProbationCheck() {
  console.log('Running daily probation check...');

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);

    if (!probationSheet) {
      console.log('Probation sheet not found');
      return;
    }

    const data = probationSheet.getDataRange().getValues();
    const headers = data[0];

    const colIndex = {
      employee_id: headers.indexOf('Employee_ID'),
      employee_name: headers.indexOf('Employee_Name'),
      current_end_date: headers.indexOf('Current_End_Date'),
      status: headers.indexOf('Status')
    };

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let completedCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (row[colIndex.status] === 'Active') {
        const endDate = new Date(row[colIndex.current_end_date]);
        endDate.setHours(0, 0, 0, 0);

        if (endDate <= today) {
          // Probation has expired - mark as completed
          const employeeId = row[colIndex.employee_id];
          const employeeName = row[colIndex.employee_name];

          probationSheet.getRange(i + 1, 7).setValue('Completed');  // Status
          probationSheet.getRange(i + 1, 13).setValue('Automatically completed - probation period ended');

          // Send notification
          sendProbationCompletedEmail(employeeId, employeeName);

          console.log(`Probation completed for ${employeeName}`);
          completedCount++;
        }
      }
    }

    console.log(`Daily check complete. ${completedCount} probations completed.`);

  } catch (error) {
    console.error('Error in daily probation check:', error.toString());
  }
}

// ============================================
// EMAIL NOTIFICATIONS
// ============================================

/**
 * Sends email when probation starts.
 */
function sendProbationStartEmail(employeeId, employeeName, startDate, endDate) {
  try {
    const settings = getSystemSettings();
    const recipients = settings.email?.terminationEmailList;

    if (!recipients) {
      console.log('No termination email list configured');
      return;
    }
    const variables = {
      employee_name: employeeName,
      employee_id: employeeId,
      threshold: String(PROBATION_THRESHOLD),
      probation_start_date: formatTemplateDateValue(startDate),
      probation_end_date: formatTemplateDateValue(endDate),
      days_remaining: getDaysUntil(endDate)
    };

    sendTemplatedEmail('probation_started', recipients, variables);
    console.log(`Probation start email processed for ${employeeName}`);

  } catch (error) {
    console.error('Error sending probation start email:', error.toString());
  }
}

/**
 * Sends email when probation is extended.
 */
function sendProbationExtendedEmail(employeeId, employeeName, extensionDays, newEndDate, extendedBy, reason) {
  try {
    const settings = getSystemSettings();
    const recipients = settings.email?.terminationEmailList;

    if (!recipients) return;
    const variables = {
      employee_name: employeeName,
      employee_id: employeeId,
      probation_end_date: formatTemplateDateValue(newEndDate),
      days_remaining: getDaysUntil(newEndDate),
      infraction_description: reason,
      points_assigned: extensionDays
    };

    sendTemplatedEmail('probation_ended', recipients, variables);

  } catch (error) {
    console.error('Error sending probation extended email:', error.toString());
  }
}

/**
 * Sends email when probation is ended early.
 */
function sendProbationEndedEarlyEmail(employeeId, employeeName, endedBy, reason) {
  try {
    const settings = getSystemSettings();
    const recipients = settings.email?.terminationEmailList;

    if (!recipients) return;
    const variables = {
      employee_name: employeeName,
      employee_id: employeeId,
      probation_end_date: formatTemplateDateValue(new Date()),
      infraction_description: reason
    };

    sendTemplatedEmail('probation_ended', recipients, variables);

  } catch (error) {
    console.error('Error sending probation ended early email:', error.toString());
  }
}

/**
 * Sends email when probation is completed.
 */
function sendProbationCompletedEmail(employeeId, employeeName) {
  try {
    const settings = getSystemSettings();
    const recipients = settings.email?.terminationEmailList;

    if (!recipients) return;
    const variables = {
      employee_name: employeeName,
      employee_id: employeeId,
      probation_end_date: formatTemplateDateValue(new Date())
    };

    sendTemplatedEmail('probation_ended', recipients, variables);

  } catch (error) {
    console.error('Error sending probation completed email:', error.toString());
  }
}

/**
 * Helper to format date for email.
 */
function formatDateForEmail(date) {
  if (!date) return 'N/A';
  const d = new Date(date);
  return d.toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Gets all employees currently on probation.
 * Used for filtering in the employee list.
 *
 * @returns {Array} List of employee IDs on probation
 */
function getEmployeesOnProbation() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const probationSheet = ss.getSheetByName(PROBATION_SHEET_NAME);

    if (!probationSheet) {
      return [];
    }

    const data = probationSheet.getDataRange().getValues();
    const headers = data[0];

    const employeeIdCol = headers.indexOf('Employee_ID');
    const statusCol = headers.indexOf('Status');

    const onProbation = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][statusCol] === 'Active') {
        onProbation.push(data[i][employeeIdCol]);
      }
    }

    return onProbation;

  } catch (error) {
    console.error('Error getting employees on probation:', error.toString());
    return [];
  }
}

/**
 * Gets probation count for dashboard display.
 *
 * @returns {Object} Probation statistics
 */
function getProbationStats() {
  try {
    const onProbation = getEmployeesOnProbation();

    return {
      success: true,
      total_on_probation: onProbation.length,
      employee_ids: onProbation
    };

  } catch (error) {
    console.error('Error getting probation stats:', error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Logs action to audit trail if the function exists.
 */
function logToAuditTrail(action, employeeId, details) {
  try {
    // Check if there's an existing audit logging function
    if (typeof logAuditEntry === 'function') {
      logAuditEntry(action, employeeId, JSON.stringify(details));
    } else {
      console.log(`AUDIT: ${action} - Employee: ${employeeId} - Details: ${JSON.stringify(details)}`);
    }
  } catch (error) {
    console.log(`AUDIT (fallback): ${action} - Employee: ${employeeId}`);
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Comprehensive test function for probation tracking.
 */
function testProbationTracking() {
  console.log('=== Testing Probation Tracking ===');
  console.log('');

  const testResults = [];

  // Ensure sheets exist
  console.log('Setting up test environment...');
  ensureProbationSheets();

  // Get a test employee
  const employees = getActiveEmployees();
  if (!employees || employees.length === 0) {
    console.log('FAIL: No employees found for testing');
    return { success: false, message: 'No employees available' };
  }

  const testEmployee = employees[0];
  console.log(`Testing with employee: ${testEmployee.full_name} (${testEmployee.employee_id})`);
  console.log('');

  // Test Case 1: Check probation status
  console.log('Test Case 1: Check probation status');
  const status1 = checkProbationStatus(testEmployee.employee_id);
  console.log(`  Is on probation: ${status1.is_on_probation}`);
  console.log(`  Current points: ${status1.current_points}`);
  testResults.push({ test: 'Check Probation Status', passed: status1.current_points !== undefined });
  console.log('');

  // Test Case 2: Get required actions
  console.log('Test Case 2: Get required actions');
  const actions = getRequiredActions(testEmployee.employee_id);
  console.log(`  Success: ${actions.success}`);
  console.log(`  Total required: ${actions.total_required}`);
  console.log(`  Total completed: ${actions.total_completed}`);
  if (actions.actions && actions.actions.length > 0) {
    console.log(`  First action: ${actions.actions[0].action_name}`);
  }
  testResults.push({ test: 'Get Required Actions', passed: actions.success });
  console.log('');

  // Test Case 3: Get employees on probation
  console.log('Test Case 3: Get employees on probation');
  const onProbation = getEmployeesOnProbation();
  console.log(`  Count: ${onProbation.length}`);
  testResults.push({ test: 'Get Employees On Probation', passed: Array.isArray(onProbation) });
  console.log('');

  // Test Case 4: Get probation stats
  console.log('Test Case 4: Get probation stats');
  const stats = getProbationStats();
  console.log(`  Success: ${stats.success}`);
  console.log(`  Total on probation: ${stats.total_on_probation}`);
  testResults.push({ test: 'Get Probation Stats', passed: stats.success });
  console.log('');

  // Test Case 5: Ensure sheets created
  console.log('Test Case 5: Verify sheets exist');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const probSheet = ss.getSheetByName(PROBATION_SHEET_NAME);
  const actSheet = ss.getSheetByName(ACTION_SHEET_NAME);
  console.log(`  Probation_Tracking exists: ${probSheet !== null}`);
  console.log(`  Action_Tracking exists: ${actSheet !== null}`);
  testResults.push({ test: 'Sheets Exist', passed: probSheet !== null && actSheet !== null });
  console.log('');

  // Summary
  console.log('=== Test Summary ===');
  const passed = testResults.filter(r => r.passed).length;
  const failed = testResults.filter(r => !r.passed).length;
  console.log(`Passed: ${passed}/${testResults.length}`);
  console.log(`Failed: ${failed}/${testResults.length}`);

  for (const result of testResults) {
    console.log(`  ${result.passed ? '✓' : '✗'} ${result.test}`);
  }

  return {
    success: failed === 0,
    message: `${passed}/${testResults.length} tests passed`,
    results: testResults
  };
}

/**
 * Test creating a probation record manually.
 */
function testCreateProbation() {
  const employees = getActiveEmployees();
  if (!employees || employees.length === 0) {
    console.log('No employees found');
    return;
  }

  const testEmployee = employees[0];
  console.log(`Creating test probation for: ${testEmployee.full_name}`);

  const result = createProbationRecord(
    testEmployee.employee_id,
    testEmployee.full_name,
    'TEST-INFRACTION-001'
  );

  console.log('Result:', JSON.stringify(result, null, 2));
}
