// ============================================
// MICRO-PHASE 33: DATA VALIDATION & CLEANUP
// ============================================

const DATA_QUALITY_SHEET = 'Data_Quality_Issues';
const DATA_QUALITY_HEADERS = [
  'Issue_ID',
  'Issue_Type',
  'Severity',
  'Description',
  'Detected_Date',
  'Affected_Record_ID',
  'Status',
  'Fixed_Date',
  'Fixed_By',
  'Fix_Action_Taken'
];

// =========================
// Helpers
// =========================
function getOrCreateDataQualityIssuesSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(DATA_QUALITY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_QUALITY_SHEET);
    sheet.getRange(1, 1, 1, DATA_QUALITY_HEADERS.length).setValues([DATA_QUALITY_HEADERS]);
    sheet.getRange(1, 1, 1, DATA_QUALITY_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateArchivedSignupsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Archived_Signups');
  if (!sheet) {
    sheet = ss.insertSheet('Archived_Signups');
    const pendingSheet = ss.getSheetByName('Pending_Signups');
    const headers = pendingSheet
      ? pendingSheet.getRange(1, 1, 1, pendingSheet.getLastColumn()).getValues()[0]
      : ['Signup_ID', 'Token', 'Employee_ID', 'Email', 'Role', 'Can_See_Directors', 'Created_Date', 'Expires_Date', 'Status', 'Created_By', 'Completed_Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateDataQualityIssueId_() {
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const rand = Math.random().toString(36).substr(2, 4).toUpperCase();
  return `DQ-${dateStr}-${rand}`;
}

function normalizeHeader_(header) {
  return String(header || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, idx) => {
    map[normalizeHeader_(header)] = idx;
  });
  return map;
}

function getSheetData_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return { headers: [], headerMap: {}, rows: [] };
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  return { headers: headers, headerMap: getHeaderMap_(headers), rows: data.slice(1) };
}

function parseDate_(value) {
  if (value instanceof Date) return value;
  if (!value) return null;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateKey_(value) {
  const date = parseDate_(value);
  if (!date) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function addIssue_(issues, keys, issue) {
  const key = `${issue.issue_type}|${issue.affected_record_id}|${issue.description}`;
  if (keys.has(key)) return;
  keys.add(key);
  issues.push(issue);
}

function requireOperatorSession_(token) {
  if (!token) return { valid: false, error: 'Missing session token' };
  const session = validateSessionToken(token);
  if (!session || !session.valid) return { valid: false, sessionExpired: true };
  if (session.role !== 'Operator') return { valid: false, error: 'Operator access required' };
  return { valid: true, session: session };
}

function getPayrollEmployeesMap_() {
  const result = { byId: {}, byIdLower: {}, allIds: new Set() };
  try {
    const payrollSpreadsheet = SpreadsheetApp.openById(PAYROLL_TRACKER_ID);
    const employeesSheet = payrollSpreadsheet.getSheetByName(PAYROLL_TAB_NAME);
    if (!employeesSheet || employeesSheet.getLastRow() < 2) return result;
    const data = employeesSheet.getRange(2, 1, employeesSheet.getLastRow() - 1, 8).getValues();
    data.forEach(row => {
      const employeeId = row[0];
      if (!employeeId) return;
      result.byId[employeeId] = {
        employee_id: row[0],
        full_name: row[1],
        first_seen: row[5]
      };
      result.byIdLower[String(employeeId).toLowerCase()] = result.byId[employeeId];
      result.allIds.add(String(employeeId));
    });
  } catch (error) {
    console.error('Error loading payroll tracker:', error.toString());
  }
  return result;
}

function findIssueRow_(issueId, sheetData) {
  if (!issueId || !sheetData) return null;
  for (let i = 0; i < sheetData.rows.length; i++) {
    if (sheetData.rows[i][0] === issueId) {
      return i + 2; // account for header row
    }
  }
  return null;
}

function updateIssueRow_(sheet, rowIndex, updates) {
  if (!sheet || !rowIndex) return;
  const data = sheet.getRange(1, 1, 1, DATA_QUALITY_HEADERS.length).getValues();
  const headers = data[0] || [];
  const headerMap = getHeaderMap_(headers);
  const current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  Object.keys(updates).forEach(key => {
    const idx = headerMap[normalizeHeader_(key)];
    if (typeof idx === 'number') current[idx] = updates[key];
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([current]);
}

function writeIssuesToSheet_(issues) {
  const sheet = getOrCreateDataQualityIssuesSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || DATA_QUALITY_HEADERS;
  const headerMap = getHeaderMap_(headers);

  const existingRows = data.length > 1 ? data.slice(1) : [];
  const retainedRows = existingRows.filter(row => {
    const status = row[headerMap.status] || 'Open';
    return status !== 'Open';
  });

  const newRows = issues.map(issue => ([
    issue.issue_id,
    issue.issue_type,
    issue.severity,
    issue.description,
    issue.detected_date,
    issue.affected_record_id,
    issue.status || 'Open',
    issue.fixed_date || '',
    issue.fixed_by || '',
    issue.fix_action_taken || ''
  ]));

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }

  const allRows = retainedRows.concat(newRows);
  if (allRows.length) {
    sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
  }
}

function buildIssueSummary_(issues) {
  const summary = { total: issues.length, bySeverity: {}, byType: {} };
  issues.forEach(issue => {
    summary.bySeverity[issue.severity] = (summary.bySeverity[issue.severity] || 0) + 1;
    summary.byType[issue.issue_type] = (summary.byType[issue.issue_type] || 0) + 1;
  });
  return summary;
}

// =========================
// Core Functions
// =========================

/**
 * Scan entire system for data quality issues.
 */
function scanDataQuality(token) {
  if (token) {
    const session = requireOperatorSession_(token);
    if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const infractionsSheet = ss.getSheetByName('Infractions');
  const permissionsSheet = ss.getSheetByName('User_Permissions');
  const pendingSheet = ss.getSheetByName('Pending_Signups');
  const payroll = getPayrollEmployeesMap_();

  const issues = [];
  const keys = new Set();
  const detectedDate = new Date();

  const infractionsData = getSheetData_(infractionsSheet);
  const infHeaders = infractionsData.headerMap;
  const infRows = infractionsData.rows || [];

  const employeeIdsWithNameVariants = {};
  const infractionGroups = {};
  const infractionsByEmployeeDate = {};
  const infractionsByEmployee = {};
  const enteredBySet = new Set();

  const statusIdx = infHeaders.status;
  const employeeIdx = infHeaders.employeeid;
  const nameIdx = infHeaders.fullname;
  const dateIdx = infHeaders.date;
  const typeIdx = infHeaders.infractiontype;
  const pointsIdx = infHeaders.pointsassigned;
  const enteredByIdx = infHeaders.enteredby;
  const infractionIdIdx = infHeaders.infractionid;

  infRows.forEach(row => {
    const status = statusIdx != null ? row[statusIdx] : '';
    if (String(status).toLowerCase() === 'deleted') return;

    const employeeId = employeeIdx != null ? row[employeeIdx] : '';
    const infractionId = infractionIdIdx != null ? row[infractionIdIdx] : '';
    const infractionType = typeIdx != null ? row[typeIdx] : '';
    const infractionDate = dateIdx != null ? row[dateIdx] : '';
    const fullName = nameIdx != null ? row[nameIdx] : '';
    const enteredBy = enteredByIdx != null ? row[enteredByIdx] : '';
    const pointsAssigned = pointsIdx != null ? Number(row[pointsIdx]) || 0 : 0;

    if (employeeId && !payroll.byId[employeeId] && !payroll.byIdLower[String(employeeId).toLowerCase()]) {
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Missing_Employee',
        severity: 'High',
        description: `Infraction ${infractionId} references employee ${employeeId} who is not in Payroll Tracker.`,
        detected_date: detectedDate,
        affected_record_id: infractionId
      });
    }

    const dateKey = formatDateKey_(infractionDate);
    if (employeeId && infractionType && dateKey) {
      const key = `${employeeId}|${infractionType}|${dateKey}`;
      if (!infractionGroups[key]) infractionGroups[key] = [];
      infractionGroups[key].push({ infractionId: infractionId, employeeId: employeeId, name: fullName, date: dateKey });
    }

    if (employeeId && dateKey) {
      const countKey = `${employeeId}|${dateKey}`;
      infractionsByEmployeeDate[countKey] = (infractionsByEmployeeDate[countKey] || 0) + 1;
    }

    if (employeeId) {
      if (!infractionsByEmployee[employeeId]) infractionsByEmployee[employeeId] = [];
      infractionsByEmployee[employeeId].push({ date: infractionDate, points: pointsAssigned, infractionId: infractionId });

      if (fullName) {
        if (!employeeIdsWithNameVariants[employeeId]) employeeIdsWithNameVariants[employeeId] = new Set();
        employeeIdsWithNameVariants[employeeId].add(String(fullName));
      }
    }

    if (enteredBy) {
      enteredBySet.add(enteredBy);
    }

    if (pointsAssigned >= 15) {
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Anomaly_Point_Jump',
        severity: 'Medium',
        description: `Infraction ${infractionId} assigned ${pointsAssigned} points in one entry.`,
        detected_date: detectedDate,
        affected_record_id: infractionId
      });
    }

    const payrollRecord = payroll.byId[employeeId] || payroll.byIdLower[String(employeeId || '').toLowerCase()];
    if (payrollRecord && payrollRecord.first_seen) {
      const firstSeen = parseDate_(payrollRecord.first_seen);
      const infDate = parseDate_(infractionDate);
      if (firstSeen && infDate && infDate < firstSeen) {
        addIssue_(issues, keys, {
          issue_id: generateDataQualityIssueId_(),
          issue_type: 'Anomaly_PreHire_Infraction',
          severity: 'Medium',
          description: `Infraction ${infractionId} is dated before employee first seen date.`,
          detected_date: detectedDate,
          affected_record_id: infractionId
        });
      }
    }
  });

  // Duplicate infractions
  Object.keys(infractionGroups).forEach(key => {
    const group = infractionGroups[key];
    if (group.length > 1) {
      const ids = group.map(item => item.infractionId).filter(Boolean);
      const sample = group[0];
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Duplicate_Infraction',
        severity: 'Medium',
        description: `Employee ${sample.name || sample.employeeId} has ${group.length} identical infractions on ${sample.date}.`,
        detected_date: detectedDate,
        affected_record_id: ids.join(', ')
      });
    }
  });

  // Invalid point calculations
  try {
    const employees = getActiveEmployees();
    const ninetyDaysAgo = new Date();
    ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
    employees.forEach(employee => {
      const items = infractionsByEmployee[employee.employee_id] || [];
      let manualTotal = 0;
      items.forEach(item => {
        const infDate = parseDate_(item.date);
        if (infDate && infDate >= ninetyDaysAgo) {
          manualTotal += Number(item.points) || 0;
        }
      });
      const calculated = calculatePoints(employee.employee_id);
      const expected = calculated && typeof calculated.total_points !== 'undefined'
        ? Number(calculated.total_points) || 0
        : manualTotal;
      if (Math.abs(manualTotal - expected) > 0.01) {
        addIssue_(issues, keys, {
          issue_id: generateDataQualityIssueId_(),
          issue_type: 'Invalid_Points',
          severity: 'Critical',
          description: `Point mismatch for ${employee.full_name}: Calculated ${manualTotal}, Expected ${expected}.`,
          detected_date: detectedDate,
          affected_record_id: employee.employee_id
        });
      }
    });
  } catch (error) {
    console.error('Error checking point totals:', error.toString());
  }

  // Orphaned user permission records
  const permissionsData = getSheetData_(permissionsSheet);
  const permHeaders = permissionsData.headerMap;
  (permissionsData.rows || []).forEach(row => {
    const employeeId = permHeaders.employeeid != null ? row[permHeaders.employeeid] : '';
    const role = permHeaders.role != null ? row[permHeaders.role] : '';
    if (!employeeId || String(role).toLowerCase() === 'operator') return;
    if (!payroll.byId[employeeId] && !payroll.byIdLower[String(employeeId).toLowerCase()]) {
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Orphaned_Record',
        severity: 'High',
        description: `User permission entry for ${employeeId} not found in Payroll Tracker.`,
        detected_date: detectedDate,
        affected_record_id: employeeId
      });
    }
  });

  // Pending signups older than 30 days still pending
  const pendingData = getSheetData_(pendingSheet);
  const pendingHeaders = pendingData.headerMap;
  const now = new Date();
  (pendingData.rows || []).forEach(row => {
    const status = pendingHeaders.status != null ? row[pendingHeaders.status] : '';
    const created = pendingHeaders.createddate != null ? row[pendingHeaders.createddate] : '';
    const signupId = pendingHeaders.signupid != null ? row[pendingHeaders.signupid] : '';
    const createdDate = parseDate_(created);
    if (String(status).toLowerCase() === 'pending' && createdDate) {
      const ageDays = (now - createdDate) / (1000 * 60 * 60 * 24);
      if (ageDays > 30) {
        addIssue_(issues, keys, {
          issue_id: generateDataQualityIssueId_(),
          issue_type: 'Stale_Signup',
          severity: 'Medium',
          description: `Signup ${signupId} has been pending for over 30 days.`,
          detected_date: detectedDate,
          affected_record_id: signupId
        });
      }
    }
  });

  // Broken references: Entered_By not in User_Permissions
  const permissionNames = new Set(
    (permissionsData.rows || [])
      .map(row => permHeaders.fullname != null ? row[permHeaders.fullname] : '')
      .filter(Boolean)
      .map(name => String(name))
  );

  infRows.forEach(row => {
    const status = statusIdx != null ? row[statusIdx] : '';
    if (String(status).toLowerCase() === 'deleted') return;
    const enteredBy = enteredByIdx != null ? row[enteredByIdx] : '';
    const infractionId = infractionIdIdx != null ? row[infractionIdIdx] : '';
    if (enteredBy && !permissionNames.has(String(enteredBy))) {
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Broken_Reference',
        severity: 'Medium',
        description: `Infraction ${infractionId} entered by ${enteredBy} who is not in User_Permissions.`,
        detected_date: detectedDate,
        affected_record_id: infractionId
      });
    }
  });

  // Anomaly: >10 infractions in one day
  Object.keys(infractionsByEmployeeDate).forEach(key => {
    if (infractionsByEmployeeDate[key] > 10) {
      const parts = key.split('|');
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Anomaly_High_Infractions',
        severity: 'Medium',
        description: `Employee ${parts[0]} has ${infractionsByEmployeeDate[key]} infractions on ${parts[1]}.`,
        detected_date: detectedDate,
        affected_record_id: parts[0]
      });
    }
  });

  // Name inconsistencies
  Object.keys(employeeIdsWithNameVariants).forEach(employeeId => {
    const variants = Array.from(employeeIdsWithNameVariants[employeeId]);
    const normalized = new Set(variants.map(name => String(name).toLowerCase()));
    if (variants.length > 1 && normalized.size === 1) {
      addIssue_(issues, keys, {
        issue_id: generateDataQualityIssueId_(),
        issue_type: 'Name_Inconsistency',
        severity: 'Low',
        description: `Employee ${employeeId} has inconsistent capitalization: ${variants.join(', ')}`,
        detected_date: detectedDate,
        affected_record_id: employeeId
      });
    }
  });

  writeIssuesToSheet_(issues);

  return {
    success: true,
    summary: buildIssueSummary_(issues),
    total_issues: issues.length
  };
}

/**
 * Auto-fix a data quality issue.
 */
function autoFixDataIssue(issue_id, fix_method, fixed_by) {
  const sheet = getOrCreateDataQualityIssuesSheet();
  const sheetData = getSheetData_(sheet);
  const rowIndex = findIssueRow_(issue_id, sheetData);
  if (!rowIndex) return { success: false, error: 'Issue not found' };

  const row = sheet.getRange(rowIndex, 1, 1, DATA_QUALITY_HEADERS.length).getValues()[0];
  const issueType = row[1];
  const affectedRecordId = row[5];

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const infractionsSheet = ss.getSheetByName('Infractions');
  const permissionsSheet = ss.getSheetByName('User_Permissions');
  const pendingSheet = ss.getSheetByName('Pending_Signups');

  const now = new Date();
  const fixAction = fix_method || '';
  let fixResult = { success: false, message: 'No fix applied' };

  if (issueType === 'Missing_Employee') {
    fixResult = fixMissingEmployee_(infractionsSheet, affectedRecordId, fix_method, fixed_by);
  } else if (issueType === 'Duplicate_Infraction') {
    fixResult = fixDuplicateInfraction_(infractionsSheet, affectedRecordId, fix_method, fixed_by);
  } else if (issueType === 'Invalid_Points') {
    fixResult = fixInvalidPoints_(affectedRecordId, fix_method, fixed_by);
  } else if (issueType === 'Orphaned_Record') {
    fixResult = fixOrphanedRecord_(permissionsSheet, pendingSheet, affectedRecordId, fix_method, fixed_by);
  } else if (issueType === 'Stale_Signup') {
    fixResult = fixOrphanedRecord_(permissionsSheet, pendingSheet, affectedRecordId, fix_method, fixed_by);
  } else if (issueType === 'Name_Inconsistency') {
    fixResult = fixNameInconsistency_(infractionsSheet, affectedRecordId, fixed_by);
  } else if (issueType === 'Broken_Reference') {
    fixResult = { success: false, message: 'Manual review required' };
  } else {
    fixResult = { success: false, message: 'No auto-fix available' };
  }

  if (!fixResult.success) return fixResult;

  updateIssueRow_(sheet, rowIndex, {
    Status: fixResult.status || 'Fixed',
    Fixed_Date: now,
    Fixed_By: fixed_by || 'Operator',
    Fix_Action_Taken: fixAction
  });

  return { success: true, message: fixResult.message || 'Issue fixed' };
}

function fixMissingEmployee_(infractionsSheet, infractionId, fixMethod, fixedBy) {
  if (!infractionsSheet) return { success: false, message: 'Infractions sheet not found' };
  const data = infractionsSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const idIdx = map.infractionid;
  if (typeof idIdx === 'undefined') return { success: false, message: 'Infraction_ID column missing' };

  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[idIdx] === infractionId);
  if (rowIndex === -1) return { success: false, message: 'Infraction not found' };

  const sheetRow = rowIndex + 1;
  const reasonIdx = map.modificationreason;
  const statusIdx = map.status;
  const modifiedByIdx = map.lastmodifiedby;
  const modifiedDateIdx = map.lastmodifiedtimestamp;
  const employeeIdIdx = map.employeeid;
  const nameIdx = map.fullname;

  if (fixMethod.indexOf('Link_To_Employee') === 0) {
    const parts = fixMethod.split(':');
    const newEmployeeId = parts[1];
    if (!newEmployeeId) return { success: false, message: 'Missing employee ID for linking' };
    const payroll = getPayrollEmployeesMap_();
    const payrollRecord = payroll.byId[newEmployeeId] || payroll.byIdLower[String(newEmployeeId).toLowerCase()];
    if (!payrollRecord) return { success: false, message: 'Employee not found in Payroll Tracker' };
    if (employeeIdIdx != null) infractionsSheet.getRange(sheetRow, employeeIdIdx + 1).setValue(newEmployeeId);
    if (nameIdx != null) infractionsSheet.getRange(sheetRow, nameIdx + 1).setValue(payrollRecord.full_name);
    if (modifiedByIdx != null) infractionsSheet.getRange(sheetRow, modifiedByIdx + 1).setValue(fixedBy || 'Operator');
    if (modifiedDateIdx != null) infractionsSheet.getRange(sheetRow, modifiedDateIdx + 1).setValue(new Date());
    if (reasonIdx != null) infractionsSheet.getRange(sheetRow, reasonIdx + 1).setValue('Linked to correct employee');
  } else {
    if (statusIdx != null) infractionsSheet.getRange(sheetRow, statusIdx + 1).setValue('Deleted');
    if (modifiedByIdx != null) infractionsSheet.getRange(sheetRow, modifiedByIdx + 1).setValue(fixedBy || 'Operator');
    if (modifiedDateIdx != null) infractionsSheet.getRange(sheetRow, modifiedDateIdx + 1).setValue(new Date());
    if (reasonIdx != null) infractionsSheet.getRange(sheetRow, reasonIdx + 1).setValue('Employee not found in Payroll');
  }

  logEditAction({
    actionType: 'data_quality_fix',
    directorEmail: fixedBy || 'Operator',
    targetType: 'infraction',
    targetId: infractionId,
    employeeId: '',
    employeeName: '',
    fieldChanged: 'employee_reference',
    originalValue: '',
    newValue: fixMethod,
    reason: 'Data quality fix',
    sessionInfo: 'Data Quality'
  });

  return { success: true, message: 'Missing employee issue resolved' };
}

function fixDuplicateInfraction_(infractionsSheet, affectedRecordId, fixMethod, fixedBy) {
  if (fixMethod === 'Keep_All') {
    return { success: true, message: 'Marked as false positive', status: 'False_Positive' };
  }
  if (!infractionsSheet) return { success: false, message: 'Infractions sheet not found' };
  const data = infractionsSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const idIdx = map.infractionid;
  const statusIdx = map.status;
  const modifiedByIdx = map.lastmodifiedby;
  const modifiedDateIdx = map.lastmodifiedtimestamp;
  const reasonIdx = map.modificationreason;
  const entryTsIdx = map.entrytimestamp;

  const ids = String(affectedRecordId || '')
    .split(',')
    .map(id => id.trim())
    .filter(Boolean);
  if (ids.length < 2) return { success: false, message: 'No duplicates found' };

  const rows = data
    .map((row, idx) => ({ row: row, index: idx }))
    .filter(item => item.index > 0 && ids.indexOf(item.row[idIdx]) !== -1);

  rows.sort((a, b) => {
    const aDate = parseDate_(a.row[entryTsIdx]) || new Date(0);
    const bDate = parseDate_(b.row[entryTsIdx]) || new Date(0);
    return aDate - bDate;
  });

  rows.slice(1).forEach(item => {
    const sheetRow = item.index + 1;
    if (statusIdx != null) infractionsSheet.getRange(sheetRow, statusIdx + 1).setValue('Deleted');
    if (modifiedByIdx != null) infractionsSheet.getRange(sheetRow, modifiedByIdx + 1).setValue(fixedBy || 'Operator');
    if (modifiedDateIdx != null) infractionsSheet.getRange(sheetRow, modifiedDateIdx + 1).setValue(new Date());
    if (reasonIdx != null) infractionsSheet.getRange(sheetRow, reasonIdx + 1).setValue('Duplicate infraction removed');
  });

  return { success: true, message: 'Duplicate infractions removed' };
}

function fixInvalidPoints_(employeeId, fixMethod, fixedBy) {
  if (fixMethod === 'Investigate') {
    return { success: true, message: 'Marked for manual review', status: 'Ignored' };
  }
  try {
    calculatePoints(employeeId);
  } catch (error) {
    console.error('Error recalculating points:', error.toString());
    return { success: false, message: 'Failed to recalculate points' };
  }

  logEditAction({
    actionType: 'data_quality_fix',
    directorEmail: fixedBy || 'Operator',
    targetType: 'points',
    targetId: employeeId,
    employeeId: employeeId,
    employeeName: '',
    fieldChanged: 'points_recalculated',
    originalValue: '',
    newValue: '',
    reason: 'Data quality recalculation',
    sessionInfo: 'Data Quality'
  });

  return { success: true, message: 'Points recalculated' };
}

function fixOrphanedRecord_(permissionsSheet, pendingSheet, recordId, fixMethod, fixedBy) {
  if (fixMethod === 'Cancel_Signup') {
    if (!pendingSheet) return { success: false, message: 'Pending_Signups sheet not found' };
    const data = pendingSheet.getDataRange().getValues();
    const headers = data[0];
    const map = getHeaderMap_(headers);
    const signupIdx = map.signupid;
    const statusIdx = map.status;
    const rowIndex = data.findIndex((row, idx) => idx > 0 && row[signupIdx] === recordId);
    if (rowIndex === -1) return { success: false, message: 'Signup not found' };
    pendingSheet.getRange(rowIndex + 1, statusIdx + 1).setValue('Cancelled');
    return { success: true, message: 'Signup cancelled' };
  }

  if (!permissionsSheet) return { success: false, message: 'User_Permissions sheet not found' };
  const data = permissionsSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const employeeIdx = map.employeeid;
  const statusIdx = map.status;
  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[employeeIdx] === recordId);
  if (rowIndex === -1) return { success: false, message: 'User permission entry not found' };
  permissionsSheet.getRange(rowIndex + 1, statusIdx + 1).setValue('Inactive');
  return { success: true, message: 'User permissions deactivated' };
}

function fixNameInconsistency_(infractionsSheet, employeeId, fixedBy) {
  if (!infractionsSheet) return { success: false, message: 'Infractions sheet not found' };
  const payroll = getPayrollEmployeesMap_();
  const payrollRecord = payroll.byId[employeeId] || payroll.byIdLower[String(employeeId).toLowerCase()];
  if (!payrollRecord) return { success: false, message: 'Employee not found in Payroll Tracker' };
  const data = infractionsSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const employeeIdx = map.employeeid;
  const nameIdx = map.fullname;
  const modifiedByIdx = map.lastmodifiedby;
  const modifiedDateIdx = map.lastmodifiedtimestamp;
  const reasonIdx = map.modificationreason;

  data.forEach((row, idx) => {
    if (idx === 0) return;
    if (row[employeeIdx] === employeeId && nameIdx != null) {
      const sheetRow = idx + 1;
      infractionsSheet.getRange(sheetRow, nameIdx + 1).setValue(payrollRecord.full_name);
      if (modifiedByIdx != null) infractionsSheet.getRange(sheetRow, modifiedByIdx + 1).setValue(fixedBy || 'Operator');
      if (modifiedDateIdx != null) infractionsSheet.getRange(sheetRow, modifiedDateIdx + 1).setValue(new Date());
      if (reasonIdx != null) infractionsSheet.getRange(sheetRow, reasonIdx + 1).setValue('Standardized name to Payroll Tracker');
    }
  });

  return { success: true, message: 'Names standardized' };
}

/**
 * Weekly scheduled cleanup.
 */
function scheduledCleanup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const settingsSheet = ss.getSheetByName('Settings');
  ensureDataQualitySettingsRows(settingsSheet);
  const settings = getDataQualitySettingsFromSheet(settingsSheet);

  const summary = {
    names_standardized: 0,
    email_log_duplicates_removed: 0,
    system_log_duplicates_removed: 0,
    signups_archived: 0,
    tokens_expired: 0
  };

  if (settings.actions.standardize_names) {
    summary.names_standardized = standardizeInfractionNames_();
  }
  if (settings.actions.remove_duplicate_logs) {
    const removed = removeDuplicateLogs_();
    summary.email_log_duplicates_removed = removed.email;
    summary.system_log_duplicates_removed = removed.system;
  }
  if (settings.actions.archive_old_signups) {
    summary.signups_archived = archiveOldSignups_();
  }
  if (settings.actions.expire_old_tokens) {
    summary.tokens_expired = expireOldSignupTokens_();
  }

  const now = new Date();
  const lastRunRow = findSettingsRowByLabel(settingsSheet, 'Last Cleanup Date');
  if (lastRunRow > 0) settingsSheet.getRange(lastRunRow, 2).setValue(now);
  const nextRunRow = findSettingsRowByLabel(settingsSheet, 'Next Cleanup Date');
  const nextRun = computeNextCleanupDate_(settings.frequency || 'Weekly');
  if (nextRunRow > 0) settingsSheet.getRange(nextRunRow, 2).setValue(nextRun || '');

  if (settings.emailReport) {
    try {
      const recipient = getStoreEmail();
      if (recipient) {
        const body = [
          'Data Quality Cleanup Report',
          `Run at: ${now.toLocaleString()}`,
          '',
          `Names standardized: ${summary.names_standardized}`,
          `Email log duplicates removed: ${summary.email_log_duplicates_removed}`,
          `System log duplicates removed: ${summary.system_log_duplicates_removed}`,
          `Signups archived: ${summary.signups_archived}`,
          `Tokens expired: ${summary.tokens_expired}`
        ].join('\n');
        MailApp.sendEmail(recipient, 'Data Quality Cleanup Report', body);
      }
    } catch (error) {
      console.error('Failed to send cleanup report:', error.toString());
    }
  }

  return {
    success: true,
    summary: summary,
    ran_at: now
  };
}

function standardizeInfractionNames_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const infractionsSheet = ss.getSheetByName('Infractions');
  if (!infractionsSheet) return 0;
  const payroll = getPayrollEmployeesMap_();
  const data = infractionsSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const employeeIdx = map.employeeid;
  const nameIdx = map.fullname;
  let updated = 0;
  data.forEach((row, idx) => {
    if (idx === 0) return;
    const employeeId = row[employeeIdx];
    const payrollRecord = payroll.byId[employeeId] || payroll.byIdLower[String(employeeId).toLowerCase()];
    if (payrollRecord && nameIdx != null && row[nameIdx] && row[nameIdx] !== payrollRecord.full_name) {
      infractionsSheet.getRange(idx + 1, nameIdx + 1).setValue(payrollRecord.full_name);
      updated += 1;
    }
  });
  return updated;
}

function removeDuplicateLogs_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const emailLog = ss.getSheetByName('Email_Log');
  const systemLog = ss.getSheetByName('System_Log');
  let emailRemoved = 0;
  let systemRemoved = 0;

  if (emailLog && emailLog.getLastRow() > 1) {
    const data = emailLog.getDataRange().getValues();
    const headers = data[0];
    const map = getHeaderMap_(headers);
    const tsIdx = map.timestamp;
    const recipientIdx = map.recipientemail;
    const typeIdx = map.emailtype;
    const seen = new Set();
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const key = `${row[tsIdx]}|${row[recipientIdx]}|${row[typeIdx]}`;
      if (seen.has(key)) {
        emailLog.deleteRow(i + 1);
        emailRemoved += 1;
      } else {
        seen.add(key);
      }
    }
  }

  if (systemLog && systemLog.getLastRow() > 1) {
    const data = systemLog.getDataRange().getValues();
    const headers = data[0];
    const map = getHeaderMap_(headers);
    const tsIdx = map.timestamp;
    const detailsIdx = map.eventdetails;
    const typeIdx = map.eventtype;
    const seen = {};
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const timestamp = parseDate_(row[tsIdx]);
      const eventType = String(row[typeIdx] || '').toLowerCase();
      if (eventType !== 'error') continue;
      const key = `${row[typeIdx]}|${row[detailsIdx]}`;
      if (!timestamp) continue;
      if (seen[key]) {
        const diff = Math.abs(seen[key] - timestamp.getTime());
        if (diff <= 60 * 1000) {
          systemLog.deleteRow(i + 1);
          systemRemoved += 1;
          continue;
        }
      }
      seen[key] = timestamp.getTime();
    }
  }

  return { email: emailRemoved, system: systemRemoved };
}

function archiveOldSignups_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pendingSheet = ss.getSheetByName('Pending_Signups');
  if (!pendingSheet || pendingSheet.getLastRow() < 2) return 0;
  const archiveSheet = getOrCreateArchivedSignupsSheet();
  const data = pendingSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const statusIdx = map.status;
  const completedIdx = map.completeddate;

  const rowsToArchive = [];
  const now = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const status = row[statusIdx];
    const completedDate = parseDate_(row[completedIdx]);
    if (String(status).toLowerCase() === 'completed' && completedDate) {
      const ageDays = (now - completedDate) / (1000 * 60 * 60 * 24);
      if (ageDays > 90) {
        rowsToArchive.unshift(row);
        pendingSheet.deleteRow(i + 1);
      }
    }
  }
  if (rowsToArchive.length) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, headers.length).setValues(rowsToArchive);
  }
  return rowsToArchive.length;
}

function expireOldSignupTokens_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pendingSheet = ss.getSheetByName('Pending_Signups');
  if (!pendingSheet || pendingSheet.getLastRow() < 2) return 0;
  const data = pendingSheet.getDataRange().getValues();
  const headers = data[0];
  const map = getHeaderMap_(headers);
  const statusIdx = map.status;
  const expiresIdx = map.expiresdate;
  let updated = 0;
  const now = new Date();
  data.forEach((row, idx) => {
    if (idx === 0) return;
    const status = row[statusIdx];
    const expires = parseDate_(row[expiresIdx]);
    if (String(status).toLowerCase() === 'pending' && expires) {
      const ageDays = (now - expires) / (1000 * 60 * 60 * 24);
      if (ageDays > 30) {
        pendingSheet.getRange(idx + 1, statusIdx + 1).setValue('Expired');
        updated += 1;
      }
    }
  });
  return updated;
}

// =========================
// API wrappers for UI
// =========================

function api_scanDataQuality(token) {
  return scanDataQuality(token);
}

function api_getDataQualityIssues(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };

  const sheet = getOrCreateDataQualityIssuesSheet();
  const data = getSheetData_(sheet);
  const issues = data.rows.map(row => ({
    issue_id: row[0],
    issue_type: row[1],
    severity: row[2],
    description: row[3],
    detected_date: row[4],
    affected_record_id: row[5],
    status: row[6],
    fixed_date: row[7],
    fixed_by: row[8],
    fix_action_taken: row[9]
  }));

  const summary = buildIssueSummary_(issues.filter(issue => String(issue.status || 'Open') === 'Open'));
  const settingsSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  ensureDataQualitySettingsRows(settingsSheet);
  const settings = getDataQualitySettingsFromSheet(settingsSheet);

  return { success: true, issues: issues, summary: summary, settings: settings };
}

function api_autoFixDataIssue(issueId, fixMethod, fixedBy, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return autoFixDataIssue(issueId, fixMethod, fixedBy || session.session.user_name || session.session.role);
}

function api_previewDataQualityFix(issueId, fixMethod, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };

  const sheet = getOrCreateDataQualityIssuesSheet();
  const data = getSheetData_(sheet);
  const rowIndex = findIssueRow_(issueId, data);
  if (!rowIndex) return { success: false, error: 'Issue not found' };
  const row = sheet.getRange(rowIndex, 1, 1, DATA_QUALITY_HEADERS.length).getValues()[0];

  return {
    success: true,
    issue: {
      issue_id: row[0],
      issue_type: row[1],
      severity: row[2],
      description: row[3],
      affected_record_id: row[5]
    },
    fix_method: fixMethod,
    preview: `Will apply ${fixMethod} to record ${row[5]}`
  };
}

function api_updateDataQualityIssueStatus(issueIds, status, fixedBy, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  const sheet = getOrCreateDataQualityIssuesSheet();
  const ids = Array.isArray(issueIds) ? issueIds : [issueIds];
  const data = getSheetData_(sheet);
  const now = new Date();
  ids.forEach(id => {
    const rowIndex = findIssueRow_(id, data);
    if (!rowIndex) return;
    updateIssueRow_(sheet, rowIndex, {
      Status: status,
      Fixed_Date: now,
      Fixed_By: fixedBy || 'Operator',
      Fix_Action_Taken: status
    });
  });
  return { success: true };
}

function api_runScheduledCleanup(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return scheduledCleanup();
}

function api_updateDataQualitySettings(config, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };

  const nextRunDate = config && config.frequency
    ? computeNextCleanupDate_(config.frequency)
    : '';
  const result = updateDataQualitySettings({
    frequency: config.frequency,
    actions: config.actions,
    emailReport: config.emailReport,
    nextCleanupDate: nextRunDate
  });

  if (result.success) {
    applyCleanupSchedule_(config.frequency);
  }
  return result;
}

function computeNextCleanupDate_(frequency) {
  const now = new Date();
  const next = new Date(now);
  switch (String(frequency || '').toLowerCase()) {
    case 'daily':
      next.setDate(next.getDate() + 1);
      break;
    case 'monthly':
      next.setMonth(next.getMonth() + 1);
      break;
    case 'disabled':
      return '';
    default:
      next.setDate(next.getDate() + 7);
      break;
  }
  return next;
}

function applyCleanupSchedule_(frequency) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduledCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const normalized = String(frequency || '').toLowerCase();
  if (normalized === 'disabled') return;

  let builder = ScriptApp.newTrigger('scheduledCleanup').timeBased().atHour(2);
  if (normalized === 'daily') {
    builder = builder.everyDays(1);
  } else if (normalized === 'monthly') {
    builder = builder.everyMonths(1);
  } else {
    builder = builder.everyWeeks(1);
  }
  builder.create();
}

// =========================
// Test function
// =========================

function testDataQuality() {
  console.log('=== testDataQuality ===');
  try {
    const scanResult = scanDataQuality();
    console.log('Scan result:', JSON.stringify(scanResult));
  } catch (error) {
    console.error('testDataQuality error:', error.toString());
  }
}
