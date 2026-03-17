/**
 * DisasterRecovery.gs
 * Micro-Phase 34: Disaster Recovery & Backup Testing
 */

const BACKUP_TEST_LOG_HEADERS = [
  'Test_ID',
  'Test_Date',
  'Test_Type',
  'Backup_File_ID',
  'Test_Status',
  'Issues_Found',
  'Validation_Results',
  'Test_Duration_Seconds',
  'Tested_By'
];

const RESTORE_HISTORY_HEADERS = [
  'Restore_ID',
  'Restore_Date',
  'Restored_From_Backup_ID',
  'Restore_Type',
  'Sheets_Restored',
  'Restored_By',
  'Reason_For_Restore',
  'Pre_Restore_Snapshot_ID',
  'Restore_Status',
  'Issues_Encountered'
];

const DAILY_INCREMENTAL_FOLDER_NAME = 'Daily_Incrementals';
const SNAPSHOT_FOLDER_NAME = 'Pre_Restore_Snapshots';
const BACKUP_TEST_FOLDER_NAME = 'Backup_Tests';

function getOrCreateBackupTestLogSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Backup_Test_Log');
  if (!sheet) {
    sheet = ss.insertSheet('Backup_Test_Log');
    sheet.getRange(1, 1, 1, BACKUP_TEST_LOG_HEADERS.length).setValues([BACKUP_TEST_LOG_HEADERS]);
    sheet.getRange(1, 1, 1, BACKUP_TEST_LOG_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateRestoreHistorySheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Restore_History');
  if (!sheet) {
    sheet = ss.insertSheet('Restore_History');
    sheet.getRange(1, 1, 1, RESTORE_HISTORY_HEADERS.length).setValues([RESTORE_HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, RESTORE_HISTORY_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getBackupSubFolder_(name) {
  const mainFolder = getBackupFolder();
  const folders = mainFolder.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return mainFolder.createFolder(name);
}

function getLatestBackupByType_(types) {
  const list = listBackups();
  if (!list.success) return null;
  const allowed = (types || []).map(t => String(t));
  return list.backups.find(b => allowed.indexOf(String(b.backup_type)) !== -1) || null;
}

function getLastBackupOfType_(type) {
  const list = listBackups();
  if (!list.success) return null;
  return list.backups.find(b => String(b.backup_type) === String(type)) || null;
}

function parseBackupLogDate_(value) {
  if (!value) return null;
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function getChangeDateFromRow_(headers, row) {
  const dateKeys = [
    'last_updated',
    'last_modified',
    'updated_at',
    'timestamp',
    'date',
    'created_at',
    'entry_date'
  ];
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase().trim();
    if (dateKeys.indexOf(header) === -1) continue;
    const cell = row[i];
    const date = cell instanceof Date ? cell : new Date(cell);
    if (!isNaN(date.getTime())) return date;
  }
  return null;
}

function createDailyIncremental() {
  const startTime = new Date();
  try {
    const lastDaily = getLastBackupOfType_('daily_incremental');
    const lastDate = lastDaily ? parseBackupLogDate_(lastDaily.backup_date) : null;
    const baseBackup = getLatestBackupByType_(['automatic', 'manual']);
    const baseBackupId = baseBackup ? baseBackup.backup_id : '';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const changes = {};
    const recordCounts = {};

    (SHEETS_TO_BACKUP || []).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (!data || data.length <= 1) {
        recordCounts[sheetName] = 0;
        return;
      }
      const headers = data[0];
      const rows = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowDate = getChangeDateFromRow_(headers, row);
        if (!lastDate || (rowDate && rowDate >= lastDate)) {
          rows.push({ row_number: i + 1, values: row });
        }
      }
      if (rows.length) {
        changes[sheetName] = { rows: rows };
      }
      recordCounts[sheetName] = rows.length;
    });

    recordCounts.base_backup_id = baseBackupId || '';

    const payload = {
      metadata: {
        backup_id: 'INCR' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss'),
        backup_type: 'daily_incremental',
        backup_date: new Date().toISOString(),
        base_backup_id: baseBackupId || '',
        last_daily_backup_date: lastDate ? lastDate.toISOString() : ''
      },
      sheets: changes
    };

    const folder = getBackupSubFolder_(DAILY_INCREMENTAL_FOLDER_NAME);
    const filename = 'Incremental_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.json';
    const file = folder.createFile(filename, JSON.stringify(payload, null, 2), MimeType.PLAIN_TEXT);

    // Retention: keep last 90 days
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 90);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getDateCreated() < cutoff) {
        f.setTrashed(true);
      }
    }

    logBackup(payload.metadata.backup_id, new Date(), 'daily_incremental', file.getId(), file.getUrl(),
      folder.getUrl(), recordCounts, 'Success', '', 'System (Daily Incremental)');

    const duration = (new Date() - startTime) / 1000;
    return {
      success: true,
      backup_id: payload.metadata.backup_id,
      backup_file_id: file.getId(),
      backup_file_url: file.getUrl(),
      base_backup_id: baseBackupId,
      record_counts: recordCounts,
      duration_seconds: duration
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function testBackupRestore(backupId) {
  const startTime = new Date();
  let testFile = null;
  const issues = [];
  let validation = {};
  let testStatus = 'Failed';

  try {
    const list = listBackups();
    if (!list.success) {
      return { success: false, error: 'Unable to read backup log' };
    }

    const targetBackup = backupId
      ? list.backups.find(b => b.backup_id === backupId || b.file_id === backupId)
      : getLatestBackupByType_(['automatic', 'manual']);

    if (!targetBackup || !targetBackup.file_id) {
      return { success: false, error: 'No valid backup file found to test' };
    }

    const backupSpreadsheet = SpreadsheetApp.openById(targetBackup.file_id);
    const testName = 'Backup_Test_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    testFile = SpreadsheetApp.create(testName);
    const testFileId = testFile.getId();

    const folder = getBackupSubFolder_(BACKUP_TEST_FOLDER_NAME);
    DriveApp.getFileById(testFileId).moveTo(folder);

    copySpreadsheetData_(backupSpreadsheet, testFile);

    validation = runBackupValidationChecks_(testFile, backupSpreadsheet, issues);

    if (validation && validation.overall === 'pass') {
      testStatus = 'Success';
    } else if (validation && validation.overall === 'partial') {
      testStatus = 'Partial';
    } else {
      testStatus = 'Failed';
    }

    const duration = (new Date() - startTime) / 1000;
    const testId = 'BTEST' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const sheet = getOrCreateBackupTestLogSheet_();
    sheet.appendRow([
      testId,
      new Date().toISOString(),
      'Quarterly_Full',
      targetBackup.file_id,
      testStatus,
      JSON.stringify(issues),
      JSON.stringify(validation),
      duration,
      'System'
    ]);

    sendBackupTestEmail_(testStatus, issues, validation, targetBackup);

    return {
      success: testStatus === 'Success',
      test_id: testId,
      status: testStatus,
      validation: validation,
      issues: issues,
      duration_seconds: duration
    };
  } catch (error) {
    issues.push(error.toString());
    return { success: false, error: error.toString(), issues: issues };
  } finally {
    if (testFile) {
      try {
        DriveApp.getFileById(testFile.getId()).setTrashed(true);
      } catch (e) {
        // ignore cleanup errors
      }
    }
  }
}

function runBackupValidationChecks_(testSpreadsheet, backupSpreadsheet, issues) {
  const results = {};
  const requiredSheets = ['Infractions', 'Settings', 'User_Permissions'];
  let failures = 0;

  requiredSheets.forEach(name => {
    const sheet = testSpreadsheet.getSheetByName(name);
    results['sheet_exists_' + name] = !!sheet;
    if (!sheet) {
      failures++;
      issues.push('Missing required sheet: ' + name);
    }
  });

  // Row count comparison based on _Backup_Info
  const infoSheet = backupSpreadsheet.getSheetByName('_Backup_Info');
  if (infoSheet) {
    const data = infoSheet.getDataRange().getValues();
    const rowCounts = {};
    let found = false;
    data.forEach(row => {
      if (row[0] === 'Record Counts by Sheet') {
        found = true;
        return;
      }
      if (found && row[0]) {
        rowCounts[row[0]] = row[1];
      }
    });
    Object.keys(rowCounts).forEach(sheetName => {
      const testSheet = testSpreadsheet.getSheetByName(sheetName);
      if (!testSheet) return;
      const count = testSheet.getLastRow() > 0 ? testSheet.getLastRow() - 1 : 0;
      if (count !== rowCounts[sheetName]) {
        failures++;
        results['row_count_' + sheetName] = 'mismatch';
        issues.push('Row count mismatch for ' + sheetName + ': ' + count + ' vs ' + rowCounts[sheetName]);
      } else {
        results['row_count_' + sheetName] = 'match';
      }
    });
  }

  // Spot check infractions to employees (if Payroll Tracker available)
  const payroll = testSpreadsheet.getSheetByName('Payroll Tracker');
  const infractions = testSpreadsheet.getSheetByName('Infractions');
  if (payroll && infractions) {
    const payrollData = payroll.getDataRange().getValues();
    const payrollIds = {};
    payrollData.slice(1).forEach(row => {
      const id = row[0];
      if (id) payrollIds[id] = true;
    });
    const infractionsData = infractions.getDataRange().getValues().slice(1);
    const invalid = infractionsData.filter(row => row[1] && !payrollIds[row[1]]);
    results.invalid_employee_references = invalid.length;
    if (invalid.length > 0) {
      failures++;
      issues.push('Infractions reference missing employees: ' + invalid.length);
    }
  }

  results.overall = failures === 0 ? 'pass' : (failures <= 2 ? 'partial' : 'fail');
  return results;
}

function sendBackupTestEmail_(status, issues, validation, backup) {
  try {
    const recipient = getStoreEmail();
    if (!recipient) return;
    const subject = 'Backup Test ' + (status === 'Success' ? 'Passed' : 'Failed');
    const body = [
      'Backup Test Result: ' + status,
      'Backup ID: ' + (backup ? backup.backup_id : ''),
      'Backup File: ' + (backup ? backup.file_id : ''),
      '',
      'Validation Results:',
      JSON.stringify(validation || {}, null, 2),
      '',
      'Issues Found:',
      JSON.stringify(issues || [], null, 2)
    ].join('\n');
    MailApp.sendEmail(recipient, subject, body);
  } catch (error) {
    console.log('Backup test email failed: ' + error.toString());
  }
}

function restoreFromBackup(backup_id, restore_options, confirmed_by_operator) {
  try {
    const session = requireOperatorSession_(restore_options && restore_options.token);
    if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };

    const operatorName = confirmed_by_operator || session.session.user_name || 'Operator';
    if (!backup_id) return { success: false, error: 'Backup ID is required' };

    const list = listBackups();
    if (!list.success) return { success: false, error: 'Backup log unavailable' };
    const backup = list.backups.find(b => b.backup_id === backup_id || b.file_id === backup_id);
    if (!backup || !backup.file_id) return { success: false, error: 'Backup not found' };

    const tested = getLatestBackupTestStatus_(backup.file_id);
    if (!tested || tested !== 'Success') {
      return { success: false, error: 'This backup has not been tested. Test it first.' };
    }

    const restoreType = restore_options && restore_options.restore_type === 'partial' ? 'Partial_Sheets' : 'Full';
    const previewOnly = !!(restore_options && restore_options.preview_only);
    const sheetsToRestore = restore_options && restore_options.sheets_to_restore
      ? restore_options.sheets_to_restore
      : SHEETS_TO_BACKUP.slice();

    if (previewOnly) {
      const preview = buildRestorePreview_(backup.file_id, sheetsToRestore);
      return { success: true, preview: preview, preview_only: true };
    }

    const snapshotId = createPreRestoreSnapshot_(operatorName);
    const backupSpreadsheet = SpreadsheetApp.openById(backup.file_id);
    const targetSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    const restoredSheets = restoreSheetsFromBackup_(backupSpreadsheet, targetSpreadsheet, sheetsToRestore);
    const issues = restoredSheets.errors || [];
    const status = issues.length === 0 ? 'Success' : 'Partial';

    logRestoreHistory_(backup.file_id, restoreType, restoredSheets.names, operatorName,
      (restore_options && restore_options.reason) || '', snapshotId, status, issues);

    sendRestoreNotificationEmail(backup.file_id, restoredSheets.details, operatorName, issues);
    forceLogoutAllUsers_();

    return {
      success: status === 'Success',
      status: status,
      restored_sheets: restoredSheets.names,
      snapshot_id: snapshotId,
      issues: issues
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function rollbackRestore(restore_id, rollback_by_operator) {
  try {
    const history = listRestoreHistory_();
    const restore = history.find(r => r.Restore_ID === restore_id);
    if (!restore) return { success: false, error: 'Restore ID not found' };
    const snapshotId = restore.Pre_Restore_Snapshot_ID;
    if (!snapshotId) return { success: false, error: 'No snapshot found for restore' };

    const snapshot = SpreadsheetApp.openById(snapshotId);
    const target = SpreadsheetApp.openById(SPREADSHEET_ID);
    const result = restoreSheetsFromBackup_(snapshot, target, SHEETS_TO_BACKUP.slice());

    logRestoreHistory_(snapshotId, 'Full', result.names, rollback_by_operator || 'Operator',
      'Rollback restore ' + restore_id, snapshotId, 'Rolled_Back', result.errors || []);

    return { success: true, rolled_back: true, issues: result.errors || [] };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function scheduleBackupTests() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const existing = triggers.find(t => t.getHandlerFunction() === 'runScheduledBackupTest');
    if (existing) {
      return { success: true, message: 'Trigger already exists', trigger_id: existing.getUniqueId() };
    }
    const trigger = ScriptApp.newTrigger('runScheduledBackupTest')
      .timeBased()
      .onMonthDay(8)
      .atHour(3)
      .create();
    return { success: true, message: 'Backup test trigger created', trigger_id: trigger.getUniqueId() };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function runScheduledBackupTest() {
  const now = new Date();
  const month = now.getMonth();
  if (month !== 0 && month !== 3 && month !== 6 && month !== 9) return;

  const result = testBackupRestore();
  if (!result.success) {
    sendBackupTestEmail_('Failed', result.issues || ['Unknown error'], result.validation || {}, null);
    createSystemBackup('automatic', 'System (Auto Retry)');
  }
}

function getLatestBackupTestStatus_(backupFileId) {
  const sheet = getOrCreateBackupTestLogSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  const headers = data[0];
  const fileIndex = headers.indexOf('Backup_File_ID');
  const statusIndex = headers.indexOf('Test_Status');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][fileIndex] === backupFileId) {
      return data[i][statusIndex];
    }
  }
  return null;
}

function listBackupTests_() {
  const sheet = getOrCreateBackupTestLogSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    const item = {};
    headers.forEach((header, idx) => item[header] = row[idx]);
    return item;
  }).reverse();
}

function listRestoreHistory_() {
  const sheet = getOrCreateRestoreHistorySheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    const item = {};
    headers.forEach((header, idx) => item[header] = row[idx]);
    return item;
  }).reverse();
}

function listSnapshots_() {
  const folder = getBackupSubFolder_(SNAPSHOT_FOLDER_NAME);
  const files = folder.getFiles();
  const snapshots = [];
  while (files.hasNext()) {
    const file = files.next();
    snapshots.push({
      id: file.getId(),
      name: file.getName(),
      created: file.getDateCreated(),
      size: file.getSize()
    });
  }
  snapshots.sort((a, b) => b.created - a.created);
  return snapshots;
}

function createPreRestoreSnapshot_(operatorName) {
  const folder = getBackupSubFolder_(SNAPSHOT_FOLDER_NAME);
  const name = 'Pre_Restore_Snapshot_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const file = DriveApp.getFileById(SPREADSHEET_ID).makeCopy(name, folder);
  return file.getId();
}

function buildRestorePreview_(backupFileId, sheetsToRestore) {
  const backup = SpreadsheetApp.openById(backupFileId);
  const current = SpreadsheetApp.openById(SPREADSHEET_ID);
  const preview = [];

  sheetsToRestore.forEach(name => {
    const b = backup.getSheetByName(name);
    const c = current.getSheetByName(name);
    const bRows = b ? Math.max(0, b.getLastRow() - 1) : 0;
    const cRows = c ? Math.max(0, c.getLastRow() - 1) : 0;
    preview.push({
      sheet: name,
      current_rows: cRows,
      backup_rows: bRows,
      difference: bRows - cRows
    });
  });

  return preview;
}

function restoreSheetsFromBackup_(backupSpreadsheet, targetSpreadsheet, sheetsToRestore) {
  const restored = [];
  const errors = [];

  sheetsToRestore.forEach(sheetName => {
    try {
      const backupSheet = backupSpreadsheet.getSheetByName(sheetName);
      if (!backupSheet) {
        errors.push(sheetName + ': Missing in backup');
        return;
      }

      let targetSheet = targetSpreadsheet.getSheetByName(sheetName);
      if (!targetSheet) {
        targetSheet = targetSpreadsheet.insertSheet(sheetName);
      }

      const backupRange = backupSheet.getDataRange();
      const backupValues = backupRange.getValues();
      const backupFormulas = backupRange.getFormulas();
      const backupFormats = backupRange.getNumberFormats();
      const backupValidations = backupRange.getDataValidations();

      targetSheet.clear();
      const merged = backupValues.map((row, r) => row.map((cell, c) => {
        const formula = backupFormulas[r][c];
        return formula ? formula : cell;
      }));

      const targetRange = targetSheet.getRange(1, 1, backupValues.length, backupValues[0].length);
      targetRange.setValues(merged);
      try {
        targetRange.setNumberFormats(backupFormats);
        targetRange.setDataValidations(backupValidations);
      } catch (e) {
        // ignore formatting errors
      }

      const frozenRows = backupSheet.getFrozenRows();
      if (frozenRows > 0) targetSheet.setFrozenRows(frozenRows);

      restored.push({ name: sheetName, rows: backupValues.length });
    } catch (error) {
      errors.push(sheetName + ': ' + error.toString());
    }
  });

  return {
    names: restored.map(r => r.name),
    details: restored,
    errors: errors
  };
}

function logRestoreHistory_(backupFileId, restoreType, sheetsRestored, restoredBy, reason, snapshotId, status, issues) {
  const sheet = getOrCreateRestoreHistorySheet_();
  const restoreId = 'RST' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  sheet.appendRow([
    restoreId,
    new Date().toISOString(),
    backupFileId,
    restoreType,
    (sheetsRestored || []).join(', '),
    restoredBy,
    reason || '',
    snapshotId || '',
    status || 'Failed',
    JSON.stringify(issues || [])
  ]);
  return restoreId;
}

function forceLogoutAllUsers_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('force_logout_after', new Date().toISOString());
}

// API wrappers
function api_createDailyIncremental(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return createDailyIncremental();
}

function api_testBackupRestore(token, backupId) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return testBackupRestore(backupId);
}

function api_listBackupTests(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return { success: true, tests: listBackupTests_() };
}

function api_listRestoreHistory(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return { success: true, restores: listRestoreHistory_() };
}

function api_listSnapshots(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return { success: true, snapshots: listSnapshots_() };
}

function api_previewRestore(backupId, sheetsToRestore, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  const preview = buildRestorePreview_(backupId, sheetsToRestore || []);
  return { success: true, preview: preview };
}

function api_restoreFromBackup(backupId, restoreOptions, confirmedBy, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  const options = restoreOptions || {};
  options.token = token;
  return restoreFromBackup(backupId, options, confirmedBy || session.session.user_name);
}

function api_rollbackRestore(restoreId, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return rollbackRestore(restoreId, session.session.user_name || 'Operator');
}

// Test function
function testDisasterRecovery() {
  console.log('=== Disaster Recovery Test ===');
  const incremental = createDailyIncremental();
  console.log('Incremental:', incremental.success);

  const test = testBackupRestore();
  console.log('Backup test:', test.status || test.success);

  const latestBackup = getLatestBackupByType_(['automatic', 'manual']);
  const preview = latestBackup ? buildRestorePreview_(latestBackup.file_id, ['Settings']) : [];
  console.log('Preview count:', preview.length);

  const schedule = scheduleBackupTests();
  console.log('Schedule test trigger:', schedule.success);

  return {
    incremental: incremental,
    test: test,
    preview_count: preview.length,
    schedule: schedule
  };
}
