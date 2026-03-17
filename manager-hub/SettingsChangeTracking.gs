/**
 * SettingsChangeTracking.gs
 * Micro-Phase 35: Change Tracking & Version Control for Settings
 */

const SETTINGS_HISTORY_HEADERS = [
  'History_ID',
  'Change_Date',
  'Changed_By',
  'Setting_Category',
  'Setting_Name',
  'Old_Value',
  'New_Value',
  'Change_Reason',
  'Approved_By',
  'Rollback_Available'
];

const SETTINGS_SNAPSHOT_HEADERS = [
  'Snapshot_ID',
  'Snapshot_Date',
  'Snapshot_Type',
  'Created_By',
  'Description',
  'Settings_JSON',
  'Active'
];

function getOrCreateSettingsHistorySheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Settings_History');
  if (!sheet) {
    sheet = ss.insertSheet('Settings_History');
    sheet.getRange(1, 1, 1, SETTINGS_HISTORY_HEADERS.length).setValues([SETTINGS_HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, SETTINGS_HISTORY_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateSettingsSnapshotsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Settings_Snapshots');
  if (!sheet) {
    sheet = ss.insertSheet('Settings_Snapshots');
    sheet.getRange(1, 1, 1, SETTINGS_SNAPSHOT_HEADERS.length).setValues([SETTINGS_SNAPSHOT_HEADERS]);
    sheet.getRange(1, 1, 1, SETTINGS_SNAPSHOT_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function captureSettingsSnapshot(snapshot_type, created_by, description) {
  try {
    const settingsObj = buildSettingsSnapshotObject_();
    if (!settingsObj.success) return settingsObj;

    const sheet = getOrCreateSettingsSnapshotsSheet_();
    const snapshotId = 'SNP' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

    // Mark previous snapshots inactive
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).setValue('FALSE');
    }

    sheet.appendRow([
      snapshotId,
      new Date().toISOString(),
      snapshot_type || 'Manual',
      created_by || 'System',
      description || '',
      JSON.stringify(settingsObj.settings),
      'TRUE'
    ]);

    return { success: true, snapshot_id: snapshotId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function buildSettingsSnapshotObject_() {
  const base = getSystemSettings();
  if (!base.success) return base;
  const templates = getEmailTemplatesForSnapshot_();

  return {
    success: true,
    settings: {
      email: base.settings.email || {},
      passwords: base.settings.passwords || {},
      thresholds: base.settings.thresholds || [],
      buckets: base.settings.buckets || [],
      system: base.settings.system || {},
      dataQuality: base.settings.dataQuality || {},
      links: base.settings.links || [],
      emailTemplates: templates || []
    }
  };
}

function getEmailTemplatesForSnapshot_() {
  const section = ensureEmailTemplatesSection();
  if (!section.success) return [];
  const sheet = section.sheet;
  const startRow = section.startRow;
  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(0, lastRow - startRow + 1);
  if (rowCount === 0) return [];
  const data = sheet.getRange(startRow, 1, rowCount, 7).getValues();
  const templates = [];
  data.forEach(row => {
    const id = String(row[0] || '').trim();
    if (!id) return;
    templates.push({
      template_id: id,
      template_name: row[1] || '',
      enabled: String(row[2] || '').toUpperCase() === 'TRUE',
      subject_template: row[3] || '',
      body_template: row[4] || '',
      updated_by: row[5] || '',
      updated_at: row[6] || ''
    });
  });
  return templates;
}

function trackSettingChange(category, setting_name, old_value, new_value, changed_by, reason) {
  try {
    const sheet = getOrCreateSettingsHistorySheet_();
    const historyId = 'HIST' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const rollbackAvailable = (old_value !== undefined && old_value !== null);

    const safeOld = old_value === undefined ? '' : String(old_value);
    const safeNew = new_value === undefined ? '' : String(new_value);
    const safeReason = reason || '';

    const entry = [
      historyId,
      new Date().toISOString(),
      changed_by || 'Unknown',
      category || '',
      setting_name || '',
      safeOld,
      safeNew,
      safeReason,
      '',
      rollbackAvailable ? 'TRUE' : 'FALSE'
    ];
    sheet.appendRow(entry);

    const major = ['Thresholds', 'Buckets', 'Email'].indexOf(String(category)) !== -1;
    if (major) {
      captureSettingsSnapshot('Pre_Change', changed_by || 'System', category + ' change: ' + setting_name);
      sendSettingsChangeEmail_(setting_name, safeOld, safeNew, changed_by, safeReason);
    }

    return { success: true, history_id: historyId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function sendSettingsChangeEmail_(settingName, oldValue, newValue, changedBy, reason) {
  try {
    const recipients = getDirectorEmails();
    if (!recipients || recipients.length === 0) return;
    const subject = 'System Settings Changed: ' + settingName;
    const body = [
      'A system setting has been updated.',
      'Changed by: ' + (changedBy || 'Unknown'),
      'Date: ' + new Date().toISOString(),
      'Setting: ' + (settingName || ''),
      'Old Value: ' + (oldValue || ''),
      'New Value: ' + (newValue || ''),
      'Reason: ' + (reason || '')
    ].join('\n');
    MailApp.sendEmail(recipients.join(','), subject, body);
  } catch (error) {
    console.error('Settings change email failed:', error.toString());
  }
}

function getSettingsHistory(filter_options) {
  const sheet = getOrCreateSettingsHistorySheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, history: [] };
  const headers = data[0];
  let rows = data.slice(1).map(row => {
    const item = {};
    headers.forEach((h, i) => item[h] = row[i]);
    return item;
  });

  const filters = filter_options || {};
  if (filters.category) {
    rows = rows.filter(r => String(r.Setting_Category) === String(filters.category));
  }
  if (filters.changed_by) {
    rows = rows.filter(r => String(r.Changed_By) === String(filters.changed_by));
  }
  if (filters.setting_name) {
    rows = rows.filter(r => String(r.Setting_Name) === String(filters.setting_name));
  }
  if (filters.date_range && (filters.date_range.from || filters.date_range.to)) {
    const from = filters.date_range.from ? new Date(filters.date_range.from) : null;
    const to = filters.date_range.to ? new Date(filters.date_range.to) : null;
    rows = rows.filter(r => {
      const d = new Date(r.Change_Date);
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
  }

  rows.sort((a, b) => new Date(b.Change_Date) - new Date(a.Change_Date));
  return { success: true, history: rows };
}

function compareSettingsVersions(snapshot_id_1, snapshot_id_2) {
  const snapshot1 = getSnapshotById_(snapshot_id_1);
  const snapshot2 = getSnapshotById_(snapshot_id_2);
  if (!snapshot1 || !snapshot2) return { success: false, message: 'Snapshot not found' };

  const settings1 = JSON.parse(snapshot1.Settings_JSON || '{}');
  const settings2 = JSON.parse(snapshot2.Settings_JSON || '{}');
  const diffs = [];

  Object.keys(settings1).forEach(category => {
    const a = settings1[category];
    const b = settings2[category];
    if (JSON.stringify(a) !== JSON.stringify(b)) {
      diffs.push({
        category: category,
        setting_name: category,
        old_value: JSON.stringify(a),
        new_value: JSON.stringify(b)
      });
    }
  });

  return { success: true, differences: diffs };
}

function rollbackToSnapshot(snapshot_id, rollback_by, reason, preview_only) {
  const role = getCurrentRole();
  if (!role.authenticated || role.role !== 'Operator') {
    return { success: false, message: 'Only Operators can rollback settings' };
  }

  const target = getSnapshotById_(snapshot_id);
  if (!target) return { success: false, message: 'Snapshot not found' };
  const settings = JSON.parse(target.Settings_JSON || '{}');

  if (preview_only) {
    const current = buildSettingsSnapshotObject_();
    if (!current.success) return current;
    const preview = compareSettingsObjects_(current.settings, settings);
    return { success: true, preview: preview, preview_only: true };
  }

  const preSnapshot = captureSettingsSnapshot('Pre_Rollback', rollback_by, 'Pre rollback snapshot');

  const applyResult = applySettingsObject_(settings, rollback_by || 'Operator', reason || 'Rollback');
  if (!applyResult.success) return applyResult;

  setActiveSnapshot_(snapshot_id);

  trackSettingChange('Rollback', 'Rollback to snapshot', '', snapshot_id, rollback_by, reason || 'Rollback');

  sendSettingsRollbackEmail_(rollback_by, snapshot_id, reason);

  return {
    success: true,
    snapshot_id: snapshot_id,
    pre_rollback_snapshot: preSnapshot.snapshot_id || ''
  };
}

function compareSettingsObjects_(current, target) {
  const diffs = [];
  Object.keys(target || {}).forEach(key => {
    const a = current[key];
    const b = target[key];
    if (JSON.stringify(a) !== JSON.stringify(b)) {
      diffs.push({ category: key, old_value: a, new_value: b });
    }
  });
  return diffs;
}

function applySettingsObject_(settings, changedBy, reason) {
  try {
    updateEmailSettings(settings.email || {}, { changedBy: changedBy, reason: reason });
    updateThresholdSettings(settings.thresholds || [], { changedBy: changedBy, reason: reason });
    updateBucketSettings(settings.buckets || [], { changedBy: changedBy, reason: reason });
    updateSystemConfig(settings.system || {}, { changedBy: changedBy, reason: reason });
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getSnapshotById_(snapshotId) {
  const sheet = getOrCreateSettingsSnapshotsSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  for (const row of data) {
    const item = {};
    headers.forEach((h, i) => item[h] = row[i]);
    if (String(item.Snapshot_ID) === String(snapshotId)) return item;
  }
  return null;
}

function setActiveSnapshot_(snapshotId) {
  const sheet = getOrCreateSettingsSnapshotsSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  for (let i = 1; i < data.length; i++) {
    const active = String(data[i][0]) === String(snapshotId) ? 'TRUE' : 'FALSE';
    sheet.getRange(i + 1, 7).setValue(active);
  }
}

function getThresholdTimeline() {
  const history = getSettingsHistory({ category: 'Thresholds' });
  if (!history.success) return history;
  const timeline = history.history.map(item => ({
    date: item.Change_Date,
    changed_by: item.Changed_By,
    setting: item.Setting_Name,
    old_value: item.Old_Value,
    new_value: item.New_Value,
    reason: item.Change_Reason
  }));
  return { success: true, timeline: timeline };
}

function sendSettingsRollbackEmail_(rollbackBy, snapshotId, reason) {
  try {
    const recipients = getDirectorEmails();
    if (!recipients || recipients.length === 0) return;
    const subject = 'Settings rolled back';
    const body = [
      'Settings were rolled back.',
      'Rollback by: ' + (rollbackBy || 'Operator'),
      'Snapshot: ' + snapshotId,
      'Reason: ' + (reason || '')
    ].join('\n');
    MailApp.sendEmail(recipients.join(','), subject, body);
  } catch (error) {
    console.error('Rollback email failed:', error.toString());
  }
}

function getSettingsSnapshots() {
  const sheet = getOrCreateSettingsSnapshotsSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, snapshots: [] };
  const headers = data[0];
  const snapshots = data.slice(1).map(row => {
    const item = {};
    headers.forEach((h, i) => item[h] = row[i]);
    return item;
  });
  snapshots.sort((a, b) => new Date(b.Snapshot_Date) - new Date(a.Snapshot_Date));
  return { success: true, snapshots: snapshots };
}

function scheduleMonthlySettingsSnapshot() {
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.find(t => t.getHandlerFunction() === 'runMonthlySettingsSnapshot');
  if (existing) return { success: true, message: 'Trigger already exists' };
  const trigger = ScriptApp.newTrigger('runMonthlySettingsSnapshot')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
  return { success: true, trigger_id: trigger.getUniqueId() };
}

function runMonthlySettingsSnapshot() {
  const label = 'Monthly_Snapshot_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  captureSettingsSnapshot('Scheduled', 'System', label);
}

// API wrappers
function api_captureSettingsSnapshot(type, createdBy, description, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return captureSettingsSnapshot(type, createdBy || session.session.user_name, description);
}

function api_getSettingsHistory(filterOptions, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return getSettingsHistory(filterOptions);
}

function api_getSettingsSnapshots(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return getSettingsSnapshots();
}

function api_compareSettingsVersions(snapshotId1, snapshotId2, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return compareSettingsVersions(snapshotId1, snapshotId2);
}

function api_getSnapshotById(snapshotId, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  const snap = getSnapshotById_(snapshotId);
  if (!snap) return { success: false, message: 'Snapshot not found' };
  return { success: true, snapshot: snap };
}

function api_rollbackToSnapshot(snapshotId, rollbackBy, reason, previewOnly, token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return rollbackToSnapshot(snapshotId, rollbackBy || session.session.user_name, reason || '', !!previewOnly);
}

function api_getThresholdTimeline(token) {
  const session = requireOperatorSession_(token);
  if (!session.valid) return session.sessionExpired ? { success: false, sessionExpired: true } : { success: false, error: session.error };
  return getThresholdTimeline();
}

function testChangeTracking() {
  const snapshot = captureSettingsSnapshot('Manual', 'Test Script', 'Initial snapshot');
  const thresholds = getSystemSettings().settings.thresholds;
  if (thresholds && thresholds.length) {
    const old = thresholds[0].points;
    thresholds[0].points = old + 1;
    updateThresholdSettings(thresholds, { changedBy: 'Test Script', reason: 'Test change' });
  }
  const history = getSettingsHistory({ category: 'Thresholds' });
  const snapshots = getSettingsSnapshots();
  const timeline = getThresholdTimeline();
  return { snapshot: snapshot, history: history, snapshots: snapshots, timeline: timeline };
}
