// ============================================
// PHASE 4: SETTINGS MANAGEMENT
// ============================================
// Functions for managing system configuration
// via the Settings sheet in the Manager Hub
// ============================================

/**
 * Gets all system settings from the Settings sheet.
 * Returns structured data for the Settings UI page.
 * @returns {Object} Complete settings object
 */
function getSystemSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    ensureDataQualitySettingsRows(settingsSheet);
    ensureBreakFoodThresholdRow(settingsSheet);

    // Get all data at once for efficiency
    const data = settingsSheet.getRange('A1:C80').getValues();

    // Parse Email Configuration (rows 2-3)
    const emailSettings = {
      storeEmail: data[1][1] || '',
      terminationEmailList: data[2][1] || ''
    };

    // Parse Password Configuration (rows 6-8)
    const passwordSettings = {
      operatorPassword: data[5][1] || '',
      directorPassword: data[6][1] || '',
      managerPassword: data[7][1] || ''
    };

    // Parse Point Thresholds (rows 13-19)
    const thresholds = [];
    for (let i = 12; i <= 18; i++) {
      if (data[i][0] !== '' && data[i][0] !== null) {
        thresholds.push({
          points: Number(data[i][0]) || 0,
          consequence: data[i][1] || ''
        });
      }
    }

    // Parse Infraction Buckets (rows 23-27)
    const buckets = [];
    for (let i = 22; i <= 26; i++) {
      if (data[i][0] !== '' && data[i][0] !== null) {
        let examples = [];
        try {
          if (data[i][2]) {
            examples = JSON.parse(data[i][2]);
          }
        } catch (e) {
          console.log('Error parsing bucket examples:', e);
        }

        buckets.push({
          name: data[i][0] || '',
          pointValue: Number(data[i][1]) || 0,
          examples: examples
        });
      }
    }

    // Parse System Configuration (rows 30-37)
    const systemConfig = {
      sessionTimeout: Number(data[29][1]) || 10,
      maxLoginAttempts: Number(data[30][1]) || 3,
      backdateLimit: Number(data[31][1]) || 7,
      maxNegativePoints: Number(data[32][1]) || -6,
      pointExpiration: Number(data[33][1]) || 90,
      probationDuration: Number(data[34][1]) || 30,
      lastBackupDate: data[35][1] || '',
      breakFoodThreshold: Number(data[36][1]) || 9
    };

    const dataQuality = getDataQualitySettingsFromSheet(settingsSheet);

    // Get Links from Links sheet
    const links = getLinksData(ss);

    return {
      success: true,
      settings: {
        email: emailSettings,
        passwords: passwordSettings,
        thresholds: thresholds,
        buckets: buckets,
        system: systemConfig,
        dataQuality: dataQuality,
        links: links
      }
    };

  } catch (error) {
    console.error('Error getting system settings:', error.toString());
    return { success: false, message: 'Error loading settings: ' + error.message };
  }
}

/**
 * Ensures Data Quality settings rows exist in Settings sheet.
 */
function ensureDataQualitySettingsRows(settingsSheet) {
  try {
    if (!settingsSheet) return;
    const headerRow = findSettingsRowByLabel(settingsSheet, 'Data Quality');
    if (headerRow > 0) return;

    const dqActions = JSON.stringify({
      standardize_names: true,
      remove_duplicate_logs: true,
      archive_old_signups: true,
      expire_old_tokens: true
    });

    const startRow = settingsSheet.getLastRow() + 1;
    const rows = [
      ['', '', ''],
      ['Data Quality', '', ''],
      ['Automatic Cleanup Frequency', 'Weekly', ''],
      ['Cleanup Actions', dqActions, ''],
      ['Email Report After Cleanup', 'TRUE', ''],
      ['Last Cleanup Date', '', ''],
      ['Next Cleanup Date', '', '']
    ];
    settingsSheet.getRange(startRow, 1, rows.length, 3).setValues(rows);

    if (typeof formatSectionHeader === 'function') {
      formatSectionHeader(settingsSheet, startRow + 1, 'Data Quality', 3);
    }
  } catch (error) {
    console.error('Error ensuring Data Quality settings rows:', error.toString());
  }
}

function ensureBreakFoodThresholdRow(settingsSheet) {
  try {
    if (!settingsSheet) return;
    const label = settingsSheet.getRange('A37').getValue();
    if (!label) {
      settingsSheet.getRange('A37').setValue('Break Food Threshold (points)');
      settingsSheet.getRange('B37').setValue(9);
    }
  } catch (error) {
    console.error('ensureBreakFoodThresholdRow error:', error.toString());
  }
}

/**
 * Reads Data Quality settings from Settings sheet.
 */
function getDataQualitySettingsFromSheet(settingsSheet) {
  if (!settingsSheet) return {};
  const frequencyRow = findSettingsRowByLabel(settingsSheet, 'Automatic Cleanup Frequency');
  const actionsRow = findSettingsRowByLabel(settingsSheet, 'Cleanup Actions');
  const emailRow = findSettingsRowByLabel(settingsSheet, 'Email Report After Cleanup');
  const lastRunRow = findSettingsRowByLabel(settingsSheet, 'Last Cleanup Date');
  const nextRunRow = findSettingsRowByLabel(settingsSheet, 'Next Cleanup Date');

  let actions = {
    standardize_names: true,
    remove_duplicate_logs: true,
    archive_old_signups: true,
    expire_old_tokens: true
  };
  if (actionsRow > 0) {
    const raw = settingsSheet.getRange(actionsRow, 2).getValue();
    try {
      if (raw) actions = JSON.parse(raw);
    } catch (error) {
      console.warn('Invalid cleanup actions JSON, using defaults.');
    }
  }

  return {
    frequency: frequencyRow > 0 ? (settingsSheet.getRange(frequencyRow, 2).getValue() || 'Weekly') : 'Weekly',
    actions: actions,
    emailReport: emailRow > 0 ? String(settingsSheet.getRange(emailRow, 2).getValue() || 'TRUE') === 'TRUE' : true,
    lastCleanupDate: lastRunRow > 0 ? settingsSheet.getRange(lastRunRow, 2).getValue() : '',
    nextCleanupDate: nextRunRow > 0 ? settingsSheet.getRange(nextRunRow, 2).getValue() : ''
  };
}

/**
 * Updates Data Quality settings in the Settings sheet.
 */
function updateDataQualitySettings(config) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    ensureDataQualitySettingsRows(settingsSheet);

    const frequencyRow = findSettingsRowByLabel(settingsSheet, 'Automatic Cleanup Frequency');
    const actionsRow = findSettingsRowByLabel(settingsSheet, 'Cleanup Actions');
    const emailRow = findSettingsRowByLabel(settingsSheet, 'Email Report After Cleanup');
    const nextRunRow = findSettingsRowByLabel(settingsSheet, 'Next Cleanup Date');

    if (config.frequency && frequencyRow > 0) {
      settingsSheet.getRange(frequencyRow, 2).setValue(config.frequency);
    }
    if (config.actions && actionsRow > 0) {
      settingsSheet.getRange(actionsRow, 2).setValue(JSON.stringify(config.actions));
    }
    if (typeof config.emailReport !== 'undefined' && emailRow > 0) {
      settingsSheet.getRange(emailRow, 2).setValue(config.emailReport ? 'TRUE' : 'FALSE');
    }
    if (config.nextCleanupDate && nextRunRow > 0) {
      settingsSheet.getRange(nextRunRow, 2).setValue(config.nextCleanupDate);
    }

    logSettingsChange('Data Quality Settings', 'Updated data quality cleanup configuration');

    return { success: true, message: 'Data quality settings updated' };
  } catch (error) {
    console.error('Error updating data quality settings:', error.toString());
    return { success: false, message: 'Error updating data quality settings: ' + error.message };
  }
}

/**
 * Gets links data from the Links sheet.
 * @param {Spreadsheet} ss - The spreadsheet object
 * @returns {Array} Array of link objects
 */
function getLinksData(ss) {
  try {
    if (typeof getAllLinks === 'function') {
      const result = getAllLinks(true);
      if (result && result.success && Array.isArray(result.links)) {
        return result.links.map(link => ({
          label: link.title || link.name || '',
          url: link.url || '',
          category: link.category || 'General'
        })).filter(link => link.label && link.url);
      }
    }

    const linksSheet = ss.getSheetByName('Link_Management') || ss.getSheetByName('Links');
    if (!linksSheet) return [];

    const lastRow = linksSheet.getLastRow();
    if (lastRow < 2) return [];

    const data = linksSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const normalizedHeaders = headers.map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
    const idxName = normalizedHeaders.indexOf('linkname') !== -1 ? normalizedHeaders.indexOf('linkname') : normalizedHeaders.indexOf('title');
    const idxUrl = normalizedHeaders.indexOf('linkurl') !== -1 ? normalizedHeaders.indexOf('linkurl') : normalizedHeaders.indexOf('url');
    const idxCategory = normalizedHeaders.indexOf('category');

    const links = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const label = idxName >= 0 ? row[idxName] : '';
      const url = idxUrl >= 0 ? row[idxUrl] : '';
      if (label && url) {
        links.push({
          label: label,
          url: url,
          category: idxCategory >= 0 ? row[idxCategory] : 'General'
        });
      }
    }

    return links;
  } catch (error) {
    console.error('Error getting links:', error.toString());
    return [];
  }
}

/**
 * Updates email settings in the Settings sheet.
 * @param {Object} emailData - Email configuration object
 * @returns {Object} Success/failure result
 */
function updateEmailSettings(emailData, meta) {
  try {
    const session = getCurrentRole();
    const changedBy = (meta && meta.changedBy) || (session.user_name || session.role || 'Unknown');
    const reason = (meta && meta.reason) || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    // Validate email format
    if (emailData.storeEmail && !isValidEmail(emailData.storeEmail)) {
      return { success: false, message: 'Invalid store email format' };
    }

    const oldStoreEmail = settingsSheet.getRange('B2').getValue();
    const oldTerminationList = settingsSheet.getRange('B3').getValue();

    // Update email settings (rows 2-3, column B)
    if (emailData.storeEmail !== undefined) {
      settingsSheet.getRange('B2').setValue(emailData.storeEmail);
      if (String(oldStoreEmail) !== String(emailData.storeEmail)) {
        trackSettingChange(
          'Email',
          'store_email',
          oldStoreEmail,
          emailData.storeEmail,
          changedBy,
          reason
        );
      }
    }

    if (emailData.terminationEmailList !== undefined) {
      // Validate each email in the list
      if (emailData.terminationEmailList) {
        const emails = emailData.terminationEmailList.split(',').map(e => e.trim());
        for (const email of emails) {
          if (email && !isValidEmail(email)) {
            return { success: false, message: `Invalid email in termination list: ${email}` };
          }
        }
      }
      settingsSheet.getRange('B3').setValue(emailData.terminationEmailList);
      if (String(oldTerminationList) !== String(emailData.terminationEmailList)) {
        trackSettingChange(
          'Email',
          'termination_email_list',
          oldTerminationList,
          emailData.terminationEmailList,
          changedBy,
          reason
        );
      }
    }

    // Log the change
    logSettingsChange('Email Settings', 'Updated email configuration');

    return { success: true, message: 'Email settings updated successfully' };

  } catch (error) {
    console.error('Error updating email settings:', error.toString());
    return { success: false, message: 'Error updating email settings: ' + error.message };
  }
}

/**
 * Validates email format.
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Updates password settings in the Settings sheet.
 * Only Operators can update passwords.
 * @param {Object} passwordData - Password configuration object
 * @returns {Object} Success/failure result
 */
function updatePasswordSettings(passwordData, meta) {
  try {
    const session = getCurrentRole();
    const changedBy = (meta && meta.changedBy) || (session.user_name || session.role || 'Unknown');
    const reason = (meta && meta.reason) || '';
    // Check if current user has Operator role
    const roleResult = getCurrentRole();
    if (!roleResult.authenticated || roleResult.role !== 'Operator') {
      return { success: false, message: 'Only Operators can change passwords' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    // Validate password requirements
    const minLength = 4;

    if (passwordData.operatorPassword !== undefined) {
      if (passwordData.operatorPassword.length < minLength) {
        return { success: false, message: `Operator password must be at least ${minLength} characters` };
      }
      settingsSheet.getRange('B6').setValue(passwordData.operatorPassword);
      trackSettingChange('Passwords', 'operatorPassword', '***', '***',
        changedBy, reason);
    }

    if (passwordData.directorPassword !== undefined) {
      if (passwordData.directorPassword.length < minLength) {
        return { success: false, message: `Director password must be at least ${minLength} characters` };
      }
      settingsSheet.getRange('B7').setValue(passwordData.directorPassword);
      trackSettingChange('Passwords', 'directorPassword', '***', '***',
        changedBy, reason);
    }

    if (passwordData.managerPassword !== undefined) {
      if (passwordData.managerPassword.length < minLength) {
        return { success: false, message: `Manager password must be at least ${minLength} characters` };
      }
      settingsSheet.getRange('B8').setValue(passwordData.managerPassword);
      trackSettingChange('Passwords', 'managerPassword', '***', '***',
        changedBy, reason);
    }

    // Log the change (don't log actual passwords)
    logSettingsChange('Password Settings', 'Passwords updated');

    return { success: true, message: 'Password settings updated successfully' };

  } catch (error) {
    console.error('Error updating password settings:', error.toString());
    return { success: false, message: 'Error updating password settings: ' + error.message };
  }
}

/**
 * Updates threshold settings in the Settings sheet.
 * @param {Array} thresholds - Array of threshold objects
 * @returns {Object} Success/failure result
 */
function updateThresholdSettings(thresholds, meta) {
  try {
    const session = getCurrentRole();
    const changedBy = (meta && meta.changedBy) || (session.user_name || session.role || 'Unknown');
    const reason = (meta && meta.reason) || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    // Validate thresholds
    if (!Array.isArray(thresholds) || thresholds.length === 0) {
      return { success: false, message: 'Invalid threshold data' };
    }

    // Capture old values
    const oldRows = settingsSheet.getRange(13, 1, 7, 2).getValues();

    // Build the data array
    const data = thresholds.map(t => [
      Number(t.points) || 0,
      t.consequence || ''
    ]);

    // Clear existing threshold rows (13-19) and write new ones
    const startRow = 13;
    const numRows = 7;

    // Clear the range first
    settingsSheet.getRange(startRow, 1, numRows, 2).clearContent();

    // Write the new data (only as many rows as we have)
    if (data.length > 0) {
      settingsSheet.getRange(startRow, 1, Math.min(data.length, numRows), 2).setValues(data.slice(0, numRows));
    }

    // Track individual changes
    data.slice(0, numRows).forEach((row, index) => {
      const oldRow = oldRows[index] || [];
      if (String(oldRow[0]) !== String(row[0]) || String(oldRow[1]) !== String(row[1])) {
        trackSettingChange(
          'Thresholds',
          'threshold_' + row[0],
          JSON.stringify({ points: oldRow[0], consequence: oldRow[1] }),
          JSON.stringify({ points: row[0], consequence: row[1] }),
          changedBy,
          reason
        );
      }
    });

    // Log the change
    logSettingsChange('Point Thresholds', `Updated ${data.length} thresholds`);

    return { success: true, message: 'Threshold settings updated successfully' };

  } catch (error) {
    console.error('Error updating threshold settings:', error.toString());
    return { success: false, message: 'Error updating threshold settings: ' + error.message };
  }
}

/**
 * Updates bucket settings in the Settings sheet.
 * @param {Array} buckets - Array of bucket objects
 * @returns {Object} Success/failure result
 */
function updateBucketSettings(buckets, meta) {
  try {
    const session = getCurrentRole();
    const changedBy = (meta && meta.changedBy) || (session.user_name || session.role || 'Unknown');
    const reason = (meta && meta.reason) || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    // Validate buckets
    if (!Array.isArray(buckets) || buckets.length === 0) {
      return { success: false, message: 'Invalid bucket data' };
    }

    // Validate bucket structure
    for (const bucket of buckets) {
      if (!bucket.name || bucket.pointValue === undefined) {
        return { success: false, message: 'Each bucket must have a name and point value' };
      }
      if (typeof bucket.pointValue !== 'number' || bucket.pointValue < 0) {
        return { success: false, message: 'Point values must be non-negative numbers' };
      }
    }

    const oldRows = settingsSheet.getRange(23, 1, 5, 3).getValues();

    // Build the data array
    const data = buckets.map(b => [
      b.name,
      b.pointValue,
      JSON.stringify(b.examples || [])
    ]);

    // Write to rows 23-27 (5 buckets)
    const startRow = 23;
    const numRows = 5;

    // Clear existing bucket rows
    settingsSheet.getRange(startRow, 1, numRows, 3).clearContent();

    // Write the new data
    if (data.length > 0) {
      settingsSheet.getRange(startRow, 1, Math.min(data.length, numRows), 3).setValues(data.slice(0, numRows));
    }

    data.slice(0, numRows).forEach((row, index) => {
      const oldRow = oldRows[index] || [];
      if (String(oldRow[0]) !== String(row[0]) ||
          String(oldRow[1]) !== String(row[1]) ||
          String(oldRow[2]) !== String(row[2])) {
        trackSettingChange(
          'Buckets',
          'bucket_' + (index + 1),
          JSON.stringify({ name: oldRow[0], points: oldRow[1], examples: oldRow[2] }),
          JSON.stringify({ name: row[0], points: row[1], examples: row[2] }),
          changedBy,
          reason
        );
      }
    });

    // Log the change
    logSettingsChange('Infraction Buckets', `Updated ${data.length} buckets`);

    return { success: true, message: 'Bucket settings updated successfully' };

  } catch (error) {
    console.error('Error updating bucket settings:', error.toString());
    return { success: false, message: 'Error updating bucket settings: ' + error.message };
  }
}

/**
 * Adds an example to a specific bucket.
 * @param {number} bucketIndex - Index of the bucket (0-4)
 * @param {string} example - Example text to add
 * @returns {Object} Success/failure result
 */
function addInfractionExample(bucketIndex, example) {
  try {
    if (bucketIndex < 0 || bucketIndex > 4) {
      return { success: false, message: 'Invalid bucket index' };
    }

    if (!example || example.trim() === '') {
      return { success: false, message: 'Example cannot be empty' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    const row = 23 + bucketIndex;
    const currentExamples = settingsSheet.getRange(row, 3).getValue();

    let examples = [];
    try {
      if (currentExamples) {
        examples = JSON.parse(currentExamples);
      }
    } catch (e) {
      examples = [];
    }

    // Check for duplicates
    if (examples.includes(example.trim())) {
      return { success: false, message: 'This example already exists' };
    }

    examples.push(example.trim());
    settingsSheet.getRange(row, 3).setValue(JSON.stringify(examples));

    // Log the change
    const bucketName = settingsSheet.getRange(row, 1).getValue();
    logSettingsChange('Bucket Examples', `Added example to ${bucketName}`);

    return { success: true, message: 'Example added successfully', examples: examples };

  } catch (error) {
    console.error('Error adding example:', error.toString());
    return { success: false, message: 'Error adding example: ' + error.message };
  }
}

/**
 * Removes an example from a specific bucket.
 * @param {number} bucketIndex - Index of the bucket (0-4)
 * @param {number} exampleIndex - Index of the example to remove
 * @returns {Object} Success/failure result
 */
function removeInfractionExample(bucketIndex, exampleIndex) {
  try {
    if (bucketIndex < 0 || bucketIndex > 4) {
      return { success: false, message: 'Invalid bucket index' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    const row = 23 + bucketIndex;
    const currentExamples = settingsSheet.getRange(row, 3).getValue();

    let examples = [];
    try {
      if (currentExamples) {
        examples = JSON.parse(currentExamples);
      }
    } catch (e) {
      return { success: false, message: 'Error parsing current examples' };
    }

    if (exampleIndex < 0 || exampleIndex >= examples.length) {
      return { success: false, message: 'Invalid example index' };
    }

    const removed = examples.splice(exampleIndex, 1);
    settingsSheet.getRange(row, 3).setValue(JSON.stringify(examples));

    // Log the change
    const bucketName = settingsSheet.getRange(row, 1).getValue();
    logSettingsChange('Bucket Examples', `Removed example from ${bucketName}`);

    return { success: true, message: 'Example removed successfully', examples: examples };

  } catch (error) {
    console.error('Error removing example:', error.toString());
    return { success: false, message: 'Error removing example: ' + error.message };
  }
}

/**
 * Updates system configuration settings.
 * @param {Object} config - System configuration object
 * @returns {Object} Success/failure result
 */
function updateSystemConfig(config, meta) {
  try {
    const session = getCurrentRole();
    const changedBy = (meta && meta.changedBy) || (session.user_name || session.role || 'Unknown');
    const reason = (meta && meta.reason) || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return { success: false, message: 'Settings sheet not found' };
    }

    const oldValues = {
      sessionTimeout: settingsSheet.getRange('B30').getValue(),
      maxLoginAttempts: settingsSheet.getRange('B31').getValue(),
      backdateLimit: settingsSheet.getRange('B32').getValue(),
      maxNegativePoints: settingsSheet.getRange('B33').getValue(),
      pointExpiration: settingsSheet.getRange('B34').getValue(),
      probationDuration: settingsSheet.getRange('B35').getValue(),
      breakFoodThreshold: settingsSheet.getRange('B37').getValue()
    };

    // Update each config value if provided
    if (config.sessionTimeout !== undefined) {
      const timeout = Number(config.sessionTimeout);
      if (timeout < 1 || timeout > 60) {
        return { success: false, message: 'Session timeout must be between 1 and 60 minutes' };
      }
      settingsSheet.getRange('B30').setValue(timeout);
      if (String(oldValues.sessionTimeout) !== String(timeout)) {
        trackSettingChange('System_Config', 'sessionTimeout', oldValues.sessionTimeout, timeout,
          changedBy, reason);
      }
    }

    if (config.maxLoginAttempts !== undefined) {
      const attempts = Number(config.maxLoginAttempts);
      if (attempts < 1 || attempts > 10) {
        return { success: false, message: 'Max login attempts must be between 1 and 10' };
      }
      settingsSheet.getRange('B31').setValue(attempts);
      if (String(oldValues.maxLoginAttempts) !== String(attempts)) {
        trackSettingChange('System_Config', 'maxLoginAttempts', oldValues.maxLoginAttempts, attempts,
          changedBy, reason);
      }
    }

    if (config.backdateLimit !== undefined) {
      const limit = Number(config.backdateLimit);
      if (limit < -30 || limit > 30) {
        return { success: false, message: 'Backdate limit must be between -30 and 30 days' };
      }
      settingsSheet.getRange('B32').setValue(limit);
      if (String(oldValues.backdateLimit) !== String(limit)) {
        trackSettingChange('System_Config', 'backdateLimit', oldValues.backdateLimit, limit,
          changedBy, reason);
      }
    }

    if (config.maxNegativePoints !== undefined) {
      const points = Number(config.maxNegativePoints);
      if (points > 0 || points < -20) {
        return { success: false, message: 'Max negative points must be between -20 and 0' };
      }
      settingsSheet.getRange('B33').setValue(points);
      if (String(oldValues.maxNegativePoints) !== String(points)) {
        trackSettingChange('System_Config', 'maxNegativePoints', oldValues.maxNegativePoints, points,
          changedBy, reason);
      }
    }

    if (config.pointExpiration !== undefined) {
      const expiration = Number(config.pointExpiration);
      if (expiration < 30 || expiration > 365) {
        return { success: false, message: 'Point expiration must be between 30 and 365 days' };
      }
      settingsSheet.getRange('B34').setValue(expiration);
      if (String(oldValues.pointExpiration) !== String(expiration)) {
        trackSettingChange('System_Config', 'pointExpiration', oldValues.pointExpiration, expiration,
          changedBy, reason);
      }
    }

    if (config.probationDuration !== undefined) {
      const duration = Number(config.probationDuration);
      if (duration < 7 || duration > 90) {
        return { success: false, message: 'Probation duration must be between 7 and 90 days' };
      }
      settingsSheet.getRange('B35').setValue(duration);
      if (String(oldValues.probationDuration) !== String(duration)) {
        trackSettingChange('System_Config', 'probationDuration', oldValues.probationDuration, duration,
          changedBy, reason);
      }
    }

    if (config.breakFoodThreshold !== undefined) {
      const threshold = Number(config.breakFoodThreshold);
      if (threshold < 1 || threshold > 50) {
        return { success: false, message: 'Break food threshold must be between 1 and 50 points' };
      }
      settingsSheet.getRange('B37').setValue(threshold);
      if (String(oldValues.breakFoodThreshold) !== String(threshold)) {
        trackSettingChange('System_Config', 'breakFoodThreshold', oldValues.breakFoodThreshold, threshold,
          changedBy, reason);
      }
    }

    // Log the change
    logSettingsChange('System Configuration', 'Updated system settings');

    return { success: true, message: 'System configuration updated successfully' };

  } catch (error) {
    console.error('Error updating system config:', error.toString());
    return { success: false, message: 'Error updating system config: ' + error.message };
  }
}

/**
 * Gets all links from the Links sheet.
 * @returns {Object} Result with links array
 * @deprecated Use getAllLinks() from LinkManagement.gs instead
 */
function getLinksSimple() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const links = getLinksData(ss);
    return { success: true, links: links };
  } catch (error) {
    console.error('Error getting links:', error.toString());
    return { success: false, message: 'Error getting links: ' + error.message };
  }
}

/**
 * Adds a new link to the Links sheet.
 * @param {Object} linkData - Link object with label, url, category
 * @returns {Object} Success/failure result
 * @deprecated Use addLink() from LinkManagement.gs instead
 */
function addLinkSimple(linkData) {
  try {
    if (!linkData.label || !linkData.url) {
      return { success: false, message: 'Label and URL are required' };
    }

    // Validate URL format
    try {
      new URL(linkData.url);
    } catch (e) {
      return { success: false, message: 'Invalid URL format' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linksSheet = ss.getSheetByName('Links');

    if (!linksSheet) {
      return { success: false, message: 'Links sheet not found' };
    }

    // Add new row
    const newRow = [
      linkData.label,
      linkData.url,
      linkData.category || 'General'
    ];

    linksSheet.appendRow(newRow);

    // Log the change
    logSettingsChange('Links', `Added link: ${linkData.label}`);

    return { success: true, message: 'Link added successfully' };

  } catch (error) {
    console.error('Error adding link:', error.toString());
    return { success: false, message: 'Error adding link: ' + error.message };
  }
}

/**
 * Updates an existing link in the Links sheet.
 * @param {number} index - Row index (0-based, excluding header)
 * @param {Object} linkData - Updated link data
 * @returns {Object} Success/failure result
 * @deprecated Use updateLink() from LinkManagement.gs instead
 */
function updateLinkSimple(index, linkData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linksSheet = ss.getSheetByName('Links');

    if (!linksSheet) {
      return { success: false, message: 'Links sheet not found' };
    }

    const row = index + 2; // +2 for header and 0-indexing
    const lastRow = linksSheet.getLastRow();

    if (row < 2 || row > lastRow) {
      return { success: false, message: 'Invalid link index' };
    }

    if (linkData.label !== undefined) {
      linksSheet.getRange(row, 1).setValue(linkData.label);
    }

    if (linkData.url !== undefined) {
      try {
        new URL(linkData.url);
      } catch (e) {
        return { success: false, message: 'Invalid URL format' };
      }
      linksSheet.getRange(row, 2).setValue(linkData.url);
    }

    if (linkData.category !== undefined) {
      linksSheet.getRange(row, 3).setValue(linkData.category);
    }

    // Log the change
    logSettingsChange('Links', `Updated link at row ${row}`);

    return { success: true, message: 'Link updated successfully' };

  } catch (error) {
    console.error('Error updating link:', error.toString());
    return { success: false, message: 'Error updating link: ' + error.message };
  }
}

/**
 * Deletes a link from the Links sheet.
 * @param {number} index - Row index (0-based, excluding header)
 * @returns {Object} Success/failure result
 * @deprecated Use deleteLink() from LinkManagement.gs instead
 */
function deleteLinkSimple(index) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linksSheet = ss.getSheetByName('Links');

    if (!linksSheet) {
      return { success: false, message: 'Links sheet not found' };
    }

    const row = index + 2; // +2 for header and 0-indexing
    const lastRow = linksSheet.getLastRow();

    if (row < 2 || row > lastRow) {
      return { success: false, message: 'Invalid link index' };
    }

    // Get label for logging before deletion
    const label = linksSheet.getRange(row, 1).getValue();

    linksSheet.deleteRow(row);

    // Log the change
    logSettingsChange('Links', `Deleted link: ${label}`);

    return { success: true, message: 'Link deleted successfully' };

  } catch (error) {
    console.error('Error deleting link:', error.toString());
    return { success: false, message: 'Error deleting link: ' + error.message };
  }
}

/**
 * Logs a settings change to the Edit_Log sheet.
 * @param {string} category - Category of the change
 * @param {string} description - Description of the change
 */
function logSettingsChange(category, description) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Edit_Log');

    if (!logSheet) {
      console.log('Edit_Log sheet not found, skipping log');
      return;
    }

    const timestamp = new Date();
    const roleResult = getCurrentRole();
    const role = roleResult.authenticated ? roleResult.role : 'Unknown';

    logSheet.appendRow([
      timestamp,
      'Settings',
      category,
      description,
      role
    ]);

  } catch (error) {
    console.error('Error logging settings change:', error.toString());
  }
}

/**
 * Creates a backup of current settings.
 * @returns {Object} Backup data and result
 */
function backupSettings() {
  try {
    const settings = getSystemSettings();

    if (!settings.success) {
      return settings;
    }

    // Update last backup date
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    settingsSheet.getRange('B36').setValue(new Date());

    // Log the backup
    logSettingsChange('Backup', 'Settings backed up');

    return {
      success: true,
      message: 'Settings backed up successfully',
      backup: settings.settings,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('Error backing up settings:', error.toString());
    return { success: false, message: 'Error backing up settings: ' + error.message };
  }
}

// ============================================
// MICRO-PHASE 27: EMAIL TEMPLATE CUSTOMIZATION
// ============================================

function requireRoleForTemplates(token, allowedRoles) {
  const session = getCurrentRole(token);
  if (!session || !session.authenticated) {
    return { ok: false, sessionExpired: true, error: 'Session expired' };
  }
  if (!allowedRoles.includes(session.role)) {
    return { ok: false, error: 'Access denied' };
  }
  return { ok: true, session: session };
}

function getTemplateIdForThreshold(threshold) {
  const map = {
    2: 'two_point_alert',
    3: 'three_point_alert',
    5: 'five_point_alert',
    6: 'six_point_alert',
    9: 'nine_point_alert',
    12: 'twelve_point_alert',
    15: 'fifteen_point_alert'
  };
  return map[threshold] || null;
}

function getTemplateVariableCatalog() {
  return [
    {
      category: 'Employee Info',
      variables: [
        { name: 'employee_name', description: "Employee's full name", example: 'John Smith', required: true },
        { name: 'employee_id', description: 'Employee ID', example: '12-1234567' },
        { name: 'location', description: 'Primary location', example: 'Cockrell Hill DTO' }
      ]
    },
    {
      category: 'Point Info',
      variables: [
        { name: 'current_points', description: 'Current point total', example: '6', required: true },
        { name: 'threshold', description: 'Threshold crossed', example: '6' },
        { name: 'consequences', description: 'Consequences from Settings', example: 'Director meeting required' }
      ]
    },
    {
      category: 'Infraction Info',
      variables: [
        { name: 'infraction_type', description: 'Triggering infraction type', example: 'Bucket 3: Major Offenses' },
        { name: 'infraction_date', description: 'Triggering infraction date', example: '01/22/2026' },
        { name: 'infraction_description', description: 'Triggering infraction description', example: 'Late return from break' },
        { name: 'points_assigned', description: 'Points assigned for infraction', example: '3' }
      ]
    },
    {
      category: 'Dates',
      variables: [
        { name: 'date', description: 'Current date', example: '01/22/2026' },
        { name: 'next_expiration_date', description: 'Next points expiration date', example: '02/15/2026' },
        { name: 'days_until_expiration', description: 'Days until next expiration', example: '23' }
      ]
    },
    {
      category: 'Probation',
      variables: [
        { name: 'probation_start_date', description: 'Probation start date', example: '01/10/2026' },
        { name: 'probation_end_date', description: 'Probation end date', example: '02/09/2026' },
        { name: 'days_remaining', description: 'Days remaining in probation', example: '12' }
      ]
    },
    {
      category: 'Links',
      variables: [
        { name: 'record_link', description: 'Link to employee record', example: 'https://example.com/employee/123' },
        { name: 'termination_link', description: 'Link to termination workflow', example: 'https://example.com/terminate/123' }
      ]
    }
  ];
}

function getAllTemplateVariableNames() {
  const catalog = getTemplateVariableCatalog();
  const names = [];
  catalog.forEach(group => {
    group.variables.forEach(v => names.push(v.name));
  });
  return names;
}

function getTemplateVariables() {
  return { success: true, variables: getTemplateVariableCatalog() };
}

function extractTemplateVariables(text) {
  const matches = String(text || '').match(/\{[a-z0-9_]+\}/gi) || [];
  return Array.from(new Set(matches.map(m => m.replace(/[{}]/g, ''))));
}

function findClosestVariableName(name, allowed) {
  let best = null;
  let bestScore = Infinity;
  for (const candidate of allowed) {
    const score = levenshteinDistance(name, candidate);
    if (score < bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  return bestScore <= 3 ? best : null;
}

function levenshteinDistance(a, b) {
  const matrix = [];
  const alen = a.length;
  const blen = b.length;
  for (let i = 0; i <= alen; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= blen; j++) {
    matrix[0][j] = j;
  }
  for (let i = 1; i <= alen; i++) {
    for (let j = 1; j <= blen; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }
  return matrix[alen][blen];
}

function validateTemplateVariables(subjectTemplate, bodyTemplate) {
  const allowed = getAllTemplateVariableNames();
  const used = Array.from(new Set([
    ...extractTemplateVariables(subjectTemplate),
    ...extractTemplateVariables(bodyTemplate)
  ]));
  const missingRequired = [];
  ['employee_name', 'current_points'].forEach(req => {
    if (!used.includes(req)) missingRequired.push(req);
  });
  const unknown = used.filter(v => !allowed.includes(v));
  const suggestions = unknown.map(name => ({
    name,
    suggestion: findClosestVariableName(name, allowed)
  }));
  return { missingRequired, unknown, suggestions };
}

function renderTemplateString(template, variables) {
  const missing = [];
  const output = String(template || '').replace(/\{[a-z0-9_]+\}/gi, match => {
    const key = match.replace(/[{}]/g, '');
    if (variables && Object.prototype.hasOwnProperty.call(variables, key)) {
      const value = variables[key];
      return value === null || value === undefined ? '' : String(value);
    }
    if (!missing.includes(key)) missing.push(key);
    return '';
  });
  return { output, missing };
}

function formatTemplateDate(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function buildSampleTemplateData() {
  return {
    employee_name: 'John Smith',
    employee_id: '12-1234567',
    current_points: '6',
    location: 'Cockrell Hill DTO',
    date: formatTemplateDate(new Date()),
    threshold: '6',
    consequences: 'Director meeting required',
    infraction_type: 'Bucket 3: Major Offenses',
    infraction_date: formatTemplateDate(new Date()),
    infraction_description: 'Late return from break',
    points_assigned: '3',
    next_expiration_date: formatTemplateDate(new Date(Date.now() + 1000 * 60 * 60 * 24 * 30)),
    days_until_expiration: '30',
    probation_start_date: formatTemplateDate(new Date()),
    probation_end_date: formatTemplateDate(new Date(Date.now() + 1000 * 60 * 60 * 24 * 30)),
    days_remaining: '30',
    record_link: '',
    termination_link: ''
  };
}

function ensureEmailTemplatesSection() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    return { success: false, message: 'Settings sheet not found' };
  }

  let headerRow = findSettingsRowByLabel(sheet, 'Email Templates');
  if (headerRow < 1) {
    const lastRow = sheet.getLastRow();
    headerRow = lastRow + 2;
    sheet.getRange(headerRow, 1, 1, 7).setValues([['Email Templates', '', '', '', '', '', '']]);
    sheet.getRange(headerRow + 1, 1, 1, 7).setValues([[
      'Template ID', 'Template Name', 'Enabled', 'Subject Template',
      'Body Template', 'Updated By', 'Updated At'
    ]]);

    formatSectionHeader(sheet, headerRow, 'Email Templates', 7);
    formatSubHeader(sheet, headerRow + 1, 7);

    const defaults = getDefaultEmailTemplatesList();
    const rows = defaults.map(t => [
      t.template_id,
      t.template_name,
      t.enabled ? 'TRUE' : 'FALSE',
      t.subject_template,
      t.body_template,
      '',
      ''
    ]);
    if (rows.length) {
      sheet.getRange(headerRow + 2, 1, rows.length, 7).setValues(rows);
    }
  }

  return { success: true, sheet: sheet, headerRow: headerRow, startRow: headerRow + 2 };
}

function getEmailTemplates(token) {
  const auth = requireRoleForTemplates(token, ['Director', 'Operator']);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const section = ensureEmailTemplatesSection();
  if (!section.success) return section;

  const sheet = section.sheet;
  const startRow = section.startRow;
  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(0, lastRow - startRow + 1);
  if (rowCount === 0) {
    return { success: true, templates: [] };
  }

  const data = sheet.getRange(startRow, 1, rowCount, 7).getValues();
  const templates = [];
  for (const row of data) {
    const templateId = String(row[0] || '').trim();
    if (!templateId) break;
    templates.push({
      template_id: templateId,
      template_name: row[1] || '',
      enabled: String(row[2] || '').toUpperCase() === 'TRUE',
      subject_template: row[3] || '',
      body_template: row[4] || '',
      updated_by: row[5] || '',
      updated_at: row[6] || ''
    });
  }
  return { success: true, templates: templates };
}

function getEmailTemplateById(templateId) {
  const section = ensureEmailTemplatesSection();
  if (!section.success) return { success: false, message: section.message };

  const sheet = section.sheet;
  const startRow = section.startRow;
  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(0, lastRow - startRow + 1);
  if (rowCount === 0) {
    return { success: false, message: 'No templates found' };
  }

  const data = sheet.getRange(startRow, 1, rowCount, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentId = String(row[0] || '').trim();
    if (!currentId) break;
    if (currentId === templateId) {
      return {
        success: true,
        rowIndex: startRow + i,
        template: {
          template_id: currentId,
          template_name: row[1] || '',
          enabled: String(row[2] || '').toUpperCase() === 'TRUE',
          subject_template: row[3] || '',
          body_template: row[4] || '',
          updated_by: row[5] || '',
          updated_at: row[6] || ''
        }
      };
    }
  }

  const defaults = getDefaultEmailTemplatesMap();
  if (defaults[templateId]) {
    return { success: true, rowIndex: null, template: defaults[templateId] };
  }
  return { success: false, message: 'Template not found' };
}

function getOrCreateTemplateHistorySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Template_History');
  if (!sheet) {
    sheet = ss.insertSheet('Template_History');
    sheet.appendRow([
      'Template_ID',
      'Version',
      'Subject_Template',
      'Body_Template',
      'Enabled',
      'Updated_By',
      'Updated_At'
    ]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 350);
    sheet.setColumnWidth(4, 700);
    sheet.setColumnWidth(5, 90);
    sheet.setColumnWidth(6, 160);
    sheet.setColumnWidth(7, 160);
  }
  return sheet;
}

function getNextTemplateVersion(templateId) {
  const sheet = getOrCreateTemplateHistorySheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let maxVersion = 0;
  for (const row of data) {
    if (row[0] === templateId) {
      const version = Number(row[1]) || 0;
      if (version > maxVersion) maxVersion = version;
    }
  }
  return maxVersion + 1;
}

function previewEmailTemplate(templateId, sampleData) {
  const templateResult = getEmailTemplateById(templateId);
  if (!templateResult.success) {
    return { success: false, message: templateResult.message || 'Template not found' };
  }
  const template = templateResult.template;
  const data = Object.assign(buildSampleTemplateData(), sampleData || {});

  const subjectResult = renderTemplateString(template.subject_template, data);
  const bodyResult = renderTemplateString(template.body_template, data);

  return {
    success: true,
    subject: subjectResult.output,
    body: bodyResult.output,
    missing: Array.from(new Set(subjectResult.missing.concat(bodyResult.missing)))
  };
}

function applyEmailTemplateUpdate(templateId, templateData, updatedByDirector, sessionInfo) {
  const defaults = getDefaultEmailTemplatesMap();
  if (!defaults[templateId]) {
    return { success: false, message: 'Invalid template_id' };
  }

  const subject = String(templateData.subject_template || '').trim();
  const body = String(templateData.body_template || '').trim();
  const enabled = Boolean(templateData.enabled);

  if (!subject) return { success: false, message: 'Subject template is required' };
  if (!body) return { success: false, message: 'Body template is required' };

  const validation = validateTemplateVariables(subject, body);
  if (validation.missingRequired.length > 0) {
    return {
      success: false,
      message: `Missing required variables: ${validation.missingRequired.join(', ')}`,
      validation: validation
    };
  }
  if (validation.unknown.length > 0) {
    return {
      success: false,
      message: `Unknown variables found: ${validation.unknown.join(', ')}`,
      validation: validation
    };
  }

  const section = ensureEmailTemplatesSection();
  if (!section.success) return section;

  const templateResult = getEmailTemplateById(templateId);
  if (!templateResult.success) {
    return { success: false, message: templateResult.message || 'Template not found' };
  }

  // Save previous version to history
  const historySheet = getOrCreateTemplateHistorySheet();
  const version = getNextTemplateVersion(templateId);
  const now = new Date();
  historySheet.appendRow([
    templateId,
    version,
    templateResult.template.subject_template || '',
    templateResult.template.body_template || '',
    templateResult.template.enabled ? 'TRUE' : 'FALSE',
    updatedByDirector || 'Unknown',
    now
  ]);

  const rowIndex = templateResult.rowIndex || (section.sheet.getLastRow() + 1);
  if (!templateResult.rowIndex) {
    section.sheet.getRange(rowIndex, 1, 1, 7).setValues([[
      templateId,
      defaults[templateId].template_name,
      enabled ? 'TRUE' : 'FALSE',
      subject,
      body,
      updatedByDirector || 'Unknown',
      now
    ]]);
  } else {
    section.sheet.getRange(rowIndex, 3).setValue(enabled ? 'TRUE' : 'FALSE');
    section.sheet.getRange(rowIndex, 4).setValue(subject);
    section.sheet.getRange(rowIndex, 5).setValue(body);
    section.sheet.getRange(rowIndex, 6).setValue(updatedByDirector || 'Unknown');
    section.sheet.getRange(rowIndex, 7).setValue(now);
  }

  trackSettingChange(
    'Email_Templates',
    templateId,
    JSON.stringify({
      enabled: templateResult.template.enabled,
      subject: templateResult.template.subject_template,
      body: templateResult.template.body_template
    }),
    JSON.stringify({ enabled: enabled, subject: subject, body: body }),
    updatedByDirector || 'Unknown',
    'Email template update'
  );

  // Log change to Edit_Log
  try {
    logEditAction({
      actionType: 'update_email_template',
      directorEmail: updatedByDirector || 'Unknown',
      targetType: 'email_template',
      targetId: templateId,
      employeeId: '',
      employeeName: '',
      fieldChanged: 'subject/body/enabled',
      originalValue: '',
      newValue: '',
      reason: 'Email template updated',
      sessionInfo: sessionInfo || ''
    });
  } catch (e) {
    console.error('Failed to log template change:', e.toString());
  }

  // Build preview using sample data
  const preview = previewEmailTemplate(templateId, {});
  return {
    success: true,
    message: 'Template updated',
    preview: preview,
    validation: validation
  };
}

function updateEmailTemplate(templateId, templateData, updatedByDirector, token) {
  const auth = requireRoleForTemplates(token, ['Director']);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  return applyEmailTemplateUpdate(
    templateId,
    templateData,
    updatedByDirector || auth.session.user_name || auth.session.role,
    auth.session.role
  );
}

function resetTemplateToDefault(templateId, resetByDirector, token) {
  const auth = requireRoleForTemplates(token, ['Director']);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const defaults = getDefaultEmailTemplatesMap();
  const defaultTemplate = defaults[templateId];
  if (!defaultTemplate) {
    return { success: false, message: 'Invalid template_id' };
  }

  return updateEmailTemplate(
    templateId,
    {
      subject_template: defaultTemplate.subject_template,
      body_template: defaultTemplate.body_template,
      enabled: defaultTemplate.enabled
    },
    resetByDirector,
    token
  );
}

function getTemplateHistory(templateId, token) {
  const auth = requireRoleForTemplates(token, ['Director', 'Operator']);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const sheet = getOrCreateTemplateHistorySheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, history: [] };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const history = data
    .filter(row => row[0] === templateId)
    .map(row => ({
      template_id: row[0],
      version: row[1],
      subject_template: row[2] || '',
      body_template: row[3] || '',
      enabled: String(row[4] || '').toUpperCase() === 'TRUE',
      updated_by: row[5] || '',
      updated_at: row[6] instanceof Date ? row[6].toISOString() : row[6]
    }))
    .sort((a, b) => Number(b.version) - Number(a.version));

  return { success: true, history: history };
}

function revertTemplateVersion(templateId, version, resetByDirector, token) {
  const auth = requireRoleForTemplates(token, ['Director']);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const sheet = getOrCreateTemplateHistorySheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: false, message: 'No history available' };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const match = data.find(row => row[0] === templateId && String(row[1]) === String(version));
  if (!match) {
    return { success: false, message: 'Template version not found' };
  }

  return applyEmailTemplateUpdate(
    templateId,
    {
      subject_template: match[2] || '',
      body_template: match[3] || '',
      enabled: String(match[4] || '').toUpperCase() === 'TRUE'
    },
    resetByDirector || auth.session.user_name || auth.session.role,
    'Revert template version'
  );
}

function sendTemplatedEmail(templateId, recipient, templateVariables) {
  try {
    const templateResult = getEmailTemplateById(templateId);
    if (!templateResult.success) {
      return { success: false, message: templateResult.message || 'Template not found' };
    }
    const template = templateResult.template;
    if (template.enabled === false) {
      return { success: true, skipped: true, message: 'Template disabled' };
    }

    const subjectResult = renderTemplateString(template.subject_template, templateVariables || {});
    const bodyResult = renderTemplateString(template.body_template, templateVariables || {});

    const subject = subjectResult.output;
    const htmlBody = bodyResult.output;
    const textBody = htmlBody.replace(/<[^>]+>/g, '');

    GmailApp.sendEmail(recipient, subject, textBody || 'Please view this email in an HTML-compatible email client.', {
      htmlBody: htmlBody,
      name: 'CFA Accountability System'
    });

    const logId = generateEmailLogId();
    writeEmailLog({
      log_id: logId,
      timestamp: new Date(),
      employee_id: templateVariables?.employee_id || '',
      employee_name: templateVariables?.employee_name || '',
      recipient_email: recipient,
      email_type: templateId,
      thresholds_crossed: templateVariables?.threshold || '',
      status: 'Sent',
      retry_count: 0,
      error_message: ''
    });

    return { success: true, skipped: false, log_id: logId };
  } catch (error) {
    console.error('Error in sendTemplatedEmail:', error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Restores settings from a backup.
 * @param {Object} backup - Backup object from backupSettings()
 * @returns {Object} Success/failure result
 */
function restoreSettings(backup) {
  try {
    // Check if current user has Operator role
    const roleResult = getCurrentRole();
    if (!roleResult.authenticated || roleResult.role !== 'Operator') {
      return { success: false, message: 'Only Operators can restore settings' };
    }

    if (!backup || !backup.email || !backup.thresholds || !backup.buckets || !backup.system) {
      return { success: false, message: 'Invalid backup data' };
    }

    // Restore each section
    const emailResult = updateEmailSettings(backup.email);
    if (!emailResult.success) return emailResult;

    const thresholdResult = updateThresholdSettings(backup.thresholds);
    if (!thresholdResult.success) return thresholdResult;

    const bucketResult = updateBucketSettings(backup.buckets);
    if (!bucketResult.success) return bucketResult;

    const systemResult = updateSystemConfig(backup.system);
    if (!systemResult.success) return systemResult;

    // Log the restore
    logSettingsChange('Restore', 'Settings restored from backup');

    return { success: true, message: 'Settings restored successfully' };

  } catch (error) {
    console.error('Error restoring settings:', error.toString());
    return { success: false, message: 'Error restoring settings: ' + error.message };
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test function for email templates.
 * Requires a valid Director session token when running in production.
 */
function testEmailTemplates(token) {
  const results = [];
  const testRecipient = Session.getActiveUser().getEmail();
  if (!token) {
    return { success: false, message: 'Missing token. Provide a Director session token to run tests.' };
  }

  // Test 1: Get all templates
  const allTemplates = getEmailTemplates(token);
  results.push({ test: 'Get all templates', passed: allTemplates.success && Array.isArray(allTemplates.templates) });

  // Test 2: Update template
  const updateResult = updateEmailTemplate(
    'six_point_alert',
    {
      subject_template: 'TEST: Director Meeting Required - {employee_name}',
      body_template: 'Test body for {employee_name} at {current_points} points.',
      enabled: true
    },
    'Test Script',
    token
  );
  results.push({ test: 'Update template', passed: updateResult.success === true });

  // Test 3: Preview template
  const previewResult = previewEmailTemplate('nine_point_alert', {
    employee_name: 'Test Employee',
    current_points: '9'
  });
  results.push({ test: 'Preview template', passed: previewResult.success === true && !/\{.+\}/.test(previewResult.subject + previewResult.body) });

  // Test 4: Send templated email
  const sendResult = sendTemplatedEmail('two_point_alert', testRecipient, buildSampleTemplateData());
  results.push({ test: 'Send templated email', passed: sendResult.success === true });

  // Test 5: Disable template
  const disableResult = updateEmailTemplate(
    'three_point_alert',
    {
      subject_template: 'Accountability Notice - {employee_name}',
      body_template: 'Employee {employee_name} has {current_points} points.',
      enabled: false
    },
    'Test Script',
    token
  );
  results.push({ test: 'Disable template', passed: disableResult.success === true });

  // Test 6: Reset to default
  const resetResult = resetTemplateToDefault('six_point_alert', 'Test Script', token);
  results.push({ test: 'Reset to default', passed: resetResult.success === true });

  // Test 7: Variable validation
  const invalidResult = updateEmailTemplate(
    'two_point_alert',
    {
      subject_template: 'Missing required variable',
      body_template: 'Body without required variables.',
      enabled: true
    },
    'Test Script',
    token
  );
  results.push({ test: 'Variable validation', passed: invalidResult.success === false });

  // Test 8: HTML formatting
  const htmlResult = updateEmailTemplate(
    'five_point_alert',
    {
      subject_template: 'HTML Test - {employee_name}',
      body_template: '<h3>Hello {employee_name}</h3><strong>Points:</strong> {current_points}<br><a href="https://example.com">Link</a>',
      enabled: true
    },
    'Test Script',
    token
  );
  results.push({ test: 'HTML formatting', passed: htmlResult.success === true });

  return {
    success: results.every(r => r.passed),
    results: results
  };
}

/**
 * Test function for settings management.
 */
function testSettingsUpdate() {
  console.log('=== Testing Settings Management ===');
  console.log('');

  // Test 1: Get current settings
  console.log('Test 1: Getting current settings...');
  const settings = getSystemSettings();
  console.log('Success:', settings.success);
  if (settings.success) {
    console.log('Email settings:', JSON.stringify(settings.settings.email));
    console.log('Thresholds count:', settings.settings.thresholds.length);
    console.log('Buckets count:', settings.settings.buckets.length);
    console.log('System config:', JSON.stringify(settings.settings.system));
  } else {
    console.log('Error:', settings.message);
  }
  console.log('');

  // Test 2: Update system config
  console.log('Test 2: Updating session timeout...');
  const configResult = updateSystemConfig({ sessionTimeout: 15 });
  console.log('Success:', configResult.success);
  console.log('Message:', configResult.message);
  console.log('');

  // Test 3: Add an example to bucket
  console.log('Test 3: Adding example to bucket 1...');
  const exampleResult = addInfractionExample(0, 'Test example - ' + new Date().getTime());
  console.log('Success:', exampleResult.success);
  console.log('Message:', exampleResult.message);
  console.log('');

  // Test 4: Get links
  console.log('Test 4: Getting links...');
  const linksResult = getLinksSimple();
  console.log('Success:', linksResult.success);
  console.log('Links count:', linksResult.links ? linksResult.links.length : 0);
  console.log('');

  console.log('=== Tests Complete ===');
  return { success: true, message: 'All tests completed' };
}
