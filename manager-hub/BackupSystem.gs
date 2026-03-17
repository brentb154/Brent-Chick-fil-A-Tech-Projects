/**
 * BackupSystem.gs
 * Micro-Phase 22: Quarterly Backup System
 *
 * Automatically creates quarterly backups of the entire Accountability System
 * and stores in Google Drive with proper permissions and organization.
 */

// Backup folder name
const BACKUP_FOLDER_NAME = 'CFA Accountability Backups';

// Sheets to backup (in order)
const SHEETS_TO_BACKUP = [
  'Infractions',
  'Settings',
  'User_Permissions',
  'Email_Log',
  'Links',
  'Terminated_Employees',
  'Action_Tracking',
  'Probation_Tracking',
  'Edit_Log',
  'Backup_Log',
  'Backup_Test_Log',
  'Restore_History',
  'Data_Quality_Issues',
  'Logs'
];

/**
 * Get or create the main backup folder
 * @returns {GoogleAppsScript.Drive.Folder} The backup folder
 */
function getBackupFolder() {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  // Create the folder
  const folder = DriveApp.createFolder(BACKUP_FOLDER_NAME);
  console.log('Created backup folder: ' + folder.getName());
  return folder;
}

/**
 * Create a dated subfolder for this backup
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - The main backup folder
 * @returns {GoogleAppsScript.Drive.Folder} The dated subfolder
 */
function createDatedBackupFolder(parentFolder) {
  const now = new Date();
  const folderName = 'Backup_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const folder = parentFolder.createFolder(folderName);
  return folder;
}

/**
 * Main function to create a system backup
 * @param {string} backupType - "automatic" or "manual"
 * @param {string} createdBy - Name of person creating backup (for manual)
 * @returns {Object} Result with backup details or error
 */
function createSystemBackup(backupType, createdBy) {
  const startTime = new Date();
  let backupFile = null;
  let backupFolder = null;
  let retryCount = 0;
  const maxRetries = 1;

  while (retryCount <= maxRetries) {
    try {
      console.log('Starting system backup (' + backupType + ')...');

      // Get source spreadsheet
      const sourceSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sourceName = sourceSpreadsheet.getName();

      // Create backup folder structure
      const mainBackupFolder = getBackupFolder();
      backupFolder = createDatedBackupFolder(mainBackupFolder);

      // Generate backup file name
      const now = new Date();
      const backupFileName = 'CFA_Accountability_Backup_' +
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');

      // Create new spreadsheet for backup
      backupFile = SpreadsheetApp.create(backupFileName);
      const backupFileId = backupFile.getId();

      // Move backup file to dated folder
      const file = DriveApp.getFileById(backupFileId);
      file.moveTo(backupFolder);

      // Track record counts
      const recordCounts = {};
      let totalRecords = 0;
      const sheetsCopied = [];

      // Copy each sheet
      for (const sheetName of SHEETS_TO_BACKUP) {
        try {
          const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);

          if (sourceSheet) {
            // Get all data including formulas
            const dataRange = sourceSheet.getDataRange();
            const data = dataRange.getValues();
            const formulas = dataRange.getFormulas();
            const formats = dataRange.getNumberFormats();
            const backgrounds = dataRange.getBackgrounds();
            const fontColors = dataRange.getFontColors();
            const fontWeights = dataRange.getFontWeights();

            // Create sheet in backup
            let backupSheet;
            if (sheetsCopied.length === 0) {
              // Rename the default sheet
              backupSheet = backupFile.getSheets()[0];
              backupSheet.setName(sheetName);
            } else {
              backupSheet = backupFile.insertSheet(sheetName);
            }

            if (data.length > 0 && data[0].length > 0) {
              // Set values (or formulas where they exist)
              const targetRange = backupSheet.getRange(1, 1, data.length, data[0].length);

              // Merge data and formulas (formulas take precedence)
              const mergedData = data.map((row, rowIndex) => {
                return row.map((cell, colIndex) => {
                  const formula = formulas[rowIndex][colIndex];
                  return formula ? formula : cell;
                });
              });

              targetRange.setValues(mergedData);

              // Apply formatting
              try {
                targetRange.setNumberFormats(formats);
                targetRange.setBackgrounds(backgrounds);
                targetRange.setFontColors(fontColors);
                targetRange.setFontWeights(fontWeights);
              } catch (formatError) {
                console.log('Warning: Could not apply all formatting for ' + sheetName);
              }

              // Set column widths
              for (let col = 1; col <= data[0].length; col++) {
                try {
                  const width = sourceSheet.getColumnWidth(col);
                  backupSheet.setColumnWidth(col, width);
                } catch (e) {
                  // Ignore column width errors
                }
              }

              // Freeze rows if source has frozen rows
              const frozenRows = sourceSheet.getFrozenRows();
              if (frozenRows > 0) {
                backupSheet.setFrozenRows(frozenRows);
              }
            }

            const rowCount = data.length > 1 ? data.length - 1 : 0; // Exclude header
            recordCounts[sheetName] = rowCount;
            totalRecords += rowCount;
            sheetsCopied.push(sheetName);

            console.log('Copied sheet: ' + sheetName + ' (' + rowCount + ' records)');
          } else {
            console.log('Sheet not found, skipping: ' + sheetName);
          }
        } catch (sheetError) {
          console.log('Error copying sheet ' + sheetName + ': ' + sheetError.toString());
          recordCounts[sheetName] = 'ERROR';
        }
      }

      // Create Backup Info sheet
      const infoSheet = backupFile.insertSheet('_Backup_Info');
      const backupId = 'BKP' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

      const infoData = [
        ['Backup Information', ''],
        ['', ''],
        ['Backup ID', backupId],
        ['Backup Date', Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')],
        ['Backup Time', Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss')],
        ['Backup Type', backupType],
        ['Created By', createdBy || (backupType === 'automatic' ? 'System (Automatic)' : 'Unknown')],
        ['Source Spreadsheet', sourceName],
        ['Source Spreadsheet ID', SPREADSHEET_ID],
        ['', ''],
        ['Sheets Backed Up', sheetsCopied.length],
        ['Total Records', totalRecords],
        ['', ''],
        ['Record Counts by Sheet', '']
      ];

      // Add record counts
      for (const [sheet, count] of Object.entries(recordCounts)) {
        infoData.push([sheet, count]);
      }

      infoData.push(['', '']);
      infoData.push(['Processing Time', ((new Date() - startTime) / 1000).toFixed(2) + ' seconds']);
      infoData.push(['Status', 'SUCCESS']);

      infoSheet.getRange(1, 1, infoData.length, 2).setValues(infoData);
      infoSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
      infoSheet.setColumnWidth(1, 200);
      infoSheet.setColumnWidth(2, 300);

      // Move info sheet to first position
      backupFile.setActiveSheet(infoSheet);
      backupFile.moveActiveSheet(1);

      // Set permissions - make read-only for editors
      try {
        // Remove edit access from anyone except owner
        const protection = file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
        console.log('Set backup file to private/view-only');
      } catch (permError) {
        console.log('Warning: Could not set file permissions: ' + permError.toString());
      }

      // Update Settings sheet with backup info
      try {
        updateBackupSettings(now, backupType, backupFile.getUrl(), backupId);
      } catch (settingsError) {
        console.log('Warning: Could not update Settings: ' + settingsError.toString());
      }

      // Log to Backup_Log
      const logResult = logBackup(backupId, now, backupType, backupFileId, backupFile.getUrl(),
                                   backupFolder.getUrl(), recordCounts, 'Success', null, createdBy);

      const endTime = new Date();
      const processingTime = (endTime - startTime) / 1000;

      console.log('Backup completed successfully in ' + processingTime.toFixed(2) + ' seconds');
      logSystemEvent(
        'success',
        `Backup completed (${backupType}) - ${backupId} in ${processingTime.toFixed(2)}s`,
        'low'
      );

      return {
        success: true,
        backup_id: backupId,
        backup_file_id: backupFileId,
        backup_file_url: backupFile.getUrl(),
        view_url: 'https://docs.google.com/spreadsheets/d/' + backupFileId + '/view',
        folder_url: backupFolder.getUrl(),
        record_counts: recordCounts,
        total_records: totalRecords,
        sheets_copied: sheetsCopied.length,
        processing_time: processingTime,
        backup_type: backupType,
        created_at: now.toISOString()
      };

    } catch (error) {
      console.error('Backup error (attempt ' + (retryCount + 1) + '): ' + error.toString());
      logSystemEvent('error', error, 'high');

      // Clean up failed backup
      if (backupFile) {
        try {
          DriveApp.getFileById(backupFile.getId()).setTrashed(true);
        } catch (e) {
          // Ignore cleanup errors
        }
      }

      if (retryCount < maxRetries) {
        console.log('Retrying backup...');
        retryCount++;
        Utilities.sleep(2000); // Wait 2 seconds before retry
        continue;
      }

      // Log failed backup
      const failedBackupId = 'BKP_FAILED_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
      logBackup(failedBackupId, new Date(), backupType, null, null, null, {}, 'Failed', error.toString(), createdBy);

      // Send error notification
      try {
        sendBackupFailureEmail(error.toString(), backupType, createdBy);
      } catch (emailError) {
        console.log('Could not send failure email: ' + emailError.toString());
      }

      return {
        success: false,
        error: error.toString(),
        backup_type: backupType
      };
    }
  }
}

/**
 * Update Settings sheet with backup information
 */
function updateBackupSettings(backupDate, backupType, backupUrl, backupId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const settingsSheet = ss.getSheetByName('Settings');

  if (!settingsSheet) {
    console.log('Settings sheet not found');
    return;
  }

  // Find or create backup settings rows
  const data = settingsSheet.getDataRange().getValues();
  let lastBackupRow = -1;
  let backupTypeRow = -1;
  let backupUrlRow = -1;
  let backupIdRow = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'last_backup_date') lastBackupRow = i + 1;
    if (data[i][0] === 'last_backup_type') backupTypeRow = i + 1;
    if (data[i][0] === 'last_backup_url') backupUrlRow = i + 1;
    if (data[i][0] === 'last_backup_id') backupIdRow = i + 1;
  }

  const dateStr = Utilities.formatDate(backupDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // Update or append backup settings
  if (lastBackupRow > 0) {
    settingsSheet.getRange(lastBackupRow, 2).setValue(dateStr);
  } else {
    settingsSheet.appendRow(['last_backup_date', dateStr]);
  }

  if (backupTypeRow > 0) {
    settingsSheet.getRange(backupTypeRow, 2).setValue(backupType);
  } else {
    settingsSheet.appendRow(['last_backup_type', backupType]);
  }

  if (backupUrlRow > 0) {
    settingsSheet.getRange(backupUrlRow, 2).setValue(backupUrl);
  } else {
    settingsSheet.appendRow(['last_backup_url', backupUrl]);
  }

  if (backupIdRow > 0) {
    settingsSheet.getRange(backupIdRow, 2).setValue(backupId);
  } else {
    settingsSheet.appendRow(['last_backup_id', backupId]);
  }
}

/**
 * Log backup to Backup_Log sheet
 */
function logBackup(backupId, backupDate, backupType, fileId, fileUrl, folderUrl, recordCounts, status, errorMessage, createdBy) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName('Backup_Log');

    if (!logSheet) {
      // Create Backup_Log sheet
      logSheet = ss.insertSheet('Backup_Log');
      const headers = ['backup_id', 'backup_date', 'backup_type', 'file_id', 'file_url',
                       'folder_url', 'record_counts', 'status', 'error_message', 'created_by'];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    const dateStr = Utilities.formatDate(backupDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const recordCountsJson = JSON.stringify(recordCounts);

    logSheet.appendRow([
      backupId,
      dateStr,
      backupType,
      fileId || '',
      fileUrl || '',
      folderUrl || '',
      recordCountsJson,
      status,
      errorMessage || '',
      createdBy || ''
    ]);

    return true;
  } catch (error) {
    console.error('Error logging backup: ' + error.toString());
    return false;
  }
}

/**
 * Get list of all backups
 * @returns {Object} Result with backups array
 */
function listBackups() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('Backup_Log');

    if (!logSheet) {
      return { success: true, backups: [] };
    }

    const data = logSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, backups: [] };
    }

    const headers = data[0];
    const backups = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const backup = {};

      headers.forEach((header, index) => {
        backup[header] = row[index];
      });

      // Parse record counts JSON
      try {
        backup.record_counts_parsed = JSON.parse(backup.record_counts || '{}');
        backup.total_records = Object.values(backup.record_counts_parsed)
          .filter(v => typeof v === 'number')
          .reduce((a, b) => a + b, 0);
      } catch (e) {
        backup.record_counts_parsed = {};
        backup.total_records = 0;
      }

      // Get file size if file exists
      if (backup.file_id) {
        try {
          const file = DriveApp.getFileById(backup.file_id);
          backup.file_size = file.getSize();
          backup.file_size_formatted = formatFileSize(file.getSize());
          backup.file_exists = true;
        } catch (e) {
          backup.file_size = 0;
          backup.file_size_formatted = 'N/A';
          backup.file_exists = false;
        }
      }

      backups.push(backup);
    }

    // Sort by date (newest first)
    backups.sort((a, b) => {
      const dateA = new Date(a.backup_date);
      const dateB = new Date(b.backup_date);
      return dateB - dateA;
    });

    return { success: true, backups: backups };

  } catch (error) {
    console.error('Error listing backups: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Format file size in human readable format
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Get last backup info for display
 * @returns {Object} Last backup details
 */
function getLastBackupInfo() {
  try {
    const backupsList = listBackups();

    if (!backupsList.success || backupsList.backups.length === 0) {
      return {
        success: true,
        has_backup: false,
        last_backup_date: null,
        last_backup_type: null,
        status: 'Never'
      };
    }

    const lastBackup = backupsList.backups[0];

    // Calculate next scheduled backup (first of next quarter)
    const nextBackup = getNextQuarterlyBackupDate();

    return {
      success: true,
      has_backup: true,
      last_backup_date: lastBackup.backup_date,
      last_backup_type: lastBackup.backup_type,
      last_backup_id: lastBackup.backup_id,
      last_backup_url: lastBackup.file_url,
      status: lastBackup.status,
      total_records: lastBackup.total_records,
      file_size: lastBackup.file_size_formatted,
      next_scheduled: nextBackup
    };

  } catch (error) {
    console.error('Error getting last backup info: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Calculate next quarterly backup date
 */
function getNextQuarterlyBackupDate() {
  const now = new Date();
  const month = now.getMonth();
  const year = now.getFullYear();

  // Quarter start months: Jan(0), Apr(3), Jul(6), Oct(9)
  let nextMonth;
  let nextYear = year;

  if (month < 3) {
    nextMonth = 3; // April
  } else if (month < 6) {
    nextMonth = 6; // July
  } else if (month < 9) {
    nextMonth = 9; // October
  } else {
    nextMonth = 0; // January
    nextYear = year + 1;
  }

  const nextDate = new Date(nextYear, nextMonth, 1, 2, 0, 0); // 2 AM on first of quarter
  return Utilities.formatDate(nextDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Set up quarterly backup trigger
 */
function scheduleQuarterlyBackup() {
  try {
    // Check if trigger already exists
    const triggers = ScriptApp.getProjectTriggers();
    const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'runAutomaticBackup');

    if (existingTrigger) {
      console.log('Quarterly backup trigger already exists');
      return {
        success: true,
        message: 'Trigger already exists',
        trigger_id: existingTrigger.getUniqueId()
      };
    }

    // Create new trigger for 1st day of each quarter at 2 AM
    // We'll use a monthly trigger and check if it's a quarter start
    const trigger = ScriptApp.newTrigger('runAutomaticBackup')
      .timeBased()
      .onMonthDay(1)
      .atHour(2)
      .create();

    console.log('Created quarterly backup trigger: ' + trigger.getUniqueId());

    return {
      success: true,
      message: 'Quarterly backup trigger created',
      trigger_id: trigger.getUniqueId()
    };

  } catch (error) {
    console.error('Error scheduling backup: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Remove quarterly backup trigger
 */
function removeQuarterlyBackupTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runAutomaticBackup') {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });

    return {
      success: true,
      removed: removed
    };
  } catch (error) {
    console.error('Error removing trigger: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Function called by trigger - runs automatic backup only on quarter starts
 */
function runAutomaticBackup() {
  const now = new Date();
  const month = now.getMonth();

  // Only run on quarter start months (Jan, Apr, Jul, Oct)
  if (month !== 0 && month !== 3 && month !== 6 && month !== 9) {
    console.log('Not a quarter start month, skipping backup');
    return;
  }

  console.log('Running automatic quarterly backup...');

  const result = createSystemBackup('automatic', 'System (Scheduled)');

  if (result.success) {
    // Send success notification
    sendBackupSuccessEmail(result, 'automatic');
  }

  return result;
}

/**
 * Legacy restore data from a backup.
 * Use restoreFromBackup in DisasterRecovery.gs for full workflow.
 */
function restoreFromBackupLegacy(backupFileId, sheetsToRestore, confirmedByDirector) {
  try {
    // Validate inputs
    if (!backupFileId) {
      return { success: false, error: 'Backup file ID is required' };
    }

    if (!sheetsToRestore || sheetsToRestore.length === 0) {
      return { success: false, error: 'No sheets selected for restore' };
    }

    if (!confirmedByDirector) {
      return { success: false, error: 'Director confirmation is required' };
    }

    console.log('Starting restore from backup: ' + backupFileId);
    console.log('Sheets to restore: ' + sheetsToRestore.join(', '));
    console.log('Confirmed by: ' + confirmedByDirector);

    // Open backup file
    let backupSpreadsheet;
    try {
      backupSpreadsheet = SpreadsheetApp.openById(backupFileId);
    } catch (e) {
      return { success: false, error: 'Could not open backup file. It may have been deleted.' };
    }

    // Open target spreadsheet
    const targetSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    const restoredSheets = [];
    const errors = [];

    // Restore each selected sheet
    for (const sheetName of sheetsToRestore) {
      try {
        const backupSheet = backupSpreadsheet.getSheetByName(sheetName);

        if (!backupSheet) {
          errors.push(sheetName + ': Not found in backup');
          continue;
        }

        let targetSheet = targetSpreadsheet.getSheetByName(sheetName);

        // Create sheet if it doesn't exist
        if (!targetSheet) {
          targetSheet = targetSpreadsheet.insertSheet(sheetName);
        }

        // Get backup data
        const backupRange = backupSheet.getDataRange();
        const backupData = backupRange.getValues();
        const backupFormulas = backupRange.getFormulas();
        const backupFormats = backupRange.getNumberFormats();

        if (backupData.length === 0) {
          errors.push(sheetName + ': No data in backup');
          continue;
        }

        // Clear target sheet
        targetSheet.clear();

        // Merge data and formulas
        const mergedData = backupData.map((row, rowIndex) => {
          return row.map((cell, colIndex) => {
            const formula = backupFormulas[rowIndex][colIndex];
            return formula ? formula : cell;
          });
        });

        // Write data to target
        const targetRange = targetSheet.getRange(1, 1, backupData.length, backupData[0].length);
        targetRange.setValues(mergedData);

        // Apply formatting
        try {
          targetRange.setNumberFormats(backupFormats);
        } catch (formatError) {
          // Ignore formatting errors
        }

        // Restore frozen rows
        const frozenRows = backupSheet.getFrozenRows();
        if (frozenRows > 0) {
          targetSheet.setFrozenRows(frozenRows);
        }

        restoredSheets.push({
          name: sheetName,
          rows: backupData.length
        });

        console.log('Restored sheet: ' + sheetName + ' (' + backupData.length + ' rows)');

      } catch (sheetError) {
        errors.push(sheetName + ': ' + sheetError.toString());
        console.error('Error restoring ' + sheetName + ': ' + sheetError.toString());
      }
    }

    // Log restore operation to Edit_Log
    logRestoreOperation(backupFileId, restoredSheets, confirmedByDirector, errors);

    // Send notification to directors
    sendRestoreNotificationEmail(backupFileId, restoredSheets, confirmedByDirector, errors);

    if (restoredSheets.length === 0) {
      return {
        success: false,
        error: 'No sheets were restored. Errors: ' + errors.join('; ')
      };
    }

    return {
      success: true,
      restored_sheets: restoredSheets,
      errors: errors,
      restored_by: confirmedByDirector,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('Restore error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Log restore operation to Edit_Log
 */
function logRestoreOperation(backupFileId, restoredSheets, confirmedBy, errors) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let editLog = ss.getSheetByName('Edit_Log');

    if (!editLog) {
      editLog = ss.insertSheet('Edit_Log');
      editLog.getRange(1, 1, 1, 6).setValues([['timestamp', 'action', 'details', 'performed_by', 'ip', 'notes']]);
      editLog.setFrozenRows(1);
    }

    const sheetsRestored = restoredSheets.map(s => s.name).join(', ');
    const details = 'Restored from backup: ' + backupFileId + ' | Sheets: ' + sheetsRestored;
    const notes = errors.length > 0 ? 'Errors: ' + errors.join('; ') : 'No errors';

    editLog.appendRow([
      new Date().toISOString(),
      'RESTORE',
      details,
      confirmedBy,
      '',
      notes
    ]);

  } catch (error) {
    console.error('Error logging restore: ' + error.toString());
  }
}

/**
 * Send backup success email notification
 */
function sendBackupSuccessEmail(backupResult, backupType) {
  try {
    // Get director emails from settings or use a default
    const recipients = getDirectorEmails();

    if (recipients.length === 0) {
      console.log('No director emails configured for backup notification');
      return;
    }

    const subject = backupType === 'automatic'
      ? 'Quarterly Backup Completed - CFA Accountability'
      : 'Manual Backup Completed - CFA Accountability';

    const recordCounts = Object.entries(backupResult.record_counts || {})
      .map(([sheet, count]) => '  - ' + sheet + ': ' + count + ' records')
      .join('\n');

    const body = `
Hello,

A ${backupType} backup of the CFA Accountability System has been completed successfully.

Backup Details:
- Backup ID: ${backupResult.backup_id}
- Date/Time: ${backupResult.created_at}
- Total Records: ${backupResult.total_records}
- Sheets Copied: ${backupResult.sheets_copied}
- Processing Time: ${backupResult.processing_time} seconds

Record Counts:
${recordCounts}

Access the backup file:
${backupResult.view_url}

Access the backup folder:
${backupResult.folder_url}

This backup is stored securely in Google Drive and is set to read-only access.

Best regards,
CFA Accountability System
    `.trim();

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });

    console.log('Backup success email sent to: ' + recipients.join(', '));
    logSystemEvent('success', 'Backup success email sent to directors', 'low');

  } catch (error) {
    console.error('Error sending backup success email: ' + error.toString());
  }
}

/**
 * Send backup failure email notification
 */
function sendBackupFailureEmail(errorMessage, backupType, attemptedBy) {
  try {
    const recipients = getDirectorEmails();

    if (recipients.length === 0) {
      console.log('No director emails configured for failure notification');
      return;
    }

    const subject = 'ALERT: Backup Failed - CFA Accountability - Action Required';

    const body = `
ALERT: System Backup Failed

A ${backupType} backup of the CFA Accountability System has FAILED.

Error Details:
${errorMessage}

Attempted By: ${attemptedBy || 'System'}
Time: ${new Date().toISOString()}

Recommended Actions:
1. Check Google Drive storage space
2. Verify spreadsheet permissions
3. Try running a manual backup
4. Contact support if issue persists

This requires immediate attention to ensure data safety.

Best regards,
CFA Accountability System
    `.trim();

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });

    console.log('Backup failure email sent');
    logSystemEvent('error', 'Backup failure email sent to directors', 'high');

  } catch (error) {
    console.error('Error sending failure email: ' + error.toString());
  }
}

/**
 * Send restore notification email
 */
function sendRestoreNotificationEmail(backupFileId, restoredSheets, confirmedBy, errors) {
  try {
    const recipients = getDirectorEmails();

    if (recipients.length === 0) {
      return;
    }

    const subject = 'Data Restored from Backup - CFA Accountability - Review Required';

    const sheetsInfo = restoredSheets
      .map(s => '  - ' + s.name + ' (' + s.rows + ' rows)')
      .join('\n');

    const errorsInfo = errors.length > 0
      ? '\nErrors Encountered:\n' + errors.map(e => '  - ' + e).join('\n')
      : '';

    const body = `
NOTICE: Data Has Been Restored

A data restore operation has been performed on the CFA Accountability System.

Restore Details:
- Performed By: ${confirmedBy}
- Backup File ID: ${backupFileId}
- Timestamp: ${new Date().toISOString()}

Sheets Restored:
${sheetsInfo}
${errorsInfo}

IMPORTANT: Please review the restored data to ensure accuracy.

If this restore was not authorized, contact the director who performed it immediately.

Best regards,
CFA Accountability System
    `.trim();

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });

    console.log('Restore notification email sent');

  } catch (error) {
    console.error('Error sending restore notification: ' + error.toString());
  }
}

/**
 * Get director email addresses from settings
 */
function getDirectorEmails() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) return [];

    const data = settingsSheet.getDataRange().getValues();

    for (const row of data) {
      if (row[0] === 'termination_email_list' || row[0] === 'director_emails') {
        if (row[1]) {
          return row[1].toString().split(',').map(e => e.trim()).filter(e => e);
        }
      }
    }

    return [];
  } catch (error) {
    console.error('Error getting director emails: ' + error.toString());
    return [];
  }
}

/**
 * Get backup settings for UI
 */
function getBackupSettings() {
  try {
    const lastBackup = getLastBackupInfo();
    const triggers = ScriptApp.getProjectTriggers();
    const hasQuarterlyTrigger = triggers.some(t => t.getHandlerFunction() === 'runAutomaticBackup');

    return {
      success: true,
      automatic_backup_enabled: hasQuarterlyTrigger,
      last_backup: lastBackup,
      retention_period: '7 years',
      backup_frequency: 'Quarterly'
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Toggle automatic backup on/off
 */
function toggleAutomaticBackup(enabled) {
  try {
    if (enabled) {
      return scheduleQuarterlyBackup();
    } else {
      return removeQuarterlyBackupTrigger();
    }
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}


// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test manual backup creation
 */
function testManualBackup() {
  console.log('=== Testing Manual Backup ===');

  const result = createSystemBackup('manual', 'Test Director');

  console.log('Result:', JSON.stringify(result, null, 2));

  if (result.success) {
    console.log('\nBackup created successfully!');
    console.log('View URL: ' + result.view_url);
    console.log('Folder URL: ' + result.folder_url);
    console.log('Total records: ' + result.total_records);
  } else {
    console.log('\nBackup FAILED: ' + result.error);
  }

  return result;
}

/**
 * Test listing backups
 */
function testListBackups() {
  console.log('=== Testing List Backups ===');

  const result = listBackups();

  console.log('Found ' + (result.backups ? result.backups.length : 0) + ' backups');

  if (result.backups) {
    result.backups.forEach((backup, index) => {
      console.log('\nBackup ' + (index + 1) + ':');
      console.log('  ID: ' + backup.backup_id);
      console.log('  Date: ' + backup.backup_date);
      console.log('  Type: ' + backup.backup_type);
      console.log('  Status: ' + backup.status);
      console.log('  Records: ' + backup.total_records);
      console.log('  Size: ' + backup.file_size_formatted);
    });
  }

  return result;
}

/**
 * Test scheduling quarterly backup
 */
function testScheduleBackup() {
  console.log('=== Testing Schedule Quarterly Backup ===');

  const result = scheduleQuarterlyBackup();

  console.log('Result:', JSON.stringify(result, null, 2));

  return result;
}

/**
 * Test getting backup settings
 */
function testGetBackupSettings() {
  console.log('=== Testing Get Backup Settings ===');

  const result = getBackupSettings();

  console.log('Result:', JSON.stringify(result, null, 2));

  return result;
}

/**
 * Test restore function (with minimal data)
 * CAUTION: This will modify data
 */
function testRestorePreview() {
  console.log('=== Testing Restore Preview (No actual restore) ===');

  // Just test getting backup list
  const backups = listBackups();

  if (backups.success && backups.backups.length > 0) {
    const latestBackup = backups.backups[0];
    console.log('\nLatest backup that could be used for restore:');
    console.log('  ID: ' + latestBackup.backup_id);
    console.log('  File ID: ' + latestBackup.file_id);
    console.log('  Date: ' + latestBackup.backup_date);
    console.log('  Status: ' + latestBackup.status);

    if (latestBackup.file_exists) {
      console.log('\n  Backup file EXISTS and can be used for restore');
    } else {
      console.log('\n  WARNING: Backup file NOT FOUND');
    }
  } else {
    console.log('No backups available for restore');
  }
}

/**
 * Full backup system test
 */
function testBackupSystem() {
  console.log('=== FULL BACKUP SYSTEM TEST ===\n');

  // Test 1: Get current settings
  console.log('Test 1: Get backup settings...');
  const settings = getBackupSettings();
  console.log('Settings loaded: ' + settings.success + '\n');

  // Test 2: List existing backups
  console.log('Test 2: List existing backups...');
  const existingBackups = listBackups();
  console.log('Existing backups: ' + (existingBackups.backups ? existingBackups.backups.length : 0) + '\n');

  // Test 3: Create manual backup
  console.log('Test 3: Create manual backup...');
  const backupResult = createSystemBackup('manual', 'System Test');
  console.log('Backup created: ' + backupResult.success);
  if (backupResult.success) {
    console.log('  - Backup ID: ' + backupResult.backup_id);
    console.log('  - Total records: ' + backupResult.total_records);
    console.log('  - Processing time: ' + backupResult.processing_time + 's');
  }
  console.log('');

  // Test 4: Verify backup in list
  console.log('Test 4: Verify backup appears in list...');
  const updatedBackups = listBackups();
  const newBackupFound = updatedBackups.backups &&
    updatedBackups.backups.some(b => b.backup_id === backupResult.backup_id);
  console.log('New backup in list: ' + newBackupFound + '\n');

  // Test 5: Get last backup info
  console.log('Test 5: Get last backup info...');
  const lastInfo = getLastBackupInfo();
  console.log('Last backup date: ' + lastInfo.last_backup_date);
  console.log('Next scheduled: ' + lastInfo.next_scheduled + '\n');

  console.log('=== BACKUP SYSTEM TEST COMPLETE ===');

  return {
    settings_ok: settings.success,
    backup_created: backupResult.success,
    backup_in_list: newBackupFound,
    backup_id: backupResult.backup_id
  };
}
