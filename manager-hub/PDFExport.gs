// ============================================
// PHASE 18: PDF EXPORT GENERATION
// ============================================
// Functions for generating formatted PDF write-up
// documents with employee infraction history
// ============================================

/**
 * ========================================
 * AUTHORIZATION FUNCTION - STEP 1
 * ========================================
 * Run this function FIRST to trigger the Drive authorization prompt.
 *
 * When you run this, you MUST click "Review Permissions" and authorize.
 * If you don't see the prompt, try clicking Run again.
 */
function authorizeDrive() {
  // These calls will trigger authorization prompts - DO NOT wrap in try/catch
  const folders = DriveApp.getFoldersByName('Authorization Test');
  const rootFolder = DriveApp.getRootFolder();
  console.log('DriveApp authorized successfully!');
  console.log('Root folder name: ' + rootFolder.getName());
  return 'DriveApp authorization complete';
}

/**
 * ========================================
 * AUTHORIZATION FUNCTION - STEP 2
 * ========================================
 * Run this function AFTER authorizeDrive to trigger Docs authorization.
 */
function authorizeDocuments() {
  // Create a temp document to trigger authorization
  const doc = DocumentApp.create('Temp Auth Doc - Safe to Delete');
  const docId = doc.getId();
  console.log('DocumentApp authorized successfully!');
  console.log('Created temp doc: ' + docId);

  // Clean up - trash the temp doc
  doc.saveAndClose();
  DriveApp.getFileById(docId).setTrashed(true);
  console.log('Temp doc deleted');

  return 'DocumentApp authorization complete';
}

/**
 * ========================================
 * AUTHORIZATION FUNCTION - STEP 3 (OPTIONAL)
 * ========================================
 * Run this function ONLY if you want to use the email PDF feature.
 */
function authorizeGmail() {
  // This triggers Gmail authorization
  const drafts = GmailApp.getDrafts();
  console.log('GmailApp authorized successfully!');
  console.log('Number of drafts: ' + drafts.length);
  return 'GmailApp authorization complete';
}

/**
 * ========================================
 * VERIFICATION FUNCTION
 * ========================================
 * Run this AFTER completing authorization to verify everything works.
 */
function verifyAllPermissions() {
  console.log('=== Verifying All Permissions ===');
  console.log('');

  let allGood = true;

  // Test DriveApp
  try {
    DriveApp.getRootFolder();
    console.log('✓ DriveApp: AUTHORIZED');
  } catch (e) {
    console.log('✗ DriveApp: NOT AUTHORIZED - Run authorizeDrive() first');
    allGood = false;
  }

  // Test DocumentApp
  try {
    const doc = DocumentApp.create('Verify Test');
    const docId = doc.getId();
    doc.saveAndClose();
    DriveApp.getFileById(docId).setTrashed(true);
    console.log('✓ DocumentApp: AUTHORIZED');
  } catch (e) {
    console.log('✗ DocumentApp: NOT AUTHORIZED - Run authorizeDocuments() first');
    allGood = false;
  }

  // Test HtmlService
  try {
    HtmlService.createHtmlOutput('<p>Test</p>');
    console.log('✓ HtmlService: AUTHORIZED');
  } catch (e) {
    console.log('✗ HtmlService: Issue - ' + e.message);
    allGood = false;
  }

  // Test GmailApp (optional)
  try {
    GmailApp.getDrafts();
    console.log('✓ GmailApp: AUTHORIZED (optional)');
  } catch (e) {
    console.log('⚠ GmailApp: Not authorized (optional - only needed for email feature)');
  }

  console.log('');
  if (allGood) {
    console.log('=== ALL REQUIRED PERMISSIONS GRANTED ===');
    console.log('PDF Export should now work from the web app!');
  } else {
    console.log('=== SOME PERMISSIONS MISSING ===');
    console.log('Run the authorization functions listed above.');
  }

  return allGood ? 'All permissions verified!' : 'Some permissions missing';
}

/**
 * Main function to generate a PDF write-up for an employee.
 * Can generate for a specific infraction or full accountability record.
 *
 * @param {string} employeeId - Employee being written up
 * @param {string} infractionId - Specific infraction ID (optional, null for full record)
 * @param {Object} options - Additional options (notes, actionsTaken, etc.)
 * @returns {Object} Result with PDF URL, file ID, or error
 */
function generateWriteUpPDF(employeeId, infractionId, options) {
  const startTime = Date.now();

  try {
    // Step 1: Validate session - only Directors and Operators can generate PDFs
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true, error: 'Session expired' };
    }

    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Only Directors and Operators can generate write-up PDFs' };
    }

    const generatedBy = session.role;

    // Step 2: Get employee detail data
    const detailData = getEmployeeDetailData(employeeId);
    if (!detailData.success) {
      return { success: false, error: detailData.error || 'Failed to get employee data' };
    }

    // Step 3: Get specific infraction if ID provided
    let specificInfraction = null;
    if (infractionId) {
      specificInfraction = detailData.infractions.find(inf => inf.infraction_id === infractionId);
      if (!specificInfraction) {
        return { success: false, error: 'Specified infraction not found' };
      }
    }

    // Step 4: Get all threshold data for consequences
    const thresholdData = getAllThresholdData();

    // Step 5: Build template data
    const templateData = buildPDFTemplateData(detailData, specificInfraction, thresholdData, options, generatedBy);

    // Step 6: Create PDF HTML template
    const htmlContent = createPDFTemplate(templateData);

    // Step 7: Generate PDF using Google Docs approach
    const pdfResult = convertHtmlToPDF(htmlContent, templateData.filename);

    if (!pdfResult.success) {
      return { success: false, error: pdfResult.error };
    }

    // Step 8: Save to Drive
    const saveResult = savePDFToDrive(pdfResult.blob, templateData.filename, employeeId);

    if (!saveResult.success) {
      return { success: false, error: saveResult.error };
    }

    // Step 9: Log PDF generation
    logPDFGeneration(employeeId, templateData.documentId, generatedBy, infractionId);

    const duration = Date.now() - startTime;
    console.log(`PDF generated in ${duration}ms for employee ${employeeId}`);

    return {
      success: true,
      message: 'PDF generated successfully',
      fileId: saveResult.fileId,
      downloadUrl: saveResult.downloadUrl,
      viewUrl: saveResult.viewUrl,
      documentId: templateData.documentId,
      filename: templateData.filename
    };

  } catch (error) {
    console.error('Error in generateWriteUpPDF:', error.toString());
    return { success: false, error: 'Error generating PDF: ' + error.message };
  }
}

/**
 * Builds the template data object for PDF generation.
 */
function buildPDFTemplateData(detailData, specificInfraction, thresholdData, options, generatedBy) {
  const now = new Date();
  const documentId = generateDocumentId();

  // Get employee info
  const employee = detailData.employee;
  const currentPoints = detailData.currentPoints;
  const infractions = detailData.infractions || [];
  const statistics = detailData.statistics || {};

  // Filter to active infractions only for history
  const activeInfractions = infractions.filter(inf =>
    inf.status === 'Active' && !inf.is_expired && !inf.is_positive
  );

  // Get last 10 infractions for history table
  const infractionHistory = activeInfractions.slice(0, 10);

  // Determine current threshold level and applicable consequences
  const applicableConsequences = [];
  for (const threshold of thresholdData) {
    if (currentPoints.total >= threshold.threshold) {
      applicableConsequences.push({
        threshold: threshold.threshold,
        consequence: threshold.consequence
      });
    }
  }

  // Build filename
  const employeeName = employee.full_name.replace(/[^a-zA-Z0-9]/g, '_');
  const dateStr = formatDateForFilename(now);
  const filename = `WriteUp_${employeeName}_${dateStr}_${documentId}.pdf`;

  return {
    documentId: documentId,
    filename: filename,
    generatedDate: formatDateFull(now),
    generatedTime: formatTime(now),
    generatedBy: generatedBy,

    // Employee info
    employee: {
      fullName: employee.full_name,
      employeeId: employee.employee_id,
      primaryLocation: employee.primary_location || 'Not Assigned',
      hireDate: employee.hire_date ? formatDateShort(new Date(employee.hire_date)) : 'N/A'
    },

    // Current status
    currentPoints: currentPoints.total,
    pointLevelColor: currentPoints.pointLevelColor,
    statusBadges: currentPoints.statusBadges || [],

    // Specific infraction (if any)
    specificInfraction: specificInfraction ? {
      date: specificInfraction.dateFormatted,
      location: specificInfraction.location,
      type: specificInfraction.infraction_type,
      points: specificInfraction.points,
      description: specificInfraction.description || 'No description provided'
    } : null,

    // Infraction history - show full descriptions
    infractionHistory: infractionHistory.map(inf => ({
      date: inf.dateFormatted,
      type: inf.infraction_type,
      points: inf.points,
      description: inf.description || '',
      location: inf.location
    })),

    // Statistics
    totalInfractions: statistics.totalInfractionsEver || 0,
    activeInfractions: statistics.activeInfractions || 0,

    // Consequences
    applicableConsequences: applicableConsequences,
    highestThreshold: applicableConsequences.length > 0
      ? applicableConsequences[applicableConsequences.length - 1].threshold
      : 0,

    // Options from modal
    coachingNotes: options?.coachingNotes || '',
    additionalNotes: options?.additionalNotes || '',
    actionsTaken: options?.actionsTaken || []
  };
}

/**
 * Creates the HTML template for the PDF.
 * Uses inline styles for PDF compatibility.
 */
function createPDFTemplate(data) {
  // Build infraction history table rows
  let historyRowsHtml = '';
  if (data.infractionHistory.length > 0) {
    historyRowsHtml = data.infractionHistory.map(inf => `
      <tr>
        <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtmlForPdf(inf.date)}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtmlForPdf(inf.type)}</td>
        <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${inf.points}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtmlForPdf(inf.description)}</td>
        <td style="padding: 8px; border: 1px solid #ddd;">${escapeHtmlForPdf(inf.location)}</td>
      </tr>
    `).join('');
  } else {
    historyRowsHtml = `
      <tr>
        <td colspan="5" style="padding: 10px; text-align: center; color: #000; border: 1px solid #ccc;">
          No active infractions on record
        </td>
      </tr>
    `;
  }

  // Build consequences list
  let consequencesHtml = '';
  if (data.applicableConsequences.length > 0) {
    consequencesHtml = data.applicableConsequences.map(c => `
      <li style="margin-bottom: 8px;">
        <strong>${c.threshold} Points:</strong> ${escapeHtmlForPdf(c.consequence)}
      </li>
    `).join('');
  } else {
    consequencesHtml = '<li>No threshold consequences currently applicable</li>';
  }

  // Build actions taken checkboxes
  const actionsHtml = buildActionCheckboxesHtml(data.actionsTaken);

  // Build specific incident section if applicable
  let incidentSectionHtml = '';
  if (data.specificInfraction) {
    incidentSectionHtml = `
      <div style="margin-bottom: 25px;">
        <h2 style="color: #DD0031; font-size: 16px; border-bottom: 2px solid #DD0031; padding-bottom: 5px; margin-bottom: 15px;">
          Incident Information
        </h2>
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="padding: 8px; width: 40%; font-weight: bold;">Date of Incident:</td>
            <td style="padding: 8px;">${escapeHtmlForPdf(data.specificInfraction.date)}</td>
          </tr>
          <tr style="background-color: #f9f9f9;">
            <td style="padding: 8px; font-weight: bold;">Location of Incident:</td>
            <td style="padding: 8px;">${escapeHtmlForPdf(data.specificInfraction.location)}</td>
          </tr>
          <tr>
            <td style="padding: 8px; font-weight: bold;">Type of Infraction:</td>
            <td style="padding: 8px;">${escapeHtmlForPdf(data.specificInfraction.type)}</td>
          </tr>
          <tr style="background-color: #f9f9f9;">
            <td style="padding: 8px; font-weight: bold;">Points Assigned:</td>
            <td style="padding: 8px; color: #DD0031; font-weight: bold;">${data.specificInfraction.points}</td>
          </tr>
        </table>
        <div style="margin-top: 15px;">
          <p style="font-weight: bold; margin-bottom: 5px;">Description of Incident:</p>
          <div style="background-color: #f5f5f5; padding: 15px; border-radius: 4px; border-left: 4px solid #DD0031;">
            ${escapeHtmlForPdf(data.specificInfraction.description)}
          </div>
        </div>
      </div>
    `;
  }

  // Main HTML template
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Disciplinary Write-Up - ${escapeHtmlForPdf(data.employee.fullName)}</title>
  <style>
    @page {
      size: letter;
      margin: 0.5in;
    }
    body {
      font-family: Arial, Helvetica, sans-serif;
      font-size: 10pt;
      line-height: 1.3;
      color: #000;
      margin: 0;
      padding: 0;
    }
    .page-header {
      text-align: center;
      margin-bottom: 12px;
      padding-bottom: 8px;
      border-bottom: 3px solid #DD0031;
    }
    .logo-placeholder {
      font-size: 22px;
      font-weight: bold;
      color: #DD0031;
      margin-bottom: 3px;
    }
    h1 {
      font-size: 18px;
      color: #000;
      margin: 5px 0 3px 0;
    }
    .doc-info {
      font-size: 9pt;
      color: #000;
    }
    .section {
      margin-bottom: 12px;
    }
    .section-title {
      color: #DD0031;
      font-size: 12px;
      font-weight: bold;
      border-bottom: 2px solid #DD0031;
      padding-bottom: 3px;
      margin-bottom: 8px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
    }
    .info-table td {
      padding: 3px 6px;
      vertical-align: top;
      color: #000;
    }
    .info-table td:first-child {
      font-weight: bold;
      width: 30%;
    }
    .history-table {
      font-size: 9pt;
    }
    .history-table th {
      background-color: #DD0031;
      color: white;
      padding: 6px 5px;
      text-align: left;
      font-weight: bold;
    }
    .history-table td {
      padding: 5px;
      border: 1px solid #ccc;
      color: #000;
    }
    .history-table tr:nth-child(even) {
      background-color: #f5f5f5;
    }
    .checkbox-item {
      margin-bottom: 4px;
      color: #000;
    }
    .checkbox {
      display: inline-block;
      width: 12px;
      height: 12px;
      border: 1px solid #000;
      margin-right: 6px;
      vertical-align: middle;
      text-align: center;
      font-size: 9px;
      line-height: 10px;
    }
    .checkbox.checked {
      background-color: #DD0031;
      color: white;
    }
    .signature-line {
      border-bottom: 1px solid #000;
      min-width: 180px;
      display: inline-block;
    }
    .signature-section {
      margin-top: 15px;
    }
    .notes-box {
      border: 1px solid #ccc;
      min-height: 50px;
      padding: 8px;
      margin-top: 5px;
      color: #000;
    }
    .disclaimer {
      font-size: 8pt;
      color: #000;
      text-align: justify;
      background-color: #f5f5f5;
      padding: 10px;
      border-radius: 4px;
      margin-top: 12px;
    }
    .footer {
      margin-top: 15px;
      padding-top: 8px;
      border-top: 1px solid #ccc;
      font-size: 8pt;
      color: #000;
      text-align: center;
    }
    .points-badge {
      display: inline-block;
      background-color: #DD0031;
      color: white;
      padding: 2px 10px;
      border-radius: 12px;
      font-weight: bold;
    }
    .status-badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 9pt;
      margin-left: 4px;
    }
    .status-badge.probation { background-color: #fd7e14; color: white; }
    .status-badge.final_warning { background-color: #dc3545; color: white; }
    .status-badge.termination { background-color: #721c24; color: white; }
    .sig-table td { padding: 8px 0; vertical-align: bottom; color: #000; }
  </style>
</head>
<body>
  <!-- Header -->
  <div class="page-header">
    <div class="logo-placeholder">Chick-fil-A</div>
    <h1>Disciplinary Write-Up</h1>
    <div class="doc-info">
      Document Date: ${escapeHtmlForPdf(data.generatedDate)} at ${escapeHtmlForPdf(data.generatedTime)}<br>
      Document ID: ${escapeHtmlForPdf(data.documentId)}
    </div>
  </div>

  <!-- Employee Information -->
  <div class="section">
    <div class="section-title">Employee Information</div>
    <table class="info-table">
      <tr>
        <td>Name:</td>
        <td><strong>${escapeHtmlForPdf(data.employee.fullName)}</strong></td>
      </tr>
      <tr>
        <td>Employee ID:</td>
        <td>${escapeHtmlForPdf(data.employee.employeeId)}</td>
      </tr>
      <tr>
        <td>Primary Location:</td>
        <td>${escapeHtmlForPdf(data.employee.primaryLocation)}</td>
      </tr>
      <tr>
        <td>Current Point Total:</td>
        <td>
          <span class="points-badge">${data.currentPoints} points</span>
          ${data.statusBadges.map(b => `<span class="status-badge ${b.type}">${b.label}</span>`).join('')}
        </td>
      </tr>
    </table>
  </div>

  ${incidentSectionHtml}

  <!-- Infraction History -->
  <div class="section">
    <div class="section-title">Infraction History (Active)</div>
    <table class="history-table">
      <thead>
        <tr>
          <th style="width: 15%;">Date</th>
          <th style="width: 20%;">Type</th>
          <th style="width: 10%;">Points</th>
          <th style="width: 35%;">Description</th>
          <th style="width: 20%;">Location</th>
        </tr>
      </thead>
      <tbody>
        ${historyRowsHtml}
      </tbody>
    </table>
    <p style="font-size: 9pt; color: #000; margin-top: 5px;">
      Total Active Infractions: ${data.activeInfractions} | Total Infractions Ever: ${data.totalInfractions}
    </p>
  </div>

  <!-- Current Status & Consequences -->
  <div class="section">
    <div class="section-title">Current Status & Applicable Consequences</div>
    <p><strong>Current Point Total:</strong> ${data.currentPoints} points
       ${data.highestThreshold > 0 ? `(Threshold Level: ${data.highestThreshold} points reached)` : '(Below first threshold)'}</p>
    <p><strong>Active Consequences:</strong></p>
    <ul style="margin: 10px 0; padding-left: 25px;">
      ${consequencesHtml}
    </ul>
  </div>

  <!-- Coaching Given -->
  <div class="section">
    <div class="section-title">Coaching Given</div>
    <div class="notes-box">
      ${data.coachingNotes ? escapeHtmlForPdf(data.coachingNotes) : '<span style="color: #000;">[Document the coaching conversation here]</span>'}
    </div>
  </div>

  <!-- Disciplinary Action Taken -->
  <div class="section">
    <div class="section-title">Disciplinary Action Taken</div>
    ${actionsHtml}
  </div>

  <!-- Progressive Discipline Statement -->
  <div class="disclaimer">
    <strong>Progressive Discipline Statement:</strong><br><br>
    Where progressive discipline is appropriate, the following types of disciplinary action may be taken, in no particular order:
    <br><br>
    Disciplinary actions will be approached on a case-by-case basis, taking into account all the relevant facts and factors of the situation. Therefore, Chick-fil-A Cockrell Hill DTO retains the rights to skip any of these steps of progressive discipline if circumstances necessitate. Chick-fil-A Cockrell Hill DTO also reserves the right to discipline an employee at any time for inappropriate conduct or behavior, whether or not such conduct is referenced or mentioned in this policy. Nothing in this policy is a guarantee that any particular disciplinary steps will be followed in any given case, or at all, and this policy does not reflect any contractual agreement or right of any team member that any particular disciplinary steps will be followed in any given case. Employment at Chick-fil-A Cockrell Hill DTO remains at-will.
  </div>

  <!-- Signature Section -->
  <div class="signature-section">
    <div class="section-title">Signatures</div>
    <table class="sig-table" style="width: 100%;">
      <tr>
        <td style="width: 50%;"><strong>Team Member Signature:</strong> <span class="signature-line" style="width: 200px;"></span></td>
        <td style="width: 25%;"><strong>Date:</strong> <span class="signature-line" style="width: 100px;"></span></td>
      </tr>
      <tr>
        <td><strong>Manager/Director (${escapeHtmlForPdf(data.generatedBy)}):</strong> <span class="signature-line" style="width: 180px;"></span></td>
        <td><strong>Date:</strong> <span class="signature-line" style="width: 100px;"></span></td>
      </tr>
      <tr>
        <td><strong>Witness Name/Signature:</strong> <span class="signature-line" style="width: 200px;"></span></td>
        <td><strong>Date:</strong> <span class="signature-line" style="width: 100px;"></span></td>
      </tr>
    </table>
  </div>

  <!-- Additional Notes -->
  <div class="section">
    <div class="section-title">Additional Notes</div>
    <div class="notes-box" style="min-height: 40px;">
      ${data.additionalNotes ? escapeHtmlForPdf(data.additionalNotes) : ''}
    </div>
  </div>

  <!-- Footer -->
  <div class="footer">
    <p style="margin: 0;">Document ID: ${escapeHtmlForPdf(data.documentId)} | Generated: ${escapeHtmlForPdf(data.generatedDate)} at ${escapeHtmlForPdf(data.generatedTime)}</p>
  </div>
</body>
</html>
  `;

  return html;
}

/**
 * Builds HTML for action checkboxes.
 */
function buildActionCheckboxesHtml(actionsTaken) {
  const actions = [
    { key: 'verbal_warning', label: 'Verbal Warning' },
    { key: 'written_warning', label: 'Written Warning' },
    { key: 'removed_break', label: 'Removed break privileges' },
    { key: 'director_meeting', label: 'Director meeting held' },
    { key: 'probation_started', label: 'Probation started' },
    { key: 'suspension', label: 'Suspension' },
    { key: 'day_removed', label: 'Day removed from schedule' },
    { key: 'other', label: 'Other' }
  ];

  const actionSet = new Set(actionsTaken || []);

  let html = '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 4px;">';

  for (const action of actions) {
    const isChecked = actionSet.has(action.key);
    html += `
      <div class="checkbox-item">
        <span class="checkbox ${isChecked ? 'checked' : ''}">${isChecked ? '&#10003;' : ''}</span>
        ${action.label}
        ${action.key === 'other' ? ': _______________________' : ''}
      </div>
    `;
  }

  html += '</div>';
  return html;
}

/**
 * Converts HTML content to PDF using Google Docs.
 */
function convertHtmlToPDF(htmlContent, filename) {
  try {
    // Create a temporary Google Doc
    const tempDoc = DocumentApp.create('TempWriteUp_' + Date.now());
    const docId = tempDoc.getId();

    // Get the document body
    const body = tempDoc.getBody();

    // Unfortunately, Google Docs doesn't directly accept HTML
    // So we use a different approach: create HTML file and convert

    // Close the temp doc
    tempDoc.saveAndClose();

    // Delete the temp doc
    DriveApp.getFileById(docId).setTrashed(true);

    // Use HtmlService to create a blob from HTML
    const htmlBlob = HtmlService.createHtmlOutput(htmlContent)
      .getBlob()
      .setName(filename);

    // Convert to PDF
    // Note: This creates a PDF directly from the HTML blob
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName(filename);

    return {
      success: true,
      blob: pdfBlob
    };

  } catch (error) {
    console.error('Error converting HTML to PDF:', error.toString());
    return {
      success: false,
      error: 'Error creating PDF: ' + error.message
    };
  }
}

/**
 * Saves PDF to Google Drive in a dedicated folder.
 */
function savePDFToDrive(pdfBlob, filename, employeeId) {
  try {
    // Get or create the Write-Up PDFs folder
    const folder = getOrCreatePDFFolder();

    // Create the file in the folder
    const file = folder.createFile(pdfBlob);
    file.setName(filename);

    // Set permissions - anyone with the link can view
    // In production, you might want to restrict this further
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      fileId: file.getId(),
      downloadUrl: file.getDownloadUrl(),
      viewUrl: file.getUrl()
    };

  } catch (error) {
    console.error('Error saving PDF to Drive:', error.toString());
    return {
      success: false,
      error: 'Error saving PDF: ' + error.message
    };
  }
}

/**
 * Gets or creates the Write-Up PDFs folder in Drive.
 */
function getOrCreatePDFFolder() {
  const folderName = 'CFA Accountability - Write-Up PDFs';

  // Search for existing folder
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  // Create new folder
  const newFolder = DriveApp.createFolder(folderName);

  // Add a description
  newFolder.setDescription('Auto-generated write-up PDFs from CFA Accountability System');

  return newFolder;
}

/**
 * Logs PDF generation to Email_Log sheet.
 */
function logPDFGeneration(employeeId, documentId, generatedBy, infractionId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName('Email_Log');

    if (!logSheet) {
      console.log('Email_Log sheet not found, skipping PDF log');
      return;
    }

    const timestamp = new Date();
    const logEntry = [
      timestamp,                                    // Timestamp
      'PDF Generated',                              // Event Type
      employeeId,                                   // Employee ID
      infractionId || 'Full Record',               // Infraction ID or Full Record
      documentId,                                   // Document ID
      generatedBy,                                  // Generated By
      'Success'                                     // Status
    ];

    logSheet.appendRow(logEntry);

  } catch (error) {
    console.error('Error logging PDF generation:', error.toString());
  }
}

/**
 * Generates a unique document ID.
 */
function generateDocumentId() {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substring(2, 8);
  return `WU-${timestamp}-${random}`.toUpperCase();
}

/**
 * Formats a date for filename (YYYYMMDD).
 */
function formatDateForFilename(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * Formats a date in full format (Month DD, YYYY).
 */
function formatDateFull(date) {
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

/**
 * Formats a date in short format (MM/DD/YYYY).
 */
function formatDateShort(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

/**
 * Formats time (HH:MM AM/PM).
 */
function formatTime(date) {
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
}

/**
 * Truncates text to a maximum length.
 */
function truncateText(text, maxLength) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

/**
 * Escapes HTML for safe PDF output.
 */
function escapeHtmlForPdf(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
    .replace(/\n/g, '<br>');
}

/**
 * Gets all threshold data (value and consequence) as an array of objects.
 * Note: This may already exist in Code.gs, this is a fallback
 */
function getAllThresholdData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');

    if (!settingsSheet) {
      return [];
    }

    const data = settingsSheet.getRange('A13:B19').getValues();
    const thresholds = [];

    for (const row of data) {
      const threshold = Number(row[0]);
      const consequence = row[1];

      if (!isNaN(threshold) && threshold > 0 && consequence) {
        thresholds.push({
          threshold: threshold,
          consequence: consequence
        });
      }
    }

    thresholds.sort((a, b) => a.threshold - b.threshold);
    return thresholds;

  } catch (error) {
    console.error('Error getting threshold data:', error.toString());
    return [];
  }
}

// ============================================
// BULK PDF EXPORT
// ============================================

/**
 * Generates PDFs for multiple employees.
 * Returns a zip file with all PDFs or individual results.
 */
function generateBulkPDFs(employeeIds, options) {
  try {
    // Validate session
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Only Directors and Operators can generate PDFs' };
    }

    if (!employeeIds || employeeIds.length === 0) {
      return { success: false, error: 'No employees selected' };
    }

    if (employeeIds.length > 20) {
      return { success: false, error: 'Maximum 20 employees per batch' };
    }

    const results = [];
    const blobs = [];

    for (const employeeId of employeeIds) {
      const result = generateWriteUpPDF(employeeId, null, options);
      results.push({
        employeeId: employeeId,
        success: result.success,
        error: result.error,
        fileId: result.fileId
      });
    }

    const successCount = results.filter(r => r.success).length;
    const failCount = results.filter(r => !r.success).length;

    return {
      success: true,
      message: `Generated ${successCount} PDFs successfully${failCount > 0 ? `, ${failCount} failed` : ''}`,
      results: results,
      successCount: successCount,
      failCount: failCount
    };

  } catch (error) {
    console.error('Error in bulk PDF generation:', error.toString());
    return { success: false, error: error.message };
  }
}

// ============================================
// EMAIL PDF OPTION
// ============================================

/**
 * Emails a generated PDF to specified recipients.
 */
function emailWriteUpPDF(fileId, recipients, employeeName, documentId) {
  try {
    // Validate session
    const session = getCurrentRole();
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    if (!fileId || !recipients || recipients.length === 0) {
      return { success: false, error: 'File ID and recipients are required' };
    }

    // Get the PDF file
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();

    // Build email
    const subject = `Write-Up Document for ${employeeName} - ${documentId}`;
    const body = `
A write-up document has been generated for ${employeeName}.

Document ID: ${documentId}
Generated By: ${session.role}
Generated On: ${new Date().toLocaleString()}

The PDF is attached to this email.

---
This is an automated message from the CFA Accountability System.
    `;

    // Send email with attachment
    for (const recipient of recipients) {
      GmailApp.sendEmail(recipient, subject, body, {
        attachments: [blob],
        name: 'CFA Accountability System'
      });
    }

    // Log the email
    logPDFEmail(employeeName, documentId, recipients.join(', '), session.role);

    return {
      success: true,
      message: `PDF emailed to ${recipients.length} recipient(s)`
    };

  } catch (error) {
    console.error('Error emailing PDF:', error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Logs PDF email to Email_Log.
 */
function logPDFEmail(employeeName, documentId, recipients, sentBy) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName('Email_Log');

    if (!logSheet) return;

    logSheet.appendRow([
      new Date(),
      'PDF Emailed',
      employeeName,
      documentId,
      recipients,
      sentBy,
      'Success'
    ]);

  } catch (error) {
    console.error('Error logging PDF email:', error.toString());
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test function for PDF generation.
 */
function testPDFGeneration() {
  console.log('=== Testing PDF Generation ===');
  console.log('');

  const testResults = [];

  // Get test employee
  const employees = getActiveEmployees();
  if (!employees || employees.length === 0) {
    console.log('FAIL: No employees found for testing');
    return { success: false, message: 'No employees available' };
  }

  const testEmployee = employees[0];
  console.log(`Testing with employee: ${testEmployee.full_name} (${testEmployee.employee_id})`);
  console.log('');

  // Test Case 1: Generate full record PDF
  console.log('Test Case 1: Generate full accountability record PDF');
  const startTime1 = Date.now();
  const result1 = generateWriteUpPDF(testEmployee.employee_id, null, {
    coachingNotes: 'Test coaching notes - discussed expectations and improvement plan.',
    additionalNotes: 'Test additional notes for the write-up document.',
    actionsTaken: ['verbal_warning', 'director_meeting']
  });
  const duration1 = Date.now() - startTime1;

  console.log(`  Result: ${result1.success ? 'SUCCESS' : 'FAILED'}`);
  console.log(`  Duration: ${duration1}ms`);
  if (result1.success) {
    console.log(`  File ID: ${result1.fileId}`);
    console.log(`  Document ID: ${result1.documentId}`);
    console.log(`  View URL: ${result1.viewUrl}`);
  } else {
    console.log(`  Error: ${result1.error}`);
  }
  testResults.push({ test: 'Full Record PDF', passed: result1.success, duration: duration1 });
  console.log('');

  // Test Case 2: Test with employee detail data
  console.log('Test Case 2: Verify employee detail data retrieval');
  const detailData = getEmployeeDetailData(testEmployee.employee_id);
  console.log(`  Detail data success: ${detailData.success}`);
  if (detailData.success) {
    console.log(`  Current points: ${detailData.currentPoints.total}`);
    console.log(`  Infractions count: ${detailData.infractions ? detailData.infractions.length : 0}`);
  }
  testResults.push({ test: 'Detail Data', passed: detailData.success });
  console.log('');

  // Test Case 3: Test folder creation
  console.log('Test Case 3: Verify PDF folder creation');
  try {
    const folder = getOrCreatePDFFolder();
    console.log(`  Folder name: ${folder.getName()}`);
    console.log(`  Folder ID: ${folder.getId()}`);
    testResults.push({ test: 'Folder Creation', passed: true });
  } catch (error) {
    console.log(`  Error: ${error.message}`);
    testResults.push({ test: 'Folder Creation', passed: false });
  }
  console.log('');

  // Test Case 4: Test document ID generation
  console.log('Test Case 4: Verify document ID generation');
  const docId1 = generateDocumentId();
  const docId2 = generateDocumentId();
  console.log(`  Generated ID 1: ${docId1}`);
  console.log(`  Generated ID 2: ${docId2}`);
  console.log(`  IDs are unique: ${docId1 !== docId2}`);
  testResults.push({ test: 'Document ID Generation', passed: docId1 !== docId2 });
  console.log('');

  // Test Case 5: Test threshold data retrieval
  console.log('Test Case 5: Verify threshold data retrieval');
  const thresholds = getAllThresholdData();
  console.log(`  Thresholds found: ${thresholds.length}`);
  if (thresholds.length > 0) {
    console.log(`  First threshold: ${thresholds[0].threshold} - ${thresholds[0].consequence}`);
  }
  testResults.push({ test: 'Threshold Data', passed: thresholds.length > 0 });
  console.log('');

  // Summary
  console.log('=== Test Summary ===');
  const passed = testResults.filter(r => r.passed).length;
  const failed = testResults.filter(r => !r.passed).length;
  console.log(`Passed: ${passed}/${testResults.length}`);
  console.log(`Failed: ${failed}/${testResults.length}`);

  for (const result of testResults) {
    console.log(`  ${result.passed ? '✓' : '✗'} ${result.test}${result.duration ? ` (${result.duration}ms)` : ''}`);
  }

  return {
    success: failed === 0,
    message: `${passed}/${testResults.length} tests passed`,
    results: testResults
  };
}

/**
 * Quick test to generate a sample PDF and open it.
 */
function quickTestPDF() {
  const employees = getActiveEmployees();
  if (!employees || employees.length === 0) {
    console.log('No employees found');
    return;
  }

  const result = generateWriteUpPDF(employees[0].employee_id, null, {
    coachingNotes: 'Employee was counseled regarding recent performance issues. Discussed expectations and created improvement plan.',
    additionalNotes: 'Follow-up meeting scheduled for next week.',
    actionsTaken: ['verbal_warning']
  });

  console.log('Result:', JSON.stringify(result, null, 2));

  if (result.success) {
    console.log('\nOpen this URL to view the PDF:');
    console.log(result.viewUrl);
  }
}
