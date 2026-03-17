// ============================================
// MICRO-PHASE 28: ADVANCED REPORTING & ANALYTICS
// ============================================

function requireDirectorSession(token) {
  if (!token) {
    return { ok: true, session: { role: 'System', user_name: 'System' } };
  }
  const session = getCurrentRole(token);
  if (!session || !session.authenticated) {
    return { ok: false, sessionExpired: true, error: 'Session expired' };
  }
  if (session.role !== 'Director' && session.role !== 'Operator') {
    return { ok: false, error: 'Access denied' };
  }
  return { ok: true, session: session };
}

function getReportsFolder() {
  const folders = DriveApp.getFoldersByName('Reports');
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder('Reports');
}

function getOrCreateReportHistorySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Report_History');
  if (!sheet) {
    sheet = ss.insertSheet('Report_History');
    sheet.appendRow([
      'Report_ID',
      'Report_Type',
      'Generated_Date',
      'Generated_By',
      'Date_Range',
      'File_URL',
      'File_Size',
      'Format'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateScheduledReportsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Scheduled_Reports');
  if (!sheet) {
    sheet = ss.insertSheet('Scheduled_Reports');
    sheet.appendRow([
      'Schedule_ID',
      'Report_Name',
      'Report_Type',
      'Frequency',
      'Day_Of_Week',
      'Day_Of_Month',
      'Time',
      'Recipients',
      'Report_Config_JSON',
      'Status',
      'Next_Run',
      'Trigger_ID'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateReportTemplatesSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Report_Templates');
  if (!sheet) {
    sheet = ss.insertSheet('Report_Templates');
    sheet.appendRow([
      'Template_ID',
      'Template_Name',
      'Created_By',
      'Created_At',
      'Report_Config_JSON'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateReportId() {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const random = Math.floor(1000 + Math.random() * 9000);
  return `RPT-${timestamp}-${random}`;
}

function parseDate(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return null;
  return date;
}

function normalizeDateRange(dateRange) {
  const start = parseDate(dateRange?.start_date);
  const end = parseDate(dateRange?.end_date);
  if (!start || !end) return null;
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);
  return { start, end };
}

function getInfractionsForRange(startDate, endDate) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Infractions');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const result = [];

  data.forEach(row => {
    const date = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (!date || isNaN(date.getTime())) return;
    if (date < startDate || date > endDate) return;

    result.push({
      infraction_id: row[0],
      employee_id: row[1],
      full_name: row[2],
      date: date,
      infraction_type: row[4],
      points: Number(row[5]) || 0,
      bucket: row[6],
      description: row[7],
      location: row[8],
      entered_by: row[9],
      status: row[14],
      expiration_date: row[15]
    });
  });

  return result;
}

function applyReportFilters(infractions, filters) {
  let output = infractions.slice();
  if (filters?.locations && filters.locations.length) {
    output = output.filter(inf => filters.locations.includes(inf.location));
  }
  if (filters?.infraction_types && filters.infraction_types.length) {
    output = output.filter(inf => filters.infraction_types.includes(inf.infraction_type));
  }
  if (filters?.managers && filters.managers.length) {
    output = output.filter(inf => filters.managers.includes(inf.entered_by));
  }
  if (filters?.point_range) {
    const min = Number(filters.point_range.min ?? -999);
    const max = Number(filters.point_range.max ?? 999);
    output = output.filter(inf => inf.points >= min && inf.points <= max);
  }
  return output;
}

function buildReportMetrics(infractions) {
  const totals = {
    total_infractions: infractions.length,
    total_points: infractions.reduce((sum, inf) => sum + (Number(inf.points) || 0), 0)
  };
  const byType = {};
  const byLocation = {};
  const byManager = {};
  infractions.forEach(inf => {
    const type = inf.infraction_type || 'Unknown';
    const loc = inf.location || 'Unknown';
    const manager = inf.entered_by || 'Unknown';
    byType[type] = (byType[type] || 0) + 1;
    byLocation[loc] = (byLocation[loc] || 0) + 1;
    byManager[manager] = (byManager[manager] || 0) + 1;
  });
  return { totals, byType, byLocation, byManager };
}

function generateInfractionTrendChart(dateRange) {
  const range = normalizeDateRange(dateRange);
  if (!range) return { success: false, message: 'Invalid date range' };

  const infractions = getInfractionsForRange(range.start, range.end);
  const counts = {};
  infractions.forEach(inf => {
    const key = Utilities.formatDate(inf.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    counts[key] = (counts[key] || 0) + 1;
  });

  const days = [];
  const current = new Date(range.start);
  while (current <= range.end) {
    const key = Utilities.formatDate(current, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    days.push([key, counts[key] || 0]);
    current.setDate(current.getDate() + 1);
  }

  const data = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Date')
    .addColumn(Charts.ColumnType.NUMBER, 'Infractions');
  days.forEach(row => data.addRow(row));
  const chartImage = Charts.newLineChart()
    .setTitle('Infractions Over Time')
    .setXAxisTitle('Date')
    .setYAxisTitle('Infractions')
    .setDataTable(data)
    .setDimensions(800, 300)
    .build()
    .getAs('image/png');

  return { success: true, image: Utilities.base64Encode(chartImage.getBytes()) };
}

function generateThresholdDistributionChart() {
  const employees = getActiveEmployees();
  const infractions = getAllActiveInfractions();
  const pointsByEmployee = {};
  infractions.forEach(inf => {
    if (!inf.employee_id) return;
    pointsByEmployee[inf.employee_id] = (pointsByEmployee[inf.employee_id] || 0) + (inf.points || 0);
  });

  const thresholds = {
    employees_at_0_2_points: 0,
    employees_at_3_5_points: 0,
    employees_at_6_8_points: 0,
    employees_at_9_plus_points: 0
  };

  employees.forEach(emp => {
    const points = pointsByEmployee[emp.employee_id] || 0;
    if (points >= 9) {
      thresholds.employees_at_9_plus_points++;
    } else if (points >= 6) {
      thresholds.employees_at_6_8_points++;
    } else if (points >= 3) {
      thresholds.employees_at_3_5_points++;
    } else {
      thresholds.employees_at_0_2_points++;
    }
  });
  const data = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Range')
    .addColumn(Charts.ColumnType.NUMBER, 'Employees');
  data.addRow(['0-2', thresholds.employees_at_0_2_points || 0]);
  data.addRow(['3-5', thresholds.employees_at_3_5_points || 0]);
  data.addRow(['6-8', thresholds.employees_at_6_8_points || 0]);
  data.addRow(['9+', thresholds.employees_at_9_plus_points || 0]);

  const chartImage = Charts.newColumnChart()
    .setTitle('Threshold Distribution')
    .setDataTable(data)
    .setDimensions(800, 300)
    .build()
    .getAs('image/png');

  return { success: true, image: Utilities.base64Encode(chartImage.getBytes()) };
}

function generateInfractionTypeChart(dateRange) {
  const range = normalizeDateRange(dateRange);
  if (!range) return { success: false, message: 'Invalid date range' };

  const infractions = getInfractionsForRange(range.start, range.end);
  const metrics = buildReportMetrics(infractions);
  const data = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Type')
    .addColumn(Charts.ColumnType.NUMBER, 'Count');
  Object.keys(metrics.byType).forEach(type => data.addRow([type, metrics.byType[type]]));

  const chartImage = Charts.newPieChart()
    .setTitle('Infraction Types')
    .setDataTable(data)
    .setDimensions(600, 300)
    .build()
    .getAs('image/png');

  return { success: true, image: Utilities.base64Encode(chartImage.getBytes()) };
}

function buildReportHtml(report, charts) {
  const chartHtml = (charts || []).map(chart => {
    if (!chart || !chart.image) return '';
    return `<img src="data:image/png;base64,${chart.image}" style="max-width:100%; margin:10px 0;" />`;
  }).join('');

  const sectionHtml = (report.sections || []).map(section => `
    <h2>${section.title}</h2>
    <ul>${(section.rows || []).map(row => `<li>${row}</li>`).join('')}</ul>
  `).join('');

  const detailRows = (report.details || []).map(row => `
    <tr>
      <td>${row.date}</td>
      <td>${row.employee_name}</td>
      <td>${row.employee_id}</td>
      <td>${row.location}</td>
      <td>${row.infraction_type}</td>
      <td>${row.points}</td>
      <td>${row.entered_by}</td>
    </tr>
  `).join('');

  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; color: #333; }
          h1 { color: #E51636; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #ddd; padding: 6px; font-size: 12px; }
          th { background: #f7f7f7; }
        </style>
      </head>
      <body>
        <h1>${report.title}</h1>
        <p>Date Range: ${report.dateRange}</p>
        <h2>Summary</h2>
        <ul>
          <li>Total Infractions: ${report.summary.total_infractions}</li>
          <li>Total Points: ${report.summary.total_points}</li>
        </ul>
        ${sectionHtml}
        ${chartHtml}
        ${report.details && report.details.length ? `
          <h2>Details</h2>
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Employee</th>
                <th>ID</th>
                <th>Location</th>
                <th>Infraction</th>
                <th>Points</th>
                <th>Manager</th>
              </tr>
            </thead>
            <tbody>${detailRows}</tbody>
          </table>
        ` : ''}
      </body>
    </html>
  `;
}

function generateDetailedReportCore(reportConfig, generatedBy) {
  if (!reportConfig || !reportConfig.report_type) {
    return { success: false, message: 'report_type is required' };
  }

  const dateRange = normalizeDateRange(reportConfig.date_range);
  if (!dateRange) {
    return { success: false, message: 'Invalid date range' };
  }

  const infractions = applyReportFilters(
    getInfractionsForRange(dateRange.start, dateRange.end),
    reportConfig.filters || {}
  );
  const metrics = buildReportMetrics(infractions);
  const reportType = reportConfig.report_type;
  const employees = getActiveEmployees();
  const employeePoints = {};
  infractions.forEach(inf => {
    employeePoints[inf.employee_id] = (employeePoints[inf.employee_id] || 0) + (inf.points || 0);
  });

  const report = {
    title: reportType,
    dateRange: `${Utilities.formatDate(dateRange.start, Session.getScriptTimeZone(), 'MM/dd/yyyy')} - ${Utilities.formatDate(dateRange.end, Session.getScriptTimeZone(), 'MM/dd/yyyy')}`,
    summary: metrics.totals,
    details: infractions.map(inf => ({
      date: Utilities.formatDate(inf.date, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      employee_name: inf.full_name,
      employee_id: inf.employee_id,
      location: inf.location,
      infraction_type: inf.infraction_type,
      points: inf.points,
      entered_by: inf.entered_by
    })),
    sections: []
  };

  if (reportType === 'Executive Summary') {
    report.sections.push({
      title: 'Top Infractions',
      rows: Object.entries(metrics.byType).slice(0, 10).map(([type, count]) => `${type}: ${count}`)
    });
    report.sections.push({
      title: 'At-Risk Employees',
      rows: employees.filter(emp => (employeePoints[emp.employee_id] || 0) >= 6)
        .map(emp => `${emp.full_name} (${employeePoints[emp.employee_id] || 0} pts)`)
    });
  }

  if (reportType === 'Manager Performance') {
    const totalDays = Math.max(1, Math.ceil((dateRange.end - dateRange.start) / (1000 * 60 * 60 * 24)));
    const managerRows = Object.entries(metrics.byManager).map(([manager, count]) => ({
      manager,
      infractions: count,
      per_day: (count / totalDays).toFixed(2)
    }));
    report.sections.push({
      title: 'Manager Activity',
      rows: managerRows.map(row => `${row.manager}: ${row.infractions} infractions (${row.per_day}/day)`)
    });
  }

  if (reportType === 'Location Comparison') {
    report.sections.push({
      title: 'Location Breakdown',
      rows: Object.entries(metrics.byLocation).map(([loc, count]) => `${loc}: ${count} infractions`)
    });
  }

  if (reportType === 'Trend Analysis') {
    const trend = generateInfractionTrendChart(reportConfig.date_range);
    if (trend.success) {
      report.sections.push({ title: 'Trend Chart', rows: ['Trend chart included in report.'] });
    }
  }

  if (reportType === 'Employee Risk Assessment') {
    report.sections.push({
      title: 'High Risk Employees',
      rows: employees.filter(emp => (employeePoints[emp.employee_id] || 0) >= 9)
        .map(emp => `${emp.full_name} (${employeePoints[emp.employee_id] || 0} pts)`)
    });
  }

  if (reportType === 'Positive Behavior Report') {
    const credits = infractions.filter(inf => inf.points < 0 || inf.infraction_type === 'Positive Behavior Credit');
    report.sections.push({
      title: 'Credits Awarded',
      rows: credits.map(inf => `${inf.full_name}: ${inf.points} (${inf.date.toLocaleDateString()})`)
    });
  }

  let charts = [];
  if (reportConfig.include_charts) {
    const trend = generateInfractionTrendChart(reportConfig.date_range);
    const types = generateInfractionTypeChart(reportConfig.date_range);
    const distribution = generateThresholdDistributionChart();
    charts = [trend, types, distribution].filter(c => c && c.success);
  }

  const reportId = generateReportId();
  const format = reportConfig.format || 'pdf';
  const folder = getReportsFolder();
  const filename = `${reportConfig.report_type.replace(/\s+/g, '_')}_${reportId}.${format}`;

  let fileUrl = '';
  let fileSize = 0;

  if (format === 'pdf') {
    const html = buildReportHtml(report, charts);
    const pdfResult = convertHtmlToPDF(html, filename);
    if (!pdfResult.success) {
      return { success: false, message: pdfResult.error || 'Failed to build PDF' };
    }
    const saved = folder.createFile(pdfResult.blob);
    fileUrl = saved.getUrl();
    fileSize = saved.getSize();
  } else if (format === 'xlsx') {
    const temp = SpreadsheetApp.create(`Report_${reportId}`);
    const summarySheet = temp.getSheets()[0];
    summarySheet.setName('Summary');
    summarySheet.getRange(1, 1, 2, 2).setValues([
      ['Metric', 'Value'],
      ['Total Infractions', metrics.totals.total_infractions]
    ]);

    const detailSheet = temp.insertSheet('Details');
    const header = ['Date', 'Employee', 'Employee ID', 'Location', 'Infraction', 'Points', 'Manager'];
    detailSheet.getRange(1, 1, 1, header.length).setValues([header]);
    if (report.details.length) {
      const rows = report.details.map(row => [
        row.date, row.employee_name, row.employee_id, row.location, row.infraction_type, row.points, row.entered_by
      ]);
      detailSheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    }

    const chartsSheet = temp.insertSheet('Charts');
    charts.forEach((chart, idx) => {
      const blob = Utilities.newBlob(Utilities.base64Decode(chart.image), 'image/png', `chart_${idx}.png`);
      chartsSheet.insertImage(blob, 1, idx * 25 + 1);
    });

    const methodology = temp.insertSheet('Methodology');
    methodology.getRange(1, 1, 3, 1).setValues([
      ['Report generated by CFA Accountability System.'],
      ['Filters applied from report configuration.'],
      ['Charts generated via Apps Script Charts service.']
    ]);

    const blob = DriveApp.getFileById(temp.getId()).getAs(MimeType.MICROSOFT_EXCEL);
    const saved = folder.createFile(blob).setName(filename);
    fileUrl = saved.getUrl();
    fileSize = saved.getSize();
    DriveApp.getFileById(temp.getId()).setTrashed(true);
  } else if (format === 'csv') {
    const csvHeader = ['Date', 'Employee', 'Employee ID', 'Location', 'Infraction', 'Points', 'Manager'];
    const csvRows = report.details.map(row => [
      row.date, row.employee_name, row.employee_id, row.location, row.infraction_type, row.points, row.entered_by
    ]);
    const csv = [csvHeader].concat(csvRows).map(row => row.map(cell => `"${String(cell || '').replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = Utilities.newBlob(csv, 'text/csv', filename);
    const saved = folder.createFile(blob);
    fileUrl = saved.getUrl();
    fileSize = saved.getSize();
  } else {
    return { success: false, message: 'Unsupported format' };
  }

  const history = getOrCreateReportHistorySheet();
  history.appendRow([
    reportId,
    reportConfig.report_type,
    new Date(),
    generatedBy || 'System',
    report.dateRange,
    fileUrl,
    fileSize,
    format
  ]);

  return {
    success: true,
    report_id: reportId,
    file_url: fileUrl,
    file_size: fileSize,
    format: format
  };
}

function generateDetailedReport(reportConfig, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }
  return generateDetailedReportCore(reportConfig, auth.session.user_name || auth.session.role);
}

function getReportHistory(token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const sheet = getOrCreateReportHistorySheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, history: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const history = data.map(row => ({
    report_id: row[0],
    report_type: row[1],
    generated_date: row[2] instanceof Date ? row[2].toISOString() : row[2],
    generated_by: row[3],
    date_range_covered: row[4],
    file_url: row[5],
    file_size: row[6],
    format: row[7]
  }));
  return { success: true, history: history };
}

function exportDataDump(dumpConfig, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const include = dumpConfig?.include || ['Infractions', 'Settings', 'Email_Log', 'Edit_Log', 'User_Permissions'];
  const format = dumpConfig?.format || 'xlsx';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const folder = getReportsFolder();

  if (format === 'xlsx') {
    const temp = SpreadsheetApp.create(`DataDump_${generateReportId()}`);
    include.forEach(name => {
      const source = ss.getSheetByName(name);
      if (!source) return;
      const target = temp.insertSheet(name);
      const data = source.getDataRange().getValues();
      target.getRange(1, 1, data.length, data[0].length).setValues(data);
    });
    const blob = DriveApp.getFileById(temp.getId()).getAs(MimeType.MICROSOFT_EXCEL);
    const saved = folder.createFile(blob).setName('DataDump.xlsx');
    DriveApp.getFileById(temp.getId()).setTrashed(true);
    return { success: true, file_url: saved.getUrl() };
  }

  if (format === 'csv') {
    const blobs = [];
    include.forEach(name => {
      const source = ss.getSheetByName(name);
      if (!source) return;
      const data = source.getDataRange().getValues();
      const csv = data.map(row => row.map(cell => `"${String(cell || '').replace(/"/g, '""')}"`).join(',')).join('\n');
      blobs.push(Utilities.newBlob(csv, 'text/csv', `${name}.csv`));
    });
    const zip = Utilities.zip(blobs, 'DataDump.zip');
    const saved = folder.createFile(zip);
    return { success: true, file_url: saved.getUrl() };
  }

  return { success: false, message: 'Unsupported format' };
}

function scheduleRecurringReport(scheduleConfig, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) {
    return { success: false, sessionExpired: auth.sessionExpired || false, message: auth.error };
  }

  const frequency = String(scheduleConfig?.frequency || '').toLowerCase();
  if (!scheduleConfig?.report_type || !frequency || !scheduleConfig?.time) {
    return { success: false, message: 'Missing required schedule fields' };
  }

  scheduleConfig.frequency = frequency;

  try {
    const sheet = getOrCreateScheduledReportsSheet();
    const scheduleId = `SCH-${new Date().getTime()}`;
    const triggerBuilder = ScriptApp.newTrigger('runScheduledReport');
    const timeParts = scheduleConfig.time.split(':');
    const hour = Number(timeParts[0] || 8);

    let triggerId = '';
    let status = 'Active';
    let warning = '';

    try {
      let trigger;
      const dayOfWeek = scheduleConfig.day_of_week
        ? ScriptApp.WeekDay[scheduleConfig.day_of_week.toUpperCase()] || ScriptApp.WeekDay.MONDAY
        : ScriptApp.WeekDay.MONDAY;
      if (scheduleConfig.frequency === 'daily') {
        trigger = triggerBuilder.timeBased().everyDays(1).atHour(hour).create();
      } else if (scheduleConfig.frequency === 'weekly') {
        trigger = triggerBuilder.timeBased()
          .everyWeeks(1)
          .onWeekDay(dayOfWeek)
          .atHour(hour)
          .create();
      } else if (scheduleConfig.frequency === 'monthly') {
        trigger = triggerBuilder.timeBased()
          .everyMonths(1)
          .onMonthDay(scheduleConfig.day_of_month || 1)
          .atHour(hour)
          .create();
      } else if (scheduleConfig.frequency === 'quarterly') {
        trigger = triggerBuilder.timeBased()
          .everyMonths(3)
          .onMonthDay(scheduleConfig.day_of_month || 1)
          .atHour(hour)
          .create();
      } else {
        return { success: false, message: 'Invalid frequency' };
      }

      triggerId = trigger.getUniqueId();
      PropertiesService.getScriptProperties().setProperty(`report_trigger_${triggerId}`, scheduleId);
    } catch (error) {
      status = 'TriggerError';
      warning = 'Scheduled report saved, but trigger could not be created. Run authorizeScheduling() and then edit/re-save this schedule.';
    }

    const nextRun = computeNextRun(scheduleConfig);
    sheet.appendRow([
      scheduleId,
      scheduleConfig.report_name || scheduleConfig.report_type,
      scheduleConfig.report_type,
      scheduleConfig.frequency,
      scheduleConfig.day_of_week || '',
      scheduleConfig.day_of_month || '',
      scheduleConfig.time,
      (scheduleConfig.recipients || []).join(','),
      JSON.stringify(scheduleConfig.report_config || {}),
      status,
      nextRun,
      triggerId
    ]);

    const result = { success: true, schedule_id: scheduleId };
    if (warning) result.warning = warning;
    return result;
  } catch (error) {
    const rawMessage = error && error.message ? error.message : String(error);
    const needsAuth = rawMessage.toLowerCase().indexOf('authorization') !== -1;
    const message = needsAuth
      ? 'Scheduling requires authorization. Please run authorizeScheduling() once in the Apps Script editor, then retry.'
      : rawMessage;
    return { success: false, message: message };
  }
}

function authorizeScheduling() {
  let trigger;
  try {
    trigger = ScriptApp.newTrigger('runScheduledReport')
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();
  } catch (error) {
    return {
      success: false,
      message: 'Failed to create trigger for authorization: ' + (error && error.message ? error.message : error)
    };
  }

  try {
    ScriptApp.deleteTrigger(trigger);
  } catch (error) {
    // Trigger creation is enough to grant permissions; deletion failure is non-fatal.
    return {
      success: true,
      warning: 'Authorization granted, but cleanup trigger delete failed: ' + (error && error.message ? error.message : error)
    };
  }

  return { success: true };
}

function computeNextRun(scheduleConfig) {
  const now = new Date();
  const timeParts = scheduleConfig.time.split(':');
  const hour = Number(timeParts[0] || 8);
  const minute = Number(timeParts[1] || 0);
  let next = new Date(now);
  next.setHours(hour, minute, 0, 0);
  if (scheduleConfig.frequency === 'daily') {
    if (next <= now) next.setDate(next.getDate() + 1);
  } else if (scheduleConfig.frequency === 'weekly') {
    const day = scheduleConfig.day_of_week || 'MONDAY';
    const dayMap = { SUNDAY: 0, MONDAY: 1, TUESDAY: 2, WEDNESDAY: 3, THURSDAY: 4, FRIDAY: 5, SATURDAY: 6 };
    const target = dayMap[day.toUpperCase()] ?? 1;
    const diff = (target - next.getDay() + 7) % 7;
    if (diff === 0 && next <= now) next.setDate(next.getDate() + 7);
    else next.setDate(next.getDate() + diff);
  } else if (scheduleConfig.frequency === 'monthly' || scheduleConfig.frequency === 'quarterly') {
    const dayOfMonth = Number(scheduleConfig.day_of_month || 1);
    next.setDate(dayOfMonth);
    if (next <= now) {
      next.setMonth(next.getMonth() + (scheduleConfig.frequency === 'quarterly' ? 3 : 1));
    }
  }
  return next;
}

function getScheduledReports(token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) return { success: false, message: auth.error };

  const sheet = getOrCreateScheduledReportsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, schedules: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  const schedules = data.map(row => ({
    schedule_id: row[0],
    report_name: row[1],
    report_type: row[2],
    frequency: row[3],
    day_of_week: row[4],
    day_of_month: row[5],
    time: row[6],
    recipients: row[7],
    status: row[9],
    next_run: row[10],
    trigger_id: row[11]
  }));
  return { success: true, schedules: schedules };
}

function updateScheduledReportStatus(scheduleId, status, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) return { success: false, message: auth.error };
  const sheet = getOrCreateScheduledReportsSheet();
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[0] === scheduleId);
  if (rowIndex < 1) return { success: false, message: 'Schedule not found' };
  sheet.getRange(rowIndex + 1, 10).setValue(status);
  return { success: true };
}

function deleteScheduledReport(scheduleId, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) return { success: false, message: auth.error };
  const sheet = getOrCreateScheduledReportsSheet();
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[0] === scheduleId);
  if (rowIndex < 1) return { success: false, message: 'Schedule not found' };

  const triggerId = data[rowIndex][11];
  if (triggerId) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }

  sheet.deleteRow(rowIndex + 1);
  return { success: true };
}

function runScheduledReportNow(scheduleId, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) return { success: false, message: auth.error };
  const sheet = getOrCreateScheduledReportsSheet();
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[0] === scheduleId);
  if (rowIndex < 1) return { success: false, message: 'Schedule not found' };

  const row = data[rowIndex];
  const recipients = row[7] ? row[7].split(',').map(r => r.trim()).filter(Boolean) : [];
  const reportConfig = JSON.parse(row[8] || '{}');
  reportConfig.report_type = row[2] || reportConfig.report_type;
  reportConfig.format = reportConfig.format || 'pdf';
  reportConfig.include_charts = reportConfig.include_charts !== false;

  const result = generateDetailedReportCore(reportConfig, auth.session.user_name || auth.session.role);
  if (result.success && recipients.length) {
    GmailApp.sendEmail(
      recipients.join(','),
      `Scheduled Report: ${row[1]}`,
      `Your scheduled report is ready: ${result.file_url}`
    );
  }
  return result;
}

function runScheduledReport(e) {
  const triggerId = e?.triggerUid;
  if (!triggerId) return;

  const scheduleId = PropertiesService.getScriptProperties().getProperty(`report_trigger_${triggerId}`);
  if (!scheduleId) return;

  const sheet = getOrCreateScheduledReportsSheet();
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, index) => index > 0 && row[0] === scheduleId);
  if (rowIndex < 1) return;

  const row = data[rowIndex];
  const recipients = row[7] ? row[7].split(',').map(r => r.trim()).filter(Boolean) : [];
  const reportConfig = JSON.parse(row[8] || '{}');
  reportConfig.report_type = row[2] || reportConfig.report_type;
  reportConfig.format = reportConfig.format || 'pdf';
  reportConfig.include_charts = reportConfig.include_charts !== false;

  const result = generateDetailedReportCore(reportConfig, 'Scheduled Report');
  if (result.success && recipients.length) {
    GmailApp.sendEmail(
      recipients.join(','),
      `Scheduled Report: ${row[1]}`,
      `Your scheduled report is ready: ${result.file_url}`
    );
  }
}

function saveReportTemplate(templateName, reportConfig, token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) return { success: false, message: auth.error };

  const sheet = getOrCreateReportTemplatesSheet();
  const templateId = `TPL-${new Date().getTime()}`;
  sheet.appendRow([
    templateId,
    templateName,
    auth.session.user_name || auth.session.role,
    new Date(),
    JSON.stringify(reportConfig || {})
  ]);
  return { success: true, template_id: templateId };
}

function getReportTemplates(token) {
  const auth = requireDirectorSession(token);
  if (!auth.ok) return { success: false, message: auth.error };

  const sheet = getOrCreateReportTemplatesSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, templates: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const templates = data.map(row => ({
    template_id: row[0],
    template_name: row[1],
    created_by: row[2],
    created_at: row[3] instanceof Date ? row[3].toISOString() : row[3],
    report_config: row[4] ? JSON.parse(row[4]) : {}
  }));
  return { success: true, templates: templates };
}

function testAdvancedReporting(token) {
  const results = [];

  const reportConfig = {
    report_type: 'Executive Summary',
    date_range: {
      start_date: new Date(new Date().getTime() - 1000 * 60 * 60 * 24 * 30),
      end_date: new Date()
    },
    filters: {
      locations: [],
      infraction_types: [],
      managers: []
    },
    format: 'pdf',
    include_charts: true
  };

  const report1 = generateDetailedReport(reportConfig, token);
  results.push({ test: 'Executive Summary', passed: report1.success });

  const report2 = generateDetailedReport({
    report_type: 'Manager Performance',
    date_range: reportConfig.date_range,
    filters: {},
    format: 'xlsx',
    include_charts: true
  }, token);
  results.push({ test: 'Manager Performance', passed: report2.success });

  const report3 = generateDetailedReport({
    report_type: 'Custom Filtered',
    date_range: reportConfig.date_range,
    filters: { locations: ['Cockrell Hill DTO'], point_range: { min: 6, max: 99 } },
    format: 'csv',
    include_charts: false
  }, token);
  results.push({ test: 'Custom Filtered', passed: report3.success });

  const history = getReportHistory(token);
  results.push({ test: 'Report History', passed: history.success });

  const dataDump = exportDataDump({ include: ['Infractions', 'Settings'], format: 'xlsx' }, token);
  results.push({ test: 'Data Dump', passed: dataDump.success });

  return { success: results.every(r => r.passed), results: results };
}
