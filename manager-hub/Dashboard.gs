// ============================================
// PHASE 20: DASHBOARD STATISTICS AND REPORTING
// ============================================
// Functions for aggregated views, statistics,
// and reporting for directors
// ============================================

/**
 * Safely serialize data for client transport.
 * Ensures no Date/undefined/function values leak into JSON responses.
 */
function safeSerialize(value) {
  if (value === undefined) return null;
  if (value === null) return null;
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(safeSerialize);
  if (typeof value === 'function') return null;
  if (typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(key => {
      out[key] = safeSerialize(value[key]);
    });
    return out;
  }
  return value;
}

function getBreakFoodThreshold_() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) return 9;
    const val = Number(sheet.getRange('B37').getValue());
    return val && !isNaN(val) ? val : 9;
  } catch (error) {
    console.error('getBreakFoodThreshold_ error:', error);
    return 9;
  }
}

function buildBreakFoodRemovalList_(employees, employeeInfractions, employeePoints, threshold) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const windowStart = new Date(today);
  windowStart.setDate(windowStart.getDate() - 7);

  const list = [];
  employees.forEach(emp => {
    const infractions = (employeeInfractions[emp.employee_id] || []).filter(inf => {
      if (!inf || !inf.date) return false;
      const exp = inf.expiration_date instanceof Date ? inf.expiration_date : new Date(inf.expiration_date);
      return exp >= today;
    });
    if (!infractions.length) return;

    infractions.sort((a, b) => a.date.getTime() - b.date.getTime());
    let running = 0;
    let crossingDate = null;
    infractions.forEach(inf => {
      const points = Number(inf.points) || 0;
      const before = running;
      running += points;
      if (before < threshold && running >= threshold) {
        crossingDate = inf.date;
      }
    });

    if (!crossingDate || crossingDate < windowStart) return;

    const removalEnd = new Date(crossingDate);
    removalEnd.setDate(removalEnd.getDate() + 7);
    const currentPoints = employeePoints[emp.employee_id] || 0;

    list.push({
      employee_id: emp.employee_id,
      employee_name: emp.full_name || '',
      current_points: currentPoints,
      crossed_on: crossingDate,
      removal_end: removalEnd
    });
  });

  list.sort((a, b) => a.removal_end.getTime() - b.removal_end.getTime());
  return list;
}

/**
 * Log server-side errors to a sheet for diagnostics.
 */
function logServerError(context, error, metadata) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Server_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('Server_Logs');
      sheet.getRange(1, 1, 1, 5).setValues([['timestamp', 'context', 'error', 'metadata', 'stack']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const metaStr = metadata ? JSON.stringify(safeSerialize(metadata)) : '';
    const stack = error && error.stack ? String(error.stack) : '';
    sheet.appendRow([new Date().toISOString(), context, String(error), metaStr, stack]);
  } catch (e) {
    console.error('logServerError failed:', e);
  }
}

/**
 * Get comprehensive dashboard statistics.
 * Returns all metrics needed for the dashboard display.
 *
 * @returns {Object} Dashboard statistics object
 */
function getDashboardStatistics(token) {
  const startTime = Date.now();

  try {
    // Validate session using token-based auth
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    // Get all data needed
    const employees = getActiveEmployees();
    const allInfractions = getAllActiveInfractions();
    const probationMap = buildProbationStatusMap();

    // Calculate employee points
    const employeePoints = {};
    const employeeInfractions = {};

    for (const inf of allInfractions) {
      if (!inf.employee_id) continue;

      // Track infractions per employee
      if (!employeeInfractions[inf.employee_id]) {
        employeeInfractions[inf.employee_id] = [];
      }
      employeeInfractions[inf.employee_id].push(inf);

      // Calculate active points
      if (inf.status === 'Active' && !inf.is_expired) {
        if (!employeePoints[inf.employee_id]) {
          employeePoints[inf.employee_id] = 0;
        }
        employeePoints[inf.employee_id] += (inf.points || 0);
      }
    }

    // ========================================
    // OVERALL METRICS
    // ========================================
    const totalActiveEmployees = employees.length;
    let employeesWithPoints = 0;
    let employeesWithZeroPoints = 0;
    let totalPoints = 0;

    for (const emp of employees) {
      const points = employeePoints[emp.employee_id] || 0;
      if (points > 0) {
        employeesWithPoints++;
        totalPoints += points;
      } else {
        employeesWithZeroPoints++;
      }
    }

    const averagePointsPerEmployee = totalActiveEmployees > 0
      ? Math.round((totalPoints / totalActiveEmployees) * 10) / 10
      : 0;

    // Count active infractions
    const activeInfractions = allInfractions.filter(inf =>
      inf.status === 'Active' && !inf.is_expired
    );
    const totalActiveInfractions = activeInfractions.length;

    // Infractions this month vs last month
    const now = new Date();
    const thisMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const lastMonthEnd = new Date(now.getFullYear(), now.getMonth(), 0);

    let infractionsThisMonth = 0;
    let infractionsLastMonth = 0;

    for (const inf of allInfractions) {
      const infDate = new Date(inf.date);
      if (infDate >= thisMonthStart) {
        infractionsThisMonth++;
      } else if (infDate >= lastMonthStart && infDate <= lastMonthEnd) {
        infractionsLastMonth++;
      }
    }

    // ========================================
    // THRESHOLD DISTRIBUTION
    // ========================================
    let employees_0_2 = 0;
    let employees_3_5 = 0;
    let employees_6_8 = 0;
    let employees_9_plus = 0;
    let employeesOnProbation = 0;
    let employeesAtTermination = 0;

    for (const emp of employees) {
      const points = employeePoints[emp.employee_id] || 0;

      if (points >= 15) {
        employeesAtTermination++;
        employees_9_plus++;
      } else if (points >= 9) {
        employees_9_plus++;
      } else if (points >= 6) {
        employees_6_8++;
      } else if (points >= 3) {
        employees_3_5++;
      } else {
        employees_0_2++;
      }

      if (probationMap[emp.employee_id]) {
        employeesOnProbation++;
      }
    }

    // ========================================
    // INFRACTION BREAKDOWN
    // ========================================
    // By type
    const byType = {};
    for (const inf of allInfractions) {
      const type = inf.infraction_type || 'Unknown';
      byType[type] = (byType[type] || 0) + 1;
    }

    // Sort by count descending
    const sortedByType = Object.entries(byType)
      .sort((a, b) => b[1] - a[1])
      .reduce((obj, [key, val]) => { obj[key] = val; return obj; }, {});

    // By location
    const byLocation = {
      'Cockrell Hill DTO': 0,
      'Dallas Baptist University OCV': 0,
      'Other': 0
    };
    for (const inf of allInfractions) {
      const loc = inf.location || '';
      if (loc.toLowerCase().includes('cockrell')) {
        byLocation['Cockrell Hill DTO']++;
      } else if (loc.toLowerCase().includes('dbu') || loc.toLowerCase().includes('baptist')) {
        byLocation['Dallas Baptist University OCV']++;
      } else {
        byLocation['Other']++;
      }
    }

    // By severity (points)
    const bySeverity = {
      'Minor (1pt)': 0,
      'Moderate (3pt)': 0,
      'Major (5pt)': 0,
      'Severe (8pt)': 0
    };
    for (const inf of allInfractions) {
      const pts = inf.points || 0;
      if (pts <= 1) {
        bySeverity['Minor (1pt)']++;
      } else if (pts <= 3) {
        bySeverity['Moderate (3pt)']++;
      } else if (pts <= 5) {
        bySeverity['Major (5pt)']++;
      } else {
        bySeverity['Severe (8pt)']++;
      }
    }

    // ========================================
    // TRENDS (Last 6 months)
    // ========================================
    const infractionsByMonth = [];
    for (let i = 5; i >= 0; i--) {
      const monthStart = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const monthEnd = new Date(now.getFullYear(), now.getMonth() - i + 1, 0);
      const monthName = monthStart.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });

      let count = 0;
      for (const inf of allInfractions) {
        const infDate = new Date(inf.date);
        if (infDate >= monthStart && infDate <= monthEnd) {
          count++;
        }
      }

      infractionsByMonth.push({ month: monthName, count: count });
    }

    // Determine trend
    let pointsTrend = 'stable';
    if (infractionsByMonth.length >= 2) {
      const recent = infractionsByMonth.slice(-2);
      if (recent[1].count > recent[0].count * 1.2) {
        pointsTrend = 'increasing';
      } else if (recent[1].count < recent[0].count * 0.8) {
        pointsTrend = 'decreasing';
      }
    }

    // Top 5 infraction types
    const topInfractionTypes = Object.entries(sortedByType).slice(0, 5);

    // ========================================
    // POSITIVE METRICS
    // ========================================
    const positiveCredits = allInfractions.filter(inf => inf.is_positive).length;
    const employeesWithNegativePoints = employees.filter(emp =>
      (employeePoints[emp.employee_id] || 0) < 0
    ).length;

    // Days since last infraction (average for employees with infractions)
    let totalDaysSinceLast = 0;
    let employeesWithInfractionCount = 0;
    const today = new Date();

    for (const emp of employees) {
      const empInfractions = employeeInfractions[emp.employee_id] || [];
      if (empInfractions.length > 0) {
        const sorted = empInfractions.sort((a, b) => new Date(b.date) - new Date(a.date));
        const lastInfDate = new Date(sorted[0].date);
        const daysSince = Math.floor((today - lastInfDate) / (1000 * 60 * 60 * 24));
        totalDaysSinceLast += daysSince;
        employeesWithInfractionCount++;
      }
    }

    const avgDaysSinceLastInfraction = employeesWithInfractionCount > 0
      ? Math.round(totalDaysSinceLast / employeesWithInfractionCount)
      : null;

    // Clean record in last 30 days
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    let cleanRecord30Days = 0;
    for (const emp of employees) {
      const empInfractions = employeeInfractions[emp.employee_id] || [];
      const recentInfractions = empInfractions.filter(inf => {
        const infDate = new Date(inf.date);
        return infDate >= thirtyDaysAgo && !inf.is_positive;
      });
      if (recentInfractions.length === 0) {
        cleanRecord30Days++;
      }
    }

    // ========================================
    // RECENT ACTIVITY
    // ========================================
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const weekStart = new Date(todayStart);
    weekStart.setDate(weekStart.getDate() - 7);

    let infractionsToday = 0;
    let infractionsThisWeek = 0;

    for (const inf of allInfractions) {
      const infDate = new Date(inf.date);
      if (infDate >= todayStart) {
        infractionsToday++;
      }
      if (infDate >= weekStart) {
        infractionsThisWeek++;
      }
    }

    // Recent threshold crossings (simulated from high-point employees)
    const recentThresholdCrossings = [];
    for (const emp of employees) {
      const points = employeePoints[emp.employee_id] || 0;
      if (points >= 6 && points < 9) {
        recentThresholdCrossings.push({
          employee_id: emp.employee_id,
          employee_name: emp.full_name,
          threshold: 6,
          current_points: points
        });
      } else if (points >= 9 && points < 15) {
        recentThresholdCrossings.push({
          employee_id: emp.employee_id,
          employee_name: emp.full_name,
          threshold: 9,
          current_points: points
        });
      } else if (points >= 15) {
        recentThresholdCrossings.push({
          employee_id: emp.employee_id,
          employee_name: emp.full_name,
          threshold: 15,
          current_points: points
        });
      }
    }

    // Recent probations
    const recentProbations = [];
    for (const emp of employees) {
      if (probationMap[emp.employee_id]) {
        recentProbations.push({
          employee_id: emp.employee_id,
          employee_name: emp.full_name
        });
      }
    }

    const breakFoodThreshold = getBreakFoodThreshold_();
    const breakFoodList = buildBreakFoodRemovalList_(employees, employeeInfractions, employeePoints, breakFoodThreshold);

    const executionTime = Date.now() - startTime;

    return safeSerialize({
      success: true,
      userRole: session.role,
      generatedAt: new Date().toISOString(),
      executionTime: executionTime,

      // Overall Metrics
      overall: {
        total_active_employees: totalActiveEmployees,
        employees_with_points: employeesWithPoints,
        employees_with_zero_points: employeesWithZeroPoints,
        average_points_per_employee: averagePointsPerEmployee,
        total_active_infractions: totalActiveInfractions,
        infractions_this_month: infractionsThisMonth,
        infractions_last_month: infractionsLastMonth
      },

      // Threshold Distribution
      thresholds: {
        employees_at_0_2_points: employees_0_2,
        employees_at_3_5_points: employees_3_5,
        employees_at_6_8_points: employees_6_8,
        employees_at_9_plus_points: employees_9_plus,
        employees_on_probation: employeesOnProbation,
        employees_at_termination_level: employeesAtTermination
      },

      // Break Food Removal List
      break_food: {
        threshold: breakFoodThreshold,
        list: breakFoodList
      },

      // Infraction Breakdown
      breakdown: {
        by_type: sortedByType,
        by_location: byLocation,
        by_severity: bySeverity
      },

      // Trends
      trends: {
        infractions_by_month: infractionsByMonth,
        points_trend: pointsTrend,
        top_infraction_types: topInfractionTypes
      },

      // Positive Metrics
      positive: {
        positive_credits_awarded: positiveCredits,
        employees_with_negative_points: employeesWithNegativePoints,
        average_days_since_last_infraction: avgDaysSinceLastInfraction,
        employees_with_clean_record_30_days: cleanRecord30Days
      },

      // Recent Activity
      recent: {
        infractions_added_today: infractionsToday,
        infractions_added_this_week: infractionsThisWeek,
        recent_threshold_crossings: recentThresholdCrossings.slice(0, 10),
        recent_probations_started: recentProbations.slice(0, 10)
      }
    });

  } catch (error) {
    console.error('Error getting dashboard statistics:', error);
    logServerError('getDashboardStatistics', error, { token_present: !!token });
    return safeSerialize({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * Get detailed manager accountability report.
 * Shows which managers are entering infractions and their activity levels.
 *
 * @returns {Object} Manager accountability data
 */
function getManagerAccountability(token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    if (session.role !== 'Director' && session.role !== 'Operator') {
      return { success: false, error: 'Access denied' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);

    // Get all infractions to analyze who entered them
    const infractionsSheet = ss.getSheetByName('Infractions');
    if (!infractionsSheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    const data = infractionsSheet.getDataRange().getValues();
    const headers = data[0];

    const enteredByCol = headers.indexOf('Entered_By');
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Infraction_Type');
    const locationCol = headers.indexOf('Location');
    const pointsCol = headers.indexOf('Points_Assigned') !== -1
      ? headers.indexOf('Points_Assigned')
      : headers.indexOf('Points');
    const pointValueCol = headers.indexOf('Point_Value_At_Time');

    // Get user permissions to know manager roles
    const permSheet = ss.getSheetByName('User_Permissions');
    const permData = permSheet ? permSheet.getDataRange().getValues() : [];
    const permHeaders = permData.length > 0 ? permData[0] : [];
    const permNameCol = permHeaders.indexOf('Full_Name');
    const permRoleCol = permHeaders.indexOf('Role');
    const permEmailCol = permHeaders.indexOf('Email');

    const managerRoles = {};
    for (let i = 1; i < permData.length; i++) {
      const email = permData[i][permEmailCol];
      const name = permData[i][permNameCol];
      const role = permData[i][permRoleCol];
      if (email) {
        managerRoles[email.toLowerCase()] = { name, role };
      }
    }

    // Analyze entries by manager
    const managerStats = {};
    const today = new Date();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const enteredBy = row[enteredByCol];
      const date = row[dateCol];
      const type = row[typeCol];
      const location = row[locationCol];
      const points = row[pointsCol] || 0;

      if (!enteredBy) continue;

      const key = enteredBy.toLowerCase();

      if (!managerStats[key]) {
        const managerInfo = managerRoles[key] || { name: enteredBy, role: 'Unknown' };
        managerStats[key] = {
          email: enteredBy,
          manager_name: managerInfo.name || enteredBy,
          role: managerInfo.role || 'Unknown',
          total_infractions_entered: 0,
          first_entry_date: null,
          last_entry_date: null,
          infraction_types: {},
          locations: {},
          points_assigned: 0,
          entries_by_month: {}
        };
      }

      const stats = managerStats[key];
      stats.total_infractions_entered++;
      stats.points_assigned += points;

      const entryDate = new Date(date);
      if (!stats.first_entry_date || entryDate < new Date(stats.first_entry_date)) {
        stats.first_entry_date = entryDate.toISOString();
      }
      if (!stats.last_entry_date || entryDate > new Date(stats.last_entry_date)) {
        stats.last_entry_date = entryDate.toISOString();
      }

      // Track by type
      stats.infraction_types[type] = (stats.infraction_types[type] || 0) + 1;

      // Track by location
      stats.locations[location] = (stats.locations[location] || 0) + 1;

      // Track by month
      const monthKey = entryDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
      stats.entries_by_month[monthKey] = (stats.entries_by_month[monthKey] || 0) + 1;
    }

    // Process and calculate additional metrics
    const managers = [];
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    for (const [email, stats] of Object.entries(managerStats)) {
      // Calculate weeks active
      const firstDate = stats.first_entry_date ? new Date(stats.first_entry_date) : today;
      const lastDate = stats.last_entry_date ? new Date(stats.last_entry_date) : today;
      const weeksActive = Math.max(1, Math.ceil((lastDate - firstDate) / (1000 * 60 * 60 * 24 * 7)));
      const avgPerWeek = Math.round((stats.total_infractions_entered / weeksActive) * 10) / 10;

      // Find most common type
      const sortedTypes = Object.entries(stats.infraction_types).sort((a, b) => b[1] - a[1]);
      const mostCommonType = sortedTypes.length > 0 ? sortedTypes[0][0] : 'None';

      // Days since last entry
      const daysSinceLastEntry = stats.last_entry_date
        ? Math.floor((today - new Date(stats.last_entry_date)) / (1000 * 60 * 60 * 24))
        : null;

      // Activity status
      let activityStatus = 'Active';
      let flags = [];

      if (daysSinceLastEntry !== null && daysSinceLastEntry > 30) {
        activityStatus = 'Inactive';
        flags.push('No entries in 30+ days');
      }

      // Check if only entering minor infractions
      const minorCount = Object.entries(stats.infraction_types)
        .filter(([type]) => type.toLowerCase().includes('tardy') || type.toLowerCase().includes('dress'))
        .reduce((sum, [, count]) => sum + count, 0);

      if (stats.total_infractions_entered > 5 && minorCount / stats.total_infractions_entered > 0.8) {
        flags.push('Mostly minor infractions');
      }

      managers.push({
        email: email,
        manager_name: stats.manager_name,
        role: stats.role,
        total_infractions_entered: stats.total_infractions_entered,
        date_range: {
          first: stats.first_entry_date, // Already ISO string
          last: stats.last_entry_date // Already ISO string
        },
        average_per_week: avgPerWeek,
        most_common_infraction_type: mostCommonType,
        location_breakdown: stats.locations,
        last_activity_date: stats.last_entry_date, // Already ISO string
        days_since_last_entry: daysSinceLastEntry,
        activity_status: activityStatus,
        flags: flags
      });
    }

    // Sort by total entries descending
    managers.sort((a, b) => b.total_infractions_entered - a.total_infractions_entered);

    return safeSerialize({
      success: true,
      managers: managers,
      total_managers: managers.length,
      active_managers: managers.filter(m => m.activity_status === 'Active').length,
      inactive_managers: managers.filter(m => m.activity_status === 'Inactive').length
    });

  } catch (error) {
    console.error('Error getting manager accountability:', error);
    logServerError('getManagerAccountability', error, { token_present: !!token });
    return safeSerialize({ success: false, error: error.toString() });
  }
}

/**
 * Get location comparison metrics.
 * Compares Cockrell Hill vs DBU statistics.
 *
 * @returns {Object} Location comparison data
 */
function getLocationComparison(token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    const employees = getActiveEmployees();
    const allInfractions = getAllActiveInfractions();

    // Initialize location data
    const locations = {
      'Cockrell Hill DTO': {
        total_employees: 0,
        employees_with_points: 0,
        total_points: 0,
        total_infractions: 0,
        infraction_types: {},
        threshold_distribution: { '0-2': 0, '3-5': 0, '6-8': 0, '9+': 0 }
      },
      'Dallas Baptist University OCV': {
        total_employees: 0,
        employees_with_points: 0,
        total_points: 0,
        total_infractions: 0,
        infraction_types: {},
        threshold_distribution: { '0-2': 0, '3-5': 0, '6-8': 0, '9+': 0 }
      }
    };

    // Calculate points per employee
    const employeePoints = {};
    for (const inf of allInfractions) {
      if (inf.status === 'Active' && !inf.is_expired && !inf.is_positive) {
        employeePoints[inf.employee_id] = (employeePoints[inf.employee_id] || 0) + (inf.points || 0);
      }
    }

    // Process employees by location
    for (const emp of employees) {
      const loc = emp.primary_location || '';
      let locationKey = null;

      if (loc.toLowerCase().includes('cockrell')) {
        locationKey = 'Cockrell Hill DTO';
      } else if (loc.toLowerCase().includes('dbu') || loc.toLowerCase().includes('baptist')) {
        locationKey = 'Dallas Baptist University OCV';
      }

      if (locationKey && locations[locationKey]) {
        const locData = locations[locationKey];
        locData.total_employees++;

        const points = employeePoints[emp.employee_id] || 0;
        if (points > 0) {
          locData.employees_with_points++;
          locData.total_points += points;
        }

        // Threshold distribution
        if (points >= 9) {
          locData.threshold_distribution['9+']++;
        } else if (points >= 6) {
          locData.threshold_distribution['6-8']++;
        } else if (points >= 3) {
          locData.threshold_distribution['3-5']++;
        } else {
          locData.threshold_distribution['0-2']++;
        }
      }
    }

    // Process infractions by location
    for (const inf of allInfractions) {
      const loc = inf.location || '';
      let locationKey = null;

      if (loc.toLowerCase().includes('cockrell')) {
        locationKey = 'Cockrell Hill DTO';
      } else if (loc.toLowerCase().includes('dbu') || loc.toLowerCase().includes('baptist')) {
        locationKey = 'Dallas Baptist University OCV';
      }

      if (locationKey && locations[locationKey]) {
        const locData = locations[locationKey];
        locData.total_infractions++;

        const type = inf.infraction_type || 'Unknown';
        locData.infraction_types[type] = (locData.infraction_types[type] || 0) + 1;
      }
    }

    // Calculate derived metrics
    const result = {};
    for (const [locName, data] of Object.entries(locations)) {
      const avgPoints = data.total_employees > 0
        ? Math.round((data.total_points / data.total_employees) * 10) / 10
        : 0;

      const infractionsPerEmployee = data.total_employees > 0
        ? Math.round((data.total_infractions / data.total_employees) * 10) / 10
        : 0;

      // Sort infraction types
      const sortedTypes = Object.entries(data.infraction_types)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

      result[locName] = {
        total_employees: data.total_employees,
        employees_with_points: data.employees_with_points,
        average_points_per_employee: avgPoints,
        total_infractions: data.total_infractions,
        infractions_per_employee_ratio: infractionsPerEmployee,
        most_common_infraction_types: sortedTypes,
        threshold_distribution: data.threshold_distribution
      };
    }

    return safeSerialize({
      success: true,
      locations: result
    });

  } catch (error) {
    console.error('Error getting location comparison:', error);
    logServerError('getLocationComparison', error, { token_present: !!token });
    return safeSerialize({ success: false, error: error.toString() });
  }
}

/**
 * Get employee risk report.
 * Identifies high-risk employees who need attention.
 *
 * @returns {Object} Risk report data
 */
function getEmployeeRiskReport(token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    const employees = getActiveEmployees();
    const allInfractions = getAllActiveInfractions();
    const probationMap = buildProbationStatusMap();

    // Build employee infraction data
    const employeeData = {};
    for (const emp of employees) {
      employeeData[emp.employee_id] = {
        employee_id: emp.employee_id,
        full_name: emp.full_name,
        location: emp.primary_location,
        infractions: [],
        current_points: 0,
        active_infractions: []
      };
    }

    for (const inf of allInfractions) {
      if (!employeeData[inf.employee_id]) continue;

      employeeData[inf.employee_id].infractions.push(inf);

      if (inf.status === 'Active' && !inf.is_expired && !inf.is_positive) {
        employeeData[inf.employee_id].current_points += (inf.points || 0);
        employeeData[inf.employee_id].active_infractions.push(inf);
      }
    }

    const riskyEmployees = [];
    const today = new Date();
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    for (const [empId, data] of Object.entries(employeeData)) {
      const riskFactors = [];
      let riskScore = 0;

      // Check: Close to threshold (within 2 points of 6, 9, or 15)
      if (data.current_points >= 4 && data.current_points < 6) {
        riskFactors.push('Within 2 points of 6-point threshold');
        riskScore += 2;
      } else if (data.current_points >= 7 && data.current_points < 9) {
        riskFactors.push('Within 2 points of 9-point threshold');
        riskScore += 3;
      } else if (data.current_points >= 13 && data.current_points < 15) {
        riskFactors.push('Within 2 points of termination level');
        riskScore += 5;
      }

      // Check: At or above 9 points
      if (data.current_points >= 9) {
        riskFactors.push('At or above final warning threshold (9+ points)');
        riskScore += 4;
      }

      // Check: At termination level
      if (data.current_points >= 15) {
        riskFactors.push('At termination level (15+ points)');
        riskScore += 6;
      }

      // Check: Rapid accumulation (3+ infractions in 30 days)
      const recentInfractions = data.infractions.filter(inf => {
        const infDate = new Date(inf.date);
        return infDate >= thirtyDaysAgo && !inf.is_positive;
      });
      if (recentInfractions.length >= 3) {
        riskFactors.push(`Rapid accumulation: ${recentInfractions.length} infractions in last 30 days`);
        riskScore += 3;
      }

      // Check: Pattern of same type
      const typeCounts = {};
      for (const inf of data.infractions) {
        if (!inf.is_positive) {
          typeCounts[inf.infraction_type] = (typeCounts[inf.infraction_type] || 0) + 1;
        }
      }
      const repeatedTypes = Object.entries(typeCounts).filter(([, count]) => count >= 3);
      if (repeatedTypes.length > 0) {
        riskFactors.push(`Repeated pattern: ${repeatedTypes[0][0]} (${repeatedTypes[0][1]} times)`);
        riskScore += 2;
      }

      // Check: On probation
      if (probationMap[empId]) {
        riskFactors.push('Currently on probation');
        riskScore += 3;
      }

      // Only include if there are risk factors
      if (riskFactors.length > 0) {
        // Find next expiration
        let nextExpiration = null;
        let daysUntilExpiration = null;

        for (const inf of data.active_infractions) {
          if (inf.expiration_date) {
            const expDate = new Date(inf.expiration_date);
            if (!nextExpiration || expDate < nextExpiration) {
              nextExpiration = expDate;
            }
          }
        }

        if (nextExpiration) {
          daysUntilExpiration = Math.ceil((nextExpiration - today) / (1000 * 60 * 60 * 24));
        }

        // Determine recommended action
        let recommendedAction = 'Monitor closely';
        if (data.current_points >= 15) {
          recommendedAction = 'Review for termination';
        } else if (data.current_points >= 9) {
          recommendedAction = 'Director meeting required';
        } else if (riskScore >= 5) {
          recommendedAction = 'Schedule coaching conversation';
        }

        riskyEmployees.push({
          employee_id: data.employee_id,
          full_name: data.full_name,
          location: data.location,
          current_points: data.current_points,
          risk_factors: riskFactors,
          risk_score: riskScore,
          days_until_next_expiration: daysUntilExpiration,
          recommended_action: recommendedAction,
          is_on_probation: !!probationMap[empId]
        });
      }
    }

    // Sort by risk score descending
    riskyEmployees.sort((a, b) => b.risk_score - a.risk_score);

    return safeSerialize({
      success: true,
      employees: riskyEmployees,
      total_flagged: riskyEmployees.length,
      critical_count: riskyEmployees.filter(e => e.current_points >= 15).length,
      high_risk_count: riskyEmployees.filter(e => e.current_points >= 9 && e.current_points < 15).length,
      moderate_risk_count: riskyEmployees.filter(e => e.current_points < 9).length
    });

  } catch (error) {
    console.error('Error getting employee risk report:', error);
    logServerError('getEmployeeRiskReport', error, { token_present: !!token });
    return safeSerialize({ success: false, error: error.toString() });
  }
}

/**
 * Get recent activity feed for dashboard.
 *
 * @param {number} limit - Maximum number of events to return
 * @returns {Object} Recent activity data
 */
function getRecentActivity(limit, token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    limit = limit || 20;

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const infractionsSheet = ss.getSheetByName('Infractions');

    if (!infractionsSheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    const data = infractionsSheet.getDataRange().getValues();
    const headers = data[0];

    const employeeIdCol = headers.indexOf('Employee_ID');
    const dateCol = headers.indexOf('Date');
    const typeCol = headers.indexOf('Infraction_Type');
    const pointsCol = headers.indexOf('Points_Assigned') !== -1
      ? headers.indexOf('Points_Assigned')
      : headers.indexOf('Points');
    const pointValueCol = headers.indexOf('Point_Value_At_Time');
    const enteredByCol = headers.indexOf('Entered_By');
    const timestampCol = headers.indexOf('Entry_Timestamp') !== -1
      ? headers.indexOf('Entry_Timestamp')
      : headers.indexOf('Timestamp');

    // Get employee names
    const employees = getActiveEmployees();
    const employeeNames = {};
    for (const emp of employees) {
      employeeNames[emp.employee_id] = emp.full_name;
    }

    const activities = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = row[employeeIdCol];
      const date = row[dateCol];
      const type = row[typeCol];
      let points = pointsCol !== -1 ? row[pointsCol] : null;
      if ((points === '' || points === null || points === undefined) && pointValueCol !== -1) {
        points = row[pointValueCol];
      }
      if (points === '' || points === null || points === undefined) {
        points = 0;
      }
      const enteredBy = row[enteredByCol];
      const timestamp = row[timestampCol];

      const employeeName = employeeNames[employeeId] || employeeId;

      // Serialize dates to prevent JSON errors
      const dateStr = date instanceof Date ? date.toISOString() : String(date);
      const timestampStr = timestamp instanceof Date ? timestamp.toISOString() : String(timestamp);

      activities.push({
        type: 'infraction_added',
        employee_id: employeeId,
        employee_name: employeeName,
        infraction_type: type,
        points: points,
        entered_by: enteredBy,
        date: dateStr,
        timestamp: timestampStr,
        description: `${employeeName} received ${points} point(s) for ${type}`
      });
    }

    // Sort by timestamp descending
    activities.sort((a, b) => {
      const dateA = new Date(a.timestamp || a.date);
      const dateB = new Date(b.timestamp || b.date);
      return dateB - dateA;
    });

    return safeSerialize({
      success: true,
      activities: activities.slice(0, limit),
      total_count: activities.length
    });

  } catch (error) {
    console.error('Error getting recent activity:', error);
    logServerError('getRecentActivity', error, { token_present: !!token });
    return safeSerialize({ success: false, error: error.toString() });
  }
}

/**
 * Get manager-specific activity counts.
 */
function getManagerActivity(token) {
  try {
    const session = getCurrentRole(token);
    if (!session.authenticated) {
      return { success: false, sessionExpired: true };
    }

    const userName = session.user_name || session.role || '';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const infractionsSheet = ss.getSheetByName('Infractions');
    if (!infractionsSheet) {
      return { success: false, error: 'Infractions sheet not found' };
    }

    const data = infractionsSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const enteredByCol = headers.indexOf('Entered_By');
    const entryTimestampCol = headers.indexOf('Entry_Timestamp');
    if (enteredByCol === -1) {
      return { success: false, error: 'Entered_By column not found' };
    }

    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    let todayCount = 0;
    let monthCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const enteredBy = String(row[enteredByCol] || '');
      if (!userName || enteredBy.toLowerCase() !== userName.toLowerCase()) {
        continue;
      }

      const ts = entryTimestampCol !== -1 ? row[entryTimestampCol] : null;
      const entryDate = ts instanceof Date ? ts : (ts ? new Date(ts) : null);
      if (!entryDate || isNaN(entryDate.getTime())) {
        continue;
      }

      if (entryDate >= startOfMonth) {
        monthCount++;
      }
      if (entryDate >= startOfDay) {
        todayCount++;
      }
    }

    return {
      success: true,
      entries_today: todayCount,
      entries_this_month: monthCount
    };
  } catch (error) {
    logServerError('getManagerActivity', error, { token_present: !!token });
    return { success: false, error: String(error) };
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test dashboard statistics function
 */
function testDashboardStatistics() {
  console.log('=== Testing Dashboard Statistics ===\n');

  const result = getDashboardStatistics();
  console.log('Result:', JSON.stringify(result, null, 2));

  if (result.success) {
    console.log('\n=== Summary ===');
    console.log('Total Employees:', result.overall.total_active_employees);
    console.log('With Points:', result.overall.employees_with_points);
    console.log('Average Points:', result.overall.average_points_per_employee);
    console.log('Infractions This Month:', result.overall.infractions_this_month);
    console.log('Trend:', result.trends.points_trend);
    console.log('Execution Time:', result.executionTime + 'ms');
  }
}

/**
 * Test manager accountability function
 */
function testManagerAccountability() {
  console.log('=== Testing Manager Accountability ===\n');

  const result = getManagerAccountability();
  console.log('Result:', JSON.stringify(result, null, 2));

  if (result.success) {
    console.log('\n=== Summary ===');
    console.log('Total Managers:', result.total_managers);
    console.log('Active:', result.active_managers);
    console.log('Inactive:', result.inactive_managers);
  }
}

/**
 * Test location comparison function
 */
function testLocationComparison() {
  console.log('=== Testing Location Comparison ===\n');

  const result = getLocationComparison();
  console.log('Result:', JSON.stringify(result, null, 2));
}

/**
 * Test employee risk report function
 */
function testEmployeeRiskReport() {
  console.log('=== Testing Employee Risk Report ===\n');

  const result = getEmployeeRiskReport();
  console.log('Result:', JSON.stringify(result, null, 2));

  if (result.success) {
    console.log('\n=== Summary ===');
    console.log('Total Flagged:', result.total_flagged);
    console.log('Critical:', result.critical_count);
    console.log('High Risk:', result.high_risk_count);
    console.log('Moderate:', result.moderate_risk_count);
  }
}
