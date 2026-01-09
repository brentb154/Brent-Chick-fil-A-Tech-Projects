/**
 * Dashboard Module
 * Aggregates data from all modules for the dashboard display
 */

/**
 * Gets all dashboard data in a single call
 * @returns {Object} Dashboard data object
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get settings for thresholds
    const settings = getSettings();
    
    // Get current period info
    const currentPeriod = getCurrentPeriodInfo(ss);
    
    // Get OT metrics
    const otMetrics = getOTMetrics(ss);
    
    // Get uniform metrics
    const uniformMetrics = getUniformMetrics(ss);
    
    // Get PTO metrics (placeholder until PTO module is built)
    const ptoMetrics = getPTOMetrics(ss);
    
    // Get alerts
    const alerts = getAlerts();
    
    // Get top OT employees
    const topOTEmployees = getTopOTEmployees(ss);
    
    // Generate action items
    const actionItems = generateActionItems(otMetrics, uniformMetrics, ptoMetrics, alerts, currentPeriod, settings);
    
    return {
      success: true,
      currentPeriod: currentPeriod,
      otMetrics: otMetrics,
      uniformMetrics: uniformMetrics,
      ptoMetrics: ptoMetrics,
      alerts: alerts,
      topOTEmployees: topOTEmployees,
      actionItems: actionItems,
      lastUpdated: new Date().toISOString(),
      // Include settings thresholds for frontend use
      thresholds: {
        highOT: settings.highThreshold || 10,
        reallyHighOT: settings.reallyHighThreshold || 15,
        pendingPTOAlert: settings.pendingPTOAlertThreshold || 5,
        payrollUrgencyDays: settings.payrollUrgencyDays || 2
      }
    };
  } catch (error) {
    console.error('Error getting dashboard data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets current period information
 */
function getCurrentPeriodInfo(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
  
  // Default values
  let periodEnd = null;
  let periodStart = null;
  
  if (sheet && sheet.getLastRow() > 1) {
    // Get the most recent period end date
    const periods = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const uniquePeriods = [...new Set(periods.map(p => p[0] ? new Date(p[0]).getTime() : null).filter(p => p))];
    uniquePeriods.sort((a, b) => b - a);
    
    if (uniquePeriods.length > 0) {
      periodEnd = new Date(uniquePeriods[0]);
      periodStart = new Date(periodEnd);
      periodStart.setDate(periodStart.getDate() - 13); // 2-week period
    }
  }
  
  // Calculate next payroll date (Friday after period end, which is Saturday)
  let nextPayrollDate = null;
  let daysUntilPayroll = null;
  
  if (periodEnd) {
    // Period ends Saturday, payroll is the following Friday
    nextPayrollDate = new Date(periodEnd);
    nextPayrollDate.setDate(nextPayrollDate.getDate() + 6); // Saturday + 6 = Friday
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    nextPayrollDate.setHours(0, 0, 0, 0);
    
    daysUntilPayroll = Math.ceil((nextPayrollDate - today) / (1000 * 60 * 60 * 24));
    
    // If payroll date has passed, calculate next one
    if (daysUntilPayroll < 0) {
      nextPayrollDate.setDate(nextPayrollDate.getDate() + 14);
      daysUntilPayroll = Math.ceil((nextPayrollDate - today) / (1000 * 60 * 60 * 24));
    }
  }
  
  return {
    periodEnd: periodEnd ? formatDateForDisplay(periodEnd) : null,
    periodStart: periodStart ? formatDateForDisplay(periodStart) : null,
    nextPayrollDate: nextPayrollDate ? formatDateForDisplay(nextPayrollDate) : null,
    daysUntilPayroll: daysUntilPayroll
  };
}

/**
 * Gets OT metrics for the latest period
 */
function getOTMetrics(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
  
  const defaults = {
    latestPeriodHours: 0,
    latestPeriodCost: 0,
    highOTCount: 0,
    employeeCount: 0,
    previousPeriodHours: 0,
    percentChange: 0
  };
  
  if (!sheet || sheet.getLastRow() < 2) {
    return defaults;
  }
  
  try {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
    
    // Get unique periods sorted by date (newest first)
    const periodDates = [...new Set(data.map(r => r[0] ? new Date(r[0]).getTime() : null).filter(p => p))];
    periodDates.sort((a, b) => b - a);
    
    if (periodDates.length === 0) return defaults;
    
    const latestPeriod = periodDates[0];
    const previousPeriod = periodDates.length > 1 ? periodDates[1] : null;
    
    // Filter to latest period
    const latestData = data.filter(r => r[0] && new Date(r[0]).getTime() === latestPeriod);
    
    // Calculate metrics
    let totalOT = 0;
    let totalCost = 0;
    const employeeNames = new Set();
    const highOTEmployees = []; // Track employees with "Really High" flag
    
    latestData.forEach(row => {
      const name = row[1];
      const ot = parseFloat(row[12]) || 0; // Column M = Total OT (index 12)
      const cost = parseFloat(row[13]) || 0; // Column N = OT Cost (index 13)
      const flag = row[14] || ''; // Column O = Flag (index 14)
      
      if (name) employeeNames.add(name);
      totalOT += ot;
      totalCost += cost;
      
      // Track employees with "Really High" flag (not just above threshold)
      if (flag === 'Really High' && name) {
        // Check if already added (employee may have multiple location rows)
        if (!highOTEmployees.find(e => e.name === name)) {
          highOTEmployees.push({ name: name, hours: ot });
        }
      }
    });
    
    // Sort by hours descending
    highOTEmployees.sort((a, b) => b.hours - a.hours);
    
    // Calculate previous period for comparison
    let previousPeriodHours = 0;
    if (previousPeriod) {
      const previousData = data.filter(r => r[0] && new Date(r[0]).getTime() === previousPeriod);
      previousData.forEach(row => {
        previousPeriodHours += parseFloat(row[12]) || 0; // Column M = Total OT (index 12)
      });
    }
    
    const percentChange = previousPeriodHours > 0 
      ? Math.round(((totalOT - previousPeriodHours) / previousPeriodHours) * 100) 
      : 0;
    
    return {
      latestPeriodHours: Math.round(totalOT * 10) / 10,
      latestPeriodCost: Math.round(totalCost * 100) / 100,
      highOTCount: highOTEmployees.length, // Only "Really High" flagged employees
      highOTEmployeeNames: highOTEmployees.slice(0, 5).map(e => e.name), // Top 5 names for action item
      employeeCount: employeeNames.size,
      previousPeriodHours: Math.round(previousPeriodHours * 10) / 10,
      percentChange: percentChange
    };
  } catch (error) {
    console.error('Error getting OT metrics:', error);
    return defaults;
  }
}

/**
 * Gets uniform metrics
 */
function getUniformMetrics(ss) {
  const ordersSheet = ss.getSheetByName(SHEET_NAMES.UNIFORM_ORDERS);
  
  const defaults = {
    nextPaycheckDeductions: 0,
    activeOrdersCount: 0,
    upcomingPayments: []
  };
  
  if (!ordersSheet || ordersSheet.getLastRow() < 2) {
    return defaults;
  }
  
  try {
    // Read all 17 columns (updated for Location and Received_Date columns)
    const data = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 17).getValues();
    
    // Get active orders (Status is now column M = index 12)
    const activeOrders = data.filter(r => r[12] === 'Active');
    
    // Get upcoming paydays
    const paydays = getUpcomingPaydays(6);
    const upcomingPayments = [];
    
    // Calculate deductions for each payday
    // Column positions updated for Location column:
    // B=1: Employee_ID, C=2: Employee_Name, D=3: Location, E=4: Order_Date
    // F=5: Total_Amount, G=6: Payment_Plan, H=7: Amount_Per_Paycheck
    // I=8: First_Deduction_Date, J=9: Payments_Made
    paydays.slice(0, 3).forEach(payday => {
      const paydayDate = new Date(payday);
      let totalAmount = 0;
      let employeeCount = 0;
      const employees = new Set();
      
      activeOrders.forEach(order => {
        const firstDeduction = order[8] ? new Date(order[8]) : null; // I: First_Deduction_Date
        const paymentPlan = parseInt(order[6]) || 1;                  // G: Payment_Plan
        const amountPerPaycheck = parseFloat(order[7]) || 0;          // H: Amount_Per_Paycheck
        const paymentsMade = parseInt(order[9]) || 0;                 // J: Payments_Made
        const employeeName = order[2];                                 // C: Employee_Name
        
        if (!firstDeduction || paymentsMade >= paymentPlan) return;
        
        // Check if this payday falls within the payment schedule
        const daysDiff = Math.round((paydayDate - firstDeduction) / (1000 * 60 * 60 * 24));
        const paymentNumber = Math.floor(daysDiff / 14) + 1;
        
        if (daysDiff >= 0 && paymentNumber > paymentsMade && paymentNumber <= paymentPlan) {
          totalAmount += amountPerPaycheck;
          employees.add(employeeName);
        }
      });
      
      employeeCount = employees.size;
      
      if (employeeCount > 0) {
        upcomingPayments.push({
          date: formatDateForDisplay(paydayDate),
          employeeCount: employeeCount,
          amount: Math.round(totalAmount * 100) / 100
        });
      }
    });
    
    return {
      nextPaycheckDeductions: upcomingPayments.length > 0 ? upcomingPayments[0].amount : 0,
      activeOrdersCount: activeOrders.length,
      upcomingPayments: upcomingPayments
    };
  } catch (error) {
    console.error('Error getting uniform metrics:', error);
    return defaults;
  }
}

/**
 * Gets PTO metrics from the PTO sheet
 */
function getPTOMetrics(ss) {
  const defaults = {
    pendingCount: 0,
    thisWeekSubmissions: 0,
    unpaidHoursTotal: 0,
    moduleAvailable: true
  };
  
  try {
    const sheet = ss.getSheetByName('PTO');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return defaults;
    }
    
    // Read PTO data: columns A-M (13 columns)
    // A=PTO_ID, B=Employee_ID, C=Employee_Name, D=Location, E=Hours_Requested
    // F=PTO_Start_Date, G=PTO_End_Date, H=Payout_Period, I=HotSchedules_Confirmed
    // J=Submission_Date, K=Paid_Out, L=Status, M=Notes
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    
    const now = new Date();
    const sevenDaysAgo = new Date(now);
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    sevenDaysAgo.setHours(0, 0, 0, 0);
    
    let pendingCount = 0;
    let thisWeekSubmissions = 0;
    let unpaidHoursTotal = 0;
    
    data.forEach(row => {
      const ptoId = row[0];
      if (!ptoId) return; // Skip empty rows
      
      const hoursRequested = parseFloat(row[4]) || 0;
      const submissionDate = row[9] ? new Date(row[9]) : null;
      const paidOut = row[10] === true || row[10] === 'TRUE';
      const status = row[11] || '';
      
      // Count unpaid (pending) requests - not paid out and not cancelled/denied
      if (!paidOut && status !== 'Cancelled' && status !== 'Denied') {
        pendingCount++;
        unpaidHoursTotal += hoursRequested;
      }
      
      // Count this week's submissions
      if (submissionDate && submissionDate >= sevenDaysAgo) {
        thisWeekSubmissions++;
      }
    });
    
    return {
      pendingCount: pendingCount,
      thisWeekSubmissions: thisWeekSubmissions,
      unpaidHoursTotal: Math.round(unpaidHoursTotal * 10) / 10,
      moduleAvailable: true
    };
    
  } catch (error) {
    console.error('Error getting PTO metrics:', error);
    return defaults;
  }
}

/**
 * Gets top 5 OT employees from latest period
 */
function getTopOTEmployees(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.OT_HISTORY);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  try {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
    
    // Get latest period
    const periodDates = [...new Set(data.map(r => r[0] ? new Date(r[0]).getTime() : null).filter(p => p))];
    periodDates.sort((a, b) => b - a);
    
    if (periodDates.length === 0) return [];
    
    const latestPeriod = periodDates[0];
    
    // Filter to latest period and sort by OT hours
    const latestData = data
      .filter(r => r[0] && new Date(r[0]).getTime() === latestPeriod)
      .map(r => ({
        name: r[1],
        hours: parseFloat(r[12]) || 0, // Column M = Total OT (index 12)
        flag: r[14] || 'Normal' // Column O = Flag (index 14)
      }))
      .filter(e => e.hours > 0)
      .sort((a, b) => b.hours - a.hours)
      .slice(0, 5);
    
    return latestData;
  } catch (error) {
    console.error('Error getting top OT employees:', error);
    return [];
  }
}

/**
 * Generates action items based on current data
 */
function generateActionItems(otMetrics, uniformMetrics, ptoMetrics, alerts, currentPeriod, settings) {
  const items = [];
  
  // Get thresholds from settings (with defaults)
  const payrollUrgencyDays = (settings && settings.payrollUrgencyDays) || 2;
  const pendingPTOThreshold = (settings && settings.pendingPTOAlertThreshold) || 5;
  
  // Uniform deductions action
  if (uniformMetrics.upcomingPayments && uniformMetrics.upcomingPayments.length > 0) {
    const next = uniformMetrics.upcomingPayments[0];
    items.push({
      id: 'uniform_deductions',
      type: 'uniform_deductions',
      description: `Process uniform deductions for ${next.date} payroll (${next.employeeCount} orders)`,
      urgent: true,
      action: 'uniforms-deductions'
    });
  }
  
  // High OT action - only for "Really High" flagged employees with names
  if (otMetrics.highOTCount > 0) {
    const names = otMetrics.highOTEmployeeNames || [];
    let description;
    
    if (names.length === 0) {
      description = `Review ${otMetrics.highOTCount} employee${otMetrics.highOTCount > 1 ? 's' : ''} with Really High OT`;
    } else if (names.length <= 2) {
      description = `Review ${otMetrics.highOTCount} Really High OT alert${otMetrics.highOTCount > 1 ? 's' : ''} (${names.join(', ')})`;
    } else {
      // Show first 2 names + "..."
      const shortName = (name) => name.split(' ')[0]; // First name only
      description = `Review ${otMetrics.highOTCount} Really High OT alerts (${shortName(names[0])}, ${shortName(names[1])}...)`;
    }
    
    items.push({
      id: 'high_ot',
      type: 'high_ot',
      description: description,
      urgent: otMetrics.highOTCount >= 3,
      action: 'ot-trends'
    });
  }
  
  // Upload OT data reminder - due Monday before payday
  const uploadDueDate = getNextOTUploadDueDate(currentPeriod);
  if (uploadDueDate) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dueDate = new Date(uploadDueDate);
    dueDate.setHours(0, 0, 0, 0);
    const daysUntil = Math.ceil((dueDate - today) / (1000 * 60 * 60 * 24));
    
    if (daysUntil <= 7 && daysUntil >= 0) {
      items.push({
        id: 'upload_ot',
        type: 'upload_ot',
        description: `Upload next period OT data (Due: ${formatDateShort(uploadDueDate)})`,
        urgent: daysUntil <= payrollUrgencyDays,
        action: 'ot-upload'
      });
    }
  }
  
  // PTO action - show when pending count exceeds threshold
  if (ptoMetrics.pendingCount > 0) {
    items.push({
      id: 'pto_pending',
      type: 'pto',
      description: `Review ${ptoMetrics.pendingCount} pending PTO request${ptoMetrics.pendingCount > 1 ? 's' : ''} (${ptoMetrics.unpaidHoursTotal} hrs)`,
      urgent: ptoMetrics.pendingCount >= pendingPTOThreshold,
      action: 'pto-records'
    });
  }
  
  // Payroll processing action (when within urgency window of payroll)
  const payrollActionWindow = payrollUrgencyDays + 1; // Show action 1 day before urgency kicks in
  if (currentPeriod && currentPeriod.daysUntilPayroll !== null && currentPeriod.daysUntilPayroll <= payrollActionWindow) {
    let description;
    if (currentPeriod.daysUntilPayroll === 0) {
      description = 'Process payroll TODAY!';
    } else if (currentPeriod.daysUntilPayroll === 1) {
      description = 'Process payroll (due tomorrow)';
    } else {
      description = `Process payroll (due in ${currentPeriod.daysUntilPayroll} days)`;
    }
    
    items.push({
      id: 'payroll_due',
      type: 'payroll',
      description: description,
      urgent: currentPeriod.daysUntilPayroll <= payrollUrgencyDays,
      action: 'payroll-processing'
    });
  }
  
  // If no items, add a positive message
  if (items.length === 0) {
    items.push({
      id: 'all_clear',
      type: 'info',
      description: 'All caught up! No urgent action items.',
      urgent: false,
      action: null
    });
  }
  
  return items;
}

/**
 * Calculate next OT upload due date (Monday before payday)
 */
function getNextOTUploadDueDate(currentPeriod) {
  if (!currentPeriod || !currentPeriod.nextPayrollDate) return null;
  
  try {
    // Parse the next payroll date (Friday)
    const parts = currentPeriod.nextPayrollDate.split('/');
    const payday = new Date(parts[2], parts[0] - 1, parts[1]);
    
    // Monday before Friday payday = 4 days before
    const monday = new Date(payday);
    monday.setDate(monday.getDate() - 4);
    
    return monday;
  } catch (e) {
    console.error('Error calculating upload due date:', e);
    return null;
  }
}

/**
 * Format date as short string (M/D)
 */
function formatDateShort(date) {
  if (!date) return '';
  const d = new Date(date);
  return `${d.getMonth() + 1}/${d.getDate()}`;
}

/**
 * Helper: Format date for display
 */
function formatDateForDisplay(date) {
  if (!date) return null;
  const d = new Date(date);
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}

