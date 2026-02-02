/**
 * Haiti School Feeding Program Dashboard
 * Google Apps Script
 *
 * This script creates and updates dashboard sheets from your existing data.
 *
 * SETUP:
 * 1. Open your Google Sheet with the data
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Save (Ctrl+S)
 * 5. Run the function "createDashboard" (select it from dropdown, click Run)
 * 6. Grant permissions when prompted
 *
 * TO AUTO-UPDATE:
 * - Run "setupAutoRefresh" once to enable daily updates
 * - Or manually run "refreshDashboard" anytime
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  // Source sheet names (adjust if your sheets have different names)
  PRESENCE_SHEET: 'Présence',
  FEEDING_SHEET: 'Taux d\'alimentation',
  SCHOOLS_SHEET: 'Info sur les écoles',
  HEADCOUNT_SHEET: 'Comptage Physique',

  // Dashboard sheet names (will be created)
  EXEC_SUMMARY: 'Dashboard - Executive Summary',
  ATTENDANCE: 'Dashboard - Attendance Analysis',
  FEEDING: 'Dashboard - Feeding Analysis',
  SCHOOL_DETAIL: 'Dashboard - School Detail',
  SUPERVISOR: 'Dashboard - Supervisor Performance',

  // Thresholds
  ATTENDANCE_TARGET: 0.9,
  FEEDING_TARGET: 0.9,
  ALERT_THRESHOLD: 0.8,

  // Colors (Google Sheets format)
  HEADER_BG: '#4472C4',
  HEADER_TEXT: '#FFFFFF',
  GREEN_BG: '#C6EFCE',
  YELLOW_BG: '#FFEB9C',
  RED_BG: '#FFC7CE',
  LIGHT_BLUE_BG: '#DDEBF7'
};

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

/**
 * Main function to create or update the entire dashboard
 */
function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ui.alert('Creating Dashboard', 'This will create/update 5 dashboard sheets. Click OK to continue.', ui.ButtonSet.OK);

  try {
    // Load data
    const data = loadAllData(ss);

    if (!data.presence || data.presence.length === 0) {
      ui.alert('Error', 'Could not find data in the Présence sheet. Please check the sheet name.', ui.ButtonSet.OK);
      return;
    }

    // Calculate metrics
    const metrics = calculateAllMetrics(data);

    // Create/update each dashboard sheet
    createExecutiveSummary(ss, metrics, data);
    createAttendanceAnalysis(ss, metrics, data);
    createFeedingAnalysis(ss, metrics, data);
    createSchoolDetail(ss, metrics, data);
    createSupervisorPerformance(ss, metrics, data);

    // Move dashboard sheets to front
    organizeDashboardSheets(ss);

    ui.alert('Success!', 'Dashboard created successfully with 5 sheets. You can refresh anytime by running "refreshDashboard".', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', 'An error occurred: ' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * Refresh dashboard data (same as create but no prompts)
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const data = loadAllData(ss);
    const metrics = calculateAllMetrics(data);

    createExecutiveSummary(ss, metrics, data);
    createAttendanceAnalysis(ss, metrics, data);
    createFeedingAnalysis(ss, metrics, data);
    createSchoolDetail(ss, metrics, data);
    createSupervisorPerformance(ss, metrics, data);

    console.log('Dashboard refreshed at ' + new Date());
  } catch (error) {
    console.error('Refresh failed: ' + error.message);
  }
}

/**
 * Setup automatic daily refresh
 */
function setupAutoRefresh() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'refreshDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create daily trigger at 6 AM
  ScriptApp.newTrigger('refreshDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert('Auto-refresh enabled. Dashboard will update daily at 6 AM.');
}

/**
 * Add custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Dashboard')
    .addItem('Create/Update Dashboard', 'createDashboard')
    .addItem('Refresh Data', 'refreshDashboard')
    .addSeparator()
    .addItem('Setup Auto-Refresh (Daily)', 'setupAutoRefresh')
    .addToUi();
}

// ============================================================================
// DATA LOADING
// ============================================================================

/**
 * Load all data from source sheets
 */
function loadAllData(ss) {
  const data = {};

  // Load Présence data
  const presenceSheet = ss.getSheetByName(CONFIG.PRESENCE_SHEET);
  if (presenceSheet) {
    const presenceData = presenceSheet.getDataRange().getValues();
    data.presenceHeaders = presenceData[0];
    data.presence = presenceData.slice(2); // Skip header and blank row
    data.presenceColIndex = mapColumns(data.presenceHeaders, {
      'month': ['Le mois'],
      'schoolName': ['Nom de l\'établissement', 'Nom de l\'etablissement'],
      'schoolId': ['ID de l\'établissement', 'ID de l\'etablissement'],
      'commune': ['Commune'],
      'department': ['Departement'],
      'supervisor': ['Supervisor'],
      'attendanceRate': ['Le taux de présence'],
      'enrollment': ['Le total Effectif'],
      'variationPct': ['% de variation'],
      'variationCategory': ['La catégorie de variation'],
      'variationReason': ['La raison de la variation']
    });
  }

  // Load Feeding data
  const feedingSheet = ss.getSheetByName(CONFIG.FEEDING_SHEET);
  if (feedingSheet) {
    const feedingData = feedingSheet.getDataRange().getValues();
    data.feedingHeaders = feedingData[0];
    data.feeding = feedingData.slice(1);
    data.feedingColIndex = mapColumns(data.feedingHeaders, {
      'weekStart': ['Semaine commencant'],
      'month': ['Le mois'],
      'schoolName': ['Nom de l\'établissement'],
      'schoolId': ['ID de l\'établissement'],
      'commune': ['Commune'],
      'department': ['Departement'],
      'supervisor': ['Supervisor'],
      'daysPlanned': ['Le nombre de jours d\'alimentation prévu'],
      'daysFed': ['Le nombre réel de jours d\'alimentation'],
      'feedingRate': ['Le taux d\'alimentation'],
      'nonfeedCat1': ['1. La catégorie de non-alimentation'],
      'nonfeedCat2': ['2. La catégorie de non-alimentation'],
      'nonfeedCat3': ['3. La catégorie de non-alimentation'],
      'nonfeedCat4': ['4. La catégorie de non-alimentation'],
      'nonfeedCat5': ['5. La catégorie de non-alimentation']
    });
  }

  // Load Schools data
  const schoolsSheet = ss.getSheetByName(CONFIG.SCHOOLS_SHEET);
  if (schoolsSheet) {
    const schoolsData = schoolsSheet.getDataRange().getValues();
    data.schoolsHeaders = schoolsData[1]; // Header is in row 2
    data.schools = schoolsData.slice(2);
    data.schoolsColIndex = mapColumns(data.schoolsHeaders, {
      'schoolName': ['Nom de l\'Etablissement', 'Nom de l\'etablissement'],
      'schoolId': ['ID de l\'Etablissement', 'ID de l\'etablissement'],
      'commune': ['Commune'],
      'department': ['Departement'],
      'supervisor': ['Supervisor'],
      'grades': ['Grades served'],
      'ownership': ['School ownership']
    });
  }

  return data;
}

/**
 * Map column names to indices (handles variations in naming)
 */
function mapColumns(headers, mappings) {
  const indices = {};

  for (const [key, patterns] of Object.entries(mappings)) {
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase().replace(/[''`]/g, "'");
      for (const pattern of patterns) {
        if (header.includes(pattern.toLowerCase().replace(/[''`]/g, "'"))) {
          indices[key] = i;
          break;
        }
      }
      if (indices[key] !== undefined) break;
    }
  }

  return indices;
}

// ============================================================================
// METRICS CALCULATION
// ============================================================================

/**
 * Calculate all dashboard metrics
 */
function calculateAllMetrics(data) {
  const metrics = {};
  const col = data.presenceColIndex;

  // Get latest month
  const months = data.presence
    .map(row => row[col.month])
    .filter(m => m !== '' && m !== null && m !== undefined);

  metrics.latestMonth = Math.max(...months.filter(m => typeof m === 'number'));
  metrics.latestMonthStr = formatExcelDate(metrics.latestMonth);

  // Filter to current month
  const currentMonthData = data.presence.filter(row => row[col.month] === metrics.latestMonth);

  // KPIs
  const schoolIds = [...new Set(currentMonthData.map(row => row[col.schoolId]).filter(id => id))];
  metrics.totalSchools = schoolIds.length;

  const attendanceRates = currentMonthData
    .map(row => row[col.attendanceRate])
    .filter(r => typeof r === 'number' && !isNaN(r));
  metrics.avgAttendance = attendanceRates.length > 0
    ? attendanceRates.reduce((a, b) => a + b, 0) / attendanceRates.length
    : 0;

  // Schools below 80%
  const schoolAttendance = {};
  currentMonthData.forEach(row => {
    const id = row[col.schoolId];
    const rate = row[col.attendanceRate];
    if (id && typeof rate === 'number') {
      if (!schoolAttendance[id]) schoolAttendance[id] = [];
      schoolAttendance[id].push(rate);
    }
  });

  metrics.schoolsBelow80 = Object.values(schoolAttendance)
    .filter(rates => rates.reduce((a, b) => a + b, 0) / rates.length < 0.8)
    .length;

  // Feeding metrics
  const fcol = data.feedingColIndex;
  if (data.feeding && fcol.feedingRate !== undefined) {
    const feedingRates = data.feeding
      .map(row => row[fcol.feedingRate])
      .filter(r => typeof r === 'number' && !isNaN(r));
    metrics.avgFeeding = feedingRates.length > 0
      ? feedingRates.reduce((a, b) => a + b, 0) / feedingRates.length
      : 0;
  } else {
    metrics.avgFeeding = 0;
  }

  // Monthly trends
  metrics.attendanceTrend = calculateMonthlyTrend(data.presence, col.month, col.attendanceRate, col.schoolId);

  // Commune stats
  metrics.communeStats = calculateCommuneStats(currentMonthData, col);

  // Alerts (schools needing attention)
  metrics.alerts = calculateAlerts(currentMonthData, col);

  // Variation reasons
  metrics.variationReasons = countCategories(currentMonthData, col.variationCategory);

  // Non-feeding reasons
  metrics.nonfeedingReasons = calculateNonfeedingReasons(data);

  // Supervisor stats
  metrics.supervisorStats = calculateSupervisorStats(currentMonthData, col);

  // Feeding trends
  metrics.feedingTrend = calculateFeedingTrend(data);

  return metrics;
}

/**
 * Calculate monthly attendance trend
 */
function calculateMonthlyTrend(presenceData, monthCol, rateCol, schoolCol) {
  const byMonth = {};

  presenceData.forEach(row => {
    const month = row[monthCol];
    const rate = row[rateCol];
    const school = row[schoolCol];

    if (month && typeof rate === 'number') {
      if (!byMonth[month]) {
        byMonth[month] = { rates: [], schools: new Set() };
      }
      byMonth[month].rates.push(rate);
      if (school) byMonth[month].schools.add(school);
    }
  });

  const trend = Object.entries(byMonth)
    .map(([month, data]) => ({
      month: Number(month),
      avgRate: data.rates.reduce((a, b) => a + b, 0) / data.rates.length,
      schoolCount: data.schools.size
    }))
    .sort((a, b) => b.month - a.month)
    .slice(0, 6);

  return trend;
}

/**
 * Calculate commune statistics
 */
function calculateCommuneStats(data, col) {
  const byCommune = {};

  data.forEach(row => {
    const commune = row[col.commune];
    const rate = row[col.attendanceRate];
    const school = row[col.schoolId];

    if (commune && typeof rate === 'number') {
      if (!byCommune[commune]) {
        byCommune[commune] = { rates: [], schools: new Set() };
      }
      byCommune[commune].rates.push(rate);
      if (school) byCommune[commune].schools.add(school);
    }
  });

  return Object.entries(byCommune)
    .map(([commune, data]) => ({
      commune,
      avgRate: data.rates.reduce((a, b) => a + b, 0) / data.rates.length,
      schoolCount: data.schools.size
    }))
    .sort((a, b) => b.avgRate - a.avgRate);
}

/**
 * Calculate alerts (schools below threshold)
 */
function calculateAlerts(data, col) {
  const bySchool = {};

  data.forEach(row => {
    const id = row[col.schoolId];
    const name = row[col.schoolName];
    const commune = row[col.commune];
    const supervisor = row[col.supervisor];
    const rate = row[col.attendanceRate];

    if (id && typeof rate === 'number') {
      if (!bySchool[id]) {
        bySchool[id] = { name, commune, supervisor, rates: [] };
      }
      bySchool[id].rates.push(rate);
    }
  });

  return Object.entries(bySchool)
    .map(([id, data]) => ({
      id,
      name: data.name,
      commune: data.commune,
      supervisor: data.supervisor,
      avgRate: data.rates.reduce((a, b) => a + b, 0) / data.rates.length
    }))
    .filter(s => s.avgRate < CONFIG.ALERT_THRESHOLD)
    .sort((a, b) => a.avgRate - b.avgRate);
}

/**
 * Count categories
 */
function countCategories(data, colIndex) {
  const counts = {};

  data.forEach(row => {
    const category = row[colIndex];
    if (category && category !== '') {
      counts[category] = (counts[category] || 0) + 1;
    }
  });

  return Object.entries(counts)
    .map(([category, count]) => ({ category, count }))
    .sort((a, b) => b.count - a.count);
}

/**
 * Calculate non-feeding reasons
 */
function calculateNonfeedingReasons(data) {
  const counts = {};
  const fcol = data.feedingColIndex;

  if (!data.feeding) return [];

  const catCols = ['nonfeedCat1', 'nonfeedCat2', 'nonfeedCat3', 'nonfeedCat4', 'nonfeedCat5'];

  data.feeding.forEach(row => {
    catCols.forEach(colKey => {
      const idx = fcol[colKey];
      if (idx !== undefined) {
        const cat = row[idx];
        if (cat && cat !== '') {
          counts[cat] = (counts[cat] || 0) + 1;
        }
      }
    });
  });

  return Object.entries(counts)
    .map(([category, count]) => ({ category, count }))
    .sort((a, b) => b.count - a.count);
}

/**
 * Calculate supervisor statistics
 */
function calculateSupervisorStats(data, col) {
  const bySupervisor = {};

  data.forEach(row => {
    const supervisor = row[col.supervisor];
    const rate = row[col.attendanceRate];
    const school = row[col.schoolId];

    if (supervisor && typeof rate === 'number') {
      if (!bySupervisor[supervisor]) {
        bySupervisor[supervisor] = { rates: [], schools: new Set(), belowThreshold: new Set() };
      }
      bySupervisor[supervisor].rates.push(rate);
      if (school) {
        bySupervisor[supervisor].schools.add(school);
        if (rate < CONFIG.ALERT_THRESHOLD) {
          bySupervisor[supervisor].belowThreshold.add(school);
        }
      }
    }
  });

  return Object.entries(bySupervisor)
    .map(([supervisor, data]) => ({
      supervisor,
      avgRate: data.rates.reduce((a, b) => a + b, 0) / data.rates.length,
      schoolCount: data.schools.size,
      schoolsBelow80: data.belowThreshold.size
    }))
    .sort((a, b) => b.avgRate - a.avgRate);
}

/**
 * Calculate feeding trend by month
 */
function calculateFeedingTrend(data) {
  const fcol = data.feedingColIndex;
  if (!data.feeding || fcol.month === undefined) return [];

  const byMonth = {};

  data.feeding.forEach(row => {
    const month = row[fcol.month];
    const rate = row[fcol.feedingRate];
    const planned = row[fcol.daysPlanned] || 0;
    const fed = row[fcol.daysFed] || 0;

    if (month) {
      if (!byMonth[month]) {
        byMonth[month] = { rates: [], planned: 0, fed: 0 };
      }
      if (typeof rate === 'number') byMonth[month].rates.push(rate);
      byMonth[month].planned += Number(planned) || 0;
      byMonth[month].fed += Number(fed) || 0;
    }
  });

  return Object.entries(byMonth)
    .map(([month, data]) => ({
      month,
      avgRate: data.rates.length > 0 ? data.rates.reduce((a, b) => a + b, 0) / data.rates.length : 0,
      planned: data.planned,
      fed: data.fed
    }))
    .sort((a, b) => String(b.month).localeCompare(String(a.month)))
    .slice(0, 6);
}

// ============================================================================
// DASHBOARD SHEET CREATION
// ============================================================================

/**
 * Create or get a dashboard sheet
 */
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }
  return sheet;
}

/**
 * Format header row
 */
function formatHeader(sheet, row, startCol, endCol) {
  const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
  range.setBackground(CONFIG.HEADER_BG);
  range.setFontColor(CONFIG.HEADER_TEXT);
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
}

/**
 * Format Excel date to readable string
 */
function formatExcelDate(excelDate) {
  if (!excelDate || typeof excelDate !== 'number') return String(excelDate || '');
  try {
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getFullYear();
  } catch (e) {
    return String(excelDate);
  }
}

/**
 * Apply conditional color based on rate
 */
function getRateColor(rate) {
  if (rate >= 0.8) return CONFIG.GREEN_BG;
  if (rate >= 0.7) return CONFIG.YELLOW_BG;
  return CONFIG.RED_BG;
}

/**
 * Create Executive Summary sheet
 */
function createExecutiveSummary(ss, metrics, data) {
  const sheet = getOrCreateSheet(ss, CONFIG.EXEC_SUMMARY);

  // Title
  sheet.getRange('A1').setValue('Haiti School Feeding Program - Executive Summary');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A1:F1').merge();

  // Current month
  sheet.getRange('A3').setValue('Current Month:').setFontWeight('bold');
  sheet.getRange('B3').setValue(metrics.latestMonthStr);
  sheet.getRange('D3').setValue('Attendance Target:').setFontWeight('bold');
  sheet.getRange('E3').setValue(CONFIG.ATTENDANCE_TARGET).setNumberFormat('0%');
  sheet.getRange('D4').setValue('Feeding Target:').setFontWeight('bold');
  sheet.getRange('E4').setValue(CONFIG.FEEDING_TARGET).setNumberFormat('0%');

  // KPIs Section
  sheet.getRange('A6').setValue('Key Performance Indicators').setFontSize(12).setFontWeight('bold');

  const kpiHeaders = ['Metric', 'Value', 'Target', 'Status'];
  sheet.getRange(7, 1, 1, 4).setValues([kpiHeaders]);
  formatHeader(sheet, 7, 1, 4);

  const kpis = [
    ['Total Schools Active', metrics.totalSchools, '-', '✓'],
    ['Average Attendance Rate', metrics.avgAttendance, CONFIG.ATTENDANCE_TARGET,
     metrics.avgAttendance >= 0.9 ? '✓' : metrics.avgAttendance >= 0.8 ? '⚠' : '✗'],
    ['Average Feeding Rate', metrics.avgFeeding, CONFIG.FEEDING_TARGET,
     metrics.avgFeeding >= 0.9 ? '✓' : metrics.avgFeeding >= 0.8 ? '⚠' : '✗'],
    ['Schools Below 80%', metrics.schoolsBelow80, 0, metrics.schoolsBelow80 === 0 ? '✓' : '⚠']
  ];

  sheet.getRange(8, 1, kpis.length, 4).setValues(kpis);
  sheet.getRange('B9:B10').setNumberFormat('0.0%');
  sheet.getRange('C9:C10').setNumberFormat('0%');

  // Color code KPIs
  sheet.getRange('B9').setBackground(getRateColor(metrics.avgAttendance));
  sheet.getRange('B10').setBackground(getRateColor(metrics.avgFeeding));

  // Alerts Section
  sheet.getRange('A14').setValue('Schools Needing Attention (Attendance < 80%)').setFontSize(12).setFontWeight('bold');

  const alertHeaders = ['School Name', 'Commune', 'Supervisor', 'Attendance'];
  sheet.getRange(15, 1, 1, 4).setValues([alertHeaders]);
  formatHeader(sheet, 15, 1, 4);

  if (metrics.alerts.length > 0) {
    const alertData = metrics.alerts.slice(0, 10).map(a => [a.name, a.commune, a.supervisor, a.avgRate]);
    sheet.getRange(16, 1, alertData.length, 4).setValues(alertData);
    sheet.getRange(16, 4, alertData.length, 1).setNumberFormat('0.0%');

    // Color code alerts
    for (let i = 0; i < alertData.length; i++) {
      sheet.getRange(16 + i, 4).setBackground(getRateColor(alertData[i][3]));
    }
  } else {
    sheet.getRange('A16').setValue('No schools below 80% threshold - Great!');
  }

  // Commune Comparison
  sheet.getRange('F6').setValue('Attendance by Commune').setFontSize(12).setFontWeight('bold');

  const communeHeaders = ['Commune', 'Avg Attendance', 'Schools'];
  sheet.getRange(7, 6, 1, 3).setValues([communeHeaders]);
  formatHeader(sheet, 7, 6, 8);

  if (metrics.communeStats.length > 0) {
    const communeData = metrics.communeStats.slice(0, 12).map(c => [c.commune, c.avgRate, c.schoolCount]);
    sheet.getRange(8, 6, communeData.length, 3).setValues(communeData);
    sheet.getRange(8, 7, communeData.length, 1).setNumberFormat('0.0%');

    // Create chart
    const chartRange = sheet.getRange(7, 6, communeData.length + 1, 2);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(chartRange)
      .setPosition(6, 10, 0, 0)
      .setOption('title', 'Attendance by Commune')
      .setOption('legend', {position: 'none'})
      .setOption('width', 400)
      .setOption('height', 300)
      .build();
    sheet.insertChart(chart);
  }

  // Trend Section
  sheet.getRange('A28').setValue('6-Month Attendance Trend').setFontSize(12).setFontWeight('bold');

  const trendHeaders = ['Month', 'Avg Attendance', 'Schools'];
  sheet.getRange(29, 1, 1, 3).setValues([trendHeaders]);
  formatHeader(sheet, 29, 1, 3);

  if (metrics.attendanceTrend.length > 0) {
    const trendData = metrics.attendanceTrend.map(t => [formatExcelDate(t.month), t.avgRate, t.schoolCount]);
    sheet.getRange(30, 1, trendData.length, 3).setValues(trendData);
    sheet.getRange(30, 2, trendData.length, 1).setNumberFormat('0.0%');

    // Create line chart
    const trendRange = sheet.getRange(29, 1, trendData.length + 1, 2);
    const trendChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(trendRange)
      .setPosition(28, 5, 0, 0)
      .setOption('title', 'Attendance Trend')
      .setOption('legend', {position: 'none'})
      .setOption('width', 350)
      .setOption('height', 200)
      .build();
    sheet.insertChart(trendChart);
  }

  // Column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 80);
}

/**
 * Create Attendance Analysis sheet
 */
function createAttendanceAnalysis(ss, metrics, data) {
  const sheet = getOrCreateSheet(ss, CONFIG.ATTENDANCE);

  // Title
  sheet.getRange('A1').setValue('Attendance Analysis');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');

  // Trend Section
  sheet.getRange('A4').setValue('Monthly Attendance Trend').setFontSize(12).setFontWeight('bold');

  const trendHeaders = ['Month', 'Avg Attendance', 'School Count'];
  sheet.getRange(5, 1, 1, 3).setValues([trendHeaders]);
  formatHeader(sheet, 5, 1, 3);

  if (metrics.attendanceTrend.length > 0) {
    const trendData = metrics.attendanceTrend.map(t => [formatExcelDate(t.month), t.avgRate, t.schoolCount]);
    sheet.getRange(6, 1, trendData.length, 3).setValues(trendData);
    sheet.getRange(6, 2, trendData.length, 1).setNumberFormat('0.0%');

    // Chart
    const chartRange = sheet.getRange(5, 1, trendData.length + 1, 2);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
      .setPosition(4, 5, 0, 0)
      .setOption('title', '6-Month Attendance Trend')
      .setOption('width', 450)
      .setOption('height', 250)
      .build();
    sheet.insertChart(chart);
  }

  // Variation Reasons
  sheet.getRange('A15').setValue('Attendance Variation Reasons').setFontSize(12).setFontWeight('bold');

  const reasonHeaders = ['Category', 'Count'];
  sheet.getRange(16, 1, 1, 2).setValues([reasonHeaders]);
  formatHeader(sheet, 16, 1, 2);

  if (metrics.variationReasons.length > 0) {
    const reasonData = metrics.variationReasons.slice(0, 10).map(r => [r.category, r.count]);
    sheet.getRange(17, 1, reasonData.length, 2).setValues(reasonData);

    // Pie chart
    const pieRange = sheet.getRange(16, 1, reasonData.length + 1, 2);
    const pieChart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(pieRange)
      .setPosition(15, 5, 0, 0)
      .setOption('title', 'Variation Reasons')
      .setOption('width', 350)
      .setOption('height', 250)
      .build();
    sheet.insertChart(pieChart);
  } else {
    sheet.getRange('A17').setValue('No variation reasons recorded');
  }

  // Column widths
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 100);
}

/**
 * Create Feeding Analysis sheet
 */
function createFeedingAnalysis(ss, metrics, data) {
  const sheet = getOrCreateSheet(ss, CONFIG.FEEDING);

  // Title
  sheet.getRange('A1').setValue('Feeding Rate Analysis');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');

  // Feeding Summary
  sheet.getRange('A4').setValue('Feeding Rate Summary').setFontSize(12).setFontWeight('bold');

  const summaryHeaders = ['Month', 'Avg Feeding Rate', 'Days Planned', 'Days Fed', 'Days Missed'];
  sheet.getRange(5, 1, 1, 5).setValues([summaryHeaders]);
  formatHeader(sheet, 5, 1, 5);

  if (metrics.feedingTrend.length > 0) {
    const feedData = metrics.feedingTrend.map(f => [
      f.month, f.avgRate, f.planned, f.fed, f.planned - f.fed
    ]);
    sheet.getRange(6, 1, feedData.length, 5).setValues(feedData);
    sheet.getRange(6, 2, feedData.length, 1).setNumberFormat('0.0%');

    // Color code feeding rates
    for (let i = 0; i < feedData.length; i++) {
      sheet.getRange(6 + i, 2).setBackground(getRateColor(feedData[i][1]));
    }

    // Stacked bar chart
    const chartRange = sheet.getRange(5, 1, feedData.length + 1, 5);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(5, 1, feedData.length + 1, 1))
      .addRange(sheet.getRange(5, 4, feedData.length + 1, 2))
      .setPosition(4, 7, 0, 0)
      .setOption('title', 'Days Fed vs Missed')
      .setOption('isStacked', true)
      .setOption('width', 400)
      .setOption('height', 250)
      .build();
    sheet.insertChart(chart);
  }

  // Non-feeding Reasons
  sheet.getRange('A15').setValue('Non-Feeding Reasons').setFontSize(12).setFontWeight('bold');

  const nfHeaders = ['Category', 'Count'];
  sheet.getRange(16, 1, 1, 2).setValues([nfHeaders]);
  formatHeader(sheet, 16, 1, 2);

  if (metrics.nonfeedingReasons.length > 0) {
    const nfData = metrics.nonfeedingReasons.slice(0, 10).map(r => [r.category, r.count]);
    sheet.getRange(17, 1, nfData.length, 2).setValues(nfData);

    // Bar chart
    const barRange = sheet.getRange(16, 1, nfData.length + 1, 2);
    const barChart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(barRange)
      .setPosition(15, 5, 0, 0)
      .setOption('title', 'Non-Feeding Reasons')
      .setOption('legend', {position: 'none'})
      .setOption('width', 350)
      .setOption('height', 200)
      .build();
    sheet.insertChart(barChart);
  } else {
    sheet.getRange('A17').setValue('No non-feeding reasons recorded');
  }

  // Column widths
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
}

/**
 * Create School Detail sheet
 */
function createSchoolDetail(ss, metrics, data) {
  const sheet = getOrCreateSheet(ss, CONFIG.SCHOOL_DETAIL);
  const scol = data.schoolsColIndex;

  // Title
  sheet.getRange('A1').setValue('School Detail View');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');

  sheet.getRange('A3').setValue('Select a school from the dropdown below:').setFontStyle('italic');

  // School dropdown
  sheet.getRange('A5').setValue('Selected School:').setFontWeight('bold');

  // Get school names for dropdown
  const schoolNames = data.schools
    .map(row => row[scol.schoolName])
    .filter(name => name && name !== '')
    .sort();

  if (schoolNames.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(schoolNames.slice(0, 500), true)
      .build();
    sheet.getRange('B5').setDataValidation(rule);
    sheet.getRange('B5').setValue(schoolNames[0]);
  }

  // School Profile (using formulas to look up selected school)
  sheet.getRange('A7').setValue('School Profile').setFontSize(12).setFontWeight('bold');
  sheet.getRange('A7:D7').merge().setBackground(CONFIG.LIGHT_BLUE_BG);

  const profileLabels = ['School ID:', 'School Name:', 'Commune:', 'Department:', 'Supervisor:', 'Grades:', 'Type:'];
  profileLabels.forEach((label, i) => {
    sheet.getRange(8 + i, 1).setValue(label).setFontWeight('bold');
  });

  // Set up VLOOKUP formulas to pull school data
  if (data.schools.length > 0) {
    const schoolsSheetName = CONFIG.SCHOOLS_SHEET;
    // Note: These formulas reference the Info sur les écoles sheet
    sheet.getRange('B8').setFormula(`=IFERROR(VLOOKUP(B5,'${schoolsSheetName}'!A:B,2,FALSE),"")`);
    sheet.getRange('B9').setValue('=B5');
    sheet.getRange('B10').setFormula(`=IFERROR(VLOOKUP(B5,'${schoolsSheetName}'!A:C,3,FALSE),"")`);
    sheet.getRange('B11').setFormula(`=IFERROR(VLOOKUP(B5,'${schoolsSheetName}'!A:D,4,FALSE),"")`);
    sheet.getRange('B12').setFormula(`=IFERROR(VLOOKUP(B5,'${schoolsSheetName}'!A:E,5,FALSE),"")`);
    sheet.getRange('B13').setFormula(`=IFERROR(VLOOKUP(B5,'${schoolsSheetName}'!A:G,7,FALSE),"")`);
    sheet.getRange('B14').setFormula(`=IFERROR(VLOOKUP(B5,'${schoolsSheetName}'!A:H,8,FALSE),"")`);
  }

  // All Schools Reference
  sheet.getRange('A18').setValue('All Schools Reference').setFontSize(12).setFontWeight('bold');

  const schoolHeaders = ['School Name', 'ID', 'Commune', 'Department', 'Supervisor'];
  sheet.getRange(19, 1, 1, 5).setValues([schoolHeaders]);
  formatHeader(sheet, 19, 1, 5);

  if (data.schools.length > 0) {
    const schoolData = data.schools.slice(0, 100).map(row => [
      row[scol.schoolName] || '',
      row[scol.schoolId] || '',
      row[scol.commune] || '',
      row[scol.department] || '',
      row[scol.supervisor] || ''
    ]);
    sheet.getRange(20, 1, schoolData.length, 5).setValues(schoolData);
  }

  // Column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
}

/**
 * Create Supervisor Performance sheet
 */
function createSupervisorPerformance(ss, metrics, data) {
  const sheet = getOrCreateSheet(ss, CONFIG.SUPERVISOR);

  // Title
  sheet.getRange('A1').setValue('Supervisor Performance');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');

  // Ranking Table
  sheet.getRange('A4').setValue('Supervisor Ranking').setFontSize(12).setFontWeight('bold');

  const headers = ['Rank', 'Supervisor', 'Schools', 'Avg Attendance', 'Schools < 80%', 'Status'];
  sheet.getRange(5, 1, 1, 6).setValues([headers]);
  formatHeader(sheet, 5, 1, 6);

  if (metrics.supervisorStats.length > 0) {
    const supData = metrics.supervisorStats.map((s, i) => {
      let status = '✓ OK';
      if (s.schoolsBelow80 >= 3) status = '⚠ Critical';
      else if (s.schoolsBelow80 >= 2) status = '⚠ Attention';
      return [i + 1, s.supervisor, s.schoolCount, s.avgRate, s.schoolsBelow80, status];
    });

    sheet.getRange(6, 1, supData.length, 6).setValues(supData);
    sheet.getRange(6, 4, supData.length, 1).setNumberFormat('0.0%');

    // Color code
    for (let i = 0; i < supData.length; i++) {
      sheet.getRange(6 + i, 4).setBackground(getRateColor(supData[i][3]));

      const below = supData[i][4];
      if (below >= 3) {
        sheet.getRange(6 + i, 6).setBackground(CONFIG.RED_BG);
      } else if (below >= 2) {
        sheet.getRange(6 + i, 6).setBackground(CONFIG.YELLOW_BG);
      } else {
        sheet.getRange(6 + i, 6).setBackground(CONFIG.GREEN_BG);
      }
    }

    // Bar chart
    const chartRange = sheet.getRange(5, 2, Math.min(supData.length + 1, 16), 3);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(5, 2, Math.min(supData.length + 1, 16), 1))
      .addRange(sheet.getRange(5, 4, Math.min(supData.length + 1, 16), 1))
      .setPosition(4, 8, 0, 0)
      .setOption('title', 'Supervisor Performance')
      .setOption('legend', {position: 'none'})
      .setOption('width', 450)
      .setOption('height', 350)
      .build();
    sheet.insertChart(chart);
  }

  // Problem Flagging
  const problemSups = metrics.supervisorStats.filter(s => s.schoolsBelow80 >= 2);
  const problemRow = Math.max(6 + metrics.supervisorStats.length, 25);

  sheet.getRange(problemRow, 1).setValue('Supervisors with 2+ Problem Schools').setFontSize(12).setFontWeight('bold');

  if (problemSups.length > 0) {
    const problemHeaders = ['Supervisor', 'Problem Schools', 'Total Schools', 'Problem Rate'];
    sheet.getRange(problemRow + 1, 1, 1, 4).setValues([problemHeaders]);
    formatHeader(sheet, problemRow + 1, 1, 4);

    const problemData = problemSups.map(s => [
      s.supervisor, s.schoolsBelow80, s.schoolCount, s.schoolsBelow80 / s.schoolCount
    ]);
    sheet.getRange(problemRow + 2, 1, problemData.length, 4).setValues(problemData);
    sheet.getRange(problemRow + 2, 4, problemData.length, 1).setNumberFormat('0%');
  } else {
    sheet.getRange(problemRow + 2, 1).setValue('No supervisors with 2+ problem schools - Great!').setBackground(CONFIG.GREEN_BG);
  }

  // Column widths
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
}

/**
 * Organize dashboard sheets (move to front)
 */
function organizeDashboardSheets(ss) {
  const dashboardNames = [
    CONFIG.EXEC_SUMMARY,
    CONFIG.ATTENDANCE,
    CONFIG.FEEDING,
    CONFIG.SCHOOL_DETAIL,
    CONFIG.SUPERVISOR
  ];

  dashboardNames.forEach((name, index) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });

  // Activate executive summary
  const execSheet = ss.getSheetByName(CONFIG.EXEC_SUMMARY);
  if (execSheet) ss.setActiveSheet(execSheet);
}
