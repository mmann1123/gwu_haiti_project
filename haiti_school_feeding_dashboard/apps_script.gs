/**
 * Haiti School Feeding Dashboard - Google Apps Script
 *
 * Optional automation scripts for:
 * 1. Auto-refresh data on open
 * 2. Email alerts for problem schools
 * 3. Create filter views programmatically
 * 4. Generate PDF reports
 */

// Configuration
const CONFIG = {
  ATTENDANCE_THRESHOLD: 0.8,
  FEEDING_THRESHOLD: 0.8,
  ALERT_EMAIL: 'program-manager@example.com',
  DASHBOARD_SHEET: 'Executive Summary'
};

/**
 * Runs when spreadsheet is opened
 * Refreshes calculations and checks for alerts
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Add custom menu
  ui.createMenu('Dashboard')
    .addItem('Refresh All Data', 'refreshAllData')
    .addItem('Send Alert Email', 'sendAlertEmail')
    .addItem('Generate PDF Report', 'generatePDFReport')
    .addSeparator()
    .addItem('Setup Dashboard', 'setupDashboard')
    .addToUi();
}

/**
 * Force refresh all formulas
 */
function refreshAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();

  // Touch a cell to force recalculation
  const sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (sheet) {
    const range = sheet.getRange('A1');
    const value = range.getValue();
    range.setValue(value);
  }

  SpreadsheetApp.getUi().alert('Data refreshed successfully!');
}

/**
 * Send email alert for schools below threshold
 */
function sendAlertEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alertSheet = ss.getSheetByName('School_Current');

  if (!alertSheet) {
    SpreadsheetApp.getUi().alert('School_Current sheet not found.');
    return;
  }

  // Get data from School_Current sheet
  const data = alertSheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indices
  const nameCol = headers.indexOf('School_Name');
  const communeCol = headers.indexOf('Commune');
  const attendanceCol = headers.indexOf('Current_Attendance_Rate');
  const feedingCol = headers.indexOf('Current_Feeding_Rate');
  const alertCol = headers.indexOf('Alert_Status');

  // Collect problem schools
  const problemSchools = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][alertCol] === 'ALERT') {
      problemSchools.push({
        name: data[i][nameCol],
        commune: data[i][communeCol],
        attendance: (data[i][attendanceCol] * 100).toFixed(1) + '%',
        feeding: (data[i][feedingCol] * 100).toFixed(1) + '%'
      });
    }
  }

  if (problemSchools.length === 0) {
    SpreadsheetApp.getUi().alert('No schools currently need attention.');
    return;
  }

  // Build email body
  let emailBody = `Haiti School Feeding Program - Alert Report\n`;
  emailBody += `Generated: ${new Date().toLocaleDateString()}\n\n`;
  emailBody += `${problemSchools.length} schools require attention:\n\n`;

  problemSchools.forEach((school, idx) => {
    emailBody += `${idx + 1}. ${school.name} (${school.commune})\n`;
    emailBody += `   Attendance: ${school.attendance}\n`;
    emailBody += `   Feeding Rate: ${school.feeding}\n\n`;
  });

  emailBody += `\nView full dashboard: ${ss.getUrl()}`;

  // Send email
  MailApp.sendEmail({
    to: CONFIG.ALERT_EMAIL,
    subject: `[ALERT] ${problemSchools.length} Schools Need Attention - Haiti Feeding Program`,
    body: emailBody
  });

  SpreadsheetApp.getUi().alert(`Alert email sent to ${CONFIG.ALERT_EMAIL}`);
}

/**
 * Generate PDF report of the Executive Summary
 */
function generatePDFReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Executive Summary sheet not found.');
    return;
  }

  // Create PDF
  const folder = DriveApp.getRootFolder();
  const fileName = `Haiti_Feeding_Report_${Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd')}.pdf`;

  // Export as PDF
  const url = ss.getUrl().replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&format=pdf'
    + '&size=letter'
    + '&portrait=true'
    + '&fitw=true'
    + '&gridlines=false'
    + '&printtitle=false'
    + '&pagenumbers=false'
    + '&fzr=true';

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  const blob = response.getBlob().setName(fileName);
  const file = folder.createFile(blob);

  SpreadsheetApp.getUi().alert(`PDF created: ${file.getUrl()}`);
}

/**
 * Initial setup - creates named ranges and formatting
 */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create named ranges
  const namedRanges = [
    { name: 'SelectedMonth', sheet: 'Executive Summary', range: 'B2' },
    { name: 'AttendanceTarget', sheet: 'Executive Summary', range: 'E2' },
    { name: 'FeedingTarget', sheet: 'Executive Summary', range: 'E3' },
    { name: 'SelectedCommune', sheet: 'Attendance Analysis', range: 'D2' },
    { name: 'SelectedSupervisor', sheet: 'Attendance Analysis', range: 'F2' },
    { name: 'SelectedSchool', sheet: 'School Detail', range: 'B2' }
  ];

  namedRanges.forEach(nr => {
    const sheet = ss.getSheetByName(nr.sheet);
    if (sheet) {
      try {
        ss.setNamedRange(nr.name, sheet.getRange(nr.range));
      } catch (e) {
        // Range might already exist
        console.log(`Could not create named range ${nr.name}: ${e.message}`);
      }
    }
  });

  SpreadsheetApp.getUi().alert('Dashboard setup complete!');
}

/**
 * Create conditional formatting rules for a sheet
 * @param {Sheet} sheet - The sheet to format
 * @param {string} range - The range to apply formatting (e.g., 'C:C')
 */
function applyConditionalFormatting(sheet, range) {
  const rangeObj = sheet.getRange(range);

  // Clear existing rules for this range
  const rules = sheet.getConditionalFormatRules();
  const newRules = rules.filter(rule => {
    const ruleRanges = rule.getRanges();
    return !ruleRanges.some(r => r.getA1Notation() === range);
  });

  // Add new rules
  // Red for < 70%
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.7)
    .setBackground('#EA4335')
    .setFontColor('#FFFFFF')
    .setRanges([rangeObj])
    .build();

  // Yellow for 70-80%
  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.7, 0.8)
    .setBackground('#FBBC04')
    .setRanges([rangeObj])
    .build();

  // Green for > 80%
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.8)
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setRanges([rangeObj])
    .build();

  newRules.push(redRule, yellowRule, greenRule);
  sheet.setConditionalFormatRules(newRules);
}

/**
 * Apply all formatting to dashboard sheets
 */
function applyAllFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Format Executive Summary
  const execSummary = ss.getSheetByName('Executive Summary');
  if (execSummary) {
    applyConditionalFormatting(execSummary, 'B7:B8'); // KPI values
  }

  // Format Attendance Analysis
  const attendance = ss.getSheetByName('Attendance Analysis');
  if (attendance) {
    applyConditionalFormatting(attendance, 'E:E'); // Attendance rates
  }

  // Format Feeding Analysis
  const feeding = ss.getSheetByName('Feeding Rate Analysis');
  if (feeding) {
    applyConditionalFormatting(feeding, 'F:F'); // Feeding rates
  }

  SpreadsheetApp.getUi().alert('Formatting applied to all sheets!');
}

/**
 * Weekly trigger - runs every Monday to send summary
 */
function weeklyReport() {
  sendAlertEmail();
}

/**
 * Set up time-based triggers
 */
function createTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Create weekly trigger (Monday at 8am)
  ScriptApp.newTrigger('weeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert('Weekly report trigger created for Mondays at 8am');
}

/**
 * Create data validation dropdowns
 */
function createDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get unique values for dropdowns
  const presenceSheet = ss.getSheetByName('Présence');
  const infoSheet = ss.getSheetByName('Info sur les écoles');

  if (!presenceSheet || !infoSheet) {
    SpreadsheetApp.getUi().alert('Source sheets not found');
    return;
  }

  // Get unique months
  const dates = presenceSheet.getRange('A:A').getValues().flat()
    .filter(d => d instanceof Date)
    .map(d => Utilities.formatDate(d, 'GMT', 'yyyy-MM'));
  const uniqueMonths = [...new Set(dates)].sort().reverse();

  // Get unique communes
  const communes = infoSheet.getRange('C:C').getValues().flat()
    .filter(c => c && c.toString().trim());
  const uniqueCommunes = ['All', ...new Set(communes)].sort();

  // Get unique schools
  const schools = infoSheet.getRange('B:B').getValues().flat()
    .filter(s => s && s.toString().trim());
  const uniqueSchools = [...new Set(schools)].sort();

  // Get unique supervisors
  const supervisors = infoSheet.getRange('E:E').getValues().flat()
    .filter(s => s && s.toString().trim());
  const uniqueSupervisors = ['All', ...new Set(supervisors)].sort();

  // Apply dropdowns to Executive Summary
  const execSheet = ss.getSheetByName('Executive Summary');
  if (execSheet) {
    const monthRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueMonths.slice(0, 24), true)
      .build();
    execSheet.getRange('B2').setDataValidation(monthRule);
  }

  // Apply dropdowns to Attendance Analysis
  const attSheet = ss.getSheetByName('Attendance Analysis');
  if (attSheet) {
    const monthRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueMonths.slice(0, 24), true)
      .build();
    attSheet.getRange('B2').setDataValidation(monthRule);

    const communeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueCommunes, true)
      .build();
    attSheet.getRange('D2').setDataValidation(communeRule);

    const supervisorRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueSupervisors, true)
      .build();
    attSheet.getRange('F2').setDataValidation(supervisorRule);
  }

  // Apply dropdowns to School Detail
  const schoolSheet = ss.getSheetByName('School Detail');
  if (schoolSheet) {
    const schoolRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueSchools, true)
      .build();
    schoolSheet.getRange('B2').setDataValidation(schoolRule);
  }

  SpreadsheetApp.getUi().alert('Dropdowns created successfully!');
}
