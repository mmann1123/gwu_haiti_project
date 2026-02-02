/**
 * Haiti School Feeding Program - Standalone Dashboard
 *
 * Creates a SEPARATE Google Sheets file for the dashboard.
 *
 * SETUP FOR ADMINISTRATORS:
 * 1. Open your data spreadsheet in Google Sheets
 * 2. Go to Extensions > Apps Script
 * 3. Paste this entire script
 * 4. Save (Ctrl+S)
 * 5. Run "setupDashboardButton" once to add a button to the first sheet
 * 6. Grant permissions when prompted
 * 7. Users can now click the button anytime to generate a new dashboard
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  // Name for the new dashboard file
  DASHBOARD_FILENAME: 'Haiti School Feeding Dashboard',

  // Source sheet names (will try variations if not found)
  PRESENCE_SHEET: 'Présence',
  FEEDING_SHEET_VARIATIONS: [
    'Taux dalimentation',
    "Taux d'alimentation",
    "Taux d\u2019alimentation"  // Unicode apostrophe
  ],
  SCHOOLS_SHEET: 'Info sur les écoles',

  // Thresholds
  ALERT_THRESHOLD: 0.8,

  // Colors
  HEADER_BG: '#4472C4',
  HEADER_TEXT: '#FFFFFF',
  GREEN_BG: '#C6EFCE',
  YELLOW_BG: '#FFEB9C',
  RED_BG: '#FFC7CE',
  LIGHT_BLUE: '#DDEBF7'
};

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Main function - creates a new dashboard file in the SAME FOLDER
 * Simplified for non-technical users - just one click!
 */
function createStandaloneDashboard() {
  const ui = SpreadsheetApp.getUi();
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Show a simple "please wait" message
  const toast = SpreadsheetApp.getActiveSpreadsheet();
  toast.toast('Building your dashboard... This may take a minute.', '⏳ Please Wait', 60);

  try {
    // Step 1: Load data from source
    const data = loadSourceData(sourceSpreadsheet);

    if (!data.presence || data.presence.length === 0) {
      ui.alert('Error', 'Could not load data from Présence sheet.\n\nPlease check that the sheet exists and has data.', ui.ButtonSet.OK);
      return;
    }

    // Warn if feeding data wasn't found
    if (!data.feeding || data.feeding.length === 0) {
      const availableSheets = sourceSpreadsheet.getSheets().map(s => s.getName()).join(', ');
      ui.alert('Warning',
        'Could not load feeding data.\n\n' +
        'The Feeding Analysis section will be empty.\n\n' +
        'Available sheets: ' + availableSheets + '\n\n' +
        'Expected: A sheet containing "Taux" and "alimentation" in the name.',
        ui.ButtonSet.OK);
    }

    // Step 2: Calculate metrics
    const metrics = calculateMetrics(data);

    // Step 3: Create new spreadsheet in the SAME FOLDER
    const dashboard = SpreadsheetApp.create(CONFIG.DASHBOARD_FILENAME + ' - ' + new Date().toLocaleDateString());
    const dashboardFile = DriveApp.getFileById(dashboard.getId());

    // Get the folder of the source spreadsheet and move dashboard there
    const sourceFile = DriveApp.getFileById(sourceSpreadsheet.getId());
    const parents = sourceFile.getParents();
    if (parents.hasNext()) {
      const sourceFolder = parents.next();
      dashboardFile.moveTo(sourceFolder);
    }

    // Step 4: Build dashboard sheets
    buildExecutiveSummary(dashboard, metrics);
    buildAttendanceSheet(dashboard, metrics);
    buildFeedingSheet(dashboard, metrics);
    buildSupervisorSheet(dashboard, metrics);
    buildSchoolList(dashboard, data);

    // Remove default empty sheet
    const defaultSheet = dashboard.getSheetByName('Sheet1');
    if (defaultSheet) {
      dashboard.deleteSheet(defaultSheet);
    }

    // Clear the toast
    toast.toast('Done!', '✓ Complete', 3);

    // Show clickable link
    showClickableLink(dashboard.getName(), dashboard.getUrl());

  } catch (error) {
    ui.alert('Error', 'Something went wrong:\n\n' + error.message + '\n\nPlease contact your administrator.', ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * ADMIN FUNCTION: Run this ONCE to set up the dashboard button
 * This adds a nice button to the first sheet that users can click
 */
function setupDashboardButton() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get the first sheet
  const firstSheet = ss.getSheets()[0];
  const sheetName = firstSheet.getName();

  // Find a good spot for the button (top-right area)
  // Check if there's already a button area
  const buttonCell = firstSheet.getRange('L1:N3');

  // Create the button appearance
  buttonCell.merge();
  buttonCell.setValue('📊 Create Dashboard');
  buttonCell.setBackground('#4285f4');
  buttonCell.setFontColor('#ffffff');
  buttonCell.setFontSize(14);
  buttonCell.setFontWeight('bold');
  buttonCell.setHorizontalAlignment('center');
  buttonCell.setVerticalAlignment('middle');
  buttonCell.setBorder(true, true, true, true, false, false, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Add instructions below
  const instructionCell = firstSheet.getRange('L4:N5');
  instructionCell.merge();
  instructionCell.setValue('👆 To use: Insert > Drawing, create a button,\nthen right-click it > Assign script > createStandaloneDashboard');
  instructionCell.setFontSize(9);
  instructionCell.setFontColor('#666666');
  instructionCell.setWrap(true);

  // Set column widths
  firstSheet.setColumnWidth(12, 60);  // L
  firstSheet.setColumnWidth(13, 60);  // M
  firstSheet.setColumnWidth(14, 60);  // N

  ui.alert('Setup Complete!',
    'A button placeholder has been added to "' + sheetName + '".\n\n' +
    'TO FINISH SETUP:\n' +
    '1. Go to Insert > Drawing\n' +
    '2. Create a button (rectangle with text "Create Dashboard")\n' +
    '3. Click "Save and Close"\n' +
    '4. Position the drawing where you want it\n' +
    '5. Click the drawing, then click the 3 dots (⋮) menu\n' +
    '6. Select "Assign script"\n' +
    '7. Type: createStandaloneDashboard\n' +
    '8. Click OK\n\n' +
    'Users can now click that button anytime to create a dashboard!',
    ui.ButtonSet.OK);
}

/**
 * Show a dialog with a clickable link to the dashboard
 */
function showClickableLink(fileName, url) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial, sans-serif; padding: 10px;">' +
    '<h3 style="color: green;">✓ Dashboard Created!</h3>' +
    '<p><strong>File:</strong> ' + fileName + '</p>' +
    '<p><strong>Location:</strong> Same folder as your data file</p>' +
    '<p style="margin-top: 20px;"><a href="' + url + '" target="_blank" ' +
    'style="background: #4285f4; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">' +
    '📊 Open Dashboard</a></p>' +
    '<p style="margin-top: 15px; font-size: 12px; color: #666;">Or copy this link:<br>' +
    '<input type="text" value="' + url + '" style="width: 100%; padding: 5px; margin-top: 5px;" onclick="this.select()"></p>' +
    '</div>'
  ).setWidth(400).setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Ready!');
}


// ============================================================================
// DATA LOADING
// ============================================================================

/**
 * Load data from source spreadsheet
 */
function loadSourceData(ss) {
  const data = {};

  // Load Présence
  const presenceSheet = ss.getSheetByName(CONFIG.PRESENCE_SHEET);
  if (presenceSheet) {
    const allData = presenceSheet.getDataRange().getValues();
    data.presenceHeaders = allData[0];
    // Skip row 0 (header) and row 1 (often blank)
    data.presence = allData.slice(2).filter(row => row[0] !== '' && row[0] !== null);

    // Map column indices - use index directly if pattern matching fails
    data.pCols = {
      month: 0,  // Le mois is always first column
      schoolName: 1,  // Nom de l'établissement is second column
      schoolId: 4,  // ID de l'établissement
      commune: 5,  // Commune
      department: 6,  // Departement
      supervisor: 7,  // Supervisor
      boysAttendance: 8,    // Garçons (attendance count)
      girlsAttendance: 9,   // Filles (attendance count)
      totalAttendance: 10,  // Présence totale
      boysEnrollment: 11,   // Garçons effectif
      girlsEnrollment: 12,  // Filles effectif
      enrollment: 13,  // Le total Effectif
      attendanceRate: 14,  // Le taux de présence
      variationPct: 17,  // % de variation
      variationCat: 18,  // La catégorie de variation
      variationReason: 19  // La raison de la variation
    };

    // Try to find columns dynamically, fall back to hardcoded if not found
    const monthIdx = findColumn(data.presenceHeaders, ['le mois']);
    if (monthIdx >= 0) data.pCols.month = monthIdx;

    const nameIdx = findColumnStartsWith(data.presenceHeaders, ['nom de']);
    if (nameIdx >= 0) data.pCols.schoolName = nameIdx;

    const idIdx = findColumnStartsWith(data.presenceHeaders, ['id de']);
    if (idIdx >= 0) data.pCols.schoolId = idIdx;

    const rateIdx = findColumn(data.presenceHeaders, ['taux de presence']);
    if (rateIdx >= 0) data.pCols.attendanceRate = rateIdx;

    // Debug: log column mappings
    console.log('Presence columns:', JSON.stringify(data.pCols));
    console.log('Headers sample:', data.presenceHeaders.slice(0, 5));
  }

  // Load Feeding - try multiple sheet name variations
  let feedingSheet = null;
  for (const sheetName of CONFIG.FEEDING_SHEET_VARIATIONS) {
    feedingSheet = ss.getSheetByName(sheetName);
    if (feedingSheet) {
      console.log('Found feeding sheet with name:', sheetName);
      break;
    }
  }

  // Also try to find by partial match if exact match fails
  if (!feedingSheet) {
    const allSheets = ss.getSheets();
    for (const sheet of allSheets) {
      const name = sheet.getName().toLowerCase();
      // Try multiple matching strategies
      if (name.includes('taux') && (name.includes('alimentation') || name.includes('aliment'))) {
        feedingSheet = sheet;
        console.log('Found feeding sheet by partial match:', sheet.getName());
        break;
      }
      // Also check for 'feeding' in English
      if (name.includes('feeding')) {
        feedingSheet = sheet;
        console.log('Found feeding sheet by English name:', sheet.getName());
        break;
      }
    }
  }

  if (feedingSheet) {
    const allData = feedingSheet.getDataRange().getValues();
    data.feedingHeaders = allData[0];
    // Filter more robustly - check first column OR if any key columns have data
    data.feeding = allData.slice(1).filter(row => {
      // Accept row if first column has any value (date, number, string)
      const firstCol = row[0];
      if (firstCol !== '' && firstCol !== null && firstCol !== undefined) return true;
      // Also accept if school name (col 2) or school ID (col 3) has data
      if (row[2] && row[2] !== '') return true;
      if (row[3] && row[3] !== '') return true;
      return false;
    });

    // Hardcoded defaults based on actual column structure
    data.fCols = {
      week: 0,        // Semaine commencant
      month: 1,       // Le mois
      schoolName: 2,  // Nom de l'établissement
      schoolId: 3,    // ID de l'établissement
      commune: 4,     // Commune
      department: 5,  // Departement
      supervisor: 6,  // Supervisor
      daysPlanned: 7, // Le nombre de jours d'alimentation prévu
      daysFed: 9,     // Le nombre réel de jours d'alimentation
      feedingRate: 10, // Le taux d'alimentation
      nonfeedCat1: 11  // 1. La catégorie de non-alimentation
    };

    // Try dynamic detection as fallback
    const fMonthIdx = findColumn(data.feedingHeaders, ['le mois']);
    if (fMonthIdx >= 0) data.fCols.month = fMonthIdx;

    const fNameIdx = findColumnStartsWith(data.feedingHeaders, ['nom de']);
    if (fNameIdx >= 0) data.fCols.schoolName = fNameIdx;

    const fRateIdx = findColumn(data.feedingHeaders, ['taux d']);
    if (fRateIdx >= 0) data.fCols.feedingRate = fRateIdx;

    console.log('Feeding columns:', JSON.stringify(data.fCols));
    console.log('Feeding data rows loaded:', data.feeding.length);
    console.log('Feeding headers:', data.feedingHeaders.slice(0, 12));
  } else {
    console.log('WARNING: Feeding sheet not found! Tried variations:', CONFIG.FEEDING_SHEET_VARIATIONS);
    console.log('Available sheets:', ss.getSheets().map(s => s.getName()));
  }

  // Load Schools
  const schoolsSheet = ss.getSheetByName(CONFIG.SCHOOLS_SHEET);
  if (schoolsSheet) {
    const allData = schoolsSheet.getDataRange().getValues();
    data.schoolsHeaders = allData[1]; // Header in row 2
    data.schools = allData.slice(2).filter(row => row[0] !== '' && row[0] !== null);

    // Schools sheet: columns are Nom, ID, Commune, Departement, Supervisor, GPS, Grades, Ownership
    data.sCols = {
      schoolName: 0,  // Nom de l'Etablissement
      schoolId: 1,    // ID de l'Etablissement
      commune: 2,     // Commune
      department: 3,  // Departement
      supervisor: 4   // Supervisor
    };

    // Try dynamic detection
    const sNameIdx = findColumnStartsWith(data.schoolsHeaders, ['nom']);
    if (sNameIdx >= 0) data.sCols.schoolName = sNameIdx;

    const sIdIdx = findColumnStartsWith(data.schoolsHeaders, ['id']);
    if (sIdIdx >= 0) data.sCols.schoolId = sIdIdx;

    console.log('Schools columns:', JSON.stringify(data.sCols));
    console.log('Schools headers:', data.schoolsHeaders.slice(0, 5));
  }

  return data;
}

/**
 * Normalize string for comparison (handle special apostrophes, accents, etc.)
 */
function normalizeString(str) {
  return String(str || '')
    .toLowerCase()
    .replace(/[\u2018\u2019\u0027\u0060\u00B4]/g, "'")  // All apostrophe variants (Unicode)
    .replace(/[éèêë]/g, 'e')
    .replace(/[àâä]/g, 'a')
    .replace(/[ùûü]/g, 'u')
    .replace(/[îï]/g, 'i')
    .replace(/[ôö]/g, 'o')
    .trim();
}

/**
 * Safely convert value to string (handles null, undefined, NaN)
 */
function safeString(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'number' && isNaN(val)) return '';
  return String(val);
}

/**
 * Find column index by name patterns
 */
function findColumn(headers, patterns) {
  for (let i = 0; i < headers.length; i++) {
    const header = normalizeString(headers[i]);
    for (const pattern of patterns) {
      const normalizedPattern = normalizeString(pattern);
      if (header.includes(normalizedPattern)) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * Find column by checking if header STARTS WITH pattern (more precise)
 */
function findColumnStartsWith(headers, patterns) {
  for (let i = 0; i < headers.length; i++) {
    const header = normalizeString(headers[i]);
    for (const pattern of patterns) {
      const normalizedPattern = normalizeString(pattern);
      if (header.startsWith(normalizedPattern)) {
        return i;
      }
    }
  }
  return -1;
}

// ============================================================================
// METRICS CALCULATION
// ============================================================================

/**
 * Calculate all metrics
 */
function calculateMetrics(data) {
  const metrics = {};
  const p = data.pCols;

  // Get latest month - handle both Date objects and numbers
  const monthValues = data.presence
    .map(row => row[p.month])
    .filter(m => m !== '' && m !== null && m !== undefined);

  // Convert to timestamps for comparison
  const monthTimestamps = monthValues.map(m => {
    if (m instanceof Date) return m.getTime();
    if (typeof m === 'number') return new Date((m - 25569) * 86400000).getTime();
    return 0;
  });

  const maxTimestamp = Math.max(...monthTimestamps);
  metrics.latestMonth = maxTimestamp;
  metrics.latestMonthStr = formatDateFromTimestamp(maxTimestamp);

  // Filter to current month (compare by year-month to handle date variations)
  const latestDate = new Date(maxTimestamp);
  const latestYearMonth = latestDate.getFullYear() * 100 + latestDate.getMonth();

  const current = data.presence.filter(row => {
    const m = row[p.month];
    if (!m) return false;
    let rowDate;
    if (m instanceof Date) {
      rowDate = m;
    } else if (typeof m === 'number') {
      rowDate = new Date((m - 25569) * 86400000);
    } else {
      return false;
    }
    const rowYearMonth = rowDate.getFullYear() * 100 + rowDate.getMonth();
    return rowYearMonth === latestYearMonth;
  });

  // Basic KPIs
  const schoolIds = [...new Set(current.map(row => row[p.schoolId]).filter(id => id))];
  metrics.totalSchools = schoolIds.length;

  const rates = current.map(row => row[p.attendanceRate]).filter(r => typeof r === 'number');
  metrics.avgAttendance = rates.length > 0 ? rates.reduce((a, b) => a + b, 0) / rates.length : 0;

  // Schools below threshold
  const schoolRates = {};
  current.forEach(row => {
    const id = row[p.schoolId];
    const rate = row[p.attendanceRate];
    if (id && typeof rate === 'number') {
      if (!schoolRates[id]) schoolRates[id] = [];
      schoolRates[id].push(rate);
    }
  });
  metrics.schoolsBelow80 = Object.values(schoolRates)
    .filter(arr => arr.reduce((a, b) => a + b, 0) / arr.length < 0.8).length;

  // Feeding rate
  if (data.feeding && data.fCols.feedingRate >= 0) {
    const feedRates = data.feeding
      .map(row => row[data.fCols.feedingRate])
      .filter(r => typeof r === 'number' && !isNaN(r));
    metrics.avgFeeding = feedRates.length > 0 ? feedRates.reduce((a, b) => a + b, 0) / feedRates.length : 0;
  } else {
    metrics.avgFeeding = 0;
  }

  // Attendance trend (by month)
  metrics.attendanceTrend = calcMonthlyTrend(data.presence, p.month, p.attendanceRate, p.schoolId);

  // Commune stats
  metrics.communeStats = calcCommuneStats(current, p);

  // Alerts
  metrics.alerts = calcAlerts(current, p);

  // Variation reasons - use both category and detailed reason
  metrics.variationCategories = countValues(current, p.variationCat);
  metrics.variationReasons = countValues(current, p.variationReason);

  // If detailed reasons exist, use those; otherwise use categories
  if (metrics.variationReasons.length === 0) {
    metrics.variationReasons = metrics.variationCategories;
  }

  // Non-feeding reasons
  metrics.nonfeedReasons = calcNonfeedReasons(data);

  // Supervisor stats
  metrics.supervisorStats = calcSupervisorStats(current, p);

  // Feeding trend
  metrics.feedingTrend = calcFeedingTrend(data);

  // School feeding stats (top/bottom performers)
  metrics.schoolFeedingStats = calcSchoolFeedingStats(data);

  // Gender-disaggregated attendance
  metrics.genderAttendance = calcGenderAttendance(data, p);

  // Month-over-month comparison
  metrics.monthComparison = calcMonthComparison(metrics);

  return metrics;
}

function formatDate(excelDate) {
  if (!excelDate) return '';
  try {
    let d;
    if (excelDate instanceof Date) {
      d = excelDate;
    } else if (typeof excelDate === 'number') {
      d = new Date((excelDate - 25569) * 86400000);
    } else {
      return String(excelDate);
    }
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return months[d.getMonth()] + ' ' + d.getFullYear();
  } catch (e) {
    return String(excelDate);
  }
}

function formatDateFromTimestamp(timestamp) {
  if (!timestamp) return '';
  try {
    const d = new Date(timestamp);
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return months[d.getMonth()] + ' ' + d.getFullYear();
  } catch (e) {
    return '';
  }
}

function getYearMonth(value) {
  if (!value) return null;
  let d;
  if (value instanceof Date) {
    d = value;
  } else if (typeof value === 'number') {
    d = new Date((value - 25569) * 86400000);
  } else {
    return null;
  }
  return d.getFullYear() * 100 + d.getMonth();
}

function calcMonthlyTrend(data, monthCol, rateCol, schoolCol) {
  const byMonth = {};
  data.forEach(row => {
    const m = row[monthCol], r = row[rateCol], s = row[schoolCol];
    const ym = getYearMonth(m);
    if (ym && typeof r === 'number') {
      if (!byMonth[ym]) byMonth[ym] = { monthVal: m, rates: [], schools: new Set() };
      byMonth[ym].rates.push(r);
      if (s) byMonth[ym].schools.add(s);
    }
  });
  return Object.entries(byMonth)
    .map(([ym, d]) => ({
      month: d.monthVal,
      yearMonth: Number(ym),
      rate: d.rates.reduce((a,b)=>a+b,0)/d.rates.length,
      count: d.schools.size
    }))
    .sort((a, b) => b.yearMonth - a.yearMonth)
    .slice(0, 6);
}

function calcCommuneStats(data, p) {
  const byCommune = {};
  data.forEach(row => {
    const c = safeString(row[p.commune]);
    const r = row[p.attendanceRate];
    const s = row[p.schoolId];
    if (c && typeof r === 'number') {
      if (!byCommune[c]) byCommune[c] = { rates: [], schools: new Set() };
      byCommune[c].rates.push(r);
      if (s) byCommune[c].schools.add(s);
    }
  });
  return Object.entries(byCommune)
    .map(([c, d]) => ({ commune: c, rate: d.rates.reduce((a,b)=>a+b,0)/d.rates.length, count: d.schools.size }))
    .sort((a, b) => b.rate - a.rate);
}

function calcAlerts(data, p) {
  const bySchool = {};
  data.forEach(row => {
    const id = row[p.schoolId];
    const name = row[p.schoolName];
    const commune = row[p.commune];
    const sup = row[p.supervisor];
    const rate = row[p.attendanceRate];
    if (id && typeof rate === 'number') {
      if (!bySchool[id]) bySchool[id] = { name: safeString(name), commune: safeString(commune), supervisor: safeString(sup), rates: [] };
      bySchool[id].rates.push(rate);
    }
  });
  return Object.entries(bySchool)
    .map(([id, d]) => ({ name: d.name, commune: d.commune, supervisor: d.supervisor, rate: d.rates.reduce((a,b)=>a+b,0)/d.rates.length }))
    .filter(s => s.rate < CONFIG.ALERT_THRESHOLD)
    .sort((a, b) => a.rate - b.rate);
}

function countValues(data, colIndex) {
  if (colIndex < 0) return [];
  const counts = {};
  data.forEach(row => {
    const v = row[colIndex];
    if (v && v !== '') counts[v] = (counts[v] || 0) + 1;
  });
  return Object.entries(counts).map(([k, v]) => ({ category: k, count: v })).sort((a, b) => b.count - a.count);
}

function calcNonfeedReasons(data) {
  if (!data.feeding || !data.fCols) return [];
  const counts = {};
  const catCol = data.fCols.nonfeedCat1;
  if (catCol < 0) return [];

  data.feeding.forEach(row => {
    const v = row[catCol];
    if (v && v !== '') counts[v] = (counts[v] || 0) + 1;
  });
  return Object.entries(counts).map(([k, v]) => ({ category: k, count: v })).sort((a, b) => b.count - a.count);
}

function calcSupervisorStats(data, p) {
  const bySup = {};
  data.forEach(row => {
    const sup = safeString(row[p.supervisor]);
    const rate = row[p.attendanceRate];
    const school = row[p.schoolId];
    if (sup && typeof rate === 'number') {
      if (!bySup[sup]) bySup[sup] = { rates: [], schools: new Set(), below: new Set() };
      bySup[sup].rates.push(rate);
      if (school) {
        bySup[sup].schools.add(school);
        if (rate < 0.8) bySup[sup].below.add(school);
      }
    }
  });
  return Object.entries(bySup)
    .map(([s, d]) => ({ supervisor: s, rate: d.rates.reduce((a,b)=>a+b,0)/d.rates.length, count: d.schools.size, below: d.below.size }))
    .sort((a, b) => b.rate - a.rate);
}

function calcFeedingTrend(data) {
  if (!data.feeding || !data.fCols) return [];
  const f = data.fCols;
  if (f.month < 0) return [];

  const byMonth = {};
  data.feeding.forEach(row => {
    const m = row[f.month], r = row[f.feedingRate];
    const planned = row[f.daysPlanned] || 0, fed = row[f.daysFed] || 0;
    if (m) {
      if (!byMonth[m]) byMonth[m] = { rates: [], planned: 0, fed: 0 };
      if (typeof r === 'number') byMonth[m].rates.push(r);
      byMonth[m].planned += Number(planned) || 0;
      byMonth[m].fed += Number(fed) || 0;
    }
  });
  return Object.entries(byMonth)
    .map(([m, d]) => ({ month: m, rate: d.rates.length > 0 ? d.rates.reduce((a,b)=>a+b,0)/d.rates.length : 0, planned: d.planned, fed: d.fed }))
    .sort((a, b) => String(b.month).localeCompare(String(a.month)))
    .slice(0, 6);
}

/**
 * Calculate school-level feeding statistics
 */
function calcSchoolFeedingStats(data) {
  if (!data.feeding || !data.fCols) return { top: [], bottom: [] };
  const f = data.fCols;

  const bySchool = {};
  data.feeding.forEach(row => {
    const id = row[f.schoolId];
    const name = safeString(row[f.schoolName]);
    const commune = safeString(row[f.commune]);
    const supervisor = safeString(row[f.supervisor]);
    const rate = row[f.feedingRate];
    const planned = Number(row[f.daysPlanned]) || 0;
    const fed = Number(row[f.daysFed]) || 0;

    if (id && typeof rate === 'number' && !isNaN(rate)) {
      if (!bySchool[id]) {
        bySchool[id] = { name, commune, supervisor, rates: [], totalPlanned: 0, totalFed: 0 };
      }
      bySchool[id].rates.push(rate);
      bySchool[id].totalPlanned += planned;
      bySchool[id].totalFed += fed;
    }
  });

  const schoolList = Object.entries(bySchool)
    .map(([id, d]) => ({
      id,
      name: d.name,
      commune: d.commune,
      supervisor: d.supervisor,
      rate: d.rates.reduce((a, b) => a + b, 0) / d.rates.length,
      daysPlanned: d.totalPlanned,
      daysFed: d.totalFed,
      weekCount: d.rates.length
    }))
    .filter(s => s.weekCount >= 2);  // Only include schools with at least 2 weeks of data

  // Sort by rate descending for top performers
  const sorted = [...schoolList].sort((a, b) => b.rate - a.rate);

  return {
    top: sorted.slice(0, 10),
    bottom: sorted.slice(-10).reverse()  // Reverse so worst is first
  };
}

/**
 * Calculate gender-disaggregated attendance rates by month
 */
function calcGenderAttendance(data, p) {
  if (!data.presence || data.presence.length === 0) return [];

  const byMonth = {};
  data.presence.forEach(row => {
    const m = row[p.month];
    const boysAtt = Number(row[p.boysAttendance]) || 0;
    const girlsAtt = Number(row[p.girlsAttendance]) || 0;
    const boysEnr = Number(row[p.boysEnrollment]) || 0;
    const girlsEnr = Number(row[p.girlsEnrollment]) || 0;

    const ym = getYearMonth(m);
    if (ym && (boysEnr > 0 || girlsEnr > 0)) {
      if (!byMonth[ym]) {
        byMonth[ym] = {
          monthVal: m,
          boysAttTotal: 0,
          girlsAttTotal: 0,
          boysEnrTotal: 0,
          girlsEnrTotal: 0,
          recordCount: 0
        };
      }
      byMonth[ym].boysAttTotal += boysAtt;
      byMonth[ym].girlsAttTotal += girlsAtt;
      byMonth[ym].boysEnrTotal += boysEnr;
      byMonth[ym].girlsEnrTotal += girlsEnr;
      byMonth[ym].recordCount++;
    }
  });

  return Object.entries(byMonth)
    .map(([ym, d]) => {
      const boysRate = d.boysEnrTotal > 0 ? d.boysAttTotal / d.boysEnrTotal : 0;
      const girlsRate = d.girlsEnrTotal > 0 ? d.girlsAttTotal / d.girlsEnrTotal : 0;
      return {
        month: d.monthVal,
        yearMonth: Number(ym),
        boysRate: boysRate,
        girlsRate: girlsRate,
        gap: boysRate - girlsRate,
        boysEnrollment: d.boysEnrTotal,
        girlsEnrollment: d.girlsEnrTotal
      };
    })
    .sort((a, b) => b.yearMonth - a.yearMonth)
    .slice(0, 6);
}

/**
 * Calculate month-over-month comparison metrics
 */
function calcMonthComparison(metrics) {
  const trend = metrics.attendanceTrend;
  if (!trend || trend.length < 2) {
    return null;
  }

  const current = trend[0];  // Most recent month
  const prior = trend[1];    // Previous month

  // Calculate changes
  const attendanceChange = current.rate - prior.rate;
  const schoolsChange = current.count - prior.count;

  // Calculate schools below 80% for prior month (would need to pass more data)
  // For now, we'll use the attendance trend data we have
  return {
    current: {
      month: formatDate(current.month),
      schools: current.count,
      attendanceRate: current.rate
    },
    prior: {
      month: formatDate(prior.month),
      schools: prior.count,
      attendanceRate: prior.rate
    },
    change: {
      schools: schoolsChange,
      schoolsPct: prior.count > 0 ? (schoolsChange / prior.count) : 0,
      attendanceRate: attendanceChange,
      attendanceRatePct: prior.rate > 0 ? (attendanceChange / prior.rate) : 0
    }
  };
}

// ============================================================================
// DASHBOARD BUILDING
// ============================================================================

function formatHeader(sheet, row, startCol, endCol) {
  const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
  range.setBackground(CONFIG.HEADER_BG).setFontColor(CONFIG.HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
}

function getRateColor(rate) {
  if (rate >= 0.8) return CONFIG.GREEN_BG;
  if (rate >= 0.7) return CONFIG.YELLOW_BG;
  return CONFIG.RED_BG;
}

/**
 * Format a change value with +/- prefix
 * @param {number} value - The change value
 * @param {boolean} isPercent - Whether to format as percentage points
 */
function formatChange(value, isPercent) {
  if (value === 0) return '0';
  const prefix = value > 0 ? '+' : '';
  if (isPercent) {
    // Format as percentage points (e.g., +6.2 pp)
    return prefix + (value * 100).toFixed(1) + ' pp';
  }
  return prefix + value;
}

/**
 * Build Executive Summary
 */
function buildExecutiveSummary(ss, metrics) {
  const sheet = ss.insertSheet('Executive Summary');

  // Title
  sheet.getRange('A1').setValue('Haiti School Feeding Program - Executive Summary').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A1:F1').merge();

  // Date
  sheet.getRange('A2').setValue('Generated: ' + new Date().toLocaleString()).setFontStyle('italic');
  sheet.getRange('A3').setValue('Data Month: ' + metrics.latestMonthStr).setFontWeight('bold');

  // KPIs
  sheet.getRange('A5').setValue('Key Performance Indicators').setFontSize(12).setFontWeight('bold');
  sheet.getRange(6, 1, 1, 4).setValues([['Metric', 'Value', 'Target', 'Status']]);
  formatHeader(sheet, 6, 1, 4);

  const kpis = [
    ['Total Schools', metrics.totalSchools, '-', '✓'],
    ['Avg Attendance Rate', metrics.avgAttendance, '90%', metrics.avgAttendance >= 0.9 ? '✓' : metrics.avgAttendance >= 0.8 ? '⚠' : '✗'],
    ['Avg Feeding Rate', metrics.avgFeeding, '90%', metrics.avgFeeding >= 0.9 ? '✓' : metrics.avgFeeding >= 0.8 ? '⚠' : '✗'],
    ['Schools Below 80%', metrics.schoolsBelow80, '0', metrics.schoolsBelow80 === 0 ? '✓' : '⚠']
  ];
  sheet.getRange(7, 1, 4, 4).setValues(kpis);
  sheet.getRange('B8').setNumberFormat('0.0%').setBackground(getRateColor(metrics.avgAttendance));
  sheet.getRange('B9').setNumberFormat('0.0%').setBackground(getRateColor(metrics.avgFeeding));

  // Color code the Status column
  for (let i = 0; i < kpis.length; i++) {
    const status = kpis[i][3];
    const cell = sheet.getRange(7 + i, 4);
    if (status === '✓') {
      cell.setBackground(CONFIG.GREEN_BG);
    } else if (status === '⚠') {
      cell.setBackground(CONFIG.YELLOW_BG);
    } else if (status === '✗') {
      cell.setBackground(CONFIG.RED_BG);
    }
  }

  // Month-over-Month Comparison (new section after KPIs)
  sheet.getRange('A12').setValue('Month-over-Month Comparison').setFontSize(12).setFontWeight('bold');

  if (metrics.monthComparison) {
    const mc = metrics.monthComparison;
    sheet.getRange(13, 1, 1, 4).setValues([['Metric', 'Current Month', 'Prior Month', 'Change']]);
    formatHeader(sheet, 13, 1, 4);

    const compData = [
      ['Schools Reporting', mc.current.schools, mc.prior.schools, formatChange(mc.change.schools, false)],
      ['Avg Attendance Rate', mc.current.attendanceRate, mc.prior.attendanceRate, formatChange(mc.change.attendanceRate, true)]
    ];
    sheet.getRange(14, 1, compData.length, 4).setValues(compData);

    // Format rates as percentages
    sheet.getRange('B15').setNumberFormat('0.0%');
    sheet.getRange('C15').setNumberFormat('0.0%');

    // Color code the change column
    for (let i = 0; i < compData.length; i++) {
      const changeVal = i === 0 ? mc.change.schools : mc.change.attendanceRate;
      const cell = sheet.getRange(14 + i, 4);
      if (changeVal > 0) {
        cell.setBackground(CONFIG.GREEN_BG);
      } else if (changeVal < 0) {
        cell.setBackground(CONFIG.RED_BG);
      } else {
        cell.setBackground(CONFIG.YELLOW_BG);
      }
    }

    // Add month labels
    sheet.getRange('B13').setValue(mc.current.month);
    sheet.getRange('C13').setValue(mc.prior.month);
  } else {
    sheet.getRange('A13').setValue('Not enough data for month comparison (need at least 2 months)');
  }

  // Alerts - moved down to accommodate comparison section
  sheet.getRange('A18').setValue('Schools Needing Attention (< 80%)').setFontSize(12).setFontWeight('bold');
  sheet.getRange(19, 1, 1, 4).setValues([['School', 'Commune', 'Supervisor', 'Attendance']]);
  formatHeader(sheet, 19, 1, 4);

  if (metrics.alerts.length > 0) {
    const alertData = metrics.alerts.slice(0, 10).map(a => [a.name, a.commune, a.supervisor, a.rate]);
    sheet.getRange(20, 1, alertData.length, 4).setValues(alertData);
    sheet.getRange(20, 4, alertData.length, 1).setNumberFormat('0.0%');
    for (let i = 0; i < alertData.length; i++) {
      sheet.getRange(20 + i, 4).setBackground(getRateColor(alertData[i][3]));
    }
  } else {
    sheet.getRange('A20').setValue('No schools below threshold!').setBackground(CONFIG.GREEN_BG);
  }

  // Commune comparison
  sheet.getRange('F5').setValue('Attendance by Commune').setFontSize(12).setFontWeight('bold');
  sheet.getRange(6, 6, 1, 3).setValues([['Commune', 'Avg Rate', 'Schools']]);
  formatHeader(sheet, 6, 6, 8);

  if (metrics.communeStats.length > 0) {
    const communeData = metrics.communeStats.slice(0, 12).map(c => [c.commune, c.rate, c.count]);
    sheet.getRange(7, 6, communeData.length, 3).setValues(communeData);
    sheet.getRange(7, 7, communeData.length, 1).setNumberFormat('0.0%');
    // Color code commune rates
    for (let i = 0; i < communeData.length; i++) {
      sheet.getRange(7 + i, 7).setBackground(getRateColor(communeData[i][1]));
    }
  }

  // Trend
  sheet.getRange('A33').setValue('6-Month Trend').setFontSize(12).setFontWeight('bold');
  sheet.getRange(34, 1, 1, 3).setValues([['Month', 'Attendance', 'Schools']]);
  formatHeader(sheet, 34, 1, 3);

  if (metrics.attendanceTrend.length > 0) {
    const trendData = metrics.attendanceTrend.map(t => [formatDate(t.month), t.rate, t.count]);
    sheet.getRange(35, 1, trendData.length, 3).setValues(trendData);
    sheet.getRange(35, 2, trendData.length, 1).setNumberFormat('0.0%');
    // Color code trend rates
    for (let i = 0; i < trendData.length; i++) {
      sheet.getRange(35 + i, 2).setBackground(getRateColor(trendData[i][1]));
    }
  }

  // Column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 80);
}

/**
 * Build Attendance Analysis
 */
function buildAttendanceSheet(ss, metrics) {
  const sheet = ss.insertSheet('Attendance Analysis');

  sheet.getRange('A1').setValue('Attendance Analysis').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A2').setValue('Data Month: ' + metrics.latestMonthStr).setFontStyle('italic');

  // Trend
  sheet.getRange('A4').setValue('Monthly Attendance Trend').setFontSize(12).setFontWeight('bold');
  sheet.getRange(5, 1, 1, 3).setValues([['Month', 'Avg Attendance', 'Schools']]);
  formatHeader(sheet, 5, 1, 3);

  if (metrics.attendanceTrend.length > 0) {
    const trendData = metrics.attendanceTrend.map(t => [formatDate(t.month), t.rate, t.count]);
    sheet.getRange(6, 1, trendData.length, 3).setValues(trendData);
    sheet.getRange(6, 2, trendData.length, 1).setNumberFormat('0.0%');
    // Color code attendance rates
    for (let i = 0; i < trendData.length; i++) {
      sheet.getRange(6 + i, 2).setBackground(getRateColor(trendData[i][1]));
    }
  } else {
    sheet.getRange('A6').setValue('No trend data available');
  }

  // Gender-Disaggregated Attendance (new section)
  sheet.getRange('E4').setValue('Attendance by Gender').setFontSize(12).setFontWeight('bold');
  sheet.getRange(5, 5, 1, 4).setValues([['Month', 'Boys Rate', 'Girls Rate', 'Gap']]);
  formatHeader(sheet, 5, 5, 8);

  if (metrics.genderAttendance && metrics.genderAttendance.length > 0) {
    const genderData = metrics.genderAttendance.map(g => [
      formatDate(g.month),
      g.boysRate,
      g.girlsRate,
      g.gap
    ]);
    sheet.getRange(6, 5, genderData.length, 4).setValues(genderData);
    sheet.getRange(6, 6, genderData.length, 2).setNumberFormat('0.0%');  // Boys and Girls rates
    sheet.getRange(6, 8, genderData.length, 1).setNumberFormat('+0.0%;-0.0%;0.0%');  // Gap with sign

    // Color code the rates
    for (let i = 0; i < genderData.length; i++) {
      sheet.getRange(6 + i, 6).setBackground(getRateColor(genderData[i][1]));  // Boys rate
      sheet.getRange(6 + i, 7).setBackground(getRateColor(genderData[i][2]));  // Girls rate
      // Color gap: green if small (<2%), yellow if moderate, red if large (>5%)
      const gap = Math.abs(genderData[i][3]);
      if (gap <= 0.02) {
        sheet.getRange(6 + i, 8).setBackground(CONFIG.GREEN_BG);
      } else if (gap <= 0.05) {
        sheet.getRange(6 + i, 8).setBackground(CONFIG.YELLOW_BG);
      } else {
        sheet.getRange(6 + i, 8).setBackground(CONFIG.RED_BG);
      }
    }
  } else {
    sheet.getRange('E6').setValue('No gender data available');
  }

  // Variation categories (Augmentation/Diminution)
  sheet.getRange('A15').setValue('Variation Direction').setFontSize(12).setFontWeight('bold');
  sheet.getRange(16, 1, 1, 2).setValues([['Direction', 'Count']]);
  formatHeader(sheet, 16, 1, 2);

  if (metrics.variationCategories && metrics.variationCategories.length > 0) {
    const catData = metrics.variationCategories.slice(0, 5).map(r => [r.category, r.count]);
    sheet.getRange(17, 1, catData.length, 2).setValues(catData);
  } else {
    sheet.getRange('A17').setValue('No variation categories recorded');
  }

  // Detailed variation reasons
  sheet.getRange('A24').setValue('Variation Reasons (Detailed)').setFontSize(12).setFontWeight('bold');
  sheet.getRange(25, 1, 1, 2).setValues([['Reason', 'Count']]);
  formatHeader(sheet, 25, 1, 2);

  if (metrics.variationReasons && metrics.variationReasons.length > 0) {
    const reasonData = metrics.variationReasons.slice(0, 10).map(r => [r.category, r.count]);
    sheet.getRange(26, 1, reasonData.length, 2).setValues(reasonData);
  } else {
    sheet.getRange('A26').setValue('No detailed variation reasons recorded');
  }

  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(5, 100);  // Month
  sheet.setColumnWidth(6, 100);  // Boys Rate
  sheet.setColumnWidth(7, 100);  // Girls Rate
  sheet.setColumnWidth(8, 80);   // Gap
}

/**
 * Build Feeding Analysis
 */
function buildFeedingSheet(ss, metrics) {
  const sheet = ss.insertSheet('Feeding Analysis');

  sheet.getRange('A1').setValue('Feeding Rate Analysis').setFontSize(14).setFontWeight('bold');

  // Trend
  sheet.getRange('A3').setValue('Monthly Feeding Summary').setFontSize(12).setFontWeight('bold');
  sheet.getRange(4, 1, 1, 5).setValues([['Month', 'Avg Rate', 'Days Planned', 'Days Fed', 'Days Missed']]);
  formatHeader(sheet, 4, 1, 5);

  if (metrics.feedingTrend.length > 0) {
    const trendData = metrics.feedingTrend.map(f => [f.month, f.rate, f.planned, f.fed, f.planned - f.fed]);
    sheet.getRange(5, 1, trendData.length, 5).setValues(trendData);
    sheet.getRange(5, 2, trendData.length, 1).setNumberFormat('0.0%');
    for (let i = 0; i < trendData.length; i++) {
      sheet.getRange(5 + i, 2).setBackground(getRateColor(trendData[i][1]));
    }
  }

  // Top Performers - right side
  sheet.getRange('G3').setValue('🏆 Top Feeding Rate Schools').setFontSize(12).setFontWeight('bold');
  sheet.getRange(4, 7, 1, 5).setValues([['Rank', 'School', 'Commune', 'Feeding Rate', 'Weeks']]);
  formatHeader(sheet, 4, 7, 11);

  const topSchools = metrics.schoolFeedingStats ? metrics.schoolFeedingStats.top : [];
  if (topSchools.length > 0) {
    const topData = topSchools.map((s, i) => [i + 1, s.name, s.commune, s.rate, s.weekCount]);
    sheet.getRange(5, 7, topData.length, 5).setValues(topData);
    sheet.getRange(5, 10, topData.length, 1).setNumberFormat('0.0%');
    for (let i = 0; i < topData.length; i++) {
      sheet.getRange(5 + i, 10).setBackground(getRateColor(topData[i][3]));
    }
  } else {
    sheet.getRange('G5').setValue('No feeding data available');
  }

  // Non-feeding reasons
  sheet.getRange('A14').setValue('Non-Feeding Reasons').setFontSize(12).setFontWeight('bold');
  sheet.getRange(15, 1, 1, 2).setValues([['Category', 'Count']]);
  formatHeader(sheet, 15, 1, 2);

  if (metrics.nonfeedReasons.length > 0) {
    const reasonData = metrics.nonfeedReasons.slice(0, 10).map(r => [r.category, r.count]);
    sheet.getRange(16, 1, reasonData.length, 2).setValues(reasonData);
  } else {
    sheet.getRange('A16').setValue('No non-feeding reasons recorded');
  }

  // Bottom Performers - right side (below top performers)
  sheet.getRange('G18').setValue('⚠️ Schools Needing Attention (Lowest Feeding Rates)').setFontSize(12).setFontWeight('bold');
  sheet.getRange(19, 7, 1, 5).setValues([['Rank', 'School', 'Commune', 'Feeding Rate', 'Weeks']]);
  formatHeader(sheet, 19, 7, 11);

  const bottomSchools = metrics.schoolFeedingStats ? metrics.schoolFeedingStats.bottom : [];
  if (bottomSchools.length > 0) {
    const bottomData = bottomSchools.map((s, i) => [i + 1, s.name, s.commune, s.rate, s.weekCount]);
    sheet.getRange(20, 7, bottomData.length, 5).setValues(bottomData);
    sheet.getRange(20, 10, bottomData.length, 1).setNumberFormat('0.0%');
    for (let i = 0; i < bottomData.length; i++) {
      sheet.getRange(20 + i, 10).setBackground(getRateColor(bottomData[i][3]));
    }
  } else {
    sheet.getRange('G20').setValue('No feeding data available');
  }

  // Column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(7, 50);   // Rank
  sheet.setColumnWidth(8, 250);  // School name
  sheet.setColumnWidth(9, 120);  // Commune
  sheet.setColumnWidth(10, 100); // Feeding rate
  sheet.setColumnWidth(11, 60);  // Weeks
}

/**
 * Build Supervisor Performance
 */
function buildSupervisorSheet(ss, metrics) {
  const sheet = ss.insertSheet('Supervisor Performance');

  sheet.getRange('A1').setValue('Supervisor Performance').setFontSize(14).setFontWeight('bold');

  sheet.getRange('A3').setValue('Supervisor Ranking').setFontSize(12).setFontWeight('bold');
  sheet.getRange(4, 1, 1, 5).setValues([['Rank', 'Supervisor', 'Schools', 'Avg Attendance', 'Schools <80%']]);
  formatHeader(sheet, 4, 1, 5);

  if (metrics.supervisorStats.length > 0) {
    const data = metrics.supervisorStats.map((s, i) => [i + 1, s.supervisor, s.count, s.rate, s.below]);
    sheet.getRange(5, 1, data.length, 5).setValues(data);
    sheet.getRange(5, 4, data.length, 1).setNumberFormat('0.0%');

    for (let i = 0; i < data.length; i++) {
      sheet.getRange(5 + i, 4).setBackground(getRateColor(data[i][3]));
      const below = data[i][4];
      if (below >= 3) sheet.getRange(5 + i, 5).setBackground(CONFIG.RED_BG);
      else if (below >= 2) sheet.getRange(5 + i, 5).setBackground(CONFIG.YELLOW_BG);
      else sheet.getRange(5 + i, 5).setBackground(CONFIG.GREEN_BG);
    }
  }

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
}

/**
 * Build School List
 */
function buildSchoolList(ss, data) {
  const sheet = ss.insertSheet('School List');

  sheet.getRange('A1').setValue('All Schools').setFontSize(14).setFontWeight('bold');

  const headers = ['School Name', 'School ID', 'Commune', 'Department', 'Supervisor'];
  sheet.getRange(3, 1, 1, 5).setValues([headers]);
  formatHeader(sheet, 3, 1, 5);

  if (data.schools && data.schools.length > 0) {
    const s = data.sCols;
    const schoolData = data.schools.map(row => [
      safeString(row[s.schoolName]),
      safeString(row[s.schoolId]),
      safeString(row[s.commune]),
      safeString(row[s.department]),
      safeString(row[s.supervisor])
    ]);
    sheet.getRange(4, 1, schoolData.length, 5).setValues(schoolData);
  }

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 180);
}

/**
 * Add menu when opening - appears automatically for all users
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 DASHBOARD')
    .addItem('🆕 Create New Dashboard', 'createStandaloneDashboard')
    .addSeparator()
    .addItem('⚙️ Admin: Setup Button (run once)', 'setupDashboardButton')
    .addItem('🔍 Debug: Check Sheet Names', 'debugSheetNames')
    .addToUi();
}

/**
 * Debug function to check sheet names and data counts
 */
function debugSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let report = 'SHEET NAMES AND DATA:\n\n';

  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    const rowCount = sheet.getLastRow();
    const colCount = sheet.getLastColumn();
    report += `• "${name}"\n  Rows: ${rowCount}, Columns: ${colCount}\n\n`;
  }

  report += '\n--- EXPECTED SHEETS ---\n';
  report += `Présence: Looking for "${CONFIG.PRESENCE_SHEET}"\n`;
  report += `Feeding: Looking for variations containing "taux" and "alimentation"\n`;
  report += `Schools: Looking for "${CONFIG.SCHOOLS_SHEET}"\n`;

  ui.alert('Debug: Sheet Information', report, ui.ButtonSet.OK);
}

/**
 * Install trigger - run this once if onOpen doesn't work automatically
 */
function installTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  SpreadsheetApp.getUi().alert('Trigger installed! The menu will now appear automatically when anyone opens this spreadsheet.');
}
