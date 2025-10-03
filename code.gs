
const CONFIG = {
  MASTER_SHEET_NAME: "attendance 2025-2026",
  MAIN_MONTH_HEADERS_ROW: 7, // Row containing "August", "Sep", "Oct", "Nov"
  DATE_AND_STUDENT_HEADERS_ROW: 8, // Row containing "W", "Date", "Student Name", "10", "24", etc.

  NAME_COL: "Student Name",
  OTF_ID_COL: "OTF ID",
  ATTENDANCE_COL: "Attendance",
  DATE_COLUMNS: ["August", "Sep", "Oct", "Nov"],

  ZOOM_NAME_HEADER: "Name",
  ZOOM_DURATION_HEADER: "Duration",
  MINUTES_THRESHOLD_FULL: 25, // Minutes required for a score of 1.0
  MINUTES_THRESHOLD_HALF: 5   // Minutes required for a score of 0.5
};

function normalizeIdentifier(identifier, isId = false) {
  if (!identifier) return "";

  let normalized = identifier.toString().trim().toLowerCase();

  normalized = normalized.replace(/\s+/g, " ");

  if (isId) {
    // Map OTF/otf/OT → ot
    normalized = normalized.replace(/^otf/i, "ot");
    normalized = normalized.replace(/^ot id/i, "ot");
    normalized = normalized.replace(/^ot\s+/i, "ot");
    normalized = normalized.replace(/[^a-z0-9\-]/g, ""); // Cleanup junk
  }

  return normalized;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("OTF Attendance Helper")
    .addItem("Process Zoom Report", "showSidebar")
    .addToUi();
}
function showSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("OTF Attendance Processor");
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getConfig() {
  return CONFIG;
}

function getDateOptions() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error(`Sheet "${CONFIG.MASTER_SHEET_NAME}" not found.`);

    const mainMonthHeaders = masterSheet.getRange(CONFIG.MAIN_MONTH_HEADERS_ROW, 1, 1, masterSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    const dateSubHeaders = masterSheet.getRange(CONFIG.DATE_AND_STUDENT_HEADERS_ROW, 1, 1, masterSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());

    Logger.log('Main Month Headers (Row ' + CONFIG.MAIN_MONTH_HEADERS_ROW + '): ' + JSON.stringify(mainMonthHeaders));
    Logger.log('Date Sub-headers (Row ' + CONFIG.DATE_AND_STUDENT_HEADERS_ROW + '): ' + JSON.stringify(dateSubHeaders));

    const dateOptions = {};
    let lastKnownMonthHeaderCol = -1; // Track column to search for next month

    CONFIG.DATE_COLUMNS.forEach((month) => {
      const monthHeaderColIdx = mainMonthHeaders.indexOf(month, lastKnownMonthHeaderCol + 1);

      if (monthHeaderColIdx === -1) {
        Logger.log(`Month "${month}" not found in main headers.`);
        return;
      }

      const dates = [];
      let currentScanColIdx = monthHeaderColIdx;
      // Determine the column range for the current month's dates
      let endOfCurrentMonthDataColIdx = mainMonthHeaders.length; // Default to end of headers
      for (let i = monthHeaderColIdx + 1; i < mainMonthHeaders.length; i++) {
        if (mainMonthHeaders[i] !== "") { // Next month header found
          endOfCurrentMonthDataColIdx = i;
          break;
        }
      }

      while (currentScanColIdx < endOfCurrentMonthDataColIdx && currentScanColIdx < dateSubHeaders.length) {
        const subHeaderValue = dateSubHeaders[currentScanColIdx];

        if (subHeaderValue && subHeaderValue !== "" && !isNaN(parseFloat(subHeaderValue))) {
          dates.push(subHeaderValue);
        } else if (subHeaderValue && subHeaderValue !== "") {
          // Stop if a non-date, non-empty header is found within the month's columns
          break;
        }
        currentScanColIdx++;
      }

      if (dates.length > 0) {
        dateOptions[month] = dates;
        Logger.log(`Dates found for "${month}": ${JSON.stringify(dates)}`);
      } else {
        Logger.log(`No valid dates found under month "${month}".`);
      }

      lastKnownMonthHeaderCol = monthHeaderColIdx;
    });

    Logger.log('Final Date options: ' + JSON.stringify(dateOptions));
    return dateOptions;
  } catch (error) {
    Logger.log(`Error in getDateOptions: ${error.message}`);
    throw new Error(`Failed to fetch date options: ${error.message}`);
  }
}

// --- Core Processing Logic ---

// Main function to process Zoom attendance data
function processAttendance(zoomCsvFileId, targetMonth, targetDate) {
  try {
    // Initialize statistics object for detailed reporting
    const stats = {
      totalStudentsInMaster: 0,
      totalZoomEntriesProcessed: 0, // Total Zoom entries considered
      studentsPresent: 0,
      studentsAbsent: 0,
      partialAttendance: 0,
      unmatchedEntries: 0,
      processingTime: Date.now(),
      successfulUpdates: 0,
      attendance: {
        full: 0,    // Students scoring 1.0
        partial: 0, // Students scoring 0.5
        absent: 0   // Students scoring 0.0
      },
      details: {
        presentStudentNames: [],
        absentStudentNames: [],
        partialStudentNames: [],
        unmatchedParticipantNames: []
      }
    };

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);

    if (!masterSheet) {
      throw new Error(`Master sheet "${CONFIG.MASTER_SHEET_NAME}" not found.`);
    }

    const masterData = readMasterSheet(masterSheet);
    stats.totalStudentsInMaster = masterData.data.length;

    const zoomData = readZoomCsv(zoomCsvFileId);
    const attendanceResults = calculateAttendance(zoomData); // Aggregated results from Zoom
    stats.totalZoomEntriesProcessed = Object.keys(attendanceResults).length; // Count unique aggregated entries

    const updateResult = updateMasterSheet(masterSheet, masterData, attendanceResults, targetMonth, targetDate, stats);

    stats.processingTime = Date.now() - stats.processingTime;

    const messageParts = formatAttendanceStats(stats, targetMonth, targetDate);

    const formattedMessage =
      `${messageParts.summary}\n\n` +
      `Attendance:\n` +
      `  - Present: ${messageParts.attendance.full}\n` +
      `  - Partial: ${messageParts.attendance.partial}\n` +
      `  - Absent: ${messageParts.attendance.absent}\n\n` +
      `Processing Info:\n` +
      `  - Time: ${messageParts.processing.time}\n` +
      `  - Updates: ${messageParts.processing.updates}\n` +
      `  - Unmatched: ${messageParts.processing.unmatched}`;

    return {
      success: true,
      message: formattedMessage, // The consolidated string message
      stats: stats // The detailed stats object for analytics
    };
  } catch (error) {
    Logger.log(`[PROCESS ATTENDANCE ERROR] ${error.message}`);
    return {
      success: false,
      message: ` Error: ${error.message}`, // Error message as a string
      stats: null
    };
  }
}
function formatAttendanceStats(stats, targetMonth, targetDate) {
  // Calculate percentages safely, avoiding division by zero
  const presentPercentage = stats.totalStudentsInMaster > 0 ? ((stats.attendance.full / stats.totalStudentsInMaster) * 100).toFixed(1) : '0.0';
  const partialPercentage = stats.totalStudentsInMaster > 0 ? ((stats.attendance.partial / stats.totalStudentsInMaster) * 100).toFixed(1) : '0.0';
  const absentPercentage = stats.totalStudentsInMaster > 0 ? ((stats.attendance.absent / stats.totalStudentsInMaster) * 100).toFixed(1) : '0.0';

  return {
    summary: `Processed ${stats.totalStudentsInMaster} students for ${targetMonth} ${targetDate}`,
    attendance: {
      full: `${stats.attendance.full} students present (${presentPercentage}%)`,
      partial: `${stats.attendance.partial} students partial (${partialPercentage}%)`,
      absent: `${stats.attendance.absent} students absent (${absentPercentage}%)`
    },
    processing: {
      time: `${(stats.processingTime / 1000).toFixed(2)} seconds`,
      updates: `${stats.successfulUpdates} records updated`,
      unmatched: `${stats.unmatchedEntries} unmatched entries`
    }
  };
}

// --- Data Reading Functions ---


function readMasterSheet(sheet) {
  Logger.log('[READ MASTER SHEET] Starting.');
  const range = sheet.getDataRange();
  const values = range.getValues();

  if (values.length < CONFIG.DATE_AND_STUDENT_HEADERS_ROW + 1) {
    throw new Error("Master sheet has insufficient rows for student data.");
  }

  
  const studentDataHeaders = values[CONFIG.DATE_AND_STUDENT_HEADERS_ROW - 1].map(h => h.toString().trim());
  const monthHeaders = values[CONFIG.MAIN_MONTH_HEADERS_ROW - 1].map(h => h.toString().trim());

  const nameColIdx = studentDataHeaders.indexOf(CONFIG.NAME_COL);
  const otfIdColIdx = studentDataHeaders.indexOf(CONFIG.OTF_ID_COL);

  if (nameColIdx === -1) {
    throw new Error(`Master sheet missing required column: "${CONFIG.NAME_COL}".`);
  }
  if (otfIdColIdx === -1) {
    Logger.log(`[READ MASTER SHEET WARNING] Master sheet missing optional column: "${CONFIG.OTF_ID_COL}".`);
  }
  Logger.log(`[READ MASTER SHEET] Name Col Index: ${nameColIdx}, OTF ID Col Index: ${otfIdColIdx}.`);

  return {
    monthHeaders: monthHeaders,
    studentDataHeaders: studentDataHeaders,
    nameColIdx: nameColIdx,
    otfIdColIdx: otfIdColIdx,
    data: values.slice(CONFIG.DATE_AND_STUDENT_HEADERS_ROW) // Slice to get only student data rows
  };
}

function readZoomCsv(fileId) {
  Logger.log(`[READ ZOOM CSV] Attempting to read file ID: ${fileId}`);
  const file = DriveApp.getFileById(fileId);
  const csvData = file.getBlob().getDataAsString();
  let rows = csvData.split("\n").map(row => row.split(","));

  if (rows[0] && rows[0][0] && rows[0][0].startsWith('\ufeff')) {
    rows[0][0] = rows[0][0].replace('\ufeff', '');
  }
  const validRows = rows.filter(row => row.length >= 2);
  if (validRows.length < 1) {
    throw new Error("Zoom CSV file is empty or has invalid format.");
  }
  Logger.log(`[READ ZOOM CSV] Successfully read ${validRows.length} valid rows from CSV.`);
  return validRows;
}

// --- Attendance Calculation Logic ---

function calculateAttendance(zoomData) {
  Logger.log('[CALCULATE ATTENDANCE] Starting calculation.');
  const headers = zoomData[0].map(h => h.toString().trim());
  Logger.log(`[CALCULATE ATTENDANCE] Zoom CSV Headers: ${JSON.stringify(headers)}`);

  let nameHeaderIdx = headers.findIndex(h => h.toLowerCase().includes(CONFIG.ZOOM_NAME_HEADER.toLowerCase()));
  let durationHeaderIdx = headers.findIndex(h => h.toLowerCase().includes(CONFIG.ZOOM_DURATION_HEADER.toLowerCase()));

  if (nameHeaderIdx === -1) {
    Logger.log('[CALCULATE ATTENDANCE WARNING] Zoom Name header not found, assuming column 0.');
    nameHeaderIdx = 0;
  }
  if (durationHeaderIdx === -1) {
    Logger.log('[CALCULATE ATTENDANCE WARNING] Zoom Duration header not found, assuming column 1.');
    durationHeaderIdx = 1;
  }
  Logger.log(`[CALCULATE ATTENDANCE] Using Name/ID Column Index: ${nameHeaderIdx}, Duration Column Index: ${durationHeaderIdx}`);

  const canonicalAttendanceMap = {};

  for (let i = 1; i < zoomData.length; i++) {
    const row = zoomData[i];
    let identifierRaw = row[nameHeaderIdx] ? row[nameHeaderIdx].toString().trim() : "";
    const duration = row[durationHeaderIdx] ? row[durationHeaderIdx].toString().trim() : "";

    if (!identifierRaw || !duration) {
      Logger.log(`[CALCULATE ATTENDANCE WARNING] Skipping row ${i + 1}: missing identifier or duration.`);
      continue;
    }

    let nameExtracted = identifierRaw; // Start with raw identifier as potential name
    let otfIdExtracted = null;

    // Regex to extract OTF ID and clean up the student name
    const otfIdRegex = /(OTF?[- ]?\d{2}[- ]?\d{3,4}|OT ID \s*\(?\d{2}-?\d{3,4}\)?|OT \s*\d{2}\s*\d{3,4})/i;
    const match = identifierRaw.match(otfIdRegex);

    if (match) {
      otfIdExtracted = match[0].replace(/[\(\)\[\]\{\}]/g, '').replace(/\s/g, '-').toUpperCase().trim();
      nameExtracted = identifierRaw.replace(match[0], '').trim(); // Remove ID from the string

      nameExtracted = nameExtracted
        .replace(/--|\s+\|\s*|[\({]$/i, '') // Remove trailing separators
        .replace(/\(.*\)/g, '')          // Remove any parenthesized text from name
        .trim();

      if (!nameExtracted || nameExtracted.match(/^OTF?[- ]?\d/i)) {
        nameExtracted = identifierRaw; // Use the original identifier as name if extraction failed badly
      }
      Logger.log(`[CALCULATE ATTENDANCE] Extracted from "${identifierRaw}": Name="${nameExtracted}", OTF ID="${otfIdExtracted}"`);
    } else {
      Logger.log(`[CALCULATE ATTENDANCE] No explicit OTF ID pattern found in "${identifierRaw}". Treating as Name/Identifier.`);
    }

    // Parse duration into minutes
    const minutes = parseDurationToMinutes(duration);
    if (minutes === 0 && duration !== "0:00" && duration !== "0") {
      Logger.log(`[CALCULATE ATTENDANCE WARNING] Could not parse duration "${duration}" for "${identifierRaw}". Parsed minutes: ${minutes}.`);
      continue;
    }

    let canonicalKey = null;
    if (otfIdExtracted && otfIdExtracted.length > 5) {
      canonicalKey = normalizeIdentifier(otfIdExtracted, true);
    } else if (nameExtracted) {
      canonicalKey = normalizeIdentifier(nameExtracted, false);
    } else {
      Logger.log(`[CALCULATE ATTENDANCE WARNING] No canonical key derivable for "${identifierRaw}". Skipping.`);
      continue;
    }

    canonicalAttendanceMap[canonicalKey] = (canonicalAttendanceMap[canonicalKey] || 0) + minutes;
    Logger.log(`[CALCULATE ATTENDANCE] Aggregated ${minutes} mins for key "${canonicalKey}". Total: ${canonicalAttendanceMap[canonicalKey]}.`);
  }

  const results = {};
  Object.keys(canonicalAttendanceMap).forEach(key => {
    results[key] = {
      totalTime: canonicalAttendanceMap[key],
      score: calculateScore(canonicalAttendanceMap[key])
    };
  });
  Logger.log(`[CALCULATE ATTENDANCE] Final attendance map created for ${Object.keys(results).length} unique keys.`);
  return results;
}

function parseDurationToMinutes(durationStr) {
  if (!durationStr) return 0;
  const parts = durationStr.split(":").map(Number);
  try {
    if (parts.length === 3) { // HH:MM:SS
      return parts[0] * 60 + parts[1] + parts[2] / 60;
    } else if (parts.length === 2) { // MM:SS
      return parts[0] + parts[1] / 60;
    } else if (parts.length === 1) { // Just minutes
      return parseFloat(durationStr) || 0;
    }
  } catch (e) {
    Logger.log(`[PARSE DURATION ERROR] Failed parsing "${durationStr}": ${e.message}`);
  }
  return 0; // Return 0 if parsing fails
}

// Calculates attendance score based on total minutes
function calculateScore(totalMinutes) {
  if (totalMinutes >= CONFIG.MINUTES_THRESHOLD_FULL) return 1.0;
  if (totalMinutes >= CONFIG.MINUTES_THRESHOLD_HALF) return 0.5;
  return 0.0;
}

function updateMasterSheet(sheet, masterData, attendanceResults, targetMonth, targetDate, stats) {
  Logger.log(`[UPDATE MASTER SHEET] Starting update for ${targetMonth} ${targetDate}.`);
  let updatedCount = 0;
  let skippedCount = 0;
  const nameColIdx = masterData.nameColIdx;
  const otfIdColIdx = masterData.otfIdColIdx;

  const mainMonthHeaders = masterData.monthHeaders;
  const dateSubHeaders = masterData.studentDataHeaders;

  const monthStartColIdxInSheet = mainMonthHeaders.indexOf(targetMonth);
  if (monthStartColIdxInSheet === -1) {
    throw new Error(`Target month "${targetMonth}" not found in master sheet's month headers.`);
  }
  Logger.log(`[UPDATE MASTER SHEET] Target month "${targetMonth}" found at column index ${monthStartColIdxInSheet}.`);

  let targetColIdxInSheet = -1;
  let searchLimitForDates = mainMonthHeaders.length; // Limit search to the end of month headers
  const currentMonthIndex = CONFIG.DATE_COLUMNS.indexOf(targetMonth);
  if (currentMonthIndex + 1 < CONFIG.DATE_COLUMNS.length) {
    const nextMonth = CONFIG.DATE_COLUMNS[currentMonthIndex + 1];
    const nextMonthStartColIdx = mainMonthHeaders.indexOf(nextMonth, monthStartColIdxInSheet + 1);
    if (nextMonthStartColIdx !== -1) {
      searchLimitForDates = nextMonthStartColIdx; // Limit search up to the start of the next month
    }
  }

  for (let i = monthStartColIdxInSheet; i < searchLimitForDates && i < dateSubHeaders.length; i++) {
    if (dateSubHeaders[i] && dateSubHeaders[i].toString().trim() === targetDate.toString().trim()) {
      targetColIdxInSheet = i;
      break;
    }
  }

  if (targetColIdxInSheet === -1) {
    throw new Error(`Target date "${targetDate}" not found under month "${targetMonth}" in the master sheet's date headers.`);
  }
  Logger.log(`[UPDATE MASTER SHEET] Target date "${targetDate}" found at column index ${targetColIdxInSheet}.`);

  masterData.data.forEach((row, rowIndex) => {
    const studentName = row[nameColIdx] ? row[nameColIdx].toString().trim() : "";
    const otfId = otfIdColIdx !== -1 && row[otfIdColIdx] ? row[otfIdColIdx].toString().trim() : null;

    if (!studentName && !otfId) {
      return; // Skip if no name or OTF ID is present
    }

    let matchKey = null;
    const studentNameNorm = studentName ? normalizeIdentifier(studentName, false) : null;
    const otfIdNorm = otfId ? normalizeIdentifier(otfId, true) : null;

    // Try OTF ID match first
    if (otfIdNorm) {
      matchKey = Object.keys(attendanceResults).find(
        key => key === otfIdNorm
      );
    }

    // If not matched by ID, fallback to name
    if (!matchKey && studentNameNorm) {
      matchKey = Object.keys(attendanceResults).find(
        key => key === studentNameNorm
      );
    }

    // If a match is found in Zoom results
    if (matchKey) {
      const score = attendanceResults[matchKey].score;
      const sheetRowToUpdate = rowIndex + CONFIG.DATE_AND_STUDENT_HEADERS_ROW + 1; // Row number in the sheet
      const sheetColToUpdate = targetColIdxInSheet + 1; // Column number in the sheet

      // Check current cell value to avoid unnecessary writes
      const currentCellValue = sheet.getRange(sheetRowToUpdate, sheetColToUpdate).getValue();
      if (currentCellValue !== score) {
        sheet.getRange(sheetRowToUpdate, sheetColToUpdate).setValue(score);
        updatedCount++;
        // Update overall stats based on the score
        if (score === 1.0) stats.attendance.full++;
        else if (score === 0.5) stats.attendance.partial++;
        else stats.attendance.absent++; // Score 0.0
      } else {
        // If score is the same, it's still considered "accounted for" in terms of attendance status
        if (score === 1.0) stats.attendance.full++;
        else if (score === 0.5) stats.attendance.partial++;
        else stats.attendance.absent++;
      }
    } else {
      // No match found in Zoom results
      skippedCount++;
      stats.unmatchedEntries++; // Increment unmatched count
      if (studentName) stats.details.unmatchedParticipantNames.push(`${studentName} (OTF ID: ${otfId || 'N/A'})`);
      else if (otfId) stats.details.unmatchedParticipantNames.push(`(OTF ID: ${otfId})`);
    }
  });

  stats.successfulUpdates = updatedCount;
  stats.unmatchedEntries = skippedCount;

  Logger.log(`[UPDATE MASTER SHEET] Finished. Updated: ${updatedCount}, Skipped (Unmatched): ${skippedCount}.`);
  return { updatedCount, skippedCount };
}

function onApiLoad() {
  console.log('Google API loaded. Picker library loading...');
  loadPickerLibrary();
}

// Loads the Google Picker library
function loadPickerLibrary() {
  try {
    gapi.load('picker', {
      'callback': () => {
        console.log('Google Picker library loaded successfully.');
        pickerApiLoaded = true; // Set flag in the client-side context
        document.getElementById('browseBtn').disabled = false;
        updateStatus('ready', 'Select a Zoom CSV file to process.');
      },
      'onerror': (error) => {
        console.error('Error loading Google Picker library:', error);
        handleApiError('Failed to load Google Picker library. Please check Cloud Console configuration.');
      },
      'timeout': 10000, // Timeout after 10 seconds
      'ontimeout': () => {
        console.warn('Google Picker library load timed out.');
        handleApiError('Google Picker API timed out. Please refresh and try again.');
      }
    });
  } catch (e) {
    console.error("Error during Picker API load:", e);
    handleApiError("An unexpected error occurred while loading the Picker API.");
  }
}

// Request configuration needed for the Google Picker
function getFilePicker() {
  const token = ScriptApp.getOAuthToken();
  const config = CONFIG; // Assuming CONFIG is globally available or passed appropriately

  const CLIENT_ID = 'YOUR_GOOGLE_API_CLIENT_ID'; // e.g., '1234567890-abcdefghijklmnopqrstuvwxyz.apps.googleusercontent.com'
  const DEVELOPER_KEY = 'YOUR_GOOGLE_API_DEVELOPER_KEY'; // e.g., 'AIzaSy...'
  const API_KEY = 'YOUR_GOOGLE_API_KEY'; // Usually same as CLIENT_ID for Apps Script related things, but can be different.

  if (CLIENT_ID === 'YOUR_GOOGLE_API_CLIENT_ID' || DEVELOPER_KEY === 'YOUR_GOOGLE_API_DEVELOPER_KEY' || API_KEY === 'YOUR_GOOGLE_API_KEY') {
      throw new Error("Google Picker API keys are not configured. Please set CLIENT_ID, DEVELOPER_KEY, and API_KEY in PickerConfig.gs or directly in this file.");
  }

  return {
    appId: CLIENT_ID,
    token: token,
    developerKey: DEVELOPER_KEY,
    apiKey: API_KEY,
    config: CONFIG // Pass config for other settings if needed
  };
}