/**
 * Class counts utility for Google Sheets.
 *
 * This script finds Google Sheets contained in configured Drive folders
 * and extracts a value from each sheet based on the header titled
 * "Current Grade Level".
 *
 * @author Alvaro Gomez, Academic Technology Coach, 210-397-9408
 * @module ClassCounts
 */

/**
 * SCRIPT CONFIGURATION
 * Central object to hold settings for each school level.
 * @type {{ES: {folderId: string, sheetName: string},
 * MS: {folderId: string, sheetName: string}, HS: {folderId: string
 * sheetName: string}}}
 */
const CONFIGS = {
  ES: {
    folderId: "1HHFNXX2Xcn57HHERowLlDfTjj4dQ20LE",
    sheetName: "ES",
  },
  MS: {
    folderId: "1s0SF6LBg14hU03wwcTSyVWT4kVwSA4CH",
    sheetName: "MS",
  },
  HS: {
    folderId: "1vv3_YMUom8ynJJpiyNKKKMXF8ZfRIQz8",
    sheetName: "HS",
  },
};

/**
 * Read campus count data from a source sheet using spreadsheet ID mapping.
 * 
 * This function extracts spreadsheet IDs and count values from a specified
 * source sheet (ES, MS, or HS), looks up the campus name from the mapping,
 * and returns them as an array of objects.
 * 
 * @param {string} sheetName - Name of the sheet (ES, MS, or HS)
 * @param {Map} campusMap - Map of campus names to spreadsheet IDs
 * @returns {Array<{campus: string, count: number}>} Array of campus-count pairs
 * @example
 * const esData = readSourceSheetData("ES", elementarySchoolCampusMap);
 * // Returns: [{campus: "Bernal #1", count: 17}, ...]
 */
function readSourceSheetData(sheetName, campusMap) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  // Return empty array if sheet doesn't exist (log warning)
  if (!sheet) {
    Logger.log(`Warning: ${sheetName} sheet not found, skipping`);
    return [];
  }

  // Get last row with data
  const lastRow = sheet.getLastRow();

  // Return empty array if no data rows exist (lastRow < 2)
  if (lastRow < 2) {
    Logger.log(`Warning: No data rows found in ${sheetName} sheet`);
    return [];
  }

  // Read spreadsheet IDs from column A (rows 2 to lastRow) in single batch operation
  const idRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const idValues = idRange.getValues();

  // Read count values from column D (rows 2 to lastRow) in single batch operation
  const countRange = sheet.getRange(2, 4, lastRow - 1, 1);
  const countValues = countRange.getValues();

  // Create reverse lookup map: spreadsheet ID -> campus name
  const idToCampusMap = new Map();
  for (const [campusName, spreadsheetId] of campusMap) {
    if (spreadsheetId) {
      idToCampusMap.set(spreadsheetId, campusName);
    }
  }

  // Build array of objects with campus and count properties
  const results = [];
  for (let i = 0; i < idValues.length; i++) {
    const spreadsheetId = idValues[i][0];
    const countValue = countValues[i][0];

    // Skip if spreadsheet ID is empty
    if (!spreadsheetId || spreadsheetId.toString().trim() === '') {
      continue;
    }

    // Look up campus name from mapping
    const campusName = idToCampusMap.get(spreadsheetId.toString());
    
    // Skip if no mapping found for this spreadsheet ID
    if (!campusName) {
      Logger.log(`Warning: No campus mapping found for spreadsheet ID ${spreadsheetId} in ${sheetName}`);
      continue;
    }

    // Convert empty or invalid count values to 0
    let count = 0;
    if (countValue !== null && countValue !== undefined && countValue !== '') {
      const parsedCount = Number(countValue);
      if (!isNaN(parsedCount)) {
        count = parsedCount;
      } else {
        Logger.log(`Warning: Invalid count value in ${sheetName} row ${i + 2}, defaulting to 0`);
      }
    }

    // Add to results array
    results.push({
      campus: campusName,
      count: count
    });
  }

  return results;
}

/**
 * Aggregate campus counts from ES, MS, HS sheets into ALE Counts sheet.
 * 
 * This function reads campus names from the ALE Counts sheet, processes
 * count data from the specified source sheet(s) using spreadsheet ID
 * mappings, aggregates the counts, and writes the totals back to the ALE Counts sheet.
 * 
 * @param {string} [level] - Optional level to aggregate ('ES', 'MS', 'HS'). If omitted, aggregates all levels.
 * @returns {number} Number of campuses updated with non-zero counts
 * @example
 * const updatedCount = aggregateCampusCounts('ES'); // Only elementary
 * const updatedCount = aggregateCampusCounts(); // All levels
 */
function aggregateCampusCounts(level) {
  // Get active spreadsheet reference
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get ALE Counts sheet reference
  const aleCountsSheet = spreadsheet.getSheetByName("ALE Counts");
  
  // Display error and return if ALE Counts sheet doesn't exist
  if (!aleCountsSheet) {
    const errorMessage = "Error: ALE Counts sheet not found";
    Logger.log(errorMessage);
    SpreadsheetApp.getUi().alert(errorMessage);
    return 0;
  }
  
  // Define row ranges for each level in ALE Counts sheet
  const rowRanges = {
    ES: { start: 4, end: 113 },   // E4:E113 (110 rows)
    MS: { start: 116, end: 160 }, // E116:E160 (45 rows) + E208:E209 (2 rows)
    HS: { start: 163, end: 206 }  // E163:E206 (44 rows) + E210:E212 (3 rows)
  };
  
  // Determine which rows to process based on level
  let rowsToProcess = [];
  if (level === 'ES') {
    rowsToProcess = [{ start: 4, end: 113 }];
  } else if (level === 'MS') {
    rowsToProcess = [
      { start: 116, end: 160 },
      { start: 208, end: 209 }
    ];
  } else if (level === 'HS') {
    rowsToProcess = [
      { start: 163, end: 206 },
      { start: 210, end: 212 }
    ];
  } else {
    // If no level specified, process all rows
    rowsToProcess = [{ start: 4, end: 212 }];
  }
  
  // Build campus lookup map with exact campus names as keys
  const campusLookup = {};
  const campusValuesByRange = [];
  
  // Read campus names from column D for each range
  for (const range of rowsToProcess) {
    const numRows = range.end - range.start + 1;
    const campusRange = aleCountsSheet.getRange(range.start, 4, numRows, 1);
    const campusValues = campusRange.getValues();
    
    campusValuesByRange.push({
      range: range,
      values: campusValues
    });
    
    for (let i = 0; i < campusValues.length; i++) {
      const campusName = campusValues[i][0];
      if (campusName && campusName.toString().trim() !== '') {
        campusLookup[campusName.toString().trim()] = {
          row: i + range.start, // Actual row number in sheet
          originalName: campusName.toString(),
          totalCount: 0,
          foundInSource: false
        };
      }
    }
  }
  
  // Track campuses with missing counts
  const missingCampuses = [];
  
  // Get the appropriate campus map based on level
  let campusMap;
  if (level === 'ES') {
    campusMap = elementarySchoolCampusMap;
  } else if (level === 'MS') {
    campusMap = middleSchoolCampusMap;
  } else if (level === 'HS') {
    campusMap = highSchoolCampusMap;
  }
  
  // Process only the specified level, or all levels if no level specified
  if (!level || level === 'ES') {
    const esData = readSourceSheetData("ES", elementarySchoolCampusMap);
    for (const record of esData) {
      if (campusLookup[record.campus]) {
        campusLookup[record.campus].totalCount += record.count;
        campusLookup[record.campus].foundInSource = true;
      } else {
        Logger.log(`Warning: Campus "${record.campus}" from ES sheet not found in ALE Counts`);
      }
    }
  }
  
  if (!level || level === 'MS') {
    const msData = readSourceSheetData("MS", middleSchoolCampusMap);
    for (const record of msData) {
      if (campusLookup[record.campus]) {
        campusLookup[record.campus].totalCount += record.count;
        campusLookup[record.campus].foundInSource = true;
      } else {
        Logger.log(`Warning: Campus "${record.campus}" from MS sheet not found in ALE Counts`);
      }
    }
  }
  
  if (!level || level === 'HS') {
    const hsData = readSourceSheetData("HS", highSchoolCampusMap);
    for (const record of hsData) {
      if (campusLookup[record.campus]) {
        campusLookup[record.campus].totalCount += record.count;
        campusLookup[record.campus].foundInSource = true;
      } else {
        Logger.log(`Warning: Campus "${record.campus}" from HS sheet not found in ALE Counts`);
      }
    }
  }
  
  // If a specific level was requested, check which campuses from the map are missing
  if (level && campusMap) {
    for (const [campusName] of campusMap) {
      if (campusLookup[campusName] && !campusLookup[campusName].foundInSource) {
        missingCampuses.push(campusName);
      }
    }
  }
  
  // Build output arrays and write to each range
  for (const rangeData of campusValuesByRange) {
    const outputArray = [];
    for (let i = 0; i < rangeData.values.length; i++) {
      const campusName = rangeData.values[i][0];
      if (campusName && campusName.toString().trim() !== '') {
        const count = campusLookup[campusName.toString().trim()] ? campusLookup[campusName.toString().trim()].totalCount : 0;
        outputArray.push([count]);
      } else {
        outputArray.push([0]);
      }
    }
    
    // Write counts to column E for this specific row range
    const numRows = rangeData.range.end - rangeData.range.start + 1;
    const outputRange = aleCountsSheet.getRange(rangeData.range.start, 5, numRows, 1);
    outputRange.setValues(outputArray);
  }
  
  // Count number of campuses with non-zero values
  let campusesUpdated = 0;
  for (const campusName in campusLookup) {
    if (campusLookup[campusName].totalCount > 0) {
      campusesUpdated++;
    }
  }
  
  // Display success message with count of updated campuses
  let successMessage = `${campusesUpdated} ${level}'s Total Enrolled values in the ALE Counts sheet were updated.`;
  
  // Add missing campuses information if any
  if (missingCampuses.length > 0) {
    successMessage += `\n\nCampuses with missing counts (${missingCampuses.length}):\n${missingCampuses.join(', ')}\n\nA value of 0 was added for those campuses.`;
  }
  
  Logger.log(successMessage);
  SpreadsheetApp.getUi().alert(successMessage);
  
  // Return number of campuses updated
  return campusesUpdated;
}

/**
 * Install the custom menu into the active spreadsheet UI.
 * This creates a top-level "Update Counts" menu with sub-menus for
 * Elementary, Middle, and High school actions.
 *
 * @returns {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Update Counts');

  // Elementary School Sub-Menu
  const esMenu = ui.createMenu('Elementary School');
  esMenu.addItem('1. Get Spreadsheet IDs', 'getESSpreadsheetIds');
  esMenu.addItem('2. Get Counts', 'extractESGradeLevelValue');
  menu.addSubMenu(esMenu);

  // Middle School Sub-Menu
  const msMenu = ui.createMenu('Middle School');
  msMenu.addItem('1. Get Spreadsheet IDs', 'getMSSpreadsheetIds');
  msMenu.addItem('2. Get Counts', 'extractMSGradeLevelValue');
  menu.addSubMenu(msMenu);

  // High School Sub-Menu
  const hsMenu = ui.createMenu('High School');
  hsMenu.addItem('1. Get Spreadsheet IDs', 'getHSSpreadsheetIds');
  hsMenu.addItem('2. Get Counts', 'extractHSGradeLevelValue');
  menu.addSubMenu(hsMenu);

  menu.addToUi();
}

/**
 * WRAPPER FUNCTIONS (Called by the menu)
 * These short functions pass the correct configuration to
 * the main functions.
 */

/**
 * Get spreadsheet IDs for Elementary School folder and write
 * them to the configured sheet.
 * @returns {void}
 */
function getESSpreadsheetIds() { getSpreadsheetIdsFromFolder(CONFIGS.ES); }

/**
 * Get spreadsheet IDs for Middle School folder and write
 * them to the configured sheet.
 * @returns {void}
 */
function getMSSpreadsheetIds() { getSpreadsheetIdsFromFolder(CONFIGS.MS); }

/**
 * Get spreadsheet IDs for High School folder and write them
 * to the configured sheet.
 * @returns {void}
 */
function getHSSpreadsheetIds() { getSpreadsheetIdsFromFolder(CONFIGS.HS); }

/**
 * Extract grade-level value for Elementary School spreadsheets
 * listed in the sheet.
 * @returns {void}
 */
function extractESGradeLevelValue() { 
  extractGradeLevelValue(CONFIGS.ES);
  aggregateCampusCounts('ES');
}

/**
 * Extract grade-level value for Middle School spreadsheets
 * listed in the sheet.
 * @returns {void}
 */
function extractMSGradeLevelValue() { 
  extractGradeLevelValue(CONFIGS.MS);
  aggregateCampusCounts('MS');
}

/**
 * Extract grade-level value for High School spreadsheets
 * listed in the sheet.
 * @returns {void}
 */
function extractHSGradeLevelValue() { 
  extractGradeLevelValue(CONFIGS.HS);
  aggregateCampusCounts('HS');
}

/**
 * MAIN GENERIC FUNCTIONS
 * These perform the core logic and are reusable for any level.
 */

/**
 * Generic: scan a Drive folder and write the IDs and file names
 * of Google Sheets to the configured sheet in the active spreadsheet.
 *
 * @param {{folderId: string, sheetName: string}} levelConfig
 * Configuration object for a level (from `CONFIGS`).
 * @throws {Error} If the configured sheet is not found or the
 * Drive folder cannot be accessed.
 * @returns {void}
 * @example
 * Called by the UI wrappers above, e.g. getESSpreadsheetIds()
 * getSpreadsheetIdsFromFolder(CONFIGS.ES);
 */
function getSpreadsheetIdsFromFolder(levelConfig) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(levelConfig.sheetName);

  if (!sheet) {
    const errorMessage = `Error: A sheet named "${levelConfig.sheetName}" could not be found.`;
    Logger.log(errorMessage);
    SpreadsheetApp.getUi().alert(errorMessage);
    return;
  }

  sheet.clearContents();
  sheet.getRange('A1:D1').setValues([['Spreadsheet ID', 'Campus', 'Error Log', 'Count']]);

  const parentFolderId = levelConfig.folderId;

  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const files = parentFolder.getFiles();
    const spreadsheetData = [];

    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        spreadsheetData.push([file.getId(), file.getName()]);
      }
    }

    // Sort by file name (column B) ascending before writing to the sheet
    if (spreadsheetData.length > 0) {
      spreadsheetData.sort(function(a, b) {
        const nameA = (a[1] || '').toString().toLowerCase();
        const nameB = (b[1] || '').toString().toLowerCase();
        if (nameA < nameB) return -1;
        if (nameA > nameB) return 1;
        return 0;
      });

      sheet.getRange(2, 1, spreadsheetData.length, 2).setValues(spreadsheetData);
    }

    SpreadsheetApp.getUi().alert(`Script finished. IDs and file names have been logged in the "${levelConfig.sheetName}" sheet.`);
  } catch (e) {
    const errorMessage = `An unexpected error occurred. Please check the folder ID for ${levelConfig.sheetName}. Error: ${e.message}`;
    Logger.log(errorMessage);
    SpreadsheetApp.getUi().alert(errorMessage);
  }
}

/**
 * Generic: open each spreadsheet ID listed in the configured sheet
 * and search for the header "Current Grade Level". When found,
 * the function writes the value appearing after the first empty
 * cell below that column into column D of the IDs sheet. Any
 * errors are written to column C.
 *
 * Behavior notes:
 * - If the extracted value is the literal header text 'Current
 *   Grade Level', it is normalized to 0 before writing.
 * - If no ID rows are found in the sheet (less than 2 rows),
 *   the function exits quietly.
 *
 * @param {{folderId: string, sheetName: string}} levelConfig
 * Configuration object for a level (from `CONFIGS`).
 * @returns {void}
 * @throws {Error} If a target spreadsheet cannot be opened or
 * required cells are missing; individual errors are
 * captured per-row and written into the sheet rather than propagated.
 */
function extractGradeLevelValue(levelConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idsSheet = ss.getSheetByName(levelConfig.sheetName);

  if (!idsSheet) {
    SpreadsheetApp.getUi().alert(`Error: A sheet named "${levelConfig.sheetName}" was not found.`);
    return;
  }

  const lastRow = idsSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No spreadsheet IDs found in column A to process.");
    return;
  }
  const idValues = idsSheet.getRange(`A2:A${lastRow}`).getValues();

  idValues.forEach((row, index) => {
    const id = row[0];
    const currentRowInSheet = index + 2;

    if (!id) return;

    try {
      const targetSpreadsheet = SpreadsheetApp.openById(id);
      let valueFound = false;
      let extractedValue = null;
      const sheets = targetSpreadsheet.getSheets();

      for (const sheet of sheets) {
        const data = sheet.getDataRange().getValues();
        if (data.length < 3) continue;

        const headerRow = data[1];
        const columnIndex = headerRow.findIndex(header => header.toString().trim() === "Current Grade Level");

        if (columnIndex !== -1) {
          let foundEmptyCell = false;
          for (let r = 2; r < data.length; r++) {
            const cellValue = data[r][columnIndex];
            if (!cellValue || cellValue.toString().trim() === '') {
              foundEmptyCell = true;
            } else if (foundEmptyCell) {
              extractedValue = cellValue;
              valueFound = true;
              break;
            }
          }
        }
        if (valueFound) break;
      }

      if (valueFound) {
        // If the extracted value is the literal header text 'Current Grade Level', treat it as 0.
        const normalized = extractedValue && extractedValue.toString().trim().toLowerCase();
        const valueToWrite = (normalized === 'current grade level') ? 0 : extractedValue;
        idsSheet.getRange(currentRowInSheet, 4).setValue(valueToWrite);
        idsSheet.getRange(currentRowInSheet, 3).clearContent();
      } else {
        throw new Error("Could not find 'Current Grade Level' header or target value.");
      }
    } catch (e) {
      // Log the error for debugging
      Logger.log(`Error processing row ${currentRowInSheet}, ID ${id}: ${e.message}`);
      
      // Write a user-friendly error message to the sheet
      let errorMessage = e.message;
      if (errorMessage.includes("permission") || errorMessage.includes("access")) {
        errorMessage = "No permission to access this spreadsheet";
      }
      
      idsSheet.getRange(currentRowInSheet, 3).setValue(errorMessage);
      idsSheet.getRange(currentRowInSheet, 4).clearContent();
    }
  });
}