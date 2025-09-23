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
function extractESGradeLevelValue() { extractGradeLevelValue(CONFIGS.ES); }

/**
 * Extract grade-level value for Middle School spreadsheets
 * listed in the sheet.
 * @returns {void}
 */
function extractMSGradeLevelValue() { extractGradeLevelValue(CONFIGS.MS); }

/**
 * Extract grade-level value for High School spreadsheets
 * listed in the sheet.
 * @returns {void}
 */
function extractHSGradeLevelValue() { extractGradeLevelValue(CONFIGS.HS); }

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
      idsSheet.getRange(currentRowInSheet, 3).setValue(e.message);
      idsSheet.getRange(currentRowInSheet, 4).clearContent();
    }
  });
}