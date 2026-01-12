/**
 * @file Core.gs
 * @description This file contains the core, foundational functions that are 
 * used across the entire application. It is the bedrock of the app.
 */


/**
 * Gets the master data hub spreadsheet.
 *
 * This is a critical gatekeeper function. All access to the master data
 * must go through here. It retrieves the masterSheetId from the settings
 * managed by Config.gs.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The master spreadsheet object.
 * @throws {Error} If the masterSheetId is not set or the spreadsheet cannot be opened.
 */
function getMasterDataHub() {
  try {
    const settings = getSettings();
    const masterSheetId = settings.masterSheetId;

    if (!masterSheetId) {
      throw new Error('CRITICAL: masterSheetId is not defined in script properties. The application cannot load.');
    }

    const spreadsheet = SpreadsheetApp.openById(masterSheetId);
    return spreadsheet;

  } catch (e) {
    console.error('Failed to open Master Data Hub: ' + e.message);
    // Propagate the error to the calling function so it can be handled gracefully
    // and reported to the user.
    throw e;
  }
}

/**
 * Retrieves a specific sheet (tab) from the master spreadsheet by its key name.
 * This function acts as a safeguard against hard-coded sheet names. It uses the
 * 'sheetTabs' property (a JSON string) stored in PropertiesService to look up
 * the actual sheet name.
 *
 * @param {string} sheetKey The key corresponding to the sheet (e.g., 'Staff_List', 'TechHub_Shifts').
 *        This key is the cleaned version of the sheet name stored during setup.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The requested sheet object.
 * @throws {Error} If the sheet key is not found, the sheetTabs property is missing,
 *         or the sheet itself does not exist in the spreadsheet.
 */
function getSheet(sheetKey) {
  try {
    const settings = getSettings();
    const sheetTabsJSON = settings.sheetTabs;

    if (!sheetTabsJSON) {
      throw new Error("The 'sheetTabs' property is missing from script properties. Please run the setup process.");
    }

    const sheetTabs = JSON.parse(sheetTabsJSON);
    const sheetName = sheetTabs[sheetKey];

    if (!sheetName) {
      throw new Error(`The sheet key "${sheetKey}" was not found in the stored sheet tabs. Please re-run the setup or check the key name.`);
    }

    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`The sheet named "${sheetName}" (referenced by key "${sheetKey}") could not be found in the master spreadsheet.`);
    }

    return sheet;

  } catch (e) {
    console.error(`Error in getSheet('${sheetKey}'): ${e.toString()}`);
    // Propagate the error to be handled by the calling function.
    throw e;
  }
}
