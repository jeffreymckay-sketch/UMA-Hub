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
    const masterSheetId = getSpreadsheetId(); // Fetches from Config.gs (secure)

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

/**
 * SECURITY: Fetches the current user's role from the Staff_List sheet.
 * This is used for server-side permission checking.
 * 
 * @returns {string[]} An array of roles assigned to the user (e.g., ['Admin', 'MST']).
 */
function getCurrentUserRoles() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const staffSheet = getSheet('Staff_List');
    if (!staffSheet) return [];

    const data = staffSheet.getDataRange().getValues();
    const headers = getColumnMap(data[0]);
    
    // Find the row for this user
    const userRow = data.find(row => String(row[headers.staffid]).toLowerCase() === userEmail);
    
    if (!userRow) return []; // Not found, no roles

    const rolesStr = userRow[headers.roles] || "";
    // Split by comma and trim
    return rolesStr.split(',').map(r => r.trim());

  } catch (e) {
    console.error("Error fetching user roles: " + e.message);
    return [];
  }
}

/**
 * SECURITY: Enforces access control. Throws an error if the user lacks the required role.
 * This should be called at the start of any sensitive function.
 * 
 * @param {string|string[]} requiredRole A single role string or an array of allowed roles.
 * @throws {Error} If the user does not have permission.
 */
function requireRole(requiredRole) {
  const userRoles = getCurrentUserRoles();
  const allowedRoles = Array.isArray(requiredRole) ? requiredRole : [requiredRole];

  // Check if user has ANY of the allowed roles
  const hasPermission = allowedRoles.some(role => userRoles.includes(role));

  if (!hasPermission) {
    const userEmail = Session.getActiveUser().getEmail();
    console.warn(`Security Alert: User ${userEmail} attempted unauthorized action requiring ${allowedRoles.join(' or ')}.`);
    throw new Error("Access Denied: You do not have permission to perform this action.");
  }
}