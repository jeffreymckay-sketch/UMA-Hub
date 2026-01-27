/**
 * @file Config.gs
 * @description This file contains the functions for managing application settings
 * using PropertiesService. It is the single source of truth for all configuration.
 */

// This is the only place the spreadsheet ID should be defined.
const MASTER_SPREADSHEET_ID = '1J0bMQamssoKD9OFO5HLVampKYWhWUfaljlUY3O--7us'; 

/**
 * Retrieves the master spreadsheet ID.
 * @returns {string} The spreadsheet ID.
 */
function getSpreadsheetId() {
    return MASTER_SPREADSHEET_ID;
}

/**
 * Retrieves all application settings from PropertiesService.
 *
 * @returns {object} The settings object. Returns a blank object if no settings are found.
 */
function getSettings() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    return properties;
  } catch (e) {
    console.error('Error retrieving settings from PropertiesService: ' + e.message);
    // Return an empty object to prevent downstream errors
    return {};
  }
}

/**
 * Saves a settings object to PropertiesService.
 * This will overwrite all existing properties.
 *
 * @param {object} settings The settings object to save.
 * @returns {void}
 */
function saveSettings(settings) {
  try {
    PropertiesService.getScriptProperties().setProperties(settings, true); // true to delete other properties
  } catch (e) {
    console.error('Error saving settings to PropertiesService: ' + e.message);
    throw new Error('Failed to save settings. ' + e.message);
  }
}

/**
 * Saves a single setting (key-value pair) to PropertiesService.
 *
 * @param {string} key The key for the setting.
 * @param {string} value The value for the setting.
 * @returns {void}
 */
function saveSetting(key, value) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
  } catch (e) {
    console.error(`Error saving setting '${key}' to PropertiesService: ${e.message}`);
    throw new Error(`Failed to save setting '${key}'. ${e.message}`);
  }
}

/**
 * ===================================================================
 *  ADMINISTRATIVE & SETUP FUNCTIONS
 * ===================================================================
 */

/**
 * Inspects the Master Spreadsheet and saves all tab names to the 'sheetTabs'
 * property in PropertiesService. This synchronizes the app's configuration
 * with the actual spreadsheet structure.
 *
 * TO RUN: Select 'admin_mapSheetTabs' from the function list in the
 * Apps Script editor and click "Run".
 */
function admin_mapSheetTabs() {
  try {
    const ss = getMasterDataHub();
    const sheets = ss.getSheets();
    const sheetMap = {};

    sheets.forEach(sheet => {
      const actualName = sheet.getName();
      // Create a standardized key (e.g., "Staff List - 2024" becomes "Staff_List_2024")
      const keyName = actualName.trim().replace(/\s+-\s+|\s+/g, '_');
      sheetMap[keyName] = actualName;
    });

    const sheetTabsJSON = JSON.stringify(sheetMap, null, 2);
    saveSetting('sheetTabs', sheetTabsJSON);

    console.log("SUCCESS: The 'sheetTabs' configuration has been updated.");
    console.log(sheetTabsJSON);
    return { success: true, message: "Sheet tabs mapped successfully!", data: sheetTabsJSON };

  } catch (e) {
    console.error("FAILED to map sheet tabs: " + e.message);
    throw e;
  }
}
