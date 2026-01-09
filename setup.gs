/**
 * @file setup.gs
 * @description This file contains the initial setup functions for the application.
 */

/**
 * ===================================================================
 *  DEVELOPER SETUP (RUN FROM THE SCRIPT EDITOR)
 * ===================================================================
 */

/**
 * Use this function to perform a complete setup from the script editor.
 */
function runFullSetup() {
  // --- USER ACTION REQUIRED ---
  // 1. Replace the placeholder URL below with your actual Master Google Sheet URL.
  const masterSheetUrl = 'https://docs.google.com/spreadsheets/d/1J0bMQamssoKD9OFO5HLVampKYWhWUfaljlUY3O--7us/edit';
  // --------------------------

  if (masterSheetUrl.includes('PASTE_YOUR')) {
    throw new Error('Please edit the `runFullSetup` function in `setup.gs` and replace the placeholder URL with your actual Google Sheet URL.');
  }

  const sheetId = _extractSheetIdFromUrl(masterSheetUrl);
  if (!sheetId) {
    throw new Error('Could not extract a valid Sheet ID from the URL provided.');
  }

  // Save the core setting
  saveSetting('masterSheetId', sheetId);
  
  // Sync the sheet names
  _syncSheetTabsToProperties(sheetId);

  // You can add other setup steps here if needed.
  SpreadsheetApp.getUi().alert('Full setup complete! The Master Sheet ID and tab names have been saved.');
}


/**
 * Extracts the Google Sheet ID from a given URL.
 * @param {string} url The Google Sheet URL.
 * @returns {string|null} The Sheet ID or null if not found.
 */
function _extractSheetIdFromUrl(url) {
  const regex = /spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

/**
 * Synchronizes all sheet (tab) names from the master spreadsheet to PropertiesService.
 * This is a developer utility and should be run during setup.
 * @param {string} sheetId The ID of the master Google Sheet.
 */
function _syncSheetTabsToProperties(sheetId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets();
    
    const sheetNames = {};
    
    sheets.forEach(sheet => {
      const originalName = sheet.getName();
      // Clean the name to create a valid key (e.g., "Staff List" -> "Staff_List")
      const cleanedName = originalName.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
      sheetNames[cleanedName] = originalName;
    });
    
    // Save the entire object to a single property
    saveSetting('sheetTabs', JSON.stringify(sheetNames));
    
    console.log('Sheet tabs successfully synchronized to PropertiesService.');
    console.log(JSON.stringify(sheetNames, null, 2));

  } catch (e) {
    console.error('Error in _syncSheetTabsToProperties: ' + e.toString());
    throw new Error('Could not synchronize sheet tabs. Check the Sheet ID and permissions. ' + e.message);
  }
}


/**
 * ===================================================================
 *  END-USER SETUP (RUN FROM A GOOGLE SHEET)
 * ===================================================================
 */


/**
 * Creates a custom menu in the spreadsheet to run setup.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Admin Setup')
    .addItem('Set Master Sheet URL', '_promptForMasterSheet')
    .addToUi();
}

/**
 * Prompts the user for the Master Sheet URL. Intended to be run from a Sheet menu.
 */
function _promptForMasterSheet() {
  // (Implementation remains the same as before)
}
