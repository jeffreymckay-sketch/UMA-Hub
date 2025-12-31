
/**
 * -------------------------------------------------------------------
 * CORE UTILITIES & SETTINGS
 * -------------------------------------------------------------------
 */

/**
 * Gets the Master Spreadsheet instance.
 * RELIABLE METHOD: Uses the ID from Config.gs.
 */
function getMasterDataHub() {
  try {
    if (!CONFIG || !CONFIG.MASTER_SHEET_ID) {
      throw new Error("MASTER_SHEET_ID is not defined in the CONFIG object.");
    }
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    if (ss) {
      return ss;
    } else {
      throw new Error("Could not open spreadsheet with the provided MASTER_SHEET_ID.");
    }
  } catch (e) {
    console.error("Connection Error: " + e.message);
    throw new Error("System Error: Could not connect to Data Hub. " + e.message);
  }
}

function api_getImageDataUrl(fileId) {
    try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const contentType = blob.getContentType();
        const base64data = Utilities.base64Encode(blob.getBytes());
        const dataUrl = `data:${contentType};base64,${base64data}`;
        return { success: true, data: dataUrl };
    } catch (e) {
        return { success: false, message: e.message };
    }
}


/**
 * Fetches data from a tab safely.
 * STABLE METHOD: Uses getDataRange() to ensure all data is captured.
 */
function getRequiredSheetData(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    console.warn(`Tab "${tabName}" missing.`);
    return []; 
  }
  
  if (sheet.getLastRow() === 0) return [];
  return sheet.getDataRange().getValues();
}

/**
 * Helper to get a sheet or create it if missing.
 */
function getOrCreateSheet(ss, tabName) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }
  return sheet;
}

/**
 * Helper to normalize header names.
 */
function normalizeHeader(header) {
  return String(header).toLowerCase().replace(/[\s_]/g, '');
}

/**
 * Helper to extract ID from a Google Sheet URL
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * -------------------------------------------------------------------
 * SETTINGS MANAGEMENT
 * -------------------------------------------------------------------
 */

function getSettings(key) {
  try {
    const props = PropertiesService.getScriptProperties();
    const data = props.getProperty(key);
    return data ? JSON.parse(data) : {};
  } catch (e) {
    console.error("Error loading settings: " + e.message);
    return {};
  }
}

function saveSettings(key, data) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(key, JSON.stringify(data));
    return { success: true, message: "Settings saved." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getAllSettings() {
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    const result = {};
    for (const key in props) {
      try {
        result[key] = JSON.parse(props[key]);
      } catch (e) {
        result[key] = props[key];
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * -------------------------------------------------------------------
 * METADATA HELPERS (Required for Template Builder)
 * -------------------------------------------------------------------
 */

function api_getDataHubTabs() {
  try {
    const ss = getMasterDataHub();
    const sheets = ss.getSheets();
    const names = sheets.map(s => s.getName());
    return { success: true, data: names };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_getTabHeaders(tabName) {
  try {
    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return { success: false, message: "Tab not found" };
    
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return { success: true, data: [] };
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    // Filter out empty headers
    const validHeaders = headers.filter(h => h !== "");
    
    return { success: true, data: validHeaders };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
