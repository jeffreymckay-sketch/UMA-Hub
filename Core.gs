/**
 * -------------------------------------------------------------------
 * CORE HELPERS & SETTINGS
 * -------------------------------------------------------------------
 */

// --- SETTINGS MANAGEMENT ---

function getSettings(key) {
  try {
    const props = PropertiesService.getScriptProperties();
    const data = props.getProperty(key);
    return data ? JSON.parse(data) : {};
  } catch (e) {
    Logger.log('Error getting settings: ' + e.message);
    return {};
  }
}

function saveSettings(key, settingsObj) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(key, JSON.stringify(settingsObj));
    return { success: true, message: 'Settings saved successfully.' };
  } catch (e) {
    return { success: false, message: 'Error saving settings: ' + e.message };
  }
}

function api_getAllSettings() {
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    const settings = {};
    for (const key in props) {
      try { settings[key] = JSON.parse(props[key]); } catch (e) { settings[key] = props[key]; }
    }
    return { success: true, data: settings };
  } catch (e) { return { success: false, message: e.message }; }
}

// --- DATA ACCESS HELPERS (ROBUST) ---

function getMasterDataHub() {
  const adminSettings = getSettings('adminSettings');
  if (!adminSettings || !adminSettings.dataHubUrl) throw new Error('Data Hub URL not set in Admin Settings.');
  const sheetId = extractFileIdFromUrl(adminSettings.dataHubUrl); 
  return SpreadsheetApp.openById(sheetId);
}

function getRequiredSheetData(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log(`Warning: Tab "${tabName}" missing.`); return []; }
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) { Logger.log(`Warning: Tab "${tabName}" is empty.`); return data; }
  return data;
}

function getOrCreateSheet(ss, tabName, headers) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

// --- DATA MAPPING HELPERS ---

function createLookupMap(data, keyHeader, valueHeader) {
  if (!data || data.length <= 1 || !data[0]) return {}; 
  const map = {};
  const headers = data[0].map(normalizeHeader);
  const keyIndex = headers.indexOf(normalizeHeader(keyHeader));
  const valueIndex = headers.indexOf(normalizeHeader(valueHeader));
  
  if (keyIndex === -1 || valueIndex === -1) return {};
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][keyIndex];
    if (key) map[key] = data[i][valueIndex];
  }
  return map;
}

function createDataMap(data, keyHeader) {
  if (!data || data.length <= 1 || !data[0]) return {};
  const map = {};
  const headers = data[0];
  const normalizedHeaders = headers.map(normalizeHeader);
  const keyIndex = normalizedHeaders.indexOf(normalizeHeader(keyHeader));
  
  if (keyIndex === -1) return {};
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][keyIndex];
    if (key) {
      const rowObject = {};
      headers.forEach((header, index) => { rowObject[header] = data[i][index]; });
      map[key] = rowObject;
    }
  }
  return map;
}

// --- UTILITIES ---

function normalizeHeader(h) {
    if (!h) return '';
    h = h.toString().toLowerCase().replace(/\xA0/g, ' ').trim(); 
    return h.replace(/\s+/g, ' '); 
}

/**
 * Extracts the File ID from a Google Drive/Sheet URL.
 * NOW SUPPORTS: Sheets (/d/), Folders (/folders/), and Parameters (id=).
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const str = url.toString();
  
  // 1. Try /d/ID (Sheets, Docs)
  let match = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];

  // 2. Try /folders/ID (Drive Folders)
  match = str.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  
  // 3. Try id=ID parameter
  match = str.match(/id=([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];

  // 4. Fallback: Raw ID (if user pasted just the ID)
  if (str.length > 20 && !str.includes('/')) return str;
  
  return null; // Failed
}

function getImageDataUrl(fileId) {
  try {
    if (!fileId) return { success: false };
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return { success: true, data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  } catch (e) { return { success: false, message: e.message }; }
}

function parseTime(timeStr) {
    try {
        const parts = timeStr.split(':');
        return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
    } catch (e) { return 0; }
}

function timesOverlap(a_start_str, a_end_str, b_start_str, b_end_str) {
    const a_start = parseTime(a_start_str);
    const a_end = parseTime(a_end_str);
    const b_start = parseTime(b_start_str);
    const b_end = parseTime(b_end_str);
    return (a_start < b_end) && (a_end > b_start);
}

function getTimeBlock(timeStr) {
    try {
        const hour = parseInt(timeStr.split(':')[0], 10);
        if (hour < 12) return "Morning"; 
        else if (hour < 17) return "Afternoon"; 
        else return "Evening"; 
    } catch (e) { return "Unknown"; }
}

// --- SHARED API ENDPOINTS ---

function api_getDataHubTabs() {
  try {
    const ss = getMasterDataHub();
    const tabs = ss.getSheets().map(s => s.getName());
    const ignored = ['Settings', 'Staff_List', 'Permissions_Matrix', 'Report_Data', 'Reference'];
    const validTabs = tabs.filter(t => !ignored.some(i => t.includes(i)));
    return { success: true, data: validTabs };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_getTabHeaders(tabName) {
  try {
    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) throw new Error(`Tab "${tabName}" not found.`);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    return { success: true, data: headers.filter(h => h && h.trim() !== "") };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_getEventTypes() {
  const defaults = [{ name: "Nursing", keywords: "NUR" }, { name: "MST", keywords: "BUA, ECO, CIS" }];
  try {
    const ss = getMasterDataHub(); 
    const sheet = ss.getSheetByName("Event_Types"); 
    if (!sheet) return { success: true, data: defaults };
    const data = sheet.getDataRange().getDisplayValues();
    const types = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) types.push({ name: data[i][0], keywords: data[i][1] || "" });
    }
    return { success: true, data: types.length ? types : defaults };
  } catch (e) { return { success: true, data: defaults }; }
}