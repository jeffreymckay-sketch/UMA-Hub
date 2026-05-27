/**
 * -------------------------------------------------------------------
 * CLASSROOM MANAGEMENT CONTROLLER
 * Handles: Course Lookup (View Builder) & Bulk Editing
 * -------------------------------------------------------------------
 */

/**
 * Utility: Extract Google Sheet ID from URL
 * Supports both full URLs and direct IDs
 */
function extractFileIdFromUrl(input) {
  if (!input) return null;
  
  const trimmed = input.trim();
  
  // If it's already just an ID (no slashes), return it
  if (!trimmed.includes('/') && !trimmed.includes('http')) {
    return trimmed;
  }
  
  // Extract from URL patterns
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,  // Standard Sheets URL
    /\/folders\/([a-zA-Z0-9-_]+)/,           // Folder URL
    /id=([a-zA-Z0-9-_]+)/,                   // Query parameter
    /^([a-zA-Z0-9-_]+)$/                     // Direct ID
  ];
  
  for (let pattern of patterns) {
    const match = trimmed.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return null;
}

/**
 * Helper: Resolves the Target Sheet
 * STRICT MODE: Requires a valid URL.
 */
function resolveClassroomTarget(customUrl, customTab) {
  console.log("Resolving Target. URL:", customUrl, "Tab:", customTab); // Debug Log

  // 1. Validate Input
  if (!customUrl || customUrl.trim() === "") {
    throw new Error("Please paste a Google Sheet URL in the box.");
  }

  // 2. Extract ID
  const id = extractFileIdFromUrl(customUrl);
  if (!id) {
    throw new Error("Could not find a valid Sheet ID in that URL. Please check the link.");
  }

  // 3. Resolve Tab Name (Default to Sheet1 if empty)
  const tab = customTab ? customTab.toString().trim() : "Sheet1";
  
  return { id, tab, name: "Custom Sheet" };
}

/**
 * LOOKUP: Fetch Headers (Step 1)
 */
function api_lookup_getHeaders(customUrl, customTab) {
  try {
    const target = resolveClassroomTarget(customUrl, customTab);
    const ss = SpreadsheetApp.openById(target.id);
    const sheet = ss.getSheetByName(target.tab);
    
    if (!sheet) {
      const available = ss.getSheets().map(s => s.getName()).join(", ");
      throw new Error(`Tab "${target.tab}" not found. Available tabs: [${available}]`);
    }
    
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) throw new Error("Sheet is empty.");
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    return { success: true, headers: headers, sourceName: target.name };
    
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * LOOKUP: Fetch Data (Step 2)
 */
function api_lookup_getData(customUrl, customTab) {
  try {
    const target = resolveClassroomTarget(customUrl, customTab);
    const ss = SpreadsheetApp.openById(target.id);
    const sheet = ss.getSheetByName(target.tab);
    
    if (!sheet) throw new Error(`Tab "${target.tab}" not found.`);
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) throw new Error("Sheet has no data rows.");
    
    return { success: true, headers: data[0], rows: data.slice(1) };
    
  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * BULK EDITOR: Fetch Data with Row Indices
 */
function api_fetchSheetData(customUrl, customTab) {
  try {
    const target = resolveClassroomTarget(customUrl, customTab);
    const ss = SpreadsheetApp.openById(target.id);
    const sheet = ss.getSheetByName(target.tab);
    
    if (!sheet) {
      const available = ss.getSheets().map(s => s.getName()).join(", ");
      throw new Error(`Tab "${target.tab}" not found. Available tabs: [${available}]`);
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return { success: true, headers: [], rows: [] };

    const headers = data[0];
    // Map rows to include their 1-based index for updating
    const rows = data.slice(1).map((r, i) => ({
      rowIndex: i + 2, // +2 because: 0-based index + 1 header row + 1 to make it 1-based
      data: r
    }));

    return { success: true, headers: headers, rows: rows };

  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * BULK EDITOR: Apply Updates
 */
function api_bulkUpdateSheet(customUrl, customTab, updates) {
    requireRole(['Lead']);
  try {
    // Security Check: Ensure user has permission (Admin or Lead)
    // Assuming Proofing Tools is restricted to Admins/Leads in the UI, 
    // but good to enforce here if possible. 
    // Since this tool connects to ANY sheet the user has access to, 
    // the Google Sheet permissions themselves act as the primary gatekeeper.
    // However, we still sanitize inputs to prevent formula injection.

    const target = resolveClassroomTarget(customUrl, customTab);
    const ss = SpreadsheetApp.openById(target.id);
    const sheet = ss.getSheetByName(target.tab);
    
    if (!sheet) throw new Error(`Tab "${target.tab}" not found.`);

    // updates = [{ rowIndex, colIndex, value }]
    updates.forEach(u => {
      // Security: Sanitize value before writing
      const safeValue = sanitizeInput(u.value);
      sheet.getRange(u.rowIndex, u.colIndex + 1).setValue(safeValue);
    });

    return { success: true, message: `Updated ${updates.length} cells.` };

  } catch (e) { return { success: false, message: e.message }; }
}

/**
 * EXPORT: Create new sheet with filtered data
 */
function exportToSheet(headers, rows) {
  try {
    if (!rows || rows.length === 0) throw new Error("No data to export.");
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const ss = SpreadsheetApp.create(`Proofing Export - ${timestamp}`);
    const sheet = ss.getActiveSheet();
    
    // Prepare data
    const dataToSave = [headers, ...rows];
    
    // Write to sheet
    sheet.getRange(1, 1, dataToSave.length, headers.length).setValues(dataToSave);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold")
               .setBackground("#003057")
               .setFontColor("#ffffff");
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    return { success: true, url: ss.getUrl() };
    
  } catch (e) { 
    return { success: false, error: e.message }; 
  }
}