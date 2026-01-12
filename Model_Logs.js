/**
 * -------------------------------------------------------------------
 * SYSTEM LOGGING MODEL
 * Handles Audit Trails for Document Generation and Updates
 * -------------------------------------------------------------------
 */

function logSystemAction(module, action, docName, docId, details) {
  try {
    const ss = getMasterDataHub();
    let sheet = ss.getSheetByName(CONFIG.TABS.SYSTEM_LOGS);
    
    // Create Log Sheet if missing
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.TABS.SYSTEM_LOGS);
      sheet.appendRow(['Timestamp', 'User Email', 'Module', 'Action', 'Document Name', 'File ID', 'Details']);
      sheet.setFrozenRows(1);
    }

    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail() || 'System/Script';
    
    sheet.appendRow([
      timestamp,
      userEmail,
      module,
      action,
      docName,
      docId,
      details || ''
    ]);
    
  } catch (e) {
    console.error("Failed to write to System Log: " + e.message);
    // We do not throw here to prevent breaking the main app flow
  }
}