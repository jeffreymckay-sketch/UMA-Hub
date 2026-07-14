/**
 * -------------------------------------------------------------------
 * CONTROLLER: PERMISSIONS MATRIX
 * Handles reading, syncing, and saving page-level access rights.
 * -------------------------------------------------------------------
 */

const PERMISSIONS_CONFIG = {
    SHEET_NAME: 'Permissions_Matrix',
    // The fixed list of roles as requested
    ROLES: ['Admin', 'Lead', 'Staff', 'Tech Hub', 'MST', 'Proctor', 'FDC'],
    // The list of Page IDs currently used in Index.html / JS_Core.html
    PAGES: [
        'Dashboard', 'MySchedule', 'Availability', 'UserGuide',
        'Logistics', 'MST', 'TechHub', 'SingleDoc',
        'Proctoring', 'Nursing', 'MLT', 'ProctorSchedule',
        'Reporting', 'ProofingTools',
        'Zoom', 'Settings'
    ]
};

/**
 * Fetches the permissions matrix.
 * Syncs the sheet structure if pages or roles are missing.
 */
function api_getPermissionsMatrix() {
    try {
        // Security: Admin only for editing, but we need a read-only version for app load
        // We handle the "Edit" check in the save function.
        const ss = getMasterDataHub();
        let sheet = ss.getSheetByName(PERMISSIONS_CONFIG.SHEET_NAME);
        let data, headers, sheetModified = false;

        // 1. Initialize Sheet if missing
        if (!sheet) {
            sheet = ss.insertSheet(PERMISSIONS_CONFIG.SHEET_NAME);
            // Create Headers
            headers = ['PageID', ...PERMISSIONS_CONFIG.ROLES];
            sheet.appendRow(headers);
            sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#003057').setFontColor('white');
            sheet.setFrozenRows(1);
            sheet.setFrozenColumns(1);
            data = [headers];
        } else {
            data = sheet.getDataRange().getValues();
            headers = data[0];
        }
        
        // 2. Sync Logic: Ensure all Roles exist as columns
        const missingRoles = PERMISSIONS_CONFIG.ROLES.filter(r => !headers.includes(r));
        if (missingRoles.length > 0) {
            // Add missing columns
            const startCol = headers.length + 1;
            sheet.getRange(1, startCol, 1, missingRoles.length).setValues([missingRoles]);
            sheetModified = true;
        }

        // Re-fetch data if columns were added
        if(sheetModified){
          data = sheet.getDataRange().getValues();
          headers = data[0];
          sheetModified = false; // Reset for next check
        }

        // 3. Sync Logic: Ensure all Pages exist as rows
        const existingPages = data.slice(1).map(r => r[0]);
        const missingPages = PERMISSIONS_CONFIG.PAGES.filter(p => !existingPages.includes(p));
        
        if (missingPages.length > 0) {
            const newRows = missingPages.map(pageId => {
                const row = [pageId];
                // Default: Admin gets TRUE, others FALSE
                // Let's default 'MySchedule', 'Availability', and 'UserGuide' to true for everyone
                PERMISSIONS_CONFIG.ROLES.forEach(role => {
                    let defaultVal = false;
                    if (role === 'Admin') defaultVal = true;
                    if (['MySchedule', 'Availability', 'UserGuide'].includes(pageId)) defaultVal = true; 
                    
                    row.push(defaultVal);
                });
                return row;
            });
            sheet.getRange(data.length + 1, 1, newRows.length, headers.length).setValues(newRows);
            sheetModified = true;
        }

        // Re-fetch data if anything changed to ensure we have the final matrix
        if(sheetModified){
          data = sheet.getDataRange().getValues();
          headers = data[0];
        }

        // 4. Build the Matrix Object
        // Structure: { "MST": { "Admin": true, "Staff": false }, ... }
        const matrix = {};
        const roleIndices = {};
        
        // Map header index to role name
        headers.forEach((h, i) => {
            if (PERMISSIONS_CONFIG.ROLES.includes(h)) roleIndices[h] = i;
        });

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const pageId = row[0];
            if (!pageId) continue;

            matrix[pageId] = {};
            PERMISSIONS_CONFIG.ROLES.forEach(role => {
                const idx = roleIndices[role];
                // Convert to boolean
                const val = row[idx];
                matrix[pageId][role] = (String(val).toLowerCase() === 'true' || val === true);
            });
        }

        return { success: true, data: matrix, roles: PERMISSIONS_CONFIG.ROLES, pages: PERMISSIONS_CONFIG.PAGES };

    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * Saves the updated permissions matrix.
 * INCLUDES FAIL-SAFE to prevent Admin Lockout.
 */
function api_savePermissionsMatrix(updates) {
    try {
        requireRole('Admin');
        // Strict Security Check

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(PERMISSIONS_CONFIG.SHEET_NAME);
        if (!sheet) throw new Error("Permissions sheet not found.");

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        // Map headers to indices
        const colMap = {};
        headers.forEach((h, i) => colMap[h] = i);

        // Prepare updates
        // updates is { "MST": { "Admin": true, "Staff": false }, ... }
        
        const output = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const pageId = row[0];
            
            if (updates[pageId]) {
                PERMISSIONS_CONFIG.ROLES.forEach(role => {
                    if (colMap[role] !== undefined) {
                        let newVal = updates[pageId][role];
                        
                        // --- CRITICAL FAIL-SAFE ---
                        // Ensure Admin ALWAYS has access to Settings
                        if (pageId === 'Settings' && role === 'Admin') {
                            newVal = true;
                        }
                        
                        row[colMap[role]] = newVal;
                    }
                });
            }
            output.push(row);
        }

        // Write back (excluding header)
        if (output.length > 0) {
            sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
        }

        return { success: true, message: "Permissions updated successfully." };

    } catch (e) {
        return { success: false, message: e.message };
    }
}