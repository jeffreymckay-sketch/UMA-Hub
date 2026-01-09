/**
 * -------------------------------------------------------------------
 * ENTRY POINT & SERVER-SIDE CONFIGURATION
 * -------------------------------------------------------------------
 */

/**
 * Main entry point for the web app.
 */
function doGet(e) {
    const settings = getSettings();
    const appName = settings.appName || "University Department Management";
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle(appName)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Allows HTML templates to include other HTML files.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns the script's deployment URL.
 */
function getScriptUrl() {
    return ScriptApp.getService().getUrl();
}

/**
 * ===================================================================
 * BOOTSTRAP FUNCTION (PRIMARY DATA LOAD)
 * ===================================================================
 */
function api_getInitialAppData() {
    try {
        const settings = getSettings();
        const ss = getMasterDataHub();
        const userEmail = Session.getActiveUser().getEmail();

        const staffData = getRequiredSheetData(ss, settings.tabNameStaffList);

        // --- FIX & DEFENSIVE CHECK ---
        const access = _getAccessControlData(ss, userEmail, staffData, settings);
        if (!access || !access.staffId) {
            throw new Error("Access control data was not returned. This is often due to the user\'s email not being found in the Staff_List tab under the \'StaffID\' column.");
        }

        const dashboard = _getDashboardData(ss, access.staffId, settings);
        const staffList = _getStaffList(staffData);

        return { 
            success: true, 
            data: {
                access: access,
                dashboard: dashboard,
                staff: staffList
            } 
        };

    } catch (e) {
        console.error(`Fatal Error in api_getInitialAppData: ${e.message} Stack: ${e.stack}`);
        return { success: false, message: `Could not load initial application data. The server reported an error: ${e.message}` };
    }
}


/**
 * -------------------------------------------------------------------
 * INTERNAL HELPER FUNCTIONS
 * -------------------------------------------------------------------
 */

/**
 * CORRECTED: Gathers user role, permissions, and ID.
 */
function _getAccessControlData(ss, userEmail, staffData, settings) {
    const permSheet = ss.getSheetByName(settings.tabNamePermissionsMatrix); // CORRECTED TAB NAME
    const matrix = {};
    if (permSheet) {
        const data = permSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            const page = data[i][0];
            if (page) {
                matrix[page] = { Admin: true, Lead: data[i][2] === true, Staff: data[i][3] === true };
            }
        }
    }
    
    const pages = JSON.parse(settings.pages || "[]");
    pages.forEach(p => { if (!matrix[p.id]) { matrix[p.id] = { Admin: true, Lead: true, Staff: true }; } });

    let userRole = "Staff";
    let staffId = null; 

    if (staffData && staffData.length > 1) {
        const headers = staffData[0].map(h => String(h).trim().toLowerCase());
        const emailIdx = headers.indexOf('staffid'); // CORRECTED HEADER
        const roleIdx = headers.indexOf('role');

        if (emailIdx > -1 && roleIdx > -1) {
            for (let i = 1; i < staffData.length; i++) {
                if (String(staffData[i][emailIdx]).toLowerCase() === userEmail.toLowerCase()) {
                    userRole = staffData[i][roleIdx] || "Staff";
                    staffId = staffData[i][emailIdx]; // Use the official ID from the sheet
                    break;
                }
            }
        }
    }

    return { userRole, matrix, email: userEmail, staffId, pages: pages };
}


/**
 * Fetches dashboard-specific data for a given user.
 */
function _getDashboardData(ss, staffId, settings) {
    try {
        let availability = [];
        let preferences = [];

        const availData = getRequiredSheetData(ss, settings.tabNameStaffAvailability);
        if (availData.length > 1) {
            const headers = availData[0].map(h => String(h).trim().toLowerCase());
            const staffIdCol = headers.indexOf('staffid'); // CORRECTED HEADER
            if (staffIdCol !== -1) {
                availability = availData.slice(1).filter(row => String(row[staffIdCol]).toLowerCase() === String(staffId).toLowerCase());
            }
        }

        const prefData = getRequiredSheetData(ss, settings.tabNameStaffPreferences);
        if (prefData.length > 1) {
            const headers = prefData[0].map(h => String(h).trim().toLowerCase());
            const staffIdCol = headers.indexOf('staffid'); // CORRECTED HEADER
            if (staffIdCol !== -1) {
                preferences = prefData.slice(1).filter(row => String(row[staffIdCol]).toLowerCase() === String(staffId).toLowerCase());
            }
        }

        return { availability, preferences };
    } catch (e) {
        console.error(`Error in _getDashboardData for staffId ${staffId}: ${e.message}`);
        return { availability: [], preferences: [] }; // Return empty data to prevent a full crash
    }
}

/**
 * Processes the raw staff data into a clean list of objects.
 */
function _getStaffList(staffData) {
    if (!staffData || staffData.length < 2) return [];
    const headers = staffData[0].map(h => String(h).trim().toLowerCase());
    const nameIdx = headers.indexOf('name');
    const idIdx = headers.indexOf('staffid'); // CORRECTED HEADER
    const roleIdx = headers.indexOf('role');

    if (nameIdx === -1 || idIdx === -1 || roleIdx === -1) return [];

    const staff = staffData.slice(1).map(row => {
        if (!row[idIdx]) return null;
        return { name: row[nameIdx], id: row[idIdx], role: row[roleIdx] || "Staff" };
    }).filter(s => s !== null);

    return staff;
}

/**
 * Utility function to get all data from a required sheet.
 * Throws an error if the sheet is not found.
 */
function getRequiredSheetData(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`Required sheet \"${sheetName}\" not found in the master spreadsheet.`);
    }
    return sheet.getDataRange().getValues();
}
