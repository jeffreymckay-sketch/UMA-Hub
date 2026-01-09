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
        const userEmail = Session.getActiveUser().getEmail();

        const staffData = getSheet('Staff_List').getDataRange().getValues();

        // --- FIX & DEFENSIVE CHECK ---
        const access = _getAccessControlData(userEmail, staffData, settings);
        if (!access || !access.staffId) {
            throw new Error("Access control data was not returned. This is often due to the user's email not being found in the Staff_List tab under the 'StaffID' column.");
        }

        const dashboard = _getDashboardData(access.staffId, settings);
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

function _getAccessControlData(userEmail, staffData, settings) {
    const permSheet = getSheet('Permissions_Matrix');
    const matrix = {};
    if (permSheet) {
        const data = permSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            const page = data[i][0];
            if (page) { matrix[page] = { Admin: true, Lead: data[i][2] === true, Staff: data[i][3] === true }; }
        }
    }
    
    const pages = JSON.parse(settings.pages || "[]");
    pages.forEach(p => { if (!matrix[p.id]) { matrix[p.id] = { Admin: true, Lead: true, Staff: true }; } });

    let userRole = "Staff";
    let staffId = null; 

    if (staffData && staffData.length > 1) {
        const headers = staffData[0].map(h => String(h).trim().toLowerCase());
        
        // --- ROBUST HEADER FINDER ---
        const possibleEmailHeaders = ['staffid', 'email', 'staff email', 'user email'];
        let emailIdx = -1;
        for(const header of possibleEmailHeaders){
            const idx = headers.indexOf(header);
            if(idx > -1){
                emailIdx = idx;
                break;
            }
        }
        
        const roleIdx = headers.indexOf('roles');

        if (emailIdx > -1) {
            for (let i = 1; i < staffData.length; i++) {
                if (staffData[i][emailIdx] && String(staffData[i][emailIdx]).trim().toLowerCase() === userEmail.toLowerCase()) {
                    userRole = (roleIdx > -1 && staffData[i][roleIdx]) ? staffData[i][roleIdx] : "Staff";
                    staffId = staffData[i][emailIdx]; 
                    break;
                }
            }
        } else {
             throw new Error("Could not find a valid user email column in the 'Staff_List' sheet. Please ensure a column with one of the following headers exists: 'staffid', 'email', 'staff email', or 'user email'.");
        }
    }

    return { userRole, matrix, email: userEmail, staffId, pages: pages };
}


function _getDashboardData(staffId, settings) {
    try {
        let availability = [], preferences = [];
        // CORRECTED: Used getDisplayValues() to prevent issues with Date objects.
        const availData = getSheet('Staff_Availability').getDataRange().getDisplayValues();
        if (availData.length > 1) {
            const headers = availData[0].map(h => String(h).trim().toLowerCase());
            const staffIdCol = headers.indexOf('staffid');
            if (staffIdCol !== -1) { availability = availData.slice(1).filter(row => row[staffIdCol] && String(row[staffIdCol]).toLowerCase() === String(staffId).toLowerCase()); }
        }
        const prefData = getSheet('Staff_Preferences').getDataRange().getDisplayValues();
        if (prefData.length > 1) {
            const headers = prefData[0].map(h => String(h).trim().toLowerCase());
            const staffIdCol = headers.indexOf('staffid');
            if (staffIdCol !== -1) { preferences = prefData.slice(1).filter(row => row[staffIdCol] && String(row[staffIdCol]).toLowerCase() === String(staffId).toLowerCase()); }
        }
        return { availability, preferences };
    } catch (e) {
        console.error(`Error in _getDashboardData for staffId ${staffId}: ${e.message}`);
        return { availability: [], preferences: [] };
    }
}

function _getStaffList(staffData) {
    if (!staffData || staffData.length < 2) return [];
    const headers = staffData[0].map(h => String(h).trim().toLowerCase());
    const nameIdx = headers.indexOf('fullname');
    const idIdx = headers.indexOf('staffid');
    const roleIdx = headers.indexOf('roles');
    if (nameIdx === -1 || idIdx === -1) return [];
    return staffData.slice(1).map(row => {
        if (!row[idIdx]) return null;
        return { name: row[nameIdx], id: row[idIdx], role: (roleIdx > -1 && row[roleIdx]) ? row[roleIdx] : "Staff" };
    }).filter(s => s !== null);
}
