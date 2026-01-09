/**
 * -------------------------------------------------------------------
 * ENTRY POINT & SERVER-SIDE CONFIGURATION
 * -------------------------------------------------------------------
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

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
    return ScriptApp.getService().getUrl();
}

/**
 * ===================================================================
 * BOOTSTRAP & CONFIGURATION API (v4 - ROBUST)
 * ===================================================================
 */

/**
 * API v4: Gets all core application data in a single, efficient call.
 * This is the primary function called on application load.
 */
function api_getCoreData() {
    try {
        const userEmail = Session.getActiveUser().getEmail();
        const settings = getSettings();
        const staffData = getSheet('Staff_List').getDataRange().getValues();

        const access = _getAccessControlData(userEmail, staffData, settings);
        if (!access || !access.staffId) {
            throw new Error("Access control data could not be resolved. User email may not be in Staff_List.");
        }

        const dashboard = _getDashboardData(access.staffId);
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
        console.error(`Fatal Error in api_getCoreData: ${e.message} Stack: ${e.stack}`);
        return { success: false, message: `Could not load core application data. The server reported an error: ${e.message}` };
    }
}

/**
 * API: Sets up the initial application configuration if it's missing.
 * This function creates a default page menu based on the View_*.html files.
 */
function api_initializeApplicationSettings() {
    try {
        let settings = getSettings();

        // Default pages to create if they don't exist
        const defaultPages = [
            { id: "dashboard", name: "Dashboard", icon: "\uD83D\uDDA5" }, // Computer icon
            { id: "analytics", name: "Analytics", icon: "\uD83D\uDCCA" }  // Chart icon
        ];

        settings.pages = JSON.stringify(defaultPages);
        
        // You can add other default settings here, for example:
        settings.appName = "University Dept Management";

        saveSettings(settings);
        
        return { success: true, message: "Application configured successfully. The page will now reload." };

    } catch (e) {
        console.error(`Error in api_initializeApplicationSettings: ${e.message}`);
        return { success: false, message: `Failed to initialize settings. Reason: ${e.message}` };
    }
}


/**
 * -------------------------------------------------------------------
 * INTERNAL HELPER & UTILITY FUNCTIONS
 * -------------------------------------------------------------------
 */

function getSheet(name) {
    // This is a placeholder for a more robust function that would handle errors
    // and potentially create sheets if they don't exist.
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getSettings() {
    try {
        const properties = PropertiesService.getUserProperties().getProperties();
        return properties;
    } catch (e) {
        console.error("Could not retrieve settings: " + e.message);
        return {};
    }
}

function saveSettings(settings) {
    try {
        PropertiesService.getUserProperties().setProperties(settings, true);
    } catch (e) {
        console.error("Could not save settings: " + e.message);
    }
}

function _getAccessControlData(userEmail, staffData, settings) {
    const permSheet = getSheet('Permissions_Matrix');
    const matrix = {};
    if (permSheet) {
        const data = permSheet.getDataRange().getValues();
        data.slice(1).forEach(row => {
            const page = row[0];
            if (page) { matrix[page] = { Admin: true, Lead: row[2] === true, Staff: row[3] === true }; }
        });
    }

    // --- ROBUSTNESS FIX --- 
    const pages = JSON.parse(settings.pages || "[]");
    pages.forEach(p => { if (!matrix[p.id]) { matrix[p.id] = { Admin: true, Lead: true, Staff: true }; } });

    let userRole = "Staff";
    let staffId = null; 

    const headers = staffData[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf('staffid'); // Assuming staffid is the email
    const roleIdx = headers.indexOf('roles');

    if (emailIdx === -1) {
        throw new Error("Critical: 'staffid' column not found in 'Staff_List' sheet.");
    }

    for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][emailIdx] && String(staffData[i][emailIdx]).trim().toLowerCase() === userEmail.toLowerCase()) {
            userRole = (roleIdx > -1 && staffData[i][roleIdx]) ? staffData[i][roleIdx] : "Staff";
            staffId = staffData[i][emailIdx]; 
            break;
        }
    }

    return { userRole, matrix, email: userEmail, staffId, pages: pages };
}

function _getDashboardData(staffId) {
    let availability = [], preferences = [];
    const availSheet = getSheet('Staff_Availability');
    if(availSheet){
      const availData = availSheet.getDataRange().getDisplayValues();
      const headers = availData[0].map(h => String(h).trim().toLowerCase());
      const staffIdCol = headers.indexOf('staffid');
      if (staffIdCol !== -1) { 
        availability = availData.slice(1).map((row, i) => row.concat(i + 2)).filter(row => row[staffIdCol] && String(row[staffIdCol]).toLowerCase() === String(staffId).toLowerCase()); 
      }
    }

    const prefSheet = getSheet('Staff_Preferences');
    if(prefSheet){
      const prefData = prefSheet.getDataRange().getDisplayValues();
      const headers = prefData[0].map(h => String(h).trim().toLowerCase());
      const staffIdCol = headers.indexOf('staffid');
      if (staffIdCol !== -1) { 
        preferences = prefData.slice(1).filter(row => row[staffIdCol] && String(row[staffIdCol]).toLowerCase() === String(staffId).toLowerCase());
      } 
    }

    return { availability, preferences };
}

function _getStaffList(staffData) {
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
