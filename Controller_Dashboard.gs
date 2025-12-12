/**
 * -------------------------------------------------------------------
 * DASHBOARD CONTROLLER
 * Handles Staff Availability and Preferences
 * -------------------------------------------------------------------
 */

function api_getMyAvailability(email) {
    try {
        // Fallback: If "Myself" is selected (empty string), use active user
        if (!email) email = Session.getActiveUser().getEmail();

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_AVAILABILITY);
        const data = sheet.getDataRange().getValues();
        
        // Find Staff ID from Email
        const staffSheet = ss.getSheetByName(CONFIG.TABS.STAFF_LIST);
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        
        // Assuming Col B is ID/Email
        for(let i=1; i<staffData.length; i++) {
            if(String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        
        if(!staffId) return { success: false, message: "Staff ID not found for email." };
        
        const results = [];
        for(let i=1; i<data.length; i++) {
            // Force String comparison for ID
            if(String(data[i][1]) === staffId) {
                let s = data[i][3];
                let e = data[i][4];
                // Format Date objects to simple strings for UI
                if (s instanceof Date) s = Utilities.formatDate(s, Session.getScriptTimeZone(), "HH:mm");
                if (e instanceof Date) e = Utilities.formatDate(e, Session.getScriptTimeZone(), "HH:mm");
                
                results.push({ id: data[i][0], day: data[i][2], start: s, end: e });
            }
        }
        return { success: true, data: results };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_addNotAvailable(day, start, end, email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_AVAILABILITY);
        
        // Resolve ID
        const staffSheet = ss.getSheetByName(CONFIG.TABS.STAFF_LIST);
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for(let i=1; i<staffData.length; i++) {
            if(String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if(!staffId) return { success: false, message: "Staff ID not found." };
        
        const id = Utilities.getUuid();
        sheet.appendRow([id, staffId, day, start, end]);
        return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_deleteAvailability(id, email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_AVAILABILITY);
        const data = sheet.getDataRange().getValues();
        
        for(let i=1; i<data.length; i++) {
            if(data[i][0] === id) {
                sheet.deleteRow(i+1);
                return { success: true };
            }
        }
        return { success: false, message: "Item not found." };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_getMyPreferences(email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_PREFERENCES);
        const data = sheet.getDataRange().getValues();
        
        // Resolve ID
        const staffSheet = ss.getSheetByName(CONFIG.TABS.STAFF_LIST);
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for(let i=1; i<staffData.length; i++) {
            if(String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if(!staffId) return { success: false, message: "Staff ID not found." };
        
        const prefs = {};
        for(let i=1; i<data.length; i++) {
            if(String(data[i][0]) === staffId) {
                prefs[data[i][1]] = data[i][2]; 
            }
        }
        return { success: true, data: prefs };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_savePreference(key, value, email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_PREFERENCES);
        const data = sheet.getDataRange().getValues();
        
        // Resolve ID
        const staffSheet = ss.getSheetByName(CONFIG.TABS.STAFF_LIST);
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for(let i=1; i<staffData.length; i++) {
            if(String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if(!staffId) return { success: false, message: "Staff ID not found." };
        
        let found = false;
        for(let i=1; i<data.length; i++) {
            if(String(data[i][0]) === staffId && data[i][1] === key) {
                sheet.getRange(i+1, 3).setValue(value);
                found = true;
                break;
            }
        }
        
        if(!found) {
            sheet.appendRow([staffId, key, value]);
        }
        
        return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
}