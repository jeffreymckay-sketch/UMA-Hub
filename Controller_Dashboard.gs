/**
 * -------------------------------------------------------------------
 * DASHBOARD CONTROLLER
 * Handles Staff Availability and Preferences
 * -------------------------------------------------------------------
 */

/**
 * NEW, EFFICIENT WRAPPER FUNCTION
 * Fetches all data required for the dashboard in a single server trip.
 * This avoids redundant lookups and multiple concurrent calls.
 */
function api_getDashboardData(email) {
    try {
        // Use active user if no email is provided (for non-admins)
        if (!email) email = Session.getActiveUser().getEmail();

        const staffListSheet = getSheet('Staff_List');
        const staffData = staffListSheet.getDataRange().getValues();

        // --- Find Staff ID (The core expensive operation) ---
        // This is now only done ONCE per dashboard load.
        let staffId = null;
        for (let i = 1; i < staffData.length; i++) {
            // Column B (index 1) is the Email/ID
            if (String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }

        if (!staffId) {
            return { success: false, message: "Could not find a staff record for the email: " + email };
        }

        // --- Now, gather the data using the resolved staffId ---
        const availability = getMyAvailability_DEPRECATED(staffId);
        const preferences = getMyPreferences_DEPRECATED(staffId);

        return {
            success: true,
            data: {
                availability: availability,
                preferences: preferences
            }
        };

    } catch (e) {
        return { success: false, message: `An error occurred: ${e.message}` };
    }
}


// --- HELPER for api_getDashboardData: Fetches Availability ---
function getMyAvailability_DEPRECATED(staffId) {
    const sheet = getSheet('Staff_Availability');
    const data = sheet.getDataRange().getValues();
    const results = [];
    for (let i = 1; i < data.length; i++) {
        // Column B (index 1) is Staff ID
        if (String(data[i][1]) === staffId) {
            let s = data[i][3]; // Start time
            let e = data[i][4]; // End time

            if (s instanceof Date) s = Utilities.formatDate(s, Session.getScriptTimeZone(), "HH:mm");
            if (e instanceof Date) e = Utilities.formatDate(e, Session.getScriptTimeZone(), "HH:mm");

            results.push({ id: data[i][0], day: data[i][2], start: s, end: e });
        }
    }
    return results;
}


// --- HELPER for api_getDashboardData: Fetches Preferences ---
function getMyPreferences_DEPRECATED(staffId) {
    const sheet = getSheet('Staff_Preferences');
    const data = sheet.getDataRange().getValues();
    const prefs = {};
    for (let i = 1; i < data.length; i++) {
        // Column A (index 0) is Staff ID
        if (String(data[i][0]) === staffId) {
            prefs[data[i][1]] = data[i][2]; // Key = Value
        }
    }
    return prefs;
}


/**
 * -------------------------------------------------------------------
 * The original, inefficient functions are left below for other parts
 * of the app that might use them. The dashboard itself will no longer
 * call them directly.
 * -------------------------------------------------------------------
 */

function api_getMyAvailability(email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();
        const staffSheet = getSheet('Staff_List');
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for (let i = 1; i < staffData.length; i++) {
            if (String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if (!staffId) return { success: false, message: "Staff ID not found for email." };

        const data = getMyAvailability_DEPRECATED(staffId);
        return { success: true, data: data };

    } catch (e) { return { success: false, message: e.message }; }
}

function api_getMyPreferences(email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();
        const staffSheet = getSheet('Staff_List');
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for (let i = 1; i < staffData.length; i++) {
            if (String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if (!staffId) return { success: false, message: "Staff ID not found." };
        
        const prefs = getMyPreferences_DEPRECATED(staffId);
        return { success: true, data: prefs };

    } catch (e) { return { success: false, message: e.message }; }
}


function api_addNotAvailable(day, start, end, email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const sheet = getSheet('Staff_Availability');

        // Resolve ID
        const staffSheet = getSheet('Staff_List');
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for (let i = 1; i < staffData.length; i++) {
            if (String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if (!staffId) return { success: false, message: "Staff ID not found." };

        const id = Utilities.getUuid();
        sheet.appendRow([id, staffId, day, start, end]);
        return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_deleteAvailability(id, email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const sheet = getSheet('Staff_Availability');
        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === id) {
                sheet.deleteRow(i + 1);
                return { success: true };
            }
        }
        return { success: false, message: "Item not found." };
    } catch (e) { return { success: false, message: e.message }; }
}


function api_savePreference(key, value, email) {
    try {
        if (!email) email = Session.getActiveUser().getEmail();

        const sheet = getSheet('Staff_Preferences');
        const data = sheet.getDataRange().getValues();

        // Resolve ID
        const staffSheet = getSheet('Staff_List');
        const staffData = staffSheet.getDataRange().getValues();
        let staffId = null;
        for (let i = 1; i < staffData.length; i++) {
            if (String(staffData[i][1]).toLowerCase() === String(email).toLowerCase()) {
                staffId = String(staffData[i][1]);
                break;
            }
        }
        if (!staffId) return { success: false, message: "Staff ID not found." };

        let found = false;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === staffId && data[i][1] === key) {
                sheet.getRange(i + 1, 3).setValue(value);
                found = true;
                break;
            }
        }

        if (!found) {
            sheet.appendRow([staffId, key, value]);
        }

        return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
}
