/**
 * -------------------------------------------------------------------
 * DASHBOARD CONTROLLER
 * Handles Staff Availability and Preferences
 * -------------------------------------------------------------------
 */

/**
 * Fetches all data required for the dashboard for the CURRENT user.
 * @returns {object} A payload with the user's availability and preferences.
 */
function api_getDashboardData() {
    try {
        const email = Session.getActiveUser().getEmail();

        // --- 1. Fetch Availability ---
        const availabilitySheet = getSheet('Staff_Availability');
        const availabilityData = availabilitySheet.getDataRange().getValues();
        const userAvailability = [];
        for (let i = 1; i < availabilityData.length; i++) {
            if (availabilityData[i][1] === email) { // Email is index 1
                let start = availabilityData[i][3];
                let end = availabilityData[i][4];
                if (start instanceof Date) start = Utilities.formatDate(start, Session.getScriptTimeZone(), "HH:mm");
                if (end instanceof Date) end = Utilities.formatDate(end, Session.getScriptTimeZone(), "HH:mm");
                userAvailability.push({ id: availabilityData[i][0], day: availabilityData[i][2], start: start, end: end });
            }
        }

        // --- 2. Fetch Time Block Preferences ---
        const preferencesSheet = getSheet('Staff_Preferences');
        const preferencesData = preferencesSheet.getDataRange().getValues();
        const userPreferences = {};
        for (let i = 1; i < preferencesData.length; i++) {
            if (preferencesData[i][0] === email) { // StaffID (email) is index 0
                userPreferences[preferencesData[i][1]] = preferencesData[i][2]; // e.g., { "Monday_Morning": "Preferred" }
            }
        }

        // --- 3. Return Combined Payload ---
        return {
            success: true,
            data: {
                availability: userAvailability,
                preferences: userPreferences,
            }
        };

    } catch (e) {
        console.error("api_getDashboardData Error: " + e.stack);
        return { success: false, message: `Failed to load dashboard data. Please refresh and try again. Error: ${e.message}` };
    }
}

/**
 * Creates a new availability record for the current user.
 * @param {object} formData The data from the client: { day, start, end }.
 * @returns {object} A status object.
 */
function api_addAvailability(formData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const { day, start, end } = formData;
        const email = Session.getActiveUser().getEmail();

        const sheet = getSheet('Staff_Availability');
        const newId = 'AV-' + new Date().getTime();
        sheet.appendRow([newId, email, day, start, end]);
        
        return { success: true, message: "Availability slot added successfully!" };
    } catch (e) {
        console.error("api_addAvailability Error: " + e.stack);
        return { success: false, message: `Failed to add availability slot. Error: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Deletes an availability record for the current user.
 * @param {string} recordId The unique ID of the availability slot to delete.
 * @returns {object} A status object.
 */
function api_deleteAvailability(recordId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const email = Session.getActiveUser().getEmail();
        const sheet = getSheet('Staff_Availability');
        const data = sheet.getDataRange().getValues();

        for (let i = data.length - 1; i >= 1; i--) {
            // Check that the record belongs to the current user
            if (data[i][0] === recordId && data[i][1] === email) {
                sheet.deleteRow(i + 1);
                return { success: true, message: "Availability slot has been deleted." };
            }
        }
        throw new Error("Record not found or permission denied.");
    } catch (e) {
        console.error("api_deleteAvailability Error: " + e.stack);
        return { success: false, message: `Failed to delete availability. Error: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Updates the current user's time block preferences.
 * @param {object} preferences The preferences data from the client.
 * @returns {object} A status object.
 */
function api_updateStaffPreferences(preferences) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
        const email = Session.getActiveUser().getEmail();
        const sheet = getSheet('Staff_Preferences');
        const data = sheet.getDataRange().getValues();

        // To keep the logic simple and robust, we remove all old preferences and add the new ones.
        for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][0] === email) {
                sheet.deleteRow(i + 1);
            }
        }
        
        // Add back the new preferences, skipping the neutral ones which are default
        for (const timeBlock in preferences) {
            const preference = preferences[timeBlock];
            if (preference !== 'Eh, Sure') {
                sheet.appendRow([email, timeBlock, preference]);
            }
        }

        return { success: true, message: "Preferences have been saved successfully!" };

    } catch (e) {
        console.error("api_updateStaffPreferences Error: " + e.stack);
        return { success: false, message: `Failed to save preferences. Error: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}
