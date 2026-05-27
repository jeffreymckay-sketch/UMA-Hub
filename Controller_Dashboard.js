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
        // Using getDisplayValues() to ensure time strings are formatted correctly.
        const availabilitySheet = getSheet('Staff_Availability');
        const availabilityData = availabilitySheet.getDataRange().getDisplayValues();
        const userAvailability = [];
        for (let i = 1; i < availabilityData.length; i++) {
            if (availabilityData[i][1] === email) { // Email is index 1
                userAvailability.push({ 
                    id: availabilityData[i][0], 
                    day: availabilityData[i][2], 
                    start: availabilityData[i][3] || 'N/A', 
                    end: availabilityData[i][4] || 'N/A' 
                });
            }
        }

        // --- 2. Fetch Time Block Preferences ---
        // Using getDisplayValues() for consistency.
        const preferencesSheet = getSheet('Staff_Preferences');
        const preferencesData = preferencesSheet.getDataRange().getDisplayValues();
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
        // Sanitize all inputs
        const day = sanitizeInput(formData.day);
        const start = sanitizeInput(formData.start);
        const end = sanitizeInput(formData.end);
        
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
        
        // Grab the entire data set as an array
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) throw new Error("No records to process.");
        
        const header = data[0];
        
        // Filter OUT the row that matches the record ID AND belongs to the user
        const originalLength = data.length;
        const newData = data.filter((row, index) => {
            if (index === 0) return true; // Always keep the header
            const isMatch = (String(row[0]) === String(recordId) && String(row[1]) === String(email));
            return !isMatch; // Keep if it's NOT a match
        });
        
        if (newData.length === originalLength) {
             throw new Error("Record not found or permission denied.");
        }

        // Clear the entire sheet and write the newly filtered array back in bulk
        sheet.clearContents();
        sheet.getRange(1, 1, newData.length, header.length).setValues(newData);

        return { success: true, message: "Availability slot has been deleted." };
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
        
        // Grab the entire data set as an array
        const data = sheet.getDataRange().getValues();
        const header = data.length > 0 ? data[0] : ["Staff ID", "Time Block", "Preference"]; 
        
        // Filter OUT all existing preferences for this user
        const newData = data.filter((row, index) => {
            if (index === 0) return true; // Always keep the header
            return String(row[0]) !== String(email); 
        });
        
        // Add back the new preferences, skipping the neutral ones which are default
        for (const timeBlock in preferences) {
            const preference = sanitizeInput(preferences[timeBlock]);
            const sanitizedBlock = sanitizeInput(timeBlock);
            
            if (preference !== 'Eh, Sure') {
                newData.push([email, sanitizedBlock, preference]);
            }
        }

        // Clear the entire sheet and write the updated array back in bulk
        sheet.clearContents();
        
        // If there's no data (just a header), we only write the header back
        if (newData.length > 0) {
            // Fill any missing columns in the new rows with blank strings so the array is perfectly rectangular
            const maxCols = header.length;
            const rectangularData = newData.map(row => {
                const newRow = [...row];
                while(newRow.length < maxCols) newRow.push("");
                return newRow;
            });
            sheet.getRange(1, 1, rectangularData.length, maxCols).setValues(rectangularData);
        }

        return { success: true, message: "Preferences have been saved successfully!" };

    } catch (e) {
        console.error("api_updateStaffPreferences Error: " + e.stack);
        return { success: false, message: `Failed to save preferences. Error: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}