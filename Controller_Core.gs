/**
 * -------------------------------------------------------------------
 * CORE CONTROLLER
 * Handles essential, app-wide data and actions.
 * -------------------------------------------------------------------
 */

/**
 * Fetches all necessary data for the initial application load.
 * This acts as a single point of entry to gather data for all tools,
 * preventing multiple server calls on navigation.
 * @returns {object} An object containing all the necessary data for the app.
 */
function api_getInitialAppData() {
    try {
        // Fetch data from various new internal helper functions
        const userInfo = getUserInfo();
        const mstViewData = getMstViewData();
        const allSettings = getAllSettings_();
        const calendars = getWritableCalendarsInternal();

        // Assemble the grand payload
        return {
            success: true,
            data: {
                userInfo: userInfo.success ? userInfo.data : { error: userInfo.message },
                mstData: mstViewData.success ? mstViewData.data : { error: mstViewData.error },
                settings: allSettings.success ? allSettings.data : { error: allSettings.message },
                writableCalendars: calendars.success ? calendars.data : { error: calendars.message }
            }
        };
    } catch (e) {
        console.error("api_getInitialAppData Error: " + e.stack);
        return {
            success: false,
            message: "A critical error occurred while loading application data: " + e.message
        };
    }
}

/**
 * NEW: Generic function to save a settings object for a specific tool.
 * @param {string} key The property key to save the settings under.
 * @param {object} settingsObject The settings object for the tool.
 * @returns {object} A response object with the full, updated settings cache.
 */
function api_saveSettings(key, settingsObject) {
    try {
        if (!key || !settingsObject) {
            throw new Error("A key and settings object are required.");
        }
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty(key, JSON.stringify(settingsObject));
        
        // Return all settings again so the client can update its cache
        const allSettings = getAllSettings_();
        if (allSettings.success) {
            return { success: true, message: 'Settings saved!', data: allSettings.data };
        } else {
            throw new Error('Could not retrieve updated settings after saving.');
        }
    } catch (e) {
        console.error("api_saveSettings Error: " + e.stack);
        return { success: false, message: 'Failed to save settings: ' + e.message };
    }
}


/**
 * Fetches all data for the user's dashboard.
 * This is called on-demand by the client.
 * @returns {object} An object containing dashboard data.
 */
function api_getDashboardData() {
    try {
        const email = Session.getActiveUser().getEmail();
        const availabilitySheet = getSheet('Staff_Availability');
        const availabilityData = availabilitySheet.getDataRange().getValues();
        const userAvailability = [];
        for (let i = 1; i < availabilityData.length; i++) {
            if (availabilityData[i][1] === email) {
                let start = availabilityData[i][3];
                let end = availabilityData[i][4];
                if (start instanceof Date) start = Utilities.formatDate(start, Session.getScriptTimeZone(), "HH:mm");
                if (end instanceof Date) end = Utilities.formatDate(end, Session.getScriptTimeZone(), "HH:mm");
                userAvailability.push({ id: availabilityData[i][0], day: availabilityData[i][2], start, end });
            }
        }

        const preferencesSheet = getSheet('Staff_Preferences');
        const preferencesData = preferencesSheet.getDataRange().getValues();
        const userPreferences = {};
        for (let i = 1; i < preferencesData.length; i++) {
            if (preferencesData[i][0] === email) {
                userPreferences[preferencesData[i][1]] = preferencesData[i][2];
            }
        }

        return { success: true, data: { availability: userAvailability, preferences: userPreferences } };
    } catch (e) {
        console.error("getDashboardData Error: " + e.stack);
        return { success: false, message: e.message };
    }
}


// --- INTERNAL DATA-FETCHING FUNCTIONS ---

/**
 * NEW: Internal function to get all settings stored in UserProperties.
 */
function getAllSettings_() {
    try {
        const userProperties = PropertiesService.getUserProperties();
        const allSettings = userProperties.getProperties();
        const parsedSettings = {};
        for (const key in allSettings) {
            try {
                // Assume settings are stored as JSON strings
                parsedSettings[key] = JSON.parse(allSettings[key]);
            } catch (e) {
                // If not a JSON string, use the value as is.
                parsedSettings[key] = allSettings[key];
            }
        }
        return { success: true, data: parsedSettings };
    } catch (e) {
        console.error('getAllSettings_ Error: ' + e.stack);
        return { success: false, message: 'Failed to retrieve settings.' };
    }
}

/**
 * Internal function to fetch user information.
 */
function getUserInfo() {
    try {
        const email = Session.getActiveUser().getEmail();
        const photoUrl = Session.getActiveUser().getPhotoUrl();
        return { success: true, data: { email, photoUrl } };
    } catch (e) {
        console.error("getUserInfo Error: " + e.stack);
        return { success: false, message: e.message };
    }
}

/**
 * Internal function to fetch MST view data.
 */
function getMstViewData() {
    try {
        const staffData = getSheet('Staff_List').getDataRange().getValues();
        const assignmentData = getSheet('Staff_Assignments').getDataRange().getValues();
        const courseData = getSheet('Course_Schedule').getDataRange().getValues();

        const staffHeaders = getColumnMap(staffData[0]);
        const assignmentHeaders = getColumnMap(assignmentData[0]);
        const courseHeaderRow = courseData.find(row => row.join('').toLowerCase().includes('eventid'));
        if (!courseHeaderRow) throw new Error("Missing 'eventID' header in Course Schedule.");
        const courseHeaders = getColumnMap(courseHeaderRow);
        const courseHeaderIndex = courseData.indexOf(courseHeaderRow);

        const allStaff = staffData.slice(1).map(row => parseStaff(row, staffHeaders)).filter(s => s && s.isActive);
        const allAssignments = assignmentData.slice(1).map(row => parseAssignment(row, assignmentHeaders)).filter(Boolean);
        const allCourses = courseData.slice(courseHeaderIndex + 1).map(row => parseCourse(row, courseHeaders)).filter(Boolean);

        const staffMap = new Map(allStaff.map(s => [String(s.id).toLowerCase(), s]));
        const assignmentMap = new Map(allAssignments.map(a => [String(a.eventId), a]));

        const courseAssignmentsView = allCourses.map(course => {
            const assignment = assignmentMap.get(String(course.id));
            const staff = assignment && assignment.staffId ? staffMap.get(String(assignment.staffId).toLowerCase()) : null;
            return {
                id: course.id,
                assignmentId: assignment ? assignment.id : null,
                itemName: course.name,
                courseFaculty: course.faculty,
                courseDay: course.daysOfWeek.join(' / '),
                courseTime: formatDate(course.startDate, 'h:mm') + ' - ' + formatDate(course.endDate, 'h:mm aa'),
                location: course.location,
                staffName: staff ? staff.name : "Unassigned",
                staffId: staff ? staff.id : null
            };
        });
        
        const mstStaffList = allStaff.filter(s => s.role && s.role.toLowerCase().includes('mst')).map(s => ({ id: s.id, name: s.name }));

        return { success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } };
    } catch (e) {
        console.error("Error in getMstViewData: " + e.stack);
        return { success: false, error: e.message };
    }
}

/**
 * Internal function to get a list of calendars the user can write to.
 */
function getWritableCalendarsInternal() {
  try {
    const allCals = Calendar.CalendarList.list({ showDeleted: false, minAccessRole: 'writer' });
    if (!allCals || !allCals.items) {
      return { success: true, data: [] }; // No calendars found, not an error
    }
    const writableCals = allCals.items.map(cal => ({ id: cal.id, name: cal.summary }));
    return { success: true, data: writableCals };
  } catch (e) {
    console.error('getWritableCalendarsInternal Error: ' + e.stack);
    return { success: false, message: 'Failed to fetch Google Calendar list: ' + e.message };
  }
}