/**
 * -------------------------------------------------------------------
 * CORE CONTROLLER
 * Handles essential, app-wide data and actions.
 * -------------------------------------------------------------------
 */

/**
 * Fetches all necessary data for the initial application load.
 */
function api_getInitialAppData() {
    var userInfo = { success: false, message: "Not loaded" };
    var mstViewData = { success: false, error: "Not loaded" };
    var assignmentsData = { success: false, message: "Not loaded" };
    var allSettings = { success: false, message: "Not loaded" };
    var writableCalendars = { success: false, message: "Not loaded" };
    var allCalendars = { success: false, message: "Not loaded" }; 
    var nursingData = { success: false, message: "Not loaded" };
    var sheetTabs = {};

    try {
        // 0. Get Sheet Names
        try {
            const settings = getSettings();
            if (settings && settings.sheetTabs) {
                sheetTabs = JSON.parse(settings.sheetTabs);
            } else {
                throw new Error("'sheetTabs' not found in settings. Please run admin_mapSheetTabs.");
            }
        } catch(e) {
             return { success: false, message: "Failed to load critical settings: " + e.message };
        }

        // 1. User Info
        try {
            var userRes = getUserInfo();
            userInfo = userRes.success ? userRes.data : { error: userRes.message };
        } catch (e) { userInfo = { error: "Failed to load user info." }; }

        // 2. Settings
        try {
            var settingsRes = getAllSettings_();
            allSettings = settingsRes.success ? settingsRes.data : { error: settingsRes.message };
        } catch (e) { allSettings = { error: "Failed to load settings." }; }

        // 3. Calendars
        try {
            var calRes = getWritableCalendarsInternal();
            writableCalendars = calRes.success ? calRes.data : { error: calRes.message };
            
            var allCalRes = getAllCalendarsInternal();
            allCalendars = allCalRes.success ? allCalRes.data : { error: allCalRes.message };
        } catch (e) { writableCalendars = { error: "Calendar Error" }; }
        
        // 4. Assignments Data
        try {
            const assignSheet = getSheet(sheetTabs.Staff_Assignments);
            if (!assignSheet) {
                throw new Error("Sheet 'Staff_Assignments' not found.");
            }
            const assignmentValues = assignSheet.getDataRange().getValues();
            const assignmentHeaders = getColumnMap(assignmentValues[0]);
            const allAssignments = assignmentValues.slice(1).map(row => parseAssignment(row, assignmentHeaders)).filter(Boolean);
            assignmentsData = { success: true, data: allAssignments };

        } catch (e) {
            assignmentsData = { success: false, message: "Assignments Load Error: " + e.message };
        }

        // 5. MST Data
        try {
            if (typeof getMstViewData === 'function') {
                const assignmentsForMst = assignmentsData.success ? assignmentsData.data : [];
                var mstRes = getMstViewData(sheetTabs, assignmentsForMst);
                mstViewData = mstRes.success ? mstRes.data : { error: mstRes.error };
            }
        } catch (e) { mstViewData = { error: "MST Data Error: " + e.message }; }

        // 6. Nursing Data
        try {
            if (typeof api_getNursingData === 'function') {
                nursingData = api_getNursingData(sheetTabs);
            }
        } catch (e) { nursingData = { success: false, message: "Nursing Load Error: " + e.message }; }

        var payload = {
            success: true,
            data: {
                userInfo: userInfo,
                mstData: mstViewData,
                assignments: assignmentsData.success ? assignmentsData.data : [],
                settings: allSettings,
                writableCalendars: writableCalendars,
                allCalendars: allCalendars, 
                nursingData: nursingData,
                sheetTabs: sheetTabs
            }
        };

        try { JSON.stringify(payload); } 
        catch (jsonError) { throw new Error("Data Serialization Failed: " + jsonError.message); }

        return payload;

    } catch (e) {
        console.error("Critical api_getInitialAppData Error: " + e.stack);
        return { success: false, message: "Critical App Load Failure: " + e.message };
    }
}

function api_saveSettings(key, settingsObject) {
    try {
        if (!key || !settingsObject) throw new Error("Key and settings object required.");
        PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(settingsObject));
        return { success: true, message: 'Settings saved!', data: getAllSettings_().data };
    } catch (e) {
        return { success: false, message: 'Failed to save settings: ' + e.message };
    }
}

function api_getDashboardData() {
    try {
        const email = Session.getActiveUser().getEmail();
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');

        const availabilitySheet = getSheet(sheetTabs.Staff_Availability);
        if (!availabilitySheet) throw new Error("Sheet 'Staff_Availability' not found.");

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

        const preferencesSheet = getSheet(sheetTabs.Staff_Preferences);
        const userPreferences = {};
        if (preferencesSheet) {
            const preferencesData = preferencesSheet.getDataRange().getValues();
            for (let i = 1; i < preferencesData.length; i++) {
                if (preferencesData[i][0] === email) {
                    userPreferences[preferencesData[i][1]] = preferencesData[i][2];
                }
            }
        }

        return { success: true, data: { availability: userAvailability, preferences: userPreferences } };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

// --- MST DATA LOGIC ---

function getMstViewData(sheetTabs, allAssignments) {
    try {
        const staffSheet = getSheet(sheetTabs.Staff_List);
        const courseSheet = getSheet(sheetTabs.Course_Schedule);

        if (!staffSheet || !courseSheet) {
            return { success: false, error: "Missing required MST sheets." };
        }

        const staffData = staffSheet.getDataRange().getValues();
        const courseData = courseSheet.getDataRange().getValues();

        const staffHeaders = getColumnMap(staffData[0]);
        
        // Find header row for courses (Row 2 usually, based on "eventID")
        const courseHeaderRow = courseData.find(row => row.join('').toLowerCase().includes('eventid'));
        if (!courseHeaderRow) throw new Error("Missing 'eventID' header in Course Schedule.");
        const courseHeaders = getColumnMap(courseHeaderRow);
        const courseHeaderIndex = courseData.indexOf(courseHeaderRow);

        const allStaff = staffData.slice(1).map(row => parseStaff(row, staffHeaders)).filter(s => s && s.isActive);
        const allCourses = courseData.slice(courseHeaderIndex + 1).map(row => parseCourse(row, courseHeaders)).filter(Boolean);

        const staffMap = new Map(allStaff.map(s => [String(s.id).toLowerCase(), s]));
        const assignmentMap = new Map(allAssignments.map(a => [String(a.eventId), a]));

        const courseAssignmentsView = allCourses.map(course => {
            const assignment = assignmentMap.get(String(course.id));
            const staff = assignment && assignment.staffId ? staffMap.get(String(assignment.staffId).toLowerCase()) : null;
            
            // Use the "Run Time" string if available, otherwise fallback to dates
            let timeDisplay = course.timeString || "TBD";
            if (!course.timeString && course.startDate && course.endDate) {
                 const fmt = (d) => (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'h:mm a') : String(d);
                 timeDisplay = fmt(course.startDate) + ' - ' + fmt(course.endDate);
            }

            return {
                id: course.id,
                assignmentId: assignment ? assignment.id : null,
                itemName: course.name,
                courseFaculty: course.faculty,
                courseDay: course.daysOfWeek.join(' / '),
                courseTime: timeDisplay,
                location: course.location,
                staffName: staff ? staff.name : "Unassigned",
                staffId: staff ? staff.id : null
            };
        });
        
        // Filter staff for dropdown (Role contains MST)
        const mstStaffList = allStaff.filter(s => s.role && s.role.toLowerCase().includes('mst')).map(s => ({ id: s.id, name: s.name }));

        return { success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } };
    } catch (e) {
        console.error("Error in getMstViewData: " + e.stack);
        return { success: false, error: e.message };
    }
}

// --- PARSING HELPERS (Customized for your Headers) ---

function parseStaff(row, map) {
    // Headers: FullName, StaffID, Roles, IsActive
    const nameIdx = map['fullname'] !== undefined ? map['fullname'] : map['name'];
    const idIdx = map['staffid'] !== undefined ? map['staffid'] : map['id'];
    const roleIdx = map['roles'] !== undefined ? map['roles'] : map['role'];
    const activeIdx = map['isactive'] !== undefined ? map['isactive'] : map['active'];

    if (nameIdx === undefined) return null;

    return {
        id: row[idIdx],
        name: row[nameIdx],
        role: roleIdx !== undefined ? row[roleIdx] : '',
        isActive: activeIdx !== undefined ? (String(row[activeIdx]).toLowerCase() === 'true' || row[activeIdx] === true) : true
    };
}

function parseAssignment(row, map) {
    // Headers: AssignmentID, StaffID, AssignmentType, ReferenceID
    const idIdx = map['assignmentid'];
    const eventIdx = map['referenceid'];
    const staffIdx = map['staffid'];

    if (eventIdx === undefined || staffIdx === undefined) return null;

    return {
        id: row[idIdx],
        eventId: row[eventIdx],
        staffId: row[staffIdx]
    };
}

function parseCourse(row, map) {
    // Headers: Session, Start Date, End Date, Day, Course, Faculty, Run Time, Time of Day, BX Location, eventID
    const idIdx = map['eventid'];
    const nameIdx = map['course']; // Header is "Course"
    const facultyIdx = map['faculty'];
    const daysIdx = map['day']; // Header is "Day"
    const runTimeIdx = map['runtime']; // Header is "Run Time"
    const locIdx = map['bxlocation']; // Header is "BX Location"
    
    // Fallbacks for dates if needed
    const startIdx = map['startdate'];
    const endIdx = map['enddate'];

    if (idIdx === undefined || nameIdx === undefined) return null;

    let days = [];
    if (daysIdx !== undefined && row[daysIdx]) {
        days = String(row[daysIdx]).split(',').map(d => d.trim());
    }

    return {
        id: row[idIdx],
        name: row[nameIdx],
        faculty: facultyIdx !== undefined ? row[facultyIdx] : '',
        daysOfWeek: days,
        timeString: runTimeIdx !== undefined ? row[runTimeIdx] : '', // Capture "9:00 - 11:00" string
        startDate: startIdx !== undefined ? row[startIdx] : null,
        endDate: endIdx !== undefined ? row[endIdx] : null,
        location: locIdx !== undefined ? row[locIdx] : ''
    };
}

// --- SHARED HELPERS ---

function getAllSettings_() {
    try {
        const props = PropertiesService.getScriptProperties().getProperties();
        const parsed = {};
        for (const key in props) {
            try { parsed[key] = JSON.parse(props[key]); } 
            catch (e) { parsed[key] = props[key]; }
        }
        return { success: true, data: parsed };
    } catch (e) { return { success: false, message: e.message }; }
}

function getUserInfo() {
    try {
        return { success: true, data: { email: Session.getActiveUser().getEmail(), photoUrl: "" } };
    } catch (e) { return { success: false, message: e.message }; }
}

function getWritableCalendarsInternal() {
  try {
    const allCals = Calendar.CalendarList.list({ showDeleted: false, minAccessRole: 'writer' });
    if (!allCals || !allCals.items) return { success: true, data: [] };
    const writableCals = allCals.items.map(cal => ({ id: cal.id, name: cal.summary }));
    return { success: true, data: writableCals };
  } catch (e) {
    return { success: false, message: 'Failed to fetch calendars: ' + e.message };
  }
}

function getAllCalendarsInternal() {
  try {
    const allCals = CalendarApp.getAllCalendars();
    if (!allCals) return { success: true, data: [] };
    
    const mappedCals = allCals.map(cal => ({ 
        id: cal.getId(), 
        name: cal.getName(),
        isOwned: cal.isOwnedByMe()
    }));
    
    mappedCals.sort((a, b) => {
        if (a.isOwned && !b.isOwned) return -1;
        if (!a.isOwned && b.isOwned) return 1;
        return a.name.localeCompare(b.name);
    });

    return { success: true, data: mappedCals };
  } catch (e) {
    return { success: false, message: 'Failed to fetch all calendars: ' + e.message };
  }
}

function getSheet(sheetName) {
    try {
        let ss;
        const spreadsheetId = getSpreadsheetId();
        if (spreadsheetId && spreadsheetId !== 'PASTE_YOUR_ID_HERE') {
            try {
                ss = SpreadsheetApp.openById(spreadsheetId);
            } catch (e) {
                console.warn("Could not open spreadsheet by ID. Falling back to Active.");
            }
        }
        if (!ss) {
            ss = SpreadsheetApp.getActiveSpreadsheet();
        }
        if (!ss) throw new Error("Script is not bound to a spreadsheet and no ID provided.");
        
        const sheet = ss.getSheetByName(sheetName);
        return sheet; 
    } catch (e) {
        console.error(`Error getting sheet '${sheetName}': ${e.message}`);
        return null;
    }
}

function createDataMap(data, keyIndex) {
    const map = {};
    if (!data || data.length < 2) return map; 
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const key = row[keyIndex];
        if (key) {
            map[key] = row;
        }
    }
    return map;
}