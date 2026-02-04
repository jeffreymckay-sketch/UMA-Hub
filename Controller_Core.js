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

        // 1. User Info & Permissions Lookup
        try {
            const userEmail = Session.getActiveUser().getEmail().toLowerCase();
            let userRole = "Staff"; // Default fallback
            
            const staffSheet = getSheet('Staff_List');
            if (staffSheet) {
                const staffData = staffSheet.getDataRange().getValues();
                const headers = getColumnMap(staffData[0]);
                const userRow = staffData.find(row => String(row[headers.staffid]).toLowerCase() === userEmail);
                
                if (userRow) {
                    userRole = userRow[headers.roles] || "Staff";
                }
            }
            userInfo = { success: true, data: { email: userEmail, role: userRole, photoUrl: "" } };
        } catch (e) { 
            userInfo = { success: true, data: { email: Session.getActiveUser().getEmail(), role: "Guest", error: "User lookup failed" } }; 
        }

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
            const assignSheet = getSheet('Staff_Assignments');
            if (assignSheet) {
                const assignmentValues = assignSheet.getDataRange().getValues();
                const assignmentHeaders = getColumnMap(assignmentValues[0]);
                const allAssignments = assignmentValues.slice(1).map(row => parseAssignment(row, assignmentHeaders)).filter(Boolean);
                assignmentsData = { success: true, data: allAssignments };
            }
        } catch (e) { assignmentsData = { success: false, message: "Assignments Load Error" }; }

        // 5. MST Data
        try {
            if (typeof getMstViewData === 'function') {
                const assignmentsForMst = assignmentsData.success ? assignmentsData.data : [];
                var mstRes = getMstViewData(sheetTabs, assignmentsForMst);
                mstViewData = mstRes.success ? mstRes.data : { error: mstRes.error };
            }
        } catch (e) { mstViewData = { error: "MST Data Error" }; }

        // 6. Nursing Data
        try {
            if (typeof api_getNursingData === 'function') {
                nursingData = api_getNursingData(sheetTabs);
            }
        } catch (e) { nursingData = { success: false, message: "Nursing Load Error" }; }

        return {
            success: true,
            data: {
                userInfo: userInfo.data,
                mstData: mstViewData,
                assignments: assignmentsData.success ? assignmentsData.data : [],
                settings: allSettings,
                writableCalendars: writableCalendars,
                allCalendars: allCalendars, 
                nursingData: nursingData,
                sheetTabs: sheetTabs
            }
        };

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
        
        const mstStaffList = allStaff.filter(s => s.role && s.role.toLowerCase().includes('mst')).map(s => ({ id: s.id, name: s.name }));

        return { success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } };
    } catch (e) {
        console.error("Error in getMstViewData: " + e.stack);
        return { success: false, error: e.message };
    }
}

// --- PARSING HELPERS ---

function parseStaff(row, map) {
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
    const idIdx = map['eventid'];
    const nameIdx = map['course']; 
    const facultyIdx = map['faculty'];
    const daysIdx = map['day']; 
    const runTimeIdx = map['runtime']; 
    const locIdx = map['bxlocation']; 
    
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
        timeString: runTimeIdx !== undefined ? row[runTimeIdx] : '', 
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

function getSheet(sheetKey) {
    try {
        const settings = getSettings();
        const sheetTabsJSON = settings.sheetTabs;
        if (!sheetTabsJSON) return null;

        const sheetTabs = JSON.parse(sheetTabsJSON);
        const sheetName = sheetTabs[sheetKey];
        if (!sheetName) return null;

        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(sheetName);
        return sheet; 
    } catch (e) {
        return null;
    }
}

function getColumnMap(headers) {
    var map = {};
    if (!headers) return map;
    headers.forEach(function(header, index) {
        var normalizedHeader = String(header).toLowerCase().replace(/[\s_]/g, '');
        if (normalizedHeader) map[normalizedHeader] = index;
    });
    return map;
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