/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING
 * Handles Course Schedule, Assignments, and MST Logic
 * -------------------------------------------------------------------
 */

// --- VIEW MODEL GENERATOR ---

function getMSTViewData(master) {
    const courseItems = [];
    const courseAssignments = [];
    let mstStaffList = [];
    let debugMsg = "";

    try {
        // 1. Build MST Staff List
        const allStaffObjects = Object.values(master.staffMap);
        mstStaffList = allStaffObjects
            .filter(s => (s.Roles || '').toLowerCase().includes('mst') && s.IsActive !== 'FALSE')
            .map(s => ({ id: s.StaffID, name: s.FullName }))
            .sort((a, b) => a.name.localeCompare(b.name));

        // 2. Check Data
        if (master.courseData.length <= 1) {
            debugMsg = "Course_Schedule tab is empty or missing.";
            return { courseItems, courseAssignments, mstStaffList, debugMsg };
        }

        // 3. Create Email -> Name Map
        const emailToNameMap = {};
        if (master.staffData.length > 1) {
            const sHeaders = master.staffData[0].map(normalizeHeader);
            const sNameIdx = sHeaders.indexOf('fullname');
            const sIdIdx = sHeaders.findIndex(h => h.includes('id') || h.includes('email'));
            if (sNameIdx > -1 && sIdIdx > -1) {
                for (let i = 1; i < master.staffData.length; i++) {
                    const email = (master.staffData[i][sIdIdx] || '').toLowerCase().trim();
                    if (email) emailToNameMap[email] = master.staffData[i][sNameIdx];
                }
            }
        }

        // 4. Map Columns (STRICT MODE)
        const cHeaders = master.courseData[0].map(normalizeHeader);
        
        const idx = { 
            uid: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.COURSE_UNIQUE_ID), // eventID
            code: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.COURSE_CODE),     // Course (HUS 236)
            faculty: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.FACULTY), 
            day: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.DAY), 
            time: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.TIME),            // Run Time
            loc: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.LOCATION),         // BX Location
            mst: findColumnIndex(cHeaders, CONFIG.MST.HEADERS.ASSIGNED_STAFF)    // MST Assigned by email
        };
        
        if (idx.uid === -1) debugMsg = "Could not find 'eventID' column.";
        if (idx.mst === -1) debugMsg += " Warning: 'MST Assigned by email' column not found.";

        // 5. Iterate Rows
        for (let i = 1; i < master.courseData.length; i++) {
            const row = master.courseData[i];
            const uniqueId = (idx.uid > -1) ? row[idx.uid] : null;
            
            if(uniqueId) {
                // Build Dropdown Item
                const info = { 
                    code: (idx.code > -1) ? row[idx.code] : 'Unknown', 
                    faculty: (idx.faculty > -1) ? row[idx.faculty] : '', 
                    day: (idx.day > -1) ? row[idx.day] : '', 
                    time: (idx.time > -1) ? row[idx.time] : '',
                    loc: (idx.loc > -1) ? row[idx.loc] : ''
                };
                
                // Label: HUS 236 | Smith | Mon | 9am-10am | Room 101
                const label = `${info.code} | ${info.faculty} | ${info.day} | ${info.time} | ${info.loc}`;
                
                // IMPORTANT: We use uniqueId (eventID) as the ID, but show the label
                courseItems.push({ id: uniqueId, name: label, type: 'Course' });

                // Build Assignment Item
                const assignedEmail = (idx.mst > -1) ? (row[idx.mst] || '').trim() : '';
                let staffName = "Unassigned";
                if (assignedEmail) {
                    const cleanEmail = assignedEmail.toLowerCase();
                    staffName = emailToNameMap[cleanEmail] || assignedEmail; 
                }

                const timeStr = (idx.time > -1) ? row[idx.time] : '';
                const duration = calculateCourseDuration(timeStr);

                courseAssignments.push({
                    id: uniqueId, // Use eventID for logic
                    staffName: staffName,
                    itemName: info.code, // Display "HUS 236" in the card
                    courseFaculty: info.faculty,
                    courseDay: info.day,
                    courseTime: timeStr,
                    duration: duration,
                    location: info.loc
                });
            }
        }

    } catch (e) {
        debugMsg = "Error in MST View: " + e.message;
    }

    return { courseItems, courseAssignments, mstStaffList, debugMsg };
}

// --- ACTIONS ---

function saveNewAssignment(staffId, courseId, itemType) {
    try {
        const ss = getMasterDataHub();
        let sheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
        if (!sheet) sheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE.replace('_', ' '));
        if (!sheet) throw new Error("Course_Schedule tab not found.");

        // Get Staff Email
        const staffSheet = ss.getSheetByName(CONFIG.TABS.STAFF_LIST) || ss.getSheetByName('Staff_List');
        const staffData = staffSheet.getDataRange().getValues();
        const sHeaders = staffData[0].map(h => h.toString().toLowerCase());
        const sIdIdx = sHeaders.findIndex(h => h.includes('id') || h.includes('email'));
        
        let staffEmail = staffId; 
        if (sIdIdx > -1) {
            for (let i = 1; i < staffData.length; i++) {
                if (staffData[i][sIdIdx] === staffId) {
                    staffEmail = staffData[i][sIdIdx]; 
                    break;
                }
            }
        }

        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(normalizeHeader);
        
        // Find Columns using Config
        const idIdx = findColumnIndex(headers, CONFIG.MST.HEADERS.COURSE_UNIQUE_ID); // Look for eventID
        const mstIdx = findColumnIndex(headers, CONFIG.MST.HEADERS.ASSIGNED_STAFF);  // Look for MST Assigned by email

        if (idIdx === -1) throw new Error("Could not find 'eventID' column.");
        
        let targetCol = mstIdx + 1;
        if (mstIdx === -1) {
            targetCol = sheet.getLastColumn() + 1;
            sheet.getRange(1, targetCol).setValue("MST Assigned by email");
        }

        // Search for the eventID
        for (let i = 1; i < data.length; i++) {
            if (data[i][idIdx] == courseId) {
                sheet.getRange(i + 1, targetCol).setValue(staffEmail);
                return getSchedulingRosterData();
            }
        }
        throw new Error("Course Event ID not found.");

    } catch (e) { return { error: e.message }; }
}

function api_unassignCourse(courseId) {
    try {
        const ss = getMasterDataHub();
        let sheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
        if (!sheet) sheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE.replace('_', ' '));
        if (!sheet) throw new Error("Course_Schedule tab not found.");

        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(normalizeHeader);
        
        const idIdx = findColumnIndex(headers, CONFIG.MST.HEADERS.COURSE_UNIQUE_ID);
        const mstIdx = findColumnIndex(headers, CONFIG.MST.HEADERS.ASSIGNED_STAFF);

        if (idIdx === -1 || mstIdx === -1) throw new Error("Could not find eventID or MST Assigned columns.");

        for (let i = 1; i < data.length; i++) {
            if (data[i][idIdx] == courseId) {
                sheet.getRange(i + 1, mstIdx + 1).clearContent();
                return getSchedulingRosterData();
            }
        }
        throw new Error("Course Event ID not found.");

    } catch (e) { return { error: e.message }; }
}

function api_exportCourseAssignments(data) {
    try {
        if (!data || data.length === 0) throw new Error("No data to export.");
        
        const ss = SpreadsheetApp.create(`MST Course Assignments ${new Date().toISOString().split('T')[0]}`);
        const sheet = ss.getActiveSheet();
        
        const headers = ["Assigned Staff", "Course", "Faculty", "Day", "Time", "Duration", "Classroom"];
        const rows = data.map(c => [
            c.staffName,
            c.itemName,
            c.courseFaculty,
            c.courseDay,
            c.courseTime,
            c.duration,
            c.location
        ]);

        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e0e0");
        if (rows.length > 0) {
            sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        }
        
        return { success: true, url: ss.getUrl() };
    } catch (e) { return { success: false, message: e.message }; }
}

// --- HELPERS ---

function calculateCourseDuration(timeStr) {
    if (!timeStr || !timeStr.includes('-')) return "";
    try {
        const parts = timeStr.split('-').map(s => s.trim());
        if (parts.length !== 2) return "";

        const parseToMinutes = (t) => {
            const match = t.match(/(\d+):(\d+)\s*(AM|PM)?/i);
            if (!match) return 0;
            let h = parseInt(match[1]);
            const m = parseInt(match[2]);
            const ampm = match[3] ? match[3].toUpperCase() : null;
            
            if (ampm === 'PM' && h < 12) h += 12;
            if (ampm === 'AM' && h === 12) h = 0;
            return h * 60 + m;
        };

        const startMins = parseToMinutes(parts[0]);
        const endMins = parseToMinutes(parts[1]);
        
        let diff = endMins - startMins;
        if (diff < 0) diff += 1440; 

        const h = Math.floor(diff / 60);
        const m = diff % 60;
        
        if (h > 0 && m > 0) return `${h}h ${m}m`;
        if (h > 0) return `${h}h`;
        return `${m}m`;

    } catch (e) { return ""; }
}

// Helper: Stricter Column Finder
function findColumnIndex(headers, possibleNames) {
    // 1. Try Exact Match first
    for (const name of possibleNames) {
        const idx = headers.indexOf(name);
        if (idx > -1) return idx;
    }
    // 2. Try "Starts With"
    for (const name of possibleNames) {
        const idx = headers.findIndex(h => h.startsWith(name));
        if (idx > -1) return idx;
    }
    // 3. Fallback to "Includes"
    for (const name of possibleNames) {
        const idx = headers.findIndex(h => h.includes(name));
        if (idx > -1) return idx;
    }
    return -1;
}