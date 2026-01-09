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
        // Fetch data from various helper functions/controllers
        const userInfo = getUserInfo(); // Internal function
        const mstViewData = getMstViewData(); // Internal function

        // We can add other data fetches here as the app grows
        // const techHubData = getTechHubViewData();

        return {
            success: true,
            data: {
                userInfo: userInfo.data,
                mstData: mstViewData.success ? mstViewData.data : { error: mstViewData.error },
                // techHubData: techHubData.success ? techHubData.data : { error: techHubData.error }
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
 * Internal function to fetch user information.
 * @returns {object} An object containing the user's email and a photo URL.
 */
function getUserInfo() {
    try {
        const email = Session.getActiveUser().getEmail();
        const photoUrl = Session.getActiveUser().getPhotoUrl();

        return {
            success: true,
            data: {
                email: email,
                photoUrl: photoUrl
            }
        };
    } catch (e) {
        console.error("getUserInfo Error: " + e.stack);
        return {
            success: false,
            message: "Could not retrieve user information.",
            data: {
                email: "Error loading user",
                photoUrl: ""
            }
        };
    }
}

/**
 * Internal function to fetch MST view data.
 * This is a direct copy of the logic from api_getMstViewData.
 * @returns {object} The MST view data or an error object.
 */
function getMstViewData() {
    try {
        var staffData = getSheet('Staff_List').getDataRange().getValues();
        var assignmentData = getSheet('Staff_Assignments').getDataRange().getValues();
        var courseData = getSheet('Course_Schedule').getDataRange().getValues();

        var staffHeaders = getColumnMap(staffData[0]);
        var assignmentHeaders = getColumnMap(assignmentData[0]);
        var courseHeaderRow = courseData.find(function(row) { return row.join('').toLowerCase().includes('eventid'); });
        if (!courseHeaderRow) throw new Error("Could not find header row in Course Schedule sheet. Please ensure 'eventID' column exists.");
        var courseHeaders = getColumnMap(courseHeaderRow);
        var courseHeaderIndex = courseData.indexOf(courseHeaderRow);

        var allStaff = staffData.slice(1).map(function(row) { return parseStaff(row, staffHeaders); }).filter(function(s) { return s && s.isActive; });
        var allAssignments = assignmentData.slice(1).map(function(row) { return parseAssignment(row, assignmentHeaders); }).filter(Boolean);
        var allCourses = courseData.slice(courseHeaderIndex + 1).map(function(row) { return parseCourse(row, courseHeaders); }).filter(Boolean);

        var staffMap = new Map(allStaff.map(function(s) { return [String(s.id).toLowerCase(), s]; }));
        var assignmentMap = new Map(allAssignments.map(function(a) { return [String(a.eventId), a]; }));

        var courseAssignmentsView = allCourses.map(function(course) {
            var assignment = assignmentMap.get(String(course.id));
            var staff = assignment && assignment.staffId ? staffMap.get(String(assignment.staffId).toLowerCase()) : null;
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
        
        var mstStaffList = allStaff.filter(function(s) { return s.role && s.role.toLowerCase().includes('mst'); }).map(function(s) { return { id: s.id, name: s.name }; });

        return { 
            success: true, 
            data: {
                courseAssignments: courseAssignmentsView,
                mstStaffList: mstStaffList
            }
        };

    } catch (e) {
        console.error("Error in getMstViewData: " + e.stack);
        return { success: false, error: e.message };
    }
}
