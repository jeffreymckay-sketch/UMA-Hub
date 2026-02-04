
/**
 * -------------------------------------------------------------------
 * SCHEDULING CONTROLLER
 * -------------------------------------------------------------------
 * This controller is responsible for loading, parsing, and combining
 * all data related to staff scheduling for Courses and Tech Hub shifts.
 * It leverages the data models in 'Model.gs' and parsing functions
 * in 'Parsers.gs' to create a clean and structured data set for the frontend.
 * -------------------------------------------------------------------
 */

/**
 * Fetches scheduling data specifically for the MST Classroom Management view.
 * This is a streamlined version of getSchedulingData() that only loads course data.
 * @returns {object} A response object with success status and data or an error.
 */
function api_getMstSchedulingData() {
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
                startDate: formatDate(course.startDate, 'yyyy-MM-dd'),
                endDate: formatDate(course.endDate, 'yyyy-MM-dd'),
                location: course.location,
                duration: calculateDuration(course.startDate, course.endDate),
                link: course.zoomLink,
                staffName: staff ? staff.name : "Unassigned",
                staffId: staff ? staff.id : null,
                session: course.session,
                runTime: course.runTime,
                timeOfDay: course.timeOfDay
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
        console.error("Error in api_getMstSchedulingData: " + e.stack);
        return { success: false, error: e.message };
    }
}

/**
 * Main entry point for the frontend to get all scheduling data.
 * This function orchestrates the loading and processing of data from
 * multiple spreadsheet tabs.
 * @returns {object} A response object with success status and data or an error.
 */
function getSchedulingData() {
    try {
        var staffData = getSheet('Staff_List').getDataRange().getValues();
        var assignmentData = getSheet('Staff_Assignments').getDataRange().getValues();
        var courseData = getSheet('Course_Schedule').getDataRange().getValues();
        var shiftData = getSheet('TechHub_Shifts').getDataRange().getValues();

        var staffHeaders = getColumnMap(staffData[0]);
        var assignmentHeaders = getColumnMap(assignmentData[0]);
        var courseHeaderRow = courseData.find(function(row) { return row.join('').toLowerCase().includes('eventid'); });
        if (!courseHeaderRow) throw new Error("Could not find header row in Course Schedule sheet. Please ensure 'eventID' column exists.");
        var courseHeaders = getColumnMap(courseHeaderRow);
        var shiftHeaders = getColumnMap(shiftData[0]);

        var courseHeaderIndex = courseData.indexOf(courseHeaderRow);

        var allStaff = staffData.slice(1).map(function(row) { return parseStaff(row, staffHeaders); }).filter(function(s) { return s && s.isActive; });
        var allAssignments = assignmentData.slice(1).map(function(row) { return parseAssignment(row, assignmentHeaders); }).filter(Boolean);
        var allCourses = courseData.slice(courseHeaderIndex + 1).map(function(row) { return parseCourse(row, courseHeaders); }).filter(Boolean);
        var allShifts = shiftData.slice(1).map(function(row) { return parseShift(row, shiftHeaders); }).filter(Boolean);

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
                startDate: formatDate(course.startDate, 'yyyy-MM-dd'),
                endDate: formatDate(course.endDate, 'yyyy-MM-dd'),
                location: course.location,
                duration: calculateDuration(course.startDate, course.endDate),
                link: course.zoomLink,
                staffName: staff ? staff.name : "Unassigned",
                staffId: staff ? staff.id : null,
                session: course.session,
                runTime: course.runTime,
                timeOfDay: course.timeOfDay
            };
        });

        var shiftsView = allShifts.map(function(shift) {
            var assignment = assignmentMap.get(String(shift.id));
            var staff = assignment && assignment.staffId ? staffMap.get(String(assignment.staffId).toLowerCase()) : null;
            return {
                shiftId: shift.id,
                description: shift.name,
                day: formatDate(shift.startDate, 'EEEE'), // e.g., 'Monday'
                start: formatDate(shift.startDate, 'h:mm aa'),
                end: formatDate(shift.endDate, 'h:mm aa'),
                zoom: shift.location === 'Zoom',
                assignedStaffId: staff ? staff.id : null,
                smartStaffList: [] // Placeholder for availability logic
            };
        });
        
        var mstStaffList = allStaff.filter(function(s) { return s.role && s.role.toLowerCase().includes('mst'); }).map(function(s) { return { id: s.id, name: s.name }; });

        var uniqueSessions = _getUniqueColumnValues(courseData, courseHeaders, 'session');
        var uniqueDays = _getUniqueColumnValues(courseData, courseHeaders, 'day');
        var uniqueTimesOfDay = _getUniqueColumnValues(courseData, courseHeaders, 'timeofday');
        var uniqueLocations = _getUniqueColumnValues(courseData, courseHeaders, 'bxlocation');

        return {
            success: true,
            data: {
                roster: shiftsView,
                courseAssignments: courseAssignmentsView,
                mstStaffList: mstStaffList,
                formOptions: {
                    sessions: uniqueSessions,
                    days: uniqueDays,
                    timesOfDay: uniqueTimesOfDay,
                    locations: uniqueLocations
                }
            }
        };

    } catch (e) {
        console.error("Error in getSchedulingData: " + e.stack);
        return { success: false, error: e.message };
    }
}

function api_updateCourseAssignment(courseId, newStaffId) {
    try {
        var sheet = getSheet('Staff_Assignments');

        var data = sheet.getDataRange().getValues();
        var headers = getColumnMap(data[0]);

        var itemIdCol = headers['referenceid'];
        var staffIdCol = headers['staffid'];
        var assignmentIdCol = headers['assignmentid'];
        var typeCol = headers['assignmenttype'];

        if (itemIdCol === undefined || staffIdCol === undefined || assignmentIdCol === undefined || typeCol === undefined) {
            throw new Error("One or more required columns (referenceid, staffid, assignmentid, assignmenttype) are missing in the assignments sheet.");
        }

        var existingRowIndex = -1;
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][itemIdCol]) == String(courseId)) {
                existingRowIndex = i + 1;
                break;
            }
        }

        if (existingRowIndex !== -1) {
            if (newStaffId) {
                sheet.getRange(existingRowIndex, staffIdCol + 1).setValue(newStaffId);
            } else {
                sheet.deleteRow(existingRowIndex);
            }
        } else {
            if (newStaffId) {
                var newAssignmentId = Utilities.getUuid();
                var newRow = new Array(data[0].length).fill('');
                newRow[assignmentIdCol] = newAssignmentId;
                newRow[staffIdCol] = newStaffId;
                newRow[itemIdCol] = courseId;
                newRow[typeCol] = 'Course';
                sheet.appendRow(newRow);
            }
        }

        return { success: true };

    } catch (e) {
        console.error("Error in api_updateCourseAssignment: " + e.stack);
        return { success: false, error: e.message };
    }
}

function api_deleteCourse(courseId) {
    try {
        var courseSheet = getSheet('Course_Schedule');
        var courseData = courseSheet.getDataRange().getValues();
        var courseHeaderRow = courseData.find(function(row) { return row.join('').toLowerCase().includes('eventid'); });
        if (!courseHeaderRow) throw new Error("Could not find header row in Course Schedule sheet.");
        var courseHeaders = getColumnMap(courseHeaderRow);
        var eventIdCol = courseHeaders['eventid'];

        if (eventIdCol === undefined) {
            throw new Error("Column 'eventID' not found in the Course Schedule sheet.");
        }

        for (var i = courseData.length - 1; i >= 1; i--) {
            if (String(courseData[i][eventIdCol]) === String(courseId)) {
                courseSheet.deleteRow(i + 1);
                break; 
            }
        }

        var assignmentSheet = getSheet('Staff_Assignments');
        var assignmentData = assignmentSheet.getDataRange().getValues();
        var assignmentHeaders = getColumnMap(assignmentData[0]);
        var refIdCol = assignmentHeaders['referenceid'];

        if (refIdCol !== undefined) {
            for (var i = assignmentData.length - 1; i >= 1; i--) {
                if (String(assignmentData[i][refIdCol]) === String(courseId)) {
                    assignmentSheet.deleteRow(i + 1);
                    break;
                }
            }
        }

        return { success: true };

    } catch (e) {
        console.error("Error in api_deleteCourse: " + e.stack);
        return { success: false, error: e.message };
    }
}

function api_addCourse(courseDetails) {
    try {
        var sheet = getSheet('Course_Schedule');

        var headerRow = sheet.getDataRange().getValues().find(function(row) { return row.join('').toLowerCase().includes('eventid'); });
        if (!headerRow) throw new Error("Could not find header row in Course Schedule sheet.");
        var headers = getColumnMap(headerRow);
        
        var newEventId = Utilities.getUuid();
        var newRow = new Array(Object.keys(headers).length).fill('');

        newRow[headers['session']] = courseDetails.session;
        newRow[headers['startdate']] = courseDetails.startDate;
        newRow[headers['enddate']] = courseDetails.endDate;
        newRow[headers['day']] = courseDetails.day;
        newRow[headers['course']] = courseDetails.course;
        newRow[headers['faculty']] = courseDetails.faculty;
        newRow[headers['runtime']] = courseDetails.runTime;
        newRow[headers['timeofday']] = courseDetails.timeOfDay;
        newRow[headers['bxlocation']] = courseDetails.bxLocation;
        newRow[headers['zoomlink']] = courseDetails.zoomLink;
        newRow[headers['eventid']] = newEventId;
        
        sheet.appendRow(newRow);

        if (courseDetails.mstAssignedByEmail) {
            api_updateCourseAssignment(newEventId, courseDetails.mstAssignedByEmail);
        }

        return { success: true, newEventId: newEventId };

    } catch (e) {
        console.error("Error in api_addCourse: " + e.stack);
        return { success: false, error: e.message };
    }
}

function api_updateCourse(courseDetails) {
    try {
        var sheet = getSheet('Course_Schedule');

        var data = sheet.getDataRange().getValues();
        var headerRow = data.find(function(row) { return row.join('').toLowerCase().includes('eventid'); });
        if (!headerRow) throw new Error("Could not find header row in Course Schedule sheet.");
        var headers = getColumnMap(headerRow);
        var eventIdCol = headers['eventid'];

        if (eventIdCol === undefined) {
            throw new Error("Column 'eventID' not found in the Course Schedule sheet.");
        }

        var existingRowIndex = -1;
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][eventIdCol]) == String(courseDetails.id)) {
                existingRowIndex = i + 1; 
                break;
            }
        }

        if (existingRowIndex === -1) {
            throw new Error("Could not find the course to update.");
        }

        var rowValues = data[existingRowIndex - 1];
        rowValues[headers['session']] = courseDetails.session;
        rowValues[headers['startdate']] = courseDetails.startDate;
        rowValues[headers['enddate']] = courseDetails.endDate;
        rowValues[headers['day']] = courseDetails.day;
        rowValues[headers['course']] = courseDetails.course;
        rowValues[headers['faculty']] = courseDetails.faculty;
        rowValues[headers['runtime']] = courseDetails.runTime;
        rowValues[headers['timeofday']] = courseDetails.timeOfDay;
        rowValues[headers['bxlocation']] = courseDetails.bxLocation;
        rowValues[headers['zoomlink']] = courseDetails.zoomLink;
        
        sheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);

        api_updateCourseAssignment(courseDetails.id, courseDetails.mstAssignedByEmail);

        return { success: true };

    } catch (e) {
        console.error("Error in api_updateCourse: " + e.stack);
        return { success: false, error: e.message };
    }
}


function api_exportCourseAssignments(courseAssignments) {
    if (!courseAssignments || !Array.isArray(courseAssignments) || courseAssignments.length === 0) {
        return { success: false, message: "No assignment data provided to export." };
    }

    try {
        var ss = SpreadsheetApp.create('Course Assignments Export - ' + new Date().toLocaleDateString());
        var sheet = ss.getSheets()[0];
        sheet.setName("Assignments");

        var headerMap = {
            "staffName": "Assigned Staff",
            "itemName": "Course",
            "courseFaculty": "Faculty",
            "courseDay": "Day",
            "courseTime": "Time",
            "duration": "Duration",
            "location": "Classroom",
            "link": "Zoom Link"
        };
        
        var orderedHeaders = Object.keys(headerMap);
        var displayHeaders = orderedHeaders.map(function(h) { return headerMap[h]; });

        var dataRows = courseAssignments.map(function(assignment) {
            return orderedHeaders.map(function(header) { return assignment[header] || ""; });
        });

        var exportData = [displayHeaders].concat(dataRows);

        sheet.getRange(1, 1, exportData.length, displayHeaders.length).setValues(exportData);

        for (var i = 1; i <= displayHeaders.length; i++) {
            sheet.autoResizeColumn(i);
        }

        return { success: true, url: ss.getUrl() };

    } catch (e) {
        console.error("Error in api_exportCourseAssignments: " + e.stack);
        return { success: false, error: 'Export failed: ' + e.message };
    }
}

// --- HELPERS ---

function _getUniqueColumnValues(data, headers, headerName) {
    var columnIndex = headers[headerName];
    if (columnIndex === undefined) return [];
    
    var valueSet = new Set();
    data.slice(1).forEach(function(row) {
        var value = row[columnIndex];
        if (value) { 
            valueSet.add(value.toString().trim());
        }
    });
    
    return Array.from(valueSet).sort();
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

function formatDate(date, format) {
    if (!date || !(date instanceof Date)) return '';
    return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}

function calculateDuration(start, end) {
    if (!start || !end) return '';
    var diffMs = end.getTime() - start.getTime();
    var diffHours = diffMs / (1000 * 60 * 60);
    return diffHours.toFixed(1) + ' hrs';
}
