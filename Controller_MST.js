/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING & SETTINGS (Corrected for Data Handling)
 * -------------------------------------------------------------------
 */

// --- SETTINGS API ---

function api_getMstSettings() {
  try {
    const allSettings = getSettings();
    let mstSettings = {};
    // FIX: Settings are stored as a JSON string; they must be parsed.
    if (allSettings.mstSettings) {
      try {
        mstSettings = JSON.parse(allSettings.mstSettings);
      } catch (e) { /* Ignore parsing errors, default to empty object */ }
    }

    const ss = getMasterDataHub();
    const sheetNames = ss.getSheets().map(s => s.getName());
    const data = {
      settings: mstSettings, // This is now a proper object
      sheetNames: sheetNames
    };
    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function api_saveMstSettings(newMstSettings) {
  try {
    const allSettings = getSettings();
    let currentMstSettings = {};
    // FIX: Retrieve the existing settings string and parse it before merging.
    if (allSettings.mstSettings) {
      try {
        currentMstSettings = JSON.parse(allSettings.mstSettings);
      } catch (e) { /* Default to empty if corrupt */ }
    }

    // Merge new settings into the existing ones
    const updatedMstSettings = Object.assign(currentMstSettings, newMstSettings);

    // FIX: Stringify the settings object back into a JSON string for storage.
    allSettings.mstSettings = JSON.stringify(updatedMstSettings);
    
    saveSettings(allSettings);
    return { success: true, message: "MST settings saved successfully." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- CALENDAR SYNC API ---

function api_runMstCalendarSync(targetCalendarId, isPreview) {
  const lock = LockService.getScriptLock();
  let originalSourceTab;
  try {
    lock.waitLock(20000);
    const allSettings = getSettings();
    originalSourceTab = allSettings.sourceTabName;

    let mstSettings = {};
    // FIX: Parse the stored JSON string to get the usable settings object.
    if (allSettings.mstSettings) {
      try {
        mstSettings = JSON.parse(allSettings.mstSettings);
      } catch (e) { /* Default to empty object */ }
    }

    // FIX: Check the parsed object for the property.
    if (!mstSettings || !mstSettings.sourceTabName) {
      return { success: false, message: "MST Source Tab not configured. Please set it in the MST Settings tab first." };
    }

    allSettings.sourceTabName = mstSettings.sourceTabName;
    saveSettings(allSettings);
    
    const result = core_syncLogic('Course', isPreview, null, targetCalendarId);
    return result;

  } catch (e) {
    console.error("Error in MST Calendar Sync: " + e.message);
    return { success: false, message: "A critical error occurred during the sync: " + e.message };
  } finally {
    if (originalSourceTab !== undefined) {
        const finalSettings = getSettings();
        finalSettings.sourceTabName = originalSourceTab;
        saveSettings(finalSettings);
    }
    lock.releaseLock();
  }
}

// --- VIEW MODEL GENERATOR (No changes needed here) ---

function api_getMstViewData() {
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
        console.error("Error in api_getMstViewData: " + e.stack);
        return { success: false, error: e.message };
    }
}
