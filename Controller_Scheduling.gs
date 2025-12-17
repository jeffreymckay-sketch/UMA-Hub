/**
 * -------------------------------------------------------------------
 * SCHEDULING CONTROLLER (Tech Hub & MST)
 * -------------------------------------------------------------------
 */

// --- ROSTER & SHIFT MANAGEMENT ---

function getSchedulingRosterData() {
  try {
    // 1. SMART SNAPSHOT: Fetch External Data if configured
    const syncResult = syncExternalCourseData(); 
    
    // If sync was attempted but failed, return the error to the UI
    if (syncResult && !syncResult.success) {
        return { error: "Sync Failed: " + syncResult.message };
    }
    
    SpreadsheetApp.flush();

    const ss = getMasterDataHub();
    const shiftsData = getRequiredSheetData(ss, CONFIG.TABS.TECH_HUB_SHIFTS);
    const assignmentsData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);
    const availData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_AVAILABILITY);
    const prefData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_PREFERENCES);
    const courseData = getRequiredSheetData(ss, CONFIG.TABS.COURSE_SCHEDULE);

    // 2. Parse Staff
    const staffList = [];
    for (let i = 1; i < staffData.length; i++) {
      if (staffData[i][1]) { 
        staffList.push({
          name: staffData[i][0],
          id: String(staffData[i][1]).trim(), 
          role: staffData[i][2],
          isActive: staffData[i][3]
        });
      }
    }

    // 3. Parse Shifts
    const roster = [];
    const manageShifts = [];
    
    for (let i = 1; i < shiftsData.length; i++) {
      const row = shiftsData[i];
      if (!row[0]) continue; 

      const shiftId = String(row[0]).trim(); 

      const shiftObj = {
        shiftId: shiftId,
        description: row[1],
        day: row[2],
        start: row[3],
        end: row[4],
        zoom: (row[5] === true || row[5] === "TRUE")
      };

      manageShifts.push({
        id: shiftId,
        desc: row[1],
        day: row[2],
        start: row[3],
        end: row[4],
        zoom: shiftObj.zoom
      });

      const assignment = assignmentsData.find(a => a[2] === 'Tech Hub' && String(a[3]).trim() === shiftId);
      shiftObj.assignedStaffId = assignment ? String(assignment[1]).trim() : "";

      shiftObj.smartStaffList = staffList.map(s => {
        let isAvailable = true;
        const unavail = availData.find(u => String(u[1]).trim() === s.id && u[2] === shiftObj.day);
        if (unavail) {
           if (timesOverlap(shiftObj.start, shiftObj.end, unavail[3], unavail[4])) {
             isAvailable = false;
           }
        }
        if (isAvailable) {
           const otherAssign = assignmentsData.find(a => String(a[1]).trim() === s.id && a[2] === 'Tech Hub' && String(a[3]).trim() !== shiftObj.shiftId);
           if (otherAssign) {
              const otherShift = shiftsData.find(sh => String(sh[0]).trim() === String(otherAssign[3]).trim());
              if (otherShift && otherShift[2] === shiftObj.day) {
                 if (timesOverlap(shiftObj.start, shiftObj.end, otherShift[3], otherShift[4])) {
                   isAvailable = false;
                 }
              }
           }
        }
        return { id: s.id, name: s.name, available: isAvailable };
      });

      roster.push(shiftObj);
    }

    // 4. Parse Course Assignments & Items (MST)
    const courseAssignments = [];
    const courseItems = []; 
    const courseAssigns = assignmentsData.filter(a => a[2] === 'Course');
    
    // --- SMART HEADER DETECTION ---
    let headerRowIndex = -1;
    for (let r = 0; r < Math.min(courseData.length, 5); r++) {
        const rowStr = courseData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        // Look for specific headers you provided
        if (rowStr.includes('session') && rowStr.includes('course') && rowStr.includes('faculty')) {
            headerRowIndex = r;
            break;
        }
    }

    if (headerRowIndex === -1) {
        // If we can't find headers, return what we have but log warning
        return { success: true, data: { roster, manageShifts, courseAssignments, mstStaffList: staffList, courseItems } };
    }

    const headers = courseData[headerRowIndex].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    
    const colMap = {
        id: headers.indexOf('eventid'),
        course: headers.indexOf('course'),
        faculty: headers.indexOf('faculty'),
        day: headers.indexOf('day'),
        runTime: headers.indexOf('runtime'),     
        timeOfDay: headers.indexOf('timeofday'), 
        location: headers.indexOf('bxlocation'), 
        duration: headers.indexOf('coveragehrs') 
    };

    const idUpdates = [];

    // Iterate Course Schedule (Start after header row)
    for(let i = headerRowIndex + 1; i < courseData.length; i++) {
        // If ID column exists, grab ID. If not, we can't link.
        let cId = (colMap.id > -1) ? String(courseData[i][colMap.id]).trim() : "";
        
        // Auto-Generate ID if missing AND we have a valid course row
        if (!cId) {
            if (colMap.course > -1 && courseData[i][colMap.course]) {
                cId = Utilities.getUuid();
                // If the column exists in memory, update it
                if (colMap.id > -1) {
                    courseData[i][colMap.id] = cId;
                    idUpdates.push({ row: i + 1, col: colMap.id + 1, val: cId });
                }
            } else {
                continue; 
            }
        }
        
        const courseName = (colMap.course > -1) ? courseData[i][colMap.course] : "Unknown Course";
        const faculty = (colMap.faculty > -1) ? courseData[i][colMap.faculty] : "Unknown Faculty";
        
        let timeStr = "";
        if (colMap.runTime > -1) timeStr += courseData[i][colMap.runTime];
        if (colMap.timeOfDay > -1) timeStr += " " + courseData[i][colMap.timeOfDay];
        timeStr = timeStr.trim();

        courseItems.push({
            id: cId,
            name: `${courseName} - ${faculty}` 
        });
        
        const assignedRow = courseAssigns.find(a => String(a[3]).trim() === cId);
        const staffId = assignedRow ? String(assignedRow[1]).trim() : null;
        const staffObj = staffList.find(s => s.id === staffId);
        
        courseAssignments.push({
            id: cId,
            assignmentId: assignedRow ? String(assignedRow[0]) : null,
            itemName: courseName,
            courseFaculty: faculty,
            courseDay: (colMap.day > -1) ? courseData[i][colMap.day] : "",
            courseTime: timeStr,
            location: (colMap.location > -1) ? courseData[i][colMap.location] : "",
            duration: (colMap.duration > -1) ? courseData[i][colMap.duration] : "",
            staffName: staffObj ? staffObj.name : "Unassigned",
            staffId: staffId
        });
    }

    // Write back generated IDs
    if (idUpdates.length > 0) {
        const courseSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
        idUpdates.forEach(u => {
            courseSheet.getRange(u.row, u.col).setValue(u.val);
        });
    }

    return { 
      success: true, 
      data: { 
        roster: roster, 
        manageShifts: manageShifts,
        courseAssignments: courseAssignments,
        mstStaffList: staffList,
        courseItems: courseItems 
      } 
    };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Fetches external data, merges with local IDs, and updates the sheet.
 * Returns {success: boolean, message: string}
 */
function syncExternalCourseData() {
  try {
    const settings = getSettings('courseImportSettings');
    if (!settings || !settings.sheetUrl || !settings.tabName) {
        return { success: true, message: "No sync configured" }; 
    }

    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    
    // 1. Read Local Data (To preserve IDs)
    const localData = localSheet.getDataRange().getValues();
    
    // Find Header Row in Local Data
    let localHeaderIdx = -1;
    for (let r = 0; r < Math.min(localData.length, 5); r++) {
        const rowStr = localData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('course') && rowStr.includes('faculty')) {
            localHeaderIdx = r;
            break;
        }
    }

    const idMap = new Map();
    
    if (localHeaderIdx > -1) {
       const headers = localData[localHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
       const idx = {
           id: headers.indexOf('eventid'),
           course: headers.indexOf('course'),
           faculty: headers.indexOf('faculty'),
           day: headers.indexOf('day'),
           time: headers.indexOf('runtime')
       };

       if (idx.id > -1) {
           for (let i = localHeaderIdx + 1; i < localData.length; i++) {
               const row = localData[i];
               const id = String(row[idx.id]).trim();
               if (id) {
                   const key = `${row[idx.course]}|${row[idx.faculty]}|${row[idx.day]}|${row[idx.time]}`.toLowerCase().replace(/\s/g, '');
                   idMap.set(key, id);
               }
           }
       }
    }

    // 2. Fetch External Data
    const sourceId = extractFileIdFromUrl(settings.sheetUrl);
    if (!sourceId) return { success: false, message: "Invalid Source URL" };

    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName(settings.tabName);
    if (!sourceSheet) return { success: false, message: `Tab "${settings.tabName}" not found in source.` };
    
    const sourceValues = sourceSheet.getDataRange().getValues();
    if (sourceValues.length === 0) return { success: false, message: "Source sheet is empty." };
    
    // Find Header Row in Source Data
    let sourceHeaderIdx = -1;
    for (let r = 0; r < Math.min(sourceValues.length, 5); r++) {
        const rowStr = sourceValues[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('course') && rowStr.includes('faculty')) {
            sourceHeaderIdx = r;
            break;
        }
    }
    
    if (sourceHeaderIdx === -1) sourceHeaderIdx = 0; // Default to 0 if not found

    let sourceHeaders = sourceValues[sourceHeaderIdx];
    let sourceIdIdx = sourceHeaders.map(h => String(h).toLowerCase().replace(/[\s_]/g, '')).indexOf('eventid');
    
    // If source doesn't have eventID column, add it
    if (sourceIdIdx === -1) {
        sourceIdIdx = sourceHeaders.length;
        sourceValues[sourceHeaderIdx].push('eventID');
        // Pad other rows to maintain rectangular shape
        for (let i = 0; i < sourceValues.length; i++) {
            if (i !== sourceHeaderIdx) sourceValues[i].push('');
        }
    }

    // Map IDs back
    const sourceHeaderNorm = sourceValues[sourceHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    const sIdx = {
           course: sourceHeaderNorm.indexOf('course'),
           faculty: sourceHeaderNorm.indexOf('faculty'),
           day: sourceHeaderNorm.indexOf('day'),
           time: sourceHeaderNorm.indexOf('runtime')
    };

    for (let i = sourceHeaderIdx + 1; i < sourceValues.length; i++) {
        const row = sourceValues[i];
        const key = `${row[sIdx.course]}|${row[sIdx.faculty]}|${row[sIdx.day]}|${row[sIdx.time]}`.toLowerCase().replace(/\s/g, '');
        
        if (idMap.has(key)) {
            row[sourceIdIdx] = idMap.get(key);
        }
    }

    // 4. Overwrite Local Sheet
    localSheet.clear();
    // Ensure rectangular data for setValues
    const maxCols = sourceValues[0].length;
    const cleanValues = sourceValues.map(r => {
        while(r.length < maxCols) r.push('');
        return r.slice(0, maxCols);
    });
    
    localSheet.getRange(1, 1, cleanValues.length, maxCols).setValues(cleanValues);
    
    return { success: true, message: "Sync successful" };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveNewAssignment(staffId, itemId, type) {
    try {
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_ASSIGNMENTS);
        const newId = Utilities.getUuid();
        sheet.appendRow([newId, String(staffId), type, String(itemId), '', '', '', '']);
        return getSchedulingRosterData();
    } catch (e) { return { error: e.message }; }
}

function api_updateCourseAssignment(courseId, newStaffId) {
    try {
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_ASSIGNMENTS);
        const data = sheet.getDataRange().getValues();
        
        const targetCourseId = String(courseId).trim();
        const targetStaffId = String(newStaffId).trim();
        
        let found = false;
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][2] === 'Course' && String(data[i][3]).trim() === targetCourseId) {
                found = true;
                if (targetStaffId === "") {
                    sheet.deleteRow(i + 1);
                } else {
                    sheet.getRange(i + 1, 2).setValue(targetStaffId);
                }
                break;
            }
        }
        
        if (!found && targetStaffId !== "") {
            const newId = Utilities.getUuid();
            sheet.appendRow([newId, targetStaffId, 'Course', targetCourseId, '', '', '', '']);
        }
        
        return getSchedulingRosterData();
        
    } catch (e) { return { error: e.message }; }
}

function api_unassignCourse(courseId) {
    return api_updateCourseAssignment(courseId, "");
}

function api_exportCourseAssignments(data) {
    try {
        const ss = SpreadsheetApp.create("MST Assignments Export " + new Date().toISOString().split('T')[0]);
        const sheet = ss.getActiveSheet();
        const headers = ["Staff Name", "Course", "Faculty", "Day", "Time", "Location"];
        const rows = data.map(d => [d.staffName, d.itemName, d.courseFaculty, d.courseDay, d.courseTime, d.location]);
        
        sheet.appendRow(headers);
        if(rows.length > 0) {
            sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        }
        return { success: true, url: ss.getUrl() };
    } catch (e) { return { success: false, message: e.message }; }
}

// --- CALENDAR SYNC (TECH HUB) ---

function api_syncTechHubToCalendar(startStr, endStr, calendarId, overwrite) {
  try {
    if (!calendarId) throw new Error("Calendar ID missing.");
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error("Calendar not found. Check permissions or ID.");

    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    endDate.setHours(23, 59, 59);

    if (overwrite) {
      const existingEvents = cal.getEvents(startDate, endDate);
      existingEvents.forEach(e => {
        if (e.getTag('AppSource') === 'StaffHub') {
          try {
            const series = e.getEventSeries();
            if (series) { series.deleteEventSeries(); } 
            else { e.deleteEvent(); }
          } catch(err) {
            e.deleteEvent();
          }
        }
      });
    }

    const ss = getMasterDataHub();
    const shiftsData = getRequiredSheetData(ss, CONFIG.TABS.TECH_HUB_SHIFTS);
    const assignmentsData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);

    const staffMap = {};
    staffData.slice(1).forEach(r => { if(r[1]) staffMap[String(r[1]).trim()] = r[0]; });

    const shiftsMap = {};
    shiftsData.slice(1).forEach(r => {
      if(r[0]) {
        shiftsMap[String(r[0]).trim()] = {
          id: r[0],
          desc: r[1],
          day: r[2],
          start: r[3],
          end: r[4],
          zoom: (r[5] === true || r[5] === "TRUE")
        };
      }
    });

    let count = 0;
    const thAssignments = assignmentsData.slice(1).filter(r => r[2] === 'Tech Hub');
    const settings = getSettings('schedulingSettings');
    const masterZoom = settings.zoomUrl || "";

    for (const assign of thAssignments) {
      const staffId = String(assign[1]).trim();
      const shiftId = String(assign[3]).trim();
      
      const shift = shiftsMap[shiftId];
      const staffName = staffMap[staffId];

      if (!shift || !staffName) continue;

      const firstOccurrence = sched_getNextDayOfWeek(startDate, shift.day);
      if (firstOccurrence > endDate) continue;

      const startDateTime = new Date(firstOccurrence);
      const endDateTime = new Date(firstOccurrence);
      
      const startParts = String(shift.start).split(':');
      const endParts = String(shift.end).split(':');
      
      startDateTime.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0);
      endDateTime.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0);

      const recurrence = CalendarApp.newRecurrence()
          .addWeeklyRule()
          .until(endDate);

      const title = `Tech Hub: ${staffName}`;
      const location = shift.zoom ? "Zoom (See Description)" : "Tech Hub (On Site)";
      let desc = `Shift: ${shift.desc}\nStaff: ${staffName}`;
      if (shift.zoom && masterZoom) {
          desc += `\n\nZoom Link: ${masterZoom}`;
      }

      const series = cal.createEventSeries(title, startDateTime, endDateTime, recurrence, {
          location: location,
          description: desc
      });
      series.setTag('AppSource', 'StaffHub');
      
      count++;
    }

    return { success: true, message: `Created ${count} recurring shift series.` };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function sched_getNextDayOfWeek(startDate, dayName) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const targetIndex = days.indexOf(dayName);
  if (targetIndex === -1) return startDate; 
  const resultDate = new Date(startDate);
  const currentDay = resultDate.getDay();
  let diff = targetIndex - currentDay;
  if (diff < 0) diff += 7; 
  resultDate.setDate(resultDate.getDate() + diff);
  return resultDate;
}