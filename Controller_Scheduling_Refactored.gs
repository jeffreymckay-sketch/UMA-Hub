/**
 * -------------------------------------------------------------------
 * NEW REFACTORED SCHEDULING CONTROLLER
 * -------------------------------------------------------------------
 */

// --- HELPER TO GET COLUMN INDICES FROM HEADERS ---
function getColumnMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const normalizedHeader = String(header).toLowerCase().replace(/[\s_]/g, '');
    map[normalizedHeader] = index;
  });
  return map;
}

// --- REFACTORED ROSTER & SHIFT MANAGEMENT ---

function getSchedulingRosterData_refactored() {
  try {
    const ss = getMasterDataHub();
    
    // 1. Load all data in bulk using a helper
    const shiftsData = _loadSheetData(ss, CONFIG.TABS.TECH_HUB_SHIFTS);
    const assignmentsData = _loadSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = _loadSheetData(ss, CONFIG.TABS.STAFF_LIST);
    const availData = _loadSheetData(ss, CONFIG.TABS.STAFF_AVAILABILITY);
    const courseData = _loadSheetData(ss, CONFIG.TABS.COURSE_SCHEDULE);

    if (!shiftsData.success || !assignmentsData.success || !staffData.success || !availData.success || !courseData.success) {
      throw new Error("Failed to load one or more required sheets.");
    }

    // 2. Pre-process data into maps
    const staffMap = _buildStaffMap(staffData.data);
    const assignmentMap = _buildAssignmentMap(assignmentsData.data);
    const availMap = _buildAvailMap(availData.data);

    // 3. Parse Shifts (Tech Hub)
    const { roster, manageShifts } = _parseShifts(shiftsData.data, Array.from(staffMap.values()), availMap, assignmentMap);

    // 4. Parse Course Assignments (MST)
    const { courseAssignments, courseItems } = _parseCourseAssignments(courseData.data, staffMap, assignmentMap);

    return { 
      success: true, 
      data: { 
        roster, 
        manageShifts,
        courseAssignments,
        mstStaffList: Array.from(staffMap.values()).map(s => ({ id: s.id, name: s.name })),
        courseItems 
      } 
    };

  } catch (e) {
    console.error(e);
    return { error: e.message };
  }
}

function _loadSheetData(ss, sheetName) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
    const data = sheet.getDataRange().getValues();
    return { success: true, data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _buildStaffMap(staffData) {
  const staffMap = new Map();
  const headers = getColumnMap(staffData[0]);
  for (let i = 1; i < staffData.length; i++) {
    const row = staffData[i];
    const id = String(row[headers['staffid']]).trim();
    if (id) {
      staffMap.set(id, {
        name: row[headers['fullname']],
        id,
        role: row[headers['roles']],
        isActive: row[headers['isactive']]
      });
    }
  }
  return staffMap;
}

function _buildAssignmentMap(assignmentsData) {
  const assignmentMap = new Map();
  const headers = getColumnMap(assignmentsData[0]);
  for (let i = 1; i < assignmentsData.length; i++) {
    const row = assignmentsData[i];
    const type = row[headers['type']];
    const itemId = String(row[headers['itemid']]).trim();
    const staffId = String(row[headers['staffid']]).trim();
    const recordId = String(row[headers['recordid']]).trim();
    assignmentMap.set(`${type}|${itemId}`, { staffId, recordId });
  }
  return assignmentMap;
}

function _buildAvailMap(availData) {
    const availMap = new Map();
    const headers = getColumnMap(availData[0]);
    for (let i = 1; i < availData.length; i++) {
        const row = availData[i];
        const sId = String(row[headers['staffid']]).trim();
        const day = row[headers['dayofweek']];
        const key = `${sId}|${day}`;
        if (!availMap.has(key)) availMap.set(key, []);
        availMap.get(key).push({ start: row[headers['starttime']], end: row[headers['endtime']] });
    }
    return availMap;
}


function _parseShifts(shiftsData, staffList, availMap, assignmentMap) {
  const roster = [];
  const manageShifts = [];
  const headers = getColumnMap(shiftsData[0]);

  for (let i = 1; i < shiftsData.length; i++) {
    const row = shiftsData[i];
    const shiftId = String(row[headers['shiftid']]).trim();
    if (!shiftId) continue;

    const startTime = safeFormatTime(row[headers['starttime']]);
    const endTime = safeFormatTime(row[headers['endtime']]);
    const zoom = row[headers['iszoom']] === true || String(row[headers['iszoom']]).toUpperCase() === 'TRUE';

    const shiftObj = {
      shiftId,
      description: row[headers['description']],
      day: row[headers['day']],
      start: startTime,
      end: endTime,
      zoom
    };

    manageShifts.push({ 
        id: shiftId, 
        desc: shiftObj.description, 
        day: shiftObj.day, 
        start: startTime, 
        end: endTime, 
        zoom
    });

    const assignEntry = assignmentMap.get(`Tech Hub|${shiftId}`);
    shiftObj.assignedStaffId = assignEntry ? assignEntry.staffId : "";

    shiftObj.smartStaffList = staffList.map(s => {
      let isAvailable = true;
      const busySlots = availMap.get(`${s.id}|${shiftObj.day}`);
      if (busySlots) {
        if (busySlots.some(slot => timesOverlap(shiftObj.start, shiftObj.end, slot.start, slot.end))) {
          isAvailable = false;
        }
      }
      return { id: s.id, name: s.name, available: isAvailable };
    });

    roster.push(shiftObj);
  }
  return { roster, manageShifts };
}

function _parseCourseAssignments(courseData, staffMap, assignmentMap) {
    const courseAssignments = [];
    const courseItems = [];
    
    let headerRowIndex = -1;
    for (let r = 0; r < Math.min(courseData.length, 5); r++) {
        const rowStr = courseData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('startdate') || (rowStr.includes('course') && rowStr.includes('faculty'))) {
            headerRowIndex = r;
            break;
        }
    }

    if (headerRowIndex === -1) {
        console.error("Header row not found in course data");
        return { courseAssignments, courseItems };
    }

    const headers = getColumnMap(courseData[headerRowIndex]);
    
    for(let i = headerRowIndex + 1; i < courseData.length; i++) {
        const row = courseData[i];
        let cId = String(row[headers['eventid']] || '').trim();
        const hasData = row[headers['course']];

        if (!cId && hasData) {
            cId = Utilities.getUuid();
            // This part is tricky without knowing the exact structure. 
            // For now, we'll just use the new ID. A complete solution would require updating the sheet.
        }

        if (!cId) continue; 
        
        const courseName = row[headers['course']] || "Unknown Course";
        const faculty = row[headers['faculty']] || "TBD";
        
        let timeStr = (safeFormatTime(row[headers['runtime']]) + " " + (row[headers['timeofday']] || '')).trim();
        let durationStr = safeFormatDuration(row[headers['coveragehrs']]);

        courseItems.push({ id: cId, name: `${courseName} - ${faculty}` });
        
        const assignEntry = assignmentMap.get(`Course|${cId}`);
        const staffObj = assignEntry ? staffMap.get(assignEntry.staffId) : null;
        
        courseAssignments.push({
            id: cId,
            assignmentId: assignEntry ? assignEntry.recordId : null,
            itemName: courseName,
            courseFaculty: faculty,
            courseDay: row[headers['day']] || "",
            courseTime: timeStr,
            location: row[headers['bxlocation']] || "",
            duration: durationStr,
            link: row[headers['zoomlink']] || "",
            staffName: staffObj ? staffObj.name : "Unassigned",
            staffId: assignEntry ? assignEntry.staffId : null
        });
    }

    return { courseAssignments, courseItems };
}
