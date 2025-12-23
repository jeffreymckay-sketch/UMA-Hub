/**
 * -------------------------------------------------------------------
 * SCHEDULING CONTROLLER (Tech Hub & MST) - COMPLETE
 * -------------------------------------------------------------------
 */

// --- ROSTER & SHIFT MANAGEMENT ---

function getSchedulingRosterData() {
  try {
    const ss = getMasterDataHub();
    
    // 1. Load all data in bulk
    const shiftsData = getRequiredSheetData(ss, CONFIG.TABS.TECH_HUB_SHIFTS);
    const assignmentsData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);
    const availData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_AVAILABILITY);
    const courseSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    
    const lastRow = courseSheet ? courseSheet.getLastRow() : 0;
    const courseData = lastRow > 0 ? courseSheet.getRange(1, 1, lastRow, courseSheet.getLastColumn()).getValues() : [];

    // --- PRE-PROCESSING INTO MAPS ---

    const staffList = [];
    const staffMap = new Map(); 
    for (let i = 1; i < staffData.length; i++) {
      if (staffData[i][1]) {
        const sObj = {
          name: staffData[i][0],
          id: String(staffData[i][1]).trim(), 
          role: staffData[i][2],
          isActive: staffData[i][3]
        };
        staffList.push(sObj);
        staffMap.set(sObj.id, sObj);
      }
    }

    const assignmentMap = new Map();
    for (let i = 1; i < assignmentsData.length; i++) {
      const type = assignmentsData[i][2];
      const itemId = String(assignmentsData[i][3]).trim();
      const staffId = String(assignmentsData[i][1]).trim();
      const recordId = String(assignmentsData[i][0]).trim();
      assignmentMap.set(`${type}|${itemId}`, { staffId, recordId });
    }

    const availMap = new Map();
    for (let i = 1; i < availData.length; i++) {
      const sId = String(availData[i][1]).trim();
      const day = availData[i][2];
      const key = `${sId}|${day}`;
      if (!availMap.has(key)) availMap.set(key, []);
      availMap.get(key).push({ start: availData[i][3], end: availData[i][4] });
    }

    // 2. Parse Shifts (Tech Hub)
    const roster = [];
    const manageShifts = [];
    
    for (let i = 1; i < shiftsData.length; i++) {
      const row = shiftsData[i];
      if (!row[0]) continue; 

      const shiftId = String(row[0]).trim(); 
      const startTime = safeFormatTime(row[3]);
      const endTime = safeFormatTime(row[4]);

      const shiftObj = {
        shiftId: shiftId,
        description: row[1],
        day: row[2],
        start: startTime,
        end: endTime,
        zoom: (row[5] === true || row[5] === "TRUE")
      };

      manageShifts.push({ 
          id: shiftId, 
          desc: row[1], 
          day: row[2], 
          start: startTime, 
          end: endTime, 
          zoom: shiftObj.zoom 
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

    // 3. Parse Course Assignments (MST)
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

    if (headerRowIndex > -1) {
        const headers = courseData[headerRowIndex].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
        const colMap = {
            id: headers.indexOf('eventid'),
            course: headers.indexOf('course'),
            faculty: headers.indexOf('faculty'),
            day: headers.indexOf('day'),
            runTime: headers.indexOf('runtime'),     
            timeOfDay: headers.indexOf('timeofday'), 
            location: headers.indexOf('bxlocation'), 
            duration: headers.indexOf('coveragehrs'),
            zoomLink: headers.indexOf('zoomlink')
        };

        let idColumnNeedsUpdate = false;
        const idColumnValues = courseData.map(row => (colMap.id > -1) ? row[colMap.id] : "");

        for(let i = headerRowIndex + 1; i < courseData.length; i++) {
            if (!courseData[i][0] && !courseData[i][1]) continue;

            let cId = (colMap.id > -1) ? String(courseData[i][colMap.id]).trim() : "";
            const hasData = (colMap.course > -1 && courseData[i][colMap.course]);

            if (!cId && hasData) {
                cId = Utilities.getUuid();
                if (colMap.id > -1) {
                    idColumnValues[i] = cId;
                    idColumnNeedsUpdate = true;
                }
            }

            if (!cId) continue; 
            
            const courseName = (colMap.course > -1) ? courseData[i][colMap.course] : "Unknown Course";
            const faculty = (colMap.faculty > -1) ? courseData[i][colMap.faculty] : "TBD";
            
            let timeStr = "";
            if (colMap.runTime > -1) timeStr += safeFormatTime(courseData[i][colMap.runTime]);
            if (colMap.timeOfDay > -1) timeStr += " " + String(courseData[i][colMap.timeOfDay]);
            timeStr = timeStr.trim();

            let durationStr = "";
            if (colMap.duration > -1) {
                const rawDur = courseData[i][colMap.duration];
                if (typeof rawDur === 'number') {
                    durationStr = rawDur + " hrs";
                } else {
                    durationStr = safeFormatDuration(rawDur);
                }
            }

            const link = (colMap.zoomLink > -1) ? String(courseData[i][colMap.zoomLink]) : "";

            courseItems.push({ id: cId, name: `${courseName} - ${faculty}` });
            
            const assignEntry = assignmentMap.get(`Course|${cId}`);
            const staffObj = assignEntry ? staffMap.get(assignEntry.staffId) : null;
            
            courseAssignments.push({
                id: cId,
                assignmentId: assignEntry ? assignEntry.recordId : null,
                itemName: courseName,
                courseFaculty: faculty,
                courseDay: (colMap.day > -1) ? String(courseData[i][colMap.day]) : "",
                courseTime: timeStr,
                location: (colMap.location > -1) ? String(courseData[i][colMap.location]) : "",
                duration: durationStr,
                link: link,
                staffName: staffObj ? staffObj.name : "Unassigned",
                staffId: assignEntry ? assignEntry.staffId : null
            });
        }

        if (idColumnNeedsUpdate && colMap.id > -1) {
            const output = idColumnValues.map(val => [val]);
            courseSheet.getRange(1, colMap.id + 1, output.length, 1).setValues(output);
        }
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
    console.error(e);
    return { error: e.message };
  }
}

// --- HELPER FUNCTIONS FOR FORMATTING ---

function safeFormatTime(val) {
    if (!val) return "";
    if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "h:mm a");
    }
    return String(val);
}

function safeFormatDuration(val) {
    if (!val) return "";
    if (val instanceof Date) {
        const h = val.getHours();
        const m = val.getMinutes();
        if (h === 0 && m === 0) return "";
        return (h > 0 ? h + "h " : "") + (m > 0 ? m + "m" : "");
    }
    return String(val);
}

function api_forceRemoteSync() {
    return syncExternalCourseData();
}

/**
 * NEW: Imports Zoom Links using BUCKET & FUZZY NAME Logic
 * Matches via: Course+Day+Time (Bucket) -> Name Containment (Fuzzy)
 */
function importCourseLinks(sheetUrl, tabName) {
  const debugLog = [];
  try {
    if (!sheetUrl || !tabName) return { success: false, message: "Missing URL or Tab Name." };

    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    if (!localSheet) return { success: false, message: "Local Course Schedule not found." };

    // 1. Prepare Local Data
    const localData = localSheet.getDataRange().getValues();
    let localHeaderIdx = -1;
    for (let r = 0; r < Math.min(localData.length, 5); r++) {
        const rowStr = localData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('course') || rowStr.includes('startdate')) {
            localHeaderIdx = r;
            break;
        }
    }
    if (localHeaderIdx === -1) return { success: false, message: "Local headers not found." };

    const lHeaders = localData[localHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    debugLog.push(`Local Headers Found: ${lHeaders.join(', ')}`);

    const lIdx = {
        course: lHeaders.indexOf('course'),
        faculty: lHeaders.indexOf('faculty'),
        day: lHeaders.indexOf('day'),
        runTime: lHeaders.indexOf('runtime'),
        timeOfDay: lHeaders.indexOf('timeofday'),
        zoomLink: lHeaders.indexOf('zoomlink')
    };

    if (lIdx.zoomLink === -1) {
        lIdx.zoomLink = localData[localHeaderIdx].length;
        localSheet.getRange(localHeaderIdx + 1, lIdx.zoomLink + 1).setValue("Zoom Link");
        debugLog.push("Created new 'Zoom Link' column.");
    }

    // 2. Fetch External Data
    const sourceId = extractFileIdFromUrl(sheetUrl);
    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName(tabName);
    if (!sourceSheet) return { success: false, message: "External tab not found." };
    
    const sourceValues = sourceSheet.getDataRange().getValues();
    const sourceRichText = sourceSheet.getDataRange().getRichTextValues();

    let sourceHeaderIdx = -1;
    for (let r = 0; r < Math.min(sourceValues.length, 5); r++) {
        const rowStr = sourceValues[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('coursenumber') && rowStr.includes('instructorname')) {
            sourceHeaderIdx = r;
            break;
        }
    }
    if (sourceHeaderIdx === -1) return { success: false, message: "External headers not found." };

    const sHeaders = sourceValues[sourceHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    debugLog.push(`External Headers Found: ${sHeaders.join(', ')}`);

    const sIdx = {
        course: sHeaders.indexOf('coursenumber'),
        instructor: sHeaders.indexOf('instructorname'),
        day: sHeaders.indexOf('day'),
        startTime: sHeaders.indexOf('starttime'),
        link: sHeaders.indexOf('zoomlink')
    };

    if (sIdx.link === -1) return { success: false, message: "'ZOOM Link' column not found in source." };

    // 3. Build Source Map (Bucket by Course|Day|Time)
    const sourceMap = new Map(); 
    let sampleExtKey = "";
    
    let lastCourse = "";
    let lastInstructor = "";

    for (let i = sourceHeaderIdx + 1; i < sourceValues.length; i++) {
        const row = sourceValues[i];
        const richRow = sourceRichText[i];

        let courseVal = row[sIdx.course];
        let instrVal = row[sIdx.instructor];

        if (courseVal) lastCourse = courseVal;
        else courseVal = lastCourse;

        if (instrVal) lastInstructor = instrVal;
        else instrVal = lastInstructor;

        if (!courseVal || String(courseVal).toLowerCase().includes('mode')) continue;

        // Generate Bucket Key (No Instructor)
        const c = String(courseVal).toLowerCase().replace(/[^a-z0-9]/g, '');
        const d = sched_normalizeDay(row[sIdx.day]);
        const t = sched_normalizeTime(row[sIdx.startTime], null);
        const bucketKey = `${c}|${d}|${t}`;
        
        if (!sampleExtKey) sampleExtKey = bucketKey; 

        let link = String(row[sIdx.link]).trim();
        if (!link.startsWith('http')) {
            const richCell = richRow[sIdx.link];
            const url = richCell.getLinkUrl();
            if (url) link = url;
        }

        if (link && link.startsWith('http')) {
            if (!sourceMap.has(bucketKey)) sourceMap.set(bucketKey, []);
            sourceMap.get(bucketKey).push({
                instructor: String(instrVal).toLowerCase().replace(/[^a-z]/g, ''), // Simple normalize for fuzzy match
                link: link
            });
        }
    }
    debugLog.push(`External Map Size: ${sourceMap.size}`);
    debugLog.push(`Sample External Bucket: ${sampleExtKey}`);

    // 4. Match and Update Local Data
    let updates = 0;
    const outputColumn = []; 
    let sampleLocalKey = "";

    for (let i = localHeaderIdx + 1; i < localData.length; i++) {
        const row = localData[i];
        
        let rawTime = String(row[lIdx.runTime]);
        if (rawTime.includes('-')) rawTime = rawTime.split('-')[0].trim();
        
        const c = String(row[lIdx.course]).toLowerCase().replace(/[^a-z0-9]/g, '');
        const d = sched_normalizeDay(row[lIdx.day]);
        const t = sched_normalizeTime(rawTime, row[lIdx.timeOfDay]);
        const bucketKey = `${c}|${d}|${t}`;
        
        if (!sampleLocalKey) sampleLocalKey = bucketKey; 

        const candidates = sourceMap.get(bucketKey);
        let link = "";

        if (candidates) {
            const localInstr = String(row[lIdx.faculty]).toLowerCase().replace(/[^a-z]/g, '');
            // Fuzzy Name Match: Containment
            const match = candidates.find(cand => cand.instructor.includes(localInstr) || localInstr.includes(cand.instructor));
            if (match) {
                link = match.link;
            }
        }
        
        if (link) {
            outputColumn.push([link]);
            updates++;
        } else {
            outputColumn.push([row[lIdx.zoomLink] || ""]);
        }
    }
    debugLog.push(`Sample Local Bucket: ${sampleLocalKey}`);

    // 5. Write Back
    if (updates > 0) {
        localSheet.getRange(localHeaderIdx + 2, lIdx.zoomLink + 1, outputColumn.length, 1).setValues(outputColumn);
    }

    // 6. Save Settings
    const props = PropertiesService.getScriptProperties();
    const savedSettings = { url: sheetUrl, tab: tabName };
    props.setProperty('mst_links_settings', JSON.stringify(savedSettings));

    const report = `Import Complete.\nMatched: ${updates} links.\n\n--- DEBUG INFO ---\n${debugLog.join('\n')}`;
    return { success: true, message: report };

  } catch (e) {
    return { success: false, message: e.message + "\n\nLog:\n" + debugLog.join('\n') };
  }
}

/**
 * Retrieves saved Link Import settings for the frontend.
 */
function api_getLinkImportSettings() {
    try {
        const props = PropertiesService.getScriptProperties();
        const json = props.getProperty('mst_links_settings');
        return json ? JSON.parse(json) : { url: '', tab: '' };
    } catch (e) {
        return { url: '', tab: '' };
    }
}

function sched_normalizeDay(dayStr) {
    if (!dayStr) return 0;
    const s = String(dayStr).toLowerCase().trim();
    if (s.includes('mon') || s === 'm') return 1;
    if (s.includes('tue') || s === 'tu' || s === 't') return 2;
    if (s.includes('wed') || s === 'w') return 3;
    if (s.includes('thu') || s === 'th' || s === 'r') return 4;
    if (s.includes('fri') || s === 'f') return 5;
    if (s.includes('sat') || s === 'sa') return 6;
    if (s.includes('sun') || s === 'su') return 0;
    return 0;
}

function sched_normalizeTime(timeVal, amPmVal) {
    if (!timeVal) return 0;
    
    let h = 0, m = 0;

    if (timeVal instanceof Date) {
        h = timeVal.getHours();
        m = timeVal.getMinutes();
    } 
    else {
        const str = String(timeVal).trim();
        const match = str.match(/(\d+):(\d+)/);
        if (match) {
            h = parseInt(match[1]);
            m = parseInt(match[2]);
        }
        
        if (!amPmVal) {
            if (str.toUpperCase().includes('PM')) amPmVal = 'PM';
            if (str.toUpperCase().includes('AM')) amPmVal = 'AM';
        }
    }

    if (amPmVal) {
        const isPm = String(amPmVal).trim().toUpperCase().includes('PM');
        const isAm = String(amPmVal).trim().toUpperCase().includes('AM');
        
        if (isPm && h < 12) h += 12;
        if (isAm && h === 12) h = 0;
    }

    return (h * 60) + m;
}

function syncExternalCourseData() {
  try {
    const settings = getSettings('courseImportSettings');
    if (!settings || !settings.sheetUrl || !settings.tabName) {
        return { success: true, message: "No sync configured" }; 
    }

    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    const assignSheet = ss.getSheetByName(CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffSheet = ss.getSheetByName(CONFIG.TABS.STAFF_LIST);

    // 1. Read Local Data (To Preserve IDs)
    const localData = localSheet.getDataRange().getValues();
    const localIdMap = new Map(); 
    
    let localHeaderIdx = -1;
    for (let r = 0; r < Math.min(localData.length, 5); r++) {
        const rowStr = localData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('course') || rowStr.includes('startdate')) {
            localHeaderIdx = r;
            break;
        }
    }

    if (localHeaderIdx > -1) {
        const h = localData[localHeaderIdx].map(x => String(x).toLowerCase().replace(/[\s_]/g, ''));
        const idx = {
            id: h.indexOf('eventid'),
            course: h.indexOf('course'),
            day: h.indexOf('day'),
            time: h.indexOf('runtime')
        };
        if (idx.id > -1) {
            for (let i = localHeaderIdx + 1; i < localData.length; i++) {
                const row = localData[i];
                const id = String(row[idx.id]).trim();
                if (id) {
                    const key = `${row[idx.course]}|${row[idx.day]}|${row[idx.time]}`.toLowerCase().replace(/\s/g, '');
                    localIdMap.set(key, id);
                }
            }
        }
    }

    // 2. Fetch External Data
    const sourceId = extractFileIdFromUrl(settings.sheetUrl);
    if (!sourceId) return { success: false, message: "Invalid Source URL" };

    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName(settings.tabName);
    if (!sourceSheet) return { success: false, message: `Tab "${settings.tabName}" not found.` };
    
    const sourceValues = sourceSheet.getDataRange().getValues();
    if (sourceValues.length === 0) return { success: false, message: "Source sheet is empty." };

    // 3. Identify Source Headers
    let headerIdx = -1;
    for (let r = 0; r < Math.min(sourceValues.length, 5); r++) {
        const rowStr = sourceValues[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('course') || rowStr.includes('startdate')) {
            headerIdx = r;
            break;
        }
    }
    if (headerIdx === -1) headerIdx = 0;

    const headers = sourceValues[headerIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    let idCol = headers.indexOf('eventid');
    const mstEmailCol = headers.findIndex(h => h.includes('mstassigned') || h.includes('assignedbyemail'));
    const sIdx = {
        course: headers.indexOf('course'),
        day: headers.indexOf('day'),
        time: headers.indexOf('runtime'),
        duration: headers.indexOf('coveragehrs'),
        ampm: headers.indexOf('timeofday')
    };

    if (idCol === -1) {
        idCol = headers.length;
        sourceValues[headerIdx].push('eventID');
        for (let i = 0; i < sourceValues.length; i++) {
            if (i !== headerIdx) sourceValues[i].push('');
        }
    }

    // 4. SMART MERGE & SANITIZATION
    const assignData = assignSheet.getDataRange().getValues();
    let assignmentsChanged = false;
    
    const assignRowMap = new Map();
    for(let i=1; i<assignData.length; i++) {
        if(assignData[i][2] === 'Course') assignRowMap.set(String(assignData[i][3]).trim(), i);
    }

    for (let i = headerIdx + 1; i < sourceValues.length; i++) {
        const row = sourceValues[i];
        
        // --- DATA SANITIZATION: Fix Duration ---
        if (sIdx.time > -1 && sIdx.duration > -1) {
            const timeStr = String(row[sIdx.time]);
            const ampmVal = (sIdx.ampm > -1) ? String(row[sIdx.ampm]) : "";
            
            if (timeStr.includes('-')) {
                const parts = timeStr.split('-').map(p => p.trim());
                if (parts.length === 2) {
                    const startMatch = parts[0].match(/(\d+):(\d+)/);
                    const endMatch = parts[1].match(/(\d+):(\d+)/);
                    
                    if (startMatch && endMatch) {
                        let h1 = parseInt(startMatch[1]);
                        const m1 = parseInt(startMatch[2]);
                        let h2 = parseInt(endMatch[1]);
                        const m2 = parseInt(endMatch[2]);
                        
                        const isPM = ampmVal.toLowerCase().includes('pm') || parts[0].toLowerCase().includes('pm');
                        if (isPM && h1 < 12) h1 += 12;
                        if (!isPM && h1 === 12) h1 = 0; 

                        let startMins = h1 * 60 + m1;
                        let endMins = h2 * 60 + m2;

                        if (endMins < startMins) endMins += 720;
                        else if (startMins >= 720 && endMins < 720) endMins += 720;
                        if (startMins >= 720 && endMins < startMins) endMins += 720;

                        const diffMins = endMins - startMins;
                        
                        if (diffMins > 0) {
                            row[sIdx.duration] = Number((diffMins / 60).toFixed(2));
                        }
                    }
                }
            }
        }
        // --- END SANITIZATION ---

        const extId = String(row[idCol]).trim();
        const key = `${row[sIdx.course]}|${row[sIdx.day]}|${row[sIdx.time]}`.toLowerCase().replace(/\s/g, '');
        
        let finalId = "";

        if (extId) {
            finalId = extId;
            if (localIdMap.has(key)) {
                const oldLocalId = localIdMap.get(key);
                if (oldLocalId !== finalId && assignRowMap.has(oldLocalId)) {
                    const rowIdx = assignRowMap.get(oldLocalId);
                    assignData[rowIdx][3] = finalId; 
                    assignmentsChanged = true;
                }
            }
        } else if (localIdMap.has(key)) {
            finalId = localIdMap.get(key);
        } else {
            finalId = Utilities.getUuid();
        }

        row[idCol] = finalId;
    }

    // 5. Overwrite Local Sheet
    localSheet.clear();
    const maxCols = sourceValues[0].length;
    const cleanValues = sourceValues.map(r => {
        while(r.length < maxCols) r.push('');
        return r.slice(0, maxCols);
    });
    localSheet.getRange(1, 1, cleanValues.length, maxCols).setValues(cleanValues);

    // 6. Sync "MST Assigned" Column
    if (mstEmailCol > -1 && idCol > -1) {
        const staffData = staffSheet.getDataRange().getValues();
        const staffEmailMap = new Map();
        for(let i=1; i<staffData.length; i++) {
            if(staffData[i][1]) staffEmailMap.set(String(staffData[i][1]).trim().toLowerCase(), String(staffData[i][1]).trim());
        }

        const currentAssignMap = new Map();
        for(let i=1; i<assignData.length; i++) {
            if(assignData[i][2] === 'Course') currentAssignMap.set(String(assignData[i][3]).trim(), i);
        }

        const newRows = [];

        for (let i = headerIdx + 1; i < sourceValues.length; i++) {
            const eventId = String(sourceValues[i][idCol]).trim();
            const assignedEmail = String(sourceValues[i][mstEmailCol]).trim().toLowerCase();

            if (eventId && assignedEmail && staffEmailMap.has(assignedEmail)) {
                const staffId = staffEmailMap.get(assignedEmail);
                
                if (currentAssignMap.has(eventId)) {
                    const rowIdx = currentAssignMap.get(eventId);
                    if (assignData[rowIdx][1] !== staffId) {
                        assignData[rowIdx][1] = staffId;
                        assignmentsChanged = true;
                    }
                } else {
                    const newId = Utilities.getUuid();
                    newRows.push([newId, staffId, 'Course', eventId, '', '', '', '']);
                    currentAssignMap.set(eventId, -1); 
                }
            }
        }

        if (newRows.length > 0) {
            assignSheet.getRange(assignSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        }
    }

    if (assignmentsChanged) {
        assignSheet.getRange(1, 1, assignData.length, assignData[0].length).setValues(assignData);
    }
    
    return { success: true, message: "Sync successful" };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- ACTIONS ---

function saveNewAssignment(staffId, itemId, type) {
    try {
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_ASSIGNMENTS);
        const data = sheet.getDataRange().getValues();
        let foundRow = -1;
        
        for(let i=1; i<data.length; i++) {
            if (data[i][2] === type && String(data[i][3]) === String(itemId)) {
                foundRow = i + 1;
                break;
            }
        }
        
        if (foundRow > -1) {
            sheet.getRange(foundRow, 2).setValue(staffId);
        } else {
            const newId = Utilities.getUuid();
            sheet.appendRow([newId, String(staffId), type, String(itemId), '', '', '', '']);
        }
        return getSchedulingRosterData();
    } catch (e) { return { error: e.message }; }
}

function api_updateCourseAssignment(courseId, newStaffId) {
    return saveNewAssignment(newStaffId, courseId, 'Course');
}

function api_unassignCourse(courseId) {
    try {
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName(CONFIG.TABS.STAFF_ASSIGNMENTS);
        const data = sheet.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
            if (data[i][2] === 'Course' && String(data[i][3]) === String(courseId)) {
                sheet.deleteRow(i+1);
                break;
            }
        }
        return getSchedulingRosterData();
    } catch (e) { return { error: e.message }; }
}

function api_exportCourseAssignments(data) {
    try {
        const ss = SpreadsheetApp.create("MST Assignments Export " + new Date().toISOString().split('T')[0]);
        const sheet = ss.getActiveSheet();
        const headers = ["Staff Name", "Course", "Faculty", "Day", "Time", "Location"];
        const rows = data.map(d => [d.staffName, d.itemName, d.courseFaculty, d.courseDay, d.courseTime, d.location]);
        sheet.appendRow(headers);
        if(rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        return { success: true, url: ss.getUrl() };
    } catch (e) { return { success: false, message: e.message }; }
}

function timesOverlap(start1, end1, start2, end2) {
    if (!start1 || !end1 || !start2 || !end2) return false;
    return (start1 < end2 && start2 < end1);
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

function api_syncTechHubToCalendar(startStr, endStr, calendarId, overwrite) {
  try {
    if (!calendarId) throw new Error("Calendar ID missing.");
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error("Calendar not found.");

    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    endDate.setHours(23, 59, 59);

    if (overwrite) {
      const existingEvents = cal.getEvents(startDate, endDate);
      existingEvents.forEach(e => {
        if (e.getTag('AppSource') === 'StaffHub') {
            try { e.getEventSeries().deleteEventSeries(); } catch(err) { e.deleteEvent(); }
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
      if(r[0]) shiftsMap[String(r[0]).trim()] = { id: r[0], desc: r[1], day: r[2], start: r[3], end: r[4], zoom: (r[5] === true || r[5] === "TRUE") };
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
      if (startParts.length < 2 || endParts.length < 2) continue;

      startDateTime.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0);
      endDateTime.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0);

      // FIX: Handle AM/PM flip (End < Start) by adding 12 hours
      if (endDateTime <= startDateTime) {
          endDateTime.setHours(endDateTime.getHours() + 12);
      }
      
      // FIX: Handle Zero Duration (Start == End)
      if (endDateTime.getTime() === startDateTime.getTime()) {
          endDateTime.setMinutes(endDateTime.getMinutes() + 60);
      }

      const recurrence = CalendarApp.newRecurrence().addWeeklyRule().until(endDate);
      const title = `Tech Hub: ${staffName}`;
      const location = shift.zoom ? "Zoom" : "Tech Hub";
      let desc = `Shift: ${shift.desc}\nStaff: ${staffName}`;
      if (shift.zoom && masterZoom) desc += `\nZoom: ${masterZoom}`;

      const series = cal.createEventSeries(title, startDateTime, endDateTime, recurrence, { location: location, description: desc });
      series.setTag('AppSource', 'StaffHub');
      count++;
    }
    return { success: true, message: `Created ${count} recurring shift series.` };
  } catch (e) { return { success: false, message: e.message }; }
}