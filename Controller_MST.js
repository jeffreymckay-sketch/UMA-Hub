/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING, SETTINGS & CALENDAR SYNC
 * -------------------------------------------------------------------
 */

function api_getMstSettings() {
  try {
    const allSettings = getSettings();
    let mstSettings = {};
    if (allSettings.mstSettings) {
      try { mstSettings = JSON.parse(allSettings.mstSettings); } catch (e) { }
    }
    const ss = getMasterDataHub();
    const sheetNames = ss.getSheets().map(s => s.getName());
    return { success: true, data: { settings: mstSettings, sheetNames: sheetNames } };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_saveMstSettings(newMstSettings) {
  try {
    const allSettings = getSettings();
    let currentMstSettings = {};
    if (allSettings.mstSettings) {
      try { currentMstSettings = JSON.parse(allSettings.mstSettings); } catch (e) { }
    }
    const updatedMstSettings = Object.assign(currentMstSettings, newMstSettings);
    allSettings.mstSettings = JSON.stringify(updatedMstSettings);
    saveSettings(allSettings);
    return { success: true, message: "MST settings saved successfully." };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_getMstViewData() {
    try {
        const settings = getSettings();
        let sourceTab = 'Course_Schedule';
        let targetCalId = settings.targetCalendarId; 
        
        if (settings.mstSettings) {
            try { 
                const parsed = JSON.parse(settings.mstSettings);
                sourceTab = parsed.sourceTabName || sourceTab; 
            } catch(e){}
        }

        // Recover IDs surgically (only if needed)
        if (targetCalId) {
            try { recoverMstEventIds_(sourceTab, targetCalId); } catch(e) { console.warn("ID Recovery warning:", e); }
        }

        var staffData = getSheet('Staff_List').getDataRange().getValues();
        var assignmentData = getSheet('Staff_Assignments').getDataRange().getValues();
        var courseData = getSheet(sourceTab).getDataRange().getValues();

        var staffHeaders = getColumnMap(staffData[0]);
        var assignmentHeaders = getColumnMap(assignmentData[0]);
        
        var courseHeaderRow = courseData.find(function(row) { return row.join('').toLowerCase().includes('eventid'); });
        if (!courseHeaderRow) throw new Error("Could not find header row in Course Schedule sheet.");
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
                zoomLink: course.zoomLink, 
                staffName: staff ? staff.name : "Unassigned",
                staffId: staff ? staff.id : null
            };
        });
        
        var mstStaffList = allStaff.filter(function(s) { return s.role && s.role.toLowerCase().includes('mst'); }).map(function(s) { return { id: s.id, name: s.name }; });

        return { success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } };
    } catch (e) {
        console.error("Error in api_getMstViewData: " + e.stack);
        return { success: false, error: e.message };
    }
}

function api_updateCourseAssignment(courseId, staffId) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const sheet = getSheet('Staff_Assignments');
        const data = sheet.getDataRange().getValues();
        const headers = getColumnMap(data[0]);
        
        const idCol = headers['assignmentid'];
        const staffCol = headers['staffid'];
        const refCol = headers['referenceid']; 
        const typeCol = headers['assignmenttype']; 

        let foundRow = -1;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][refCol]) === String(courseId)) {
                foundRow = i + 1;
                break;
            }
        }

        if (foundRow > -1) {
            sheet.getRange(foundRow, staffCol + 1).setValue(staffId);
        } else {
            const newId = Utilities.getUuid();
            const newRow = [];
            for(let k=0; k<data[0].length; k++) newRow.push("");
            newRow[idCol] = newId;
            newRow[staffCol] = staffId;
            newRow[refCol] = courseId;
            if(typeCol !== undefined) newRow[typeCol] = "Course";
            sheet.appendRow(newRow);
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    } finally {
        lock.releaseLock();
    }
}

function api_updateCourseZoom(courseId, zoomLink) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        
        const settings = getSettings();
        let sourceTab = 'Course_Schedule';
        if (settings.mstSettings) {
            try { 
                const parsed = JSON.parse(settings.mstSettings);
                sourceTab = parsed.sourceTabName || sourceTab; 
            } catch(e){}
        }
        
        const sheet = getSheet(sourceTab);
        if(!sheet) return { success: false, message: "Sheet not found" };
        
        const data = sheet.getDataRange().getValues();
        
        let headerRowIdx = -1;
        for(let r=0; r<Math.min(data.length, 10); r++) {
            const rowStr = data[r].join(' ').toLowerCase();
            if(rowStr.includes('start date') && (rowStr.includes('time of day') || rowStr.includes('run time'))) {
                headerRowIdx = r; break;
            }
        }
        if (headerRowIdx === -1) return { success: false, message: "Header row not found" };

        const headers = data[headerRowIdx];
        const hMap = headers.map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
        const idColIdx = hMap.indexOf('eventid');
        const zoomColIdx = hMap.indexOf('zoomlink');

        if (idColIdx === -1 || zoomColIdx === -1) return { success: false, message: "Columns not found" };

        let foundRow = -1;
        for (let i = headerRowIdx + 1; i < data.length; i++) {
            if (String(data[i][idColIdx]).trim() === String(courseId).trim()) {
                foundRow = i + 1;
                break;
            }
        }

        if (foundRow > -1) {
            sheet.getRange(foundRow, zoomColIdx + 1).setValue(zoomLink);
            return { success: true };
        } else {
            return { success: false, message: "Course ID not found" };
        }
    } catch (e) {
        return { success: false, error: e.message };
    } finally {
        lock.releaseLock();
    }
}

function api_previewMstCalendarSync(targetCalendarId) {
  try {
    const allSettings = getSettings();
    let mstSettings = {};
    try { mstSettings = JSON.parse(allSettings.mstSettings || '{}'); } catch (e) {}
    
    if (!mstSettings.sourceTabName) return { success: false, message: "MST Source Tab not configured." };
    if (!targetCalendarId) return { success: false, message: "No Calendar Selected." };

    // 1. Recover IDs (Optimized)
    recoverMstEventIds_(mstSettings.sourceTabName, targetCalendarId);

    const sheet = getSheet(mstSettings.sourceTabName);
    if (!sheet) return { success: false, message: `Tab '${mstSettings.sourceTabName}' not found.` };

    const data = sheet.getDataRange().getValues();
    let headerRowIdx = 0;
    for(let r=0; r<Math.min(data.length, 10); r++) {
        const rowStr = data[r].join(' ').toLowerCase();
        if(rowStr.includes('start date') && (rowStr.includes('time of day') || rowStr.includes('run time'))) {
            headerRowIdx = r; break;
        }
    }
    const headers = data[headerRowIdx];
    const rows = data.slice(headerRowIdx + 1);

    const assignData = getSheet('Staff_Assignments').getDataRange().getValues();
    const assignmentMap = new Map();
    for (let i = 1; i < assignData.length; i++) {
        if (assignData[i][2] === 'Course') {
            const staffId = String(assignData[i][1]).trim().toLowerCase();
            const eventId = String(assignData[i][3]).trim(); 
            if (staffId && eventId) assignmentMap.set(eventId, staffId);
        }
    }

    const hMap = headers.map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    const colIdx = {
        id: hMap.indexOf('eventid'),
        course: hMap.indexOf('course'),
        faculty: hMap.indexOf('faculty'),
        startDate: hMap.findIndex(h => h.includes('startdate')),
        endDate: hMap.findIndex(h => h.includes('enddate')),
        day: hMap.indexOf('day'),
        startTime: hMap.indexOf('runtime'), 
        ampm: hMap.indexOf('timeofday'),    
        location: hMap.indexOf('bxlocation'),
        duration: hMap.indexOf('coveragehrs'),
        zoomLink: hMap.indexOf('zoomlink') 
    };
    if (colIdx.startTime === -1) colIdx.startTime = hMap.findIndex(h => h.includes('time'));
    if (colIdx.location === -1) colIdx.location = hMap.findIndex(h => h.includes('location') || h.includes('room'));

    const calendar = CalendarApp.getCalendarById(targetCalendarId);
    if (!calendar) return { success: false, message: "Target Calendar not found." };
    
    // 2. Calculate Precise Date Range from Sheet Data
    let minDate = new Date(8640000000000000);
    let maxDate = new Date(-8640000000000000);
    let hasValidDates = false;

    rows.forEach(row => {
        if (colIdx.startDate > -1 && row[colIdx.startDate]) {
            const d = new Date(row[colIdx.startDate]);
            if (!isNaN(d.getTime())) {
                hasValidDates = true;
                if (d < minDate) minDate = d;
                const ed = (colIdx.endDate > -1 && row[colIdx.endDate]) ? new Date(row[colIdx.endDate]) : d;
                if (ed > maxDate) maxDate = ed;
            }
        }
    });

    const eventIdMap = new Map();
    
    // 3. Fetch Calendar Events (Only if we have dates, and only for that range)
    if (hasValidDates && minDate < maxDate) {
        // Add buffer to ensure we catch events on the boundary
        minDate.setHours(0,0,0,0);
        maxDate.setHours(23,59,59,999);
        
        const existingEvents = calendar.getEvents(minDate, maxDate);
        existingEvents.forEach(e => {
            let tagId = e.getTag('StaffHub_EventID');
            if (!tagId) {
                try {
                    const series = e.getEventSeries();
                    if (series) tagId = series.getTag('StaffHub_EventID');
                } catch(err) {}
            }
            if (tagId) eventIdMap.set(tagId, e);
        });
    }

    const proposals = [];
    const pattern = "{{Course}} - {{Faculty}}";

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowId = (colIdx.id > -1) ? String(row[colIdx.id]).trim() : null;
        if (!rowId) continue;

        let title = pattern;
        headers.forEach((h, idx) => { title = title.replace(new RegExp(`{{${h}}}`, 'gi'), row[idx]); });
        title = title.replace(/{{.*?}}/g, '').trim();
        const locationVal = (colIdx.location > -1) ? String(row[colIdx.location]).trim() : "";
        if (locationVal) title += ` (${locationVal})`;
        if (!title || title === '-') title = "Untitled Event";

        const courseName = (colIdx.course > -1) ? row[colIdx.course] : "";
        const faculty = (colIdx.faculty > -1) ? row[colIdx.faculty] : "";
        const zoomLink = (colIdx.zoomLink > -1) ? String(row[colIdx.zoomLink]).trim() : "";
        
        let description = `Course: ${courseName}\nFaculty: ${faculty}`;
        if (zoomLink) description += `\n\n--- RESOURCES ---\nZoom Link: ${zoomLink}`;

        let startDt = null;
        let endDt = null;
        let seriesEndDate = null;

        if (colIdx.startDate > -1 && row[colIdx.startDate]) {
             const dVal = row[colIdx.startDate];
             let tStr = (colIdx.startTime > -1) ? String(row[colIdx.startTime]) : "09:00";
             const ampmVal = (colIdx.ampm > -1) ? String(row[colIdx.ampm]).trim().toUpperCase() : "";
             startDt = new Date(dVal);
             if (!isNaN(startDt.getTime())) {
                 let startPart = tStr;
                 let endPart = null;
                 if (tStr.includes('-')) {
                     const parts = tStr.split('-');
                     startPart = parts[0].trim();
                     endPart = parts[1].trim();
                 }
                 const timeMatch = startPart.match(/(\d+):(\d+)/);
                 if (timeMatch) {
                     let h = parseInt(timeMatch[1]);
                     const m = parseInt(timeMatch[2]);
                     const isPm = ampmVal === 'PM';
                     if (isPm && h !== 12) h += 12;
                     else if (!isPm && h === 12) h = 0;
                     startDt.setHours(h, m, 0, 0);
                 } else { startDt.setHours(9, 0, 0, 0); }
                 
                 let durationMinutes = 60;
                 if (endPart) {
                     const endMatch = endPart.match(/(\d+):(\d+)/);
                     if (endMatch) {
                         let eh = parseInt(endMatch[1]);
                         const em = parseInt(endMatch[2]);
                         const isPm = ampmVal === 'PM';
                         let eh_abs = eh;
                         if (isPm && eh !== 12) eh_abs += 12;
                         else if (!isPm && eh === 12) eh_abs = 0;
                         if (eh_abs < startDt.getHours()) eh_abs += 12;
                         const tempEnd = new Date(startDt);
                         tempEnd.setHours(eh_abs, em, 0, 0);
                         const diff = (tempEnd - startDt) / 60000;
                         if (diff > 0) durationMinutes = diff;
                     }
                 } else if (colIdx.duration > -1 && row[colIdx.duration]) {
                     const durVal = row[colIdx.duration];
                     if (typeof durVal === 'number') durationMinutes = durVal * 60;
                 }
                 endDt = new Date(startDt.getTime() + (durationMinutes * 60000));
                 if (colIdx.endDate > -1 && row[colIdx.endDate]) {
                     seriesEndDate = new Date(row[colIdx.endDate]);
                     seriesEndDate.setHours(23, 59, 59);
                 }
             }
        }

        if (startDt && !isNaN(startDt.getTime())) {
            const dayStr = (colIdx.day > -1) ? String(row[colIdx.day]) : "";
            const targetEmail = assignmentMap.get(rowId);
            const targetEmailLower = targetEmail ? targetEmail.toLowerCase() : null;
            
            let status = "NEW";
            let diffs = [];
            let existingId = null;
            let currentGuests = [];
            let currentData = {};

            if (eventIdMap.has(rowId)) {
                const existing = eventIdMap.get(rowId);
                existingId = existing.getId();
                status = "SYNCED"; 
                currentGuests = existing.getGuestList().map(g => g.getEmail());
                
                currentData = {
                    title: existing.getTitle(),
                    location: existing.getLocation() || "",
                    description: existing.getDescription() || "",
                    guests: currentGuests
                };

                if (String(existing.getTitle()).trim() !== String(title).trim()) {
                    status = "UPDATE";
                    diffs.push({ key: 'title', type: 'update', text: `Title: "${existing.getTitle()}" -> "${title}"` });
                }
                const loc1 = String(existing.getLocation() || "").trim();
                const loc2 = String(locationVal || "").trim();
                if (loc1 !== loc2) {
                    status = "UPDATE";
                    diffs.push({ key: 'location', type: 'update', text: `Location: "${loc1}" -> "${loc2}"` });
                }
                
                const desc1 = String(existing.getDescription() || "").trim();
                const desc2 = String(description || "").trim();
                if (desc1 !== desc2) {
                    status = "UPDATE";
                    diffs.push({ key: 'description', type: 'update', text: "Description/Zoom Updated" });
                }
                
                let storedSig = existing.getTag('StaffHub_TimeSignature');
                if(!storedSig) {
                    try { storedSig = existing.getEventSeries().getTag('StaffHub_TimeSignature'); } catch(e){}
                }

                const timeSig = `${startDt.toISOString()}_${endDt.toISOString()}_${dayStr}`;
                if (storedSig !== timeSig) {
                     status = "UPDATE";
                     diffs.push({ key: 'time', type: 'update', text: "Time/Schedule Changed" });
                }

                const currentGuestsLower = currentGuests.map(e => e.toLowerCase());
                if (targetEmailLower && !currentGuestsLower.includes(targetEmailLower)) {
                    status = "UPDATE";
                    diffs.push({ key: 'guest_add', type: 'add', value: targetEmail, text: `Add Guest: ${targetEmail}` });
                }
                
                if (targetEmailLower) {
                     currentGuestsLower.forEach(g => {
                         if (g !== targetEmailLower) {
                             const originalEmail = currentGuests.find(e => e.toLowerCase() === g);
                             diffs.push({ key: 'guest_remove', type: 'remove', value: originalEmail, text: `Remove Guest: ${originalEmail}` });
                         }
                     });
                }
            }

            proposals.push({
                rowId: rowId,
                status: status, 
                diffs: diffs,
                existingEventId: existingId,
                currentData: currentData,
                currentGuests: currentGuests,
                seriesStartStr: formatDate(startDt, 'MM/dd/yyyy'),
                seriesEndStr: seriesEndDate ? formatDate(seriesEndDate, 'MM/dd/yyyy') : 'Single Event',
                payload: {
                    title: title,
                    startTime: startDt.getTime(), 
                    endTime: endDt.getTime(),
                    location: locationVal,
                    description: description,
                    zoomLink: zoomLink, 
                    dayStr: dayStr,
                    seriesEndDate: seriesEndDate ? seriesEndDate.getTime() : null,
                    guests: targetEmail ? [targetEmail] : []
                }
            });
        }
    }
    return { success: true, data: proposals };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_commitMstCalendarEvents(targetCalendarId, eventsToSync) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(60000); 
        const calendar = CalendarApp.getCalendarById(targetCalendarId);
        if (!calendar) return { success: false, message: "Calendar not found." };

        const staffSheet = getSheet('Staff_List');
        const staffEmails = new Set();
        if (staffSheet) {
            const data = staffSheet.getDataRange().getValues();
            for(let i=1; i<data.length; i++) {
                if(data[i][1]) staffEmails.add(String(data[i][1]).toLowerCase().trim());
            }
        }

        const stats = { created: 0, updated: 0, errors: 0 };

        eventsToSync.forEach(evt => {
            try {
                const p = evt.payload;
                const startDt = new Date(p.startTime);
                const endDt = new Date(p.endTime);
                
                const skipTitle = p.title === "SKIP";
                const skipLocation = p.location === "SKIP";
                const skipGuests = p.guests === "SKIP";
                const skipTime = p.startTime === "SKIP"; 
                
                const isSeriesRow = !!p.seriesEndDate;

                let recurrence = null;
                if (p.seriesEndDate && p.dayStr) {
                    const weekday = mstHelper_parseDayOfWeek(p.dayStr);
                    const seriesEnd = new Date(p.seriesEndDate);
                    if (weekday && seriesEnd > startDt) {
                        recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(weekday).until(seriesEnd);
                    }
                }

                const options = { description: p.description };
                if (!skipLocation) options.location = p.location;
                if (!skipGuests && p.guests && p.guests.length > 0 && p.guests !== "SKIP") options.guests = p.guests.join(',');

                const timeSig = `${startDt.toISOString()}_${endDt.toISOString()}_${p.dayStr || ''}`;

                if (evt.status === 'NEW' || !evt.existingEventId) {
                    let newEvent;
                    if (recurrence) newEvent = calendar.createEventSeries(p.title, startDt, endDt, recurrence, options);
                    else newEvent = calendar.createEvent(p.title, startDt, endDt, options);
                    
                    let tagTarget = newEvent;
                    if (newEvent.getEventSeries) { } 
                    tagTarget.setTag('StaffHub_EventID', evt.rowId);
                    tagTarget.setTag('StaffHub_TimeSignature', timeSig);
                    stats.created++;
                    Utilities.sleep(500); 

                } else {
                    let eventObj = calendar.getEventById(evt.existingEventId);
                    if (!eventObj) {
                        if (recurrence) calendar.createEventSeries(p.title, startDt, endDt, recurrence, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                        else calendar.createEvent(p.title, startDt, endDt, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                        stats.created++;
                    } else {
                        let currentSig = eventObj.getTag('StaffHub_TimeSignature');
                        if(!currentSig) {
                            try { currentSig = eventObj.getEventSeries().getTag('StaffHub_TimeSignature'); } catch(e){}
                        }
                        
                        if (!skipTime && currentSig && currentSig !== timeSig) {
                            try { eventObj.getEventSeries().deleteEventSeries(); } catch(e) { eventObj.deleteEvent(); }
                            if (recurrence) calendar.createEventSeries(p.title, startDt, endDt, recurrence, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                            else calendar.createEvent(p.title, startDt, endDt, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                            stats.updated++;
                        } else {
                            let target = eventObj;
                            if (isSeriesRow) {
                                try { target = eventObj.getEventSeries() || eventObj; } catch(e){}
                            }

                            if (!skipTitle) target.setTitle(p.title);
                            if (!skipLocation) target.setLocation(p.location);
                            target.setDescription(p.description);

                            if (!skipGuests) {
                                const desiredGuests = (p.guests || []);
                                const desiredGuestsLower = desiredGuests.map(e => e.toLowerCase());
                                const guestTarget = eventObj; 
                                const currentGuestList = guestTarget.getGuestList();
                                const currentEmailsLower = currentGuestList.map(g => g.getEmail().toLowerCase());

                                desiredGuests.forEach(email => {
                                    if (!currentEmailsLower.includes(email.toLowerCase())) guestTarget.addGuest(email);
                                });

                                currentGuestList.forEach(g => {
                                    const gEmail = g.getEmail().toLowerCase();
                                    if (!desiredGuestsLower.includes(gEmail)) {
                                        if (staffEmails.has(gEmail)) {
                                            try { guestTarget.removeGuest(gEmail); } catch(e) {}
                                        }
                                    }
                                });
                            }
                            stats.updated++;
                        }
                    }
                    Utilities.sleep(200);
                }
            } catch (err) {
                console.error("Sync Error for " + evt.rowId, err);
                stats.errors++;
            }
        });

        return { success: true, stats: stats };
    } catch (e) { return { success: false, message: e.message }; } finally { lock.releaseLock(); }
}

function recoverMstEventIds_(sheetName, calendarId) {
    if (!sheetName || !calendarId) return;
    const lock = LockService.getScriptLock(); 
    try {
        lock.waitLock(30000);
        
        const sheet = getSheet(sheetName);
        if (!sheet) return;
        
        const data = sheet.getDataRange().getValues();
        
        let headerRowIdx = -1;
        for(let r=0; r<Math.min(data.length, 10); r++) {
            const rowStr = data[r].join(' ').toLowerCase();
            if(rowStr.includes('start date') && (rowStr.includes('time of day') || rowStr.includes('run time'))) {
                headerRowIdx = r; break;
            }
        }
        if (headerRowIdx === -1) return;

        const headers = data[headerRowIdx];
        const hMap = headers.map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
        const idColIdx = hMap.indexOf('eventid');
        const courseIdx = hMap.indexOf('course');
        const startIdx = hMap.findIndex(h => h.includes('startdate'));
        const locIdx = hMap.indexOf('bxlocation');

        if (idColIdx === -1) return; 

        // --- SURGICAL RECOVERY: Only fetch calendar if IDs are missing ---
        const rowsMissingIds = [];
        for (let i = headerRowIdx + 1; i < data.length; i++) {
            if (!data[i][idColIdx] && data[i][courseIdx] && data[i][startIdx]) {
                rowsMissingIds.push({ rowIndex: i, rowData: data[i] });
            }
        }

        if (rowsMissingIds.length === 0) return; // Exit immediately if no work needed!

        // If we have missing IDs, we fetch ONLY the specific days needed
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) return;

        const idsToWrite = [];
        let hasUpdates = false;

        // Pre-fill idsToWrite with existing data so we can update specific indices
        for(let i=0; i<data.length - (headerRowIdx + 1); i++) {
            idsToWrite.push([data[headerRowIdx + 1 + i][idColIdx]]);
        }

        rowsMissingIds.forEach(item => {
            const row = item.rowData;
            const startDate = new Date(row[startIdx]);
            
            if (!isNaN(startDate.getTime())) {
                // Fetch events ONLY for this specific day
                const dayEvents = calendar.getEventsForDay(startDate);
                
                const courseName = String(row[courseIdx]).toLowerCase().trim();
                const sheetLoc = (locIdx > -1 && row[locIdx]) ? String(row[locIdx]).toLowerCase().trim() : "";
                
                let foundId = null;
                
                // Match logic
                const match = dayEvents.find(e => {
                    let tag = e.getTag('StaffHub_EventID');
                    if (!tag) { try { tag = e.getEventSeries().getTag('StaffHub_EventID'); } catch(err){} }
                    
                    // If tag exists, verify title matches. If no tag, verify title matches.
                    const eTitle = e.getTitle().toLowerCase();
                    if (!eTitle.includes(courseName)) return false;
                    
                    if (sheetLoc) {
                        const calLoc = (e.getLocation() || "").toLowerCase();
                        return calLoc.includes(sheetLoc) || sheetLoc.includes(calLoc);
                    }
                    return true;
                });

                if (match) {
                    let tag = match.getTag('StaffHub_EventID');
                    if (!tag) { try { tag = match.getEventSeries().getTag('StaffHub_EventID'); } catch(err){} }
                    foundId = tag;
                }

                if (!foundId) {
                    foundId = 'MST_' + Utilities.getUuid().split('-')[0].toUpperCase();
                }

                // Update our write buffer
                const writeIndex = item.rowIndex - (headerRowIdx + 1);
                idsToWrite[writeIndex][0] = foundId;
                hasUpdates = true;
            }
        });

        if (hasUpdates) {
            sheet.getRange(headerRowIdx + 2, idColIdx + 1, idsToWrite.length, 1).setValues(idsToWrite);
            SpreadsheetApp.flush();
        }
    } finally {
        lock.releaseLock();
    }
}

function mstHelper_parseDayOfWeek(dayStr) {
    if (!dayStr) return null;
    const s = dayStr.toLowerCase().trim();
    if (s === 'm' || s.includes('mon')) return CalendarApp.Weekday.MONDAY;
    if (s === 'tu' || s === 't' || s.includes('tue')) return CalendarApp.Weekday.TUESDAY;
    if (s === 'w' || s.includes('wed')) return CalendarApp.Weekday.WEDNESDAY;
    if (s === 'th' || s === 'r' || s.includes('thu')) return CalendarApp.Weekday.THURSDAY;
    if (s === 'f' || s.includes('fri')) return CalendarApp.Weekday.FRIDAY;
    if (s === 'sa' || s.includes('sat')) return CalendarApp.Weekday.SATURDAY;
    if (s === 'su' || s.includes('sun')) return CalendarApp.Weekday.SUNDAY;
    return null;
}

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
    return { id: row[idIdx], eventId: row[eventIdx], staffId: row[staffIdx] };
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
    const zoomIdx = map['zoomlink']; 
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
        startDate: startIdx !== undefined ? row[startIdx] : null,
        endDate: endIdx !== undefined ? row[endIdx] : null,
        location: locIdx !== undefined ? row[locIdx] : '',
        zoomLink: zoomIdx !== undefined ? row[zoomIdx] : ''
    };
}

function formatDate(date, format) {
    if (!date || !(date instanceof Date)) return '';
    return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}