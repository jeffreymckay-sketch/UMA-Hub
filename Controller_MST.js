/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING, SETTINGS & CALENDAR SYNC
 * -------------------------------------------------------------------
 */

function getMergedMstData() {
    const ss = getMasterDataHub();
    const tz = ss.getSpreadsheetTimeZone(); 
    
    let baseSheet = null;
    try { baseSheet = getSheet('Course_Schedule'); } catch(e) { baseSheet = ss.getSheetByName('Course_Schedule'); }
    if (!baseSheet) throw new Error("Course_Schedule tab not found.");

    const baseData = baseSheet.getDataRange().getValues();
    const baseHeaders = baseData[0];
    const baseObjs =[];
    
    if (baseData.length > 1) {
        for(let i = 1; i < baseData.length; i++) {
            const obj = {};
            baseHeaders.forEach((h, idx) => {
                const key = String(h).toLowerCase().replace(/[\s_]/g, '');
                let val = baseData[i][idx];
                
                if (val instanceof Date) {
                    if (key === 'starttime' || key === 'endtime') {
                        val = Utilities.formatDate(val, tz, "HH:mm");
                    } else if (key === 'startdate' || key === 'enddate') {
                        val = Utilities.formatDate(val, tz, "MM/dd/yyyy");
                    }
                }
                
                if (key) obj[key] = val;
            });
            baseObjs.push(obj);
        }
    }

    let editsSheet = null;
    try { editsSheet = getSheet('Course_Schedule_Edits'); } catch(e) { editsSheet = ss.getSheetByName('Course_Schedule_Edits'); }
    
    if (!editsSheet) {
        editsSheet = ss.insertSheet('Course_Schedule_Edits');
        editsSheet.appendRow(['Action_Type', 'Session', 'Start Date', 'End Date', 'Day', 'Course', 'Faculty', 'Start Time', 'End Time', 'BX Location', 'MST Assigned by email', 'Coverage', 'Note', 'eventID', 'Zoom Link']);
    }

    const editObjs =[];
    const editsData = editsSheet.getDataRange().getValues();
    
    if (editsData.length > 1) {
        const eHeaders = editsData[0];
        for(let i = 1; i < editsData.length; i++) {
            const obj = {};
            eHeaders.forEach((h, idx) => {
                const key = String(h).toLowerCase().replace(/[\s_]/g, '');
                let val = editsData[i][idx];
                
                if (val instanceof Date) {
                    if (key === 'starttime' || key === 'endtime') {
                        val = Utilities.formatDate(val, tz, "HH:mm");
                    } else if (key === 'startdate' || key === 'enddate') {
                        val = Utilities.formatDate(val, tz, "MM/dd/yyyy");
                    }
                }
                
                if (key) obj[key] = val;
            });
            editObjs.push(obj);
        }
    }

    const mergedMap = new Map();
    
    baseObjs.forEach(obj => {
        if (obj.eventid) mergedMap.set(String(obj.eventid).trim(), obj);
    });

    editObjs.forEach(obj => {
        const id = obj.eventid ? String(obj.eventid).trim() : null;
        if (!id) return;
        const action = String(obj.actiontype).toUpperCase();
        
        if (action === 'DELETE') {
            mergedMap.delete(id);
        } else if (action === 'ADD' || action === 'EDIT') {
            const existing = mergedMap.get(id) || {};
            mergedMap.set(id, Object.assign({}, existing, obj));
        }
    });

    return Array.from(mergedMap.values());
}

function combineDateAndTime(dateVal, timeVal) {
    if (!dateVal) return null;
    let d = new Date(dateVal);
    if (isNaN(d.getTime())) return null;

    if (timeVal) {
        if (timeVal instanceof Date) {
            d.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
        } else {
            const match = String(timeVal).trim().match(/(\d+):(\d+)(?::\d+)?\s*(AM|PM|am|pm)?/i);
            if (match) {
                let h = parseInt(match[1], 10);
                const m = parseInt(match[2], 10);
                const ampm = (match[3] || '').toUpperCase();
                if (ampm === 'PM' && h < 12) h += 12;
                if (ampm === 'AM' && h === 12) h = 0;
                d.setHours(h, m, 0, 0);
            }
        }
    }
    return d;
}

function api_getMstViewData() {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const mergedData = getMergedMstData();

        let staffData =[];
        let assignData =[];
        try { staffData = getSheet('Staff_List').getDataRange().getValues(); } catch(e) {}
        try { assignData = getSheet('Staff_Assignments').getDataRange().getValues(); } catch(e) {}

        const staffHeaders = staffData.length > 0 ? getColumnMap(staffData[0]) : {};
        const assignHeaders = assignData.length > 0 ? getColumnMap(assignData[0]) : {};

        const allStaff = staffData.length > 1 ? staffData.slice(1).map(row => {
            const idIdx = staffHeaders['staffid'] !== undefined ? staffHeaders['staffid'] : staffHeaders['id'];
            const nameIdx = staffHeaders['fullname'] !== undefined ? staffHeaders['fullname'] : staffHeaders['name'];
            const roleIdx = staffHeaders['roles'];
            const actIdx = staffHeaders['isactive'];
            return {
                id: String(row[idIdx] || ''), 
                name: String(row[nameIdx] || ''), 
                role: String(row[roleIdx] || ''), 
                isActive: (String(row[actIdx]).toUpperCase() === 'TRUE' || row[actIdx] === true)
            };
        }).filter(s => s && s.id && s.isActive) :[];

        const staffMap = new Map(allStaff.map(s => [String(s.id).toLowerCase(), s]));

        const assignmentMap = new Map();
        if (assignData.length > 1) {
            for(let i=1; i<assignData.length; i++) {
                const r = assignData[i];
                if (r[assignHeaders.assignmenttype] === 'Course') {
                    assignmentMap.set(String(r[assignHeaders.referenceid]), String(r[assignHeaders.staffid]));
                }
            }
        }

        const courseAssignmentsView = mergedData.map(courseObj => {
            const id = String(courseObj.eventid);
            
            let staffId = null;
            if (assignmentMap.has(id)) {
                staffId = assignmentMap.get(id); 
            } else if (courseObj.mstassignedbyemail) {
                staffId = String(courseObj.mstassignedbyemail).trim();
            }
            
            const staff = staffId ? staffMap.get(staffId.toLowerCase()) : null;

            const startD = combineDateAndTime(courseObj.startdate, courseObj.starttime);
            const endD = combineDateAndTime(courseObj.startdate, courseObj.endtime);
            const seriesEndD = new Date(courseObj.enddate);

            let timeDisplay = "TBD";
            if (startD && endD && !isNaN(startD.getTime()) && !isNaN(endD.getTime())) {
                timeDisplay = Utilities.formatDate(startD, Session.getScriptTimeZone(), 'h:mm a') + ' - ' + Utilities.formatDate(endD, Session.getScriptTimeZone(), 'h:mm a');
            }

            let sStr = "?"; if (startD && !isNaN(startD.getTime())) sStr = Utilities.formatDate(startD, Session.getScriptTimeZone(), 'M/d/yy');
            let eStr = "?"; if (seriesEndD && !isNaN(seriesEndD.getTime())) eStr = Utilities.formatDate(seriesEndD, Session.getScriptTimeZone(), 'M/d/yy');

            return {
                id: id,
                itemName: String(courseObj.course || "Untitled"),
                courseFaculty: String(courseObj.faculty || ""),
                courseDay: String(courseObj.day || ""),
                courseTime: timeDisplay,
                startDateStr: sStr,
                endDateStr: eStr,
                location: String(courseObj.bxlocation || ""),
                zoomLink: String(courseObj.zoomlink || ""),
                staffName: staff ? staff.name : "Unassigned",
                staffId: staff ? staff.id : null,
                raw: courseObj
            };
        });

        const mstStaffList = allStaff.filter(s => s.role && s.role.toLowerCase().includes('mst')).map(s => ({ id: s.id, name: s.name }));

        const payload = { success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } };

        return JSON.parse(JSON.stringify(payload));

    } catch(e) {
        console.error(e);
        return { success: false, message: e.message || String(e) };
    } finally {
        lock.releaseLock();
    }
}

function saveCourseEditState(actionType, courseObj) {
    const ss = getMasterDataHub();
    let sheet = ss.getSheetByName('Course_Schedule_Edits');
    
    if (!sheet) {
        sheet = ss.insertSheet('Course_Schedule_Edits');
    }
    
    let data = sheet.getDataRange().getValues();
    let headers = data[0];
    let headersChanged = false;
    
    if (data.length === 1 && headers.join('').trim() === '') {
        headers =['Action_Type', 'Session', 'Start Date', 'End Date', 'Day', 'Course', 'Faculty', 'Start Time', 'End Time', 'BX Location', 'MST Assigned by email', 'Coverage', 'Note', 'eventID', 'Zoom Link'];
        headersChanged = true;
    }
    
    const map = {};
    headers.forEach((h, i) => map[String(h).toLowerCase().replace(/[\s_]/g, '')] = i);

    const newRow = new Array(headers.length).fill('');
    
    if (map['actiontype'] !== undefined) {
        newRow[map['actiontype']] = actionType;
    } else {
        map['actiontype'] = headers.length;
        headers.push('Action_Type');
        newRow.push(actionType);
        headersChanged = true;
    }

    for (const key in courseObj) {
        const cleanKey = key.toLowerCase().replace(/[\s_]/g, '');
        let val = courseObj[key];
        
        if (typeof val === 'string' && val.includes('T12:00:00Z')) {
            val = new Date(val);
        }

        if (map[cleanKey] !== undefined) {
            newRow[map[cleanKey]] = val;
        } else if (cleanKey !== 'actiontype') { 
            map[cleanKey] = headers.length;
            headers.push(key);
            newRow.push(val);
            headersChanged = true;
        }
    }
    
    sheet.appendRow(newRow);
    if (headersChanged) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function api_mst_addCourse(courseDetails) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        courseDetails.eventid = Utilities.getUuid();
        saveCourseEditState('ADD', courseDetails);
        return { success: true };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_updateCourse(courseDetails) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        const merged = getMergedMstData().find(c => String(c.eventid) === String(courseDetails.eventid)) || {};
        saveCourseEditState('EDIT', Object.assign({}, merged, courseDetails));
        return { success: true };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_deleteCourse(courseId) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        saveCourseEditState('DELETE', { eventid: courseId });
        return { success: true };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_updateCourseZoom(courseId, zoomLink) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        const merged = getMergedMstData().find(c => String(c.eventid) === String(courseId)) || {};
        merged.zoomlink = sanitizeInput(zoomLink);
        merged.eventid = courseId;
        saveCourseEditState('EDIT', merged);
        return { success: true };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_updateCourseAssignment(courseId, staffId) {
    const lock = LockService.getScriptLock();
    try {
        requireRole(['MST', 'Lead', 'Admin']);
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
            sheet.getRange(foundRow, staffCol + 1).setValue(sanitizeInput(staffId));
        } else {
            const newRow = new Array(data[0].length).fill("");
            newRow[idCol] = Utilities.getUuid();
            newRow[staffCol] = sanitizeInput(staffId);
            newRow[refCol] = courseId;
            if(typeCol !== undefined) newRow[typeCol] = "Course";
            sheet.appendRow(newRow);
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message || String(e) };
    } finally {
        lock.releaseLock();
    }
}

function api_previewMstCalendarSync(targetCalendarId) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        if (!targetCalendarId) return { success: false, message: "No Calendar Selected." };

        const mergedData = getMergedMstData();
        if (mergedData.length === 0) return { success: false, message: "No course data found." };

        const assignData = getSheet('Staff_Assignments').getDataRange().getValues();
        const assignmentMap = new Map();
        for (let i = 1; i < assignData.length; i++) {
            if (assignData[i][2] === 'Course') {
                const staffId = String(assignData[i][1]).trim().toLowerCase();
                const eventId = String(assignData[i][3]).trim(); 
                if (eventId) assignmentMap.set(eventId, staffId);
            }
        }

        const calendar = CalendarApp.getCalendarById(targetCalendarId);
        if (!calendar) return { success: false, message: "Target Calendar not found." };
        
        let minDate = new Date(8640000000000000);
        let maxDate = new Date(-8640000000000000);
        let hasValidDates = false;

        mergedData.forEach(row => {
            const startDt = combineDateAndTime(row.startdate, row.starttime);
            const endDt = new Date(row.enddate);
            if (startDt && !isNaN(startDt.getTime())) {
                hasValidDates = true;
                if (startDt < minDate) minDate = startDt;
                if (!isNaN(endDt.getTime()) && endDt > maxDate) maxDate = endDt;
            }
        });

        const eventIdMap = new Map();
        
        if (hasValidDates && minDate < maxDate) {
            minDate.setHours(0,0,0,0);
            maxDate.setHours(23,59,59,999);
            const existingEvents = calendar.getEvents(minDate, maxDate);
            const seriesCache = {};
            
            existingEvents.forEach(e => {
                let tagId = e.getTag('StaffHub_EventID');
                let timeSig = e.getTag('StaffHub_TimeSignature');
                
                if (!tagId || !timeSig) {
                    const baseId = (e.getId() || "").split('_')[0];
                    if (seriesCache[baseId] !== undefined) {
                        if (!tagId) tagId = seriesCache[baseId].tagId;
                        if (!timeSig) timeSig = seriesCache[baseId].timeSig;
                    } else {
                        try {
                            const series = e.getEventSeries();
                            if (series) {
                                const sTag = series.getTag('StaffHub_EventID');
                                const sSig = series.getTag('StaffHub_TimeSignature');
                                seriesCache[baseId] = { tagId: sTag, timeSig: sSig };
                                if (!tagId) tagId = sTag;
                                if (!timeSig) timeSig = sSig;
                            } else { seriesCache[baseId] = { tagId: null, timeSig: null }; }
                        } catch(err) { seriesCache[baseId] = { tagId: null, timeSig: null }; }
                    }
                }
                if (tagId) {
                    e._cachedTimeSig = timeSig; 
                    eventIdMap.set(tagId, e);
                }
            });
        }

        const proposals =[];

        mergedData.forEach(row => {
            const rowId = String(row.eventid);
            if (!rowId) return;

            let title = "MST: " + (row.course || "Untitled") + " - " + (row.faculty || "No Faculty");
            const locationVal = String(row.bxlocation || "").trim();
            if (locationVal) title += " (" + locationVal + ")";

            const zoomLink = String(row.zoomlink || "").trim();
            let description = "Course: " + (row.course || "") + "\nFaculty: " + (row.faculty || "");
            if (zoomLink) description += "\n\n--- RESOURCES ---\nZoom Link: " + zoomLink;

            let startDt = combineDateAndTime(row.startdate, row.starttime);
            let endDt = combineDateAndTime(row.startdate, row.endtime);
            let seriesEndDate = new Date(row.enddate);
            
            if (startDt && !isNaN(startDt.getTime()) && endDt && !isNaN(endDt.getTime())) {
                seriesEndDate.setHours(23, 59, 59);
                const dayStr = String(row.day || "");
                
                let targetEmail = undefined;
                if (assignmentMap.has(rowId)) {
                    targetEmail = assignmentMap.get(rowId);
                } else if (row.mstassignedbyemail) {
                    targetEmail = String(row.mstassignedbyemail).trim();
                }
                const targetEmailLower = targetEmail ? targetEmail.toLowerCase() : null;
                
                let status = "NEW";
                let diffs =[];
                let existingId = null;
                let currentGuests =[];
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
                        diffs.push({ key: 'title', type: 'update', text: "Title: '" + existing.getTitle() + "' -> '" + title + "'" });
                    }
                    const loc1 = String(existing.getLocation() || "").trim();
                    if (loc1 !== locationVal) {
                        status = "UPDATE";
                        diffs.push({ key: 'location', type: 'update', text: "Location: '" + loc1 + "' -> '" + locationVal + "'" });
                    }
                    if (String(existing.getDescription() || "").trim() !== String(description).trim()) {
                        status = "UPDATE";
                        diffs.push({ key: 'description', type: 'update', text: "Description/Zoom Updated" });
                    }
                    
                    const timeSig = startDt.toISOString() + "_" + endDt.toISOString() + "_" + dayStr;
                    if ((existing._cachedTimeSig || existing.getTag('StaffHub_TimeSignature')) !== timeSig) {
                         status = "UPDATE";
                         diffs.push({ key: 'time', type: 'update', text: "Time/Schedule Changed" });
                    }

                    const currentGuestsLower = currentGuests.map(e => e.toLowerCase());
                    if (targetEmailLower && !currentGuestsLower.includes(targetEmailLower)) {
                        status = "UPDATE";
                        diffs.push({ key: 'guest_add', type: 'add', value: targetEmail, text: "Add Guest: " + targetEmail });
                    }
                    
                    if (targetEmailLower) {
                         currentGuestsLower.forEach(g => {
                             if (g !== targetEmailLower) {
                                 const originalEmail = currentGuests.find(e => e.toLowerCase() === g);
                                 diffs.push({ key: 'guest_remove', type: 'remove', value: originalEmail, text: "Remove Guest: " + originalEmail });
                             }
                         });
                    }
                }

                proposals.push({
                    rowId: rowId, status: status, diffs: diffs, existingEventId: existingId,
                    currentData: currentData, currentGuests: currentGuests,
                    seriesStartStr: Utilities.formatDate(startDt, Session.getScriptTimeZone(), 'M/d/yy'),
                    seriesEndStr: !isNaN(seriesEndDate.getTime()) ? Utilities.formatDate(seriesEndDate, Session.getScriptTimeZone(), 'M/d/yy') : 'Single Event',
                    payload: {
                        title: title, startTime: startDt.getTime(), endTime: endDt.getTime(),
                        location: locationVal, description: description, zoomLink: zoomLink, 
                        dayStr: dayStr, seriesEndDate: !isNaN(seriesEndDate.getTime()) ? seriesEndDate.getTime() : null,
                        guests: targetEmailLower ? [targetEmail] :[]
                    }
                });
            }
        });
        return { success: true, data: proposals };
    } catch(e) { 
        return { success: false, message: e.message || String(e) }; 
    } finally { 
        lock.releaseLock(); 
    }
}

function api_commitMstCalendarEvents(targetCalendarId, eventsToSync) {
    const lock = LockService.getScriptLock();
    try {
        requireRole(['MST', 'Lead', 'Admin']);
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

                const options = { description: sanitizeInput(p.description) };
                if (!skipLocation) options.location = sanitizeInput(p.location);
                if (!skipGuests && p.guests && p.guests.length > 0 && p.guests !== "SKIP") options.guests = p.guests.join(',');

                const timeSig = startDt.toISOString() + "_" + endDt.toISOString() + "_" + (p.dayStr || '');

                if (evt.status === 'NEW' || !evt.existingEventId) {
                    let newEvent;
                    if (recurrence) newEvent = calendar.createEventSeries(sanitizeInput(p.title), startDt, endDt, recurrence, options);
                    else newEvent = calendar.createEvent(sanitizeInput(p.title), startDt, endDt, options);
                    
                    newEvent.setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                    stats.created++;
                    Utilities.sleep(500); 

                } else {
                    let eventObj = calendar.getEventById(evt.existingEventId);
                    if (!eventObj) {
                        if (recurrence) calendar.createEventSeries(sanitizeInput(p.title), startDt, endDt, recurrence, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                        else calendar.createEvent(sanitizeInput(p.title), startDt, endDt, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                        stats.created++;
                    } else {
                        let currentSig = eventObj.getTag('StaffHub_TimeSignature');
                        if(!currentSig) { try { currentSig = eventObj.getEventSeries().getTag('StaffHub_TimeSignature'); } catch(e){} }
                        
                        if (!skipTime && currentSig && currentSig !== timeSig) {
                            try { eventObj.getEventSeries().deleteEventSeries(); } catch(e) { eventObj.deleteEvent(); }
                            if (recurrence) calendar.createEventSeries(sanitizeInput(p.title), startDt, endDt, recurrence, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                            else calendar.createEvent(sanitizeInput(p.title), startDt, endDt, options).setTag('StaffHub_EventID', evt.rowId).setTag('StaffHub_TimeSignature', timeSig);
                            stats.updated++;
                        } else {
                            let target = eventObj;
                            if (isSeriesRow) { try { target = eventObj.getEventSeries() || eventObj; } catch(e){} }

                            if (!skipTitle) target.setTitle(sanitizeInput(p.title));
                            if (!skipLocation) target.setLocation(sanitizeInput(p.location));
                            target.setDescription(sanitizeInput(p.description));

                            if (!skipGuests) {
                                const desiredGuests = (p.guests ||[]);
                                const desiredGuestsLower = desiredGuests.map(e => e.toLowerCase());
                                const currentGuestList = target.getGuestList();
                                const currentEmailsLower = currentGuestList.map(g => g.getEmail().toLowerCase());

                                desiredGuests.forEach(email => { if (!currentEmailsLower.includes(email.toLowerCase())) target.addGuest(email); });
                                currentGuestList.forEach(g => {
                                    const gEmail = g.getEmail().toLowerCase();
                                    if (!desiredGuestsLower.includes(gEmail) && staffEmails.has(gEmail)) {
                                        try { target.removeGuest(gEmail); } catch(e) {}
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
    } catch (e) { return { success: false, message: e.message || String(e) }; } 
    finally { lock.releaseLock(); }
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

function api_getExternalTabs(url) {
    try {
        requireRole('MST', 'Lead', 'Admin');
        const match = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        const id = match ? match[1] : url;
        const ss = SpreadsheetApp.openById(id);
        return { success: true, data: ss.getSheets().map(s => s.getName()) };
    } catch(e) {
        return { success: false, message: "Could not read external sheet. Ensure you have permission to view it. Error: " + (e.message || String(e)) };
    }
}

function api_applyImportRange(url, tabName) {
    try {
        requireRole('MST', 'Lead', 'Admin');
        const ss = getMasterDataHub();
        let sheet = ss.getSheetByName('Course_Schedule');
        if (!sheet) sheet = ss.insertSheet('Course_Schedule');
        
        sheet.clear();
        const formula = '=IMPORTRANGE("' + url + '", "' + tabName + '!A:ZZ")';
        sheet.getRange("A1").setFormula(formula);
        
        return { success: true, message: "Formula applied successfully. Please remember to click 'Allow access' on the cell if this is a new sheet." };
    } catch(e) {
        return { success: false, message: e.message || String(e) };
    }
}