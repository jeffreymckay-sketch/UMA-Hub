/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING, SETTINGS & CALENDAR SYNC
 * -------------------------------------------------------------------
 */

/**
 * Core Sync Engine: Runs automatically to pull external data, generate 
 * fingerprints, and merge updates into the local database without losing IDs.
 */
function syncExternalMstData() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return; // Prevent simultaneous syncs

    try {
        const settings = getSettings();
        const extUrl = settings.mstExternalUrl;
        const extTab = settings.mstExternalTab;
        if (!extUrl || !extTab) return; // Sync not configured yet

        const ss = getMasterDataHub();
        let localSheet = ss.getSheetByName('Course_Schedule');
        
        // 1. Standardize Headers
        const standardHeaders = ['Source', 'eventID', 'Zoom Link', 'Session', 'Start Date', 'End Date', 'Day', 'Course', 'Faculty', 'Start Time', 'End Time', 'BX Location', 'MST Assigned by email', 'Coverage', 'Note'];
        
        // Setup local sheet if missing or if it has the old IMPORTRANGE formula
        if (!localSheet) {
            localSheet = ss.insertSheet('Course_Schedule');
            localSheet.appendRow(standardHeaders);
        } else {
            const formula = localSheet.getRange("A1").getFormula();
            if (formula && formula.toUpperCase().includes("IMPORTRANGE")) {
                localSheet.clear();
                localSheet.appendRow(standardHeaders);
            }
        }

        const localData = localSheet.getDataRange().getValues();
        const localHeaders = getColumnMap(localData[0] || standardHeaders);
        
        const localMap = new Map();
        const manualClasses = [];

        // Helper: Creates a unique composite key for a class
        function createFingerprint(course, faculty, day, starttime) {
            let sTimeStr = "";
            if (starttime instanceof Date) {
                sTimeStr = starttime.toTimeString().substring(0,5); // HH:mm
            } else {
                sTimeStr = String(starttime);
            }
            return String(course + faculty + day + sTimeStr).toLowerCase().replace(/[^a-z0-9]/g, '');
        }

        // 2. Parse existing local data
        if (localData.length > 1) {
            for (let i = 1; i < localData.length; i++) {
                const row = localData[i];
                const source = row[localHeaders['source']];
                const evtId = row[localHeaders['eventid']];
                if (!evtId) continue;

                const obj = {};
                for (const key in localHeaders) obj[key] = row[localHeaders[key]];

                if (source === 'Manual') {
                    manualClasses.push(obj);
                } else {
                    const fp = createFingerprint(obj.course, obj.faculty, obj.day, obj.starttime);
                    localMap.set(fp, obj);
                }
            }
        }

        // 3. Fetch External Data
        let extSS;
        try {
            const match = extUrl.match(/[-\w]{25,}/);
            extSS = SpreadsheetApp.openById(match ? match[0] : extUrl);
        } catch(e) {
            console.error("MST Sync: Could not open external sheet.");
            return;
        }

        const extSheet = extSS.getSheetByName(extTab);
        if (!extSheet) return;

        const extData = extSheet.getDataRange().getValues();
        if (extData.length < 2) return;
        const extHeaders = getColumnMap(extData[0]);

        const newLocalData = [standardHeaders];
        
        // 4. Merge External Data into Local Database
        for (let i = 1; i < extData.length; i++) {
            const extRow = extData[i];
            const course = extRow[extHeaders['course']];
            if (!course) continue; // Skip empty rows

            const faculty = extRow[extHeaders['faculty']];
            const day = extRow[extHeaders['day']];
            const startTime = extRow[extHeaders['starttime']];
            
            const fp = createFingerprint(course, faculty, day, startTime);
            
            let localObj = localMap.get(fp);
            if (localObj) {
                // UPDATE: Preserve eventID and Zoom, update dynamic fields
                localObj.session = extRow[extHeaders['session']];
                localObj.startdate = extRow[extHeaders['startdate']];
                localObj.enddate = extRow[extHeaders['enddate']];
                localObj.endtime = extRow[extHeaders['endtime']];
                localObj.bxlocation = extRow[extHeaders['bxlocation']];
                localObj.mstassignedbyemail = extRow[extHeaders['mstassignedbyemail']];
                
                const covKey = Object.keys(extHeaders).find(k => k.includes('coverage'));
                localObj.coverage = covKey ? extRow[extHeaders[covKey]] : '';
                localObj.note = extRow[extHeaders['note']];
                
                newLocalData.push(mapObjToArray(localObj, standardHeaders));
                localMap.delete(fp); // Mark as processed
            } else {
                // CREATE NEW: Generate a permanent UUID
                const covKey = Object.keys(extHeaders).find(k => k.includes('coverage'));
                const newObj = {
                    source: 'External',
                    eventid: Utilities.getUuid(),
                    zoomlink: '',
                    session: extRow[extHeaders['session']],
                    startdate: extRow[extHeaders['startdate']],
                    enddate: extRow[extHeaders['enddate']],
                    day: day,
                    course: course,
                    faculty: faculty,
                    starttime: startTime,
                    endtime: extRow[extHeaders['endtime']],
                    bxlocation: extRow[extHeaders['bxlocation']],
                    mstassignedbyemail: extRow[extHeaders['mstassignedbyemail']],
                    coverage: covKey ? extRow[extHeaders[covKey]] : '',
                    note: extRow[extHeaders['note']]
                };
                newLocalData.push(mapObjToArray(newObj, standardHeaders));
            }
        }

        // 5. Re-append Manual Classes (Keeps them safe from sync deletion)
        manualClasses.forEach(obj => newLocalData.push(mapObjToArray(obj, standardHeaders)));

        // 6. Write back to database
        localSheet.clearContents();
        localSheet.getRange(1, 1, newLocalData.length, newLocalData[0].length).setValues(newLocalData);

    } catch (e) {
        console.error("MST Sync Error: " + e.stack);
    } finally {
        lock.releaseLock();
    }
}

// Helper for array mapping
function mapObjToArray(obj, headers) {
    return headers.map(h => {
        const key = String(h).toLowerCase().replace(/[\s_]/g, '');
        return obj[key] !== undefined ? obj[key] : '';
    });
}

function getLocalMstData() {
    const ss = getMasterDataHub();
    const tz = ss.getSpreadsheetTimeZone(); 
    let sheet = ss.getSheetByName('Course_Schedule');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    const objs = [];
    
    for(let i = 1; i < data.length; i++) {
        const obj = {};
        headers.forEach((h, idx) => {
            const key = String(h).toLowerCase().replace(/[\s_]/g, '');
            let val = data[i][idx];
            
            if (val instanceof Date) {
                if (key === 'starttime' || key === 'endtime') {
                    val = Utilities.formatDate(val, tz, "HH:mm");
                } else if (key === 'startdate' || key === 'enddate') {
                    val = Utilities.formatDate(val, tz, "MM/dd/yyyy");
                }
            }
            if (key) obj[key] = val;
        });
        objs.push(obj);
    }
    return objs;
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
    // 1. Run the sync engine first
    syncExternalMstData();

    // 2. Fetch local data for UI
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        const localData = getLocalMstData();

        let staffData = [];
        let assignData = [];
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
        }).filter(s => s && s.id && s.isActive) : [];

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

        const courseAssignmentsView = localData.map(courseObj => {
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

        return JSON.parse(JSON.stringify({ success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } }));

    } catch(e) {
        console.error(e);
        return { success: false, message: e.message || String(e) };
    } finally {
        lock.releaseLock();
    }
}

// --- MANUAL EDITS (Directly to Database) ---

function api_mst_addCourse(courseDetails) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName('Course_Schedule');
        
        courseDetails.source = 'Manual';
        courseDetails.eventid = Utilities.getUuid();
        
        const standardHeaders = ['Source', 'eventID', 'Zoom Link', 'Session', 'Start Date', 'End Date', 'Day', 'Course', 'Faculty', 'Start Time', 'End Time', 'BX Location', 'MST Assigned by email', 'Coverage', 'Note'];
        sheet.appendRow(mapObjToArray(courseDetails, standardHeaders));
        return { success: true };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_updateCourse(courseDetails) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName('Course_Schedule');
        const data = sheet.getDataRange().getValues();
        const headers = getColumnMap(data[0]);
        
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][headers.eventid]) === String(courseDetails.eventid)) {
                // Update cells directly
                for (const key in courseDetails) {
                    if (headers[key] !== undefined && key !== 'eventid' && key !== 'source') {
                        sheet.getRange(i + 1, headers[key] + 1).setValue(courseDetails[key]);
                    }
                }
                return { success: true };
            }
        }
        return { success: false, message: "Class not found in database." };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_deleteCourse(courseId) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName('Course_Schedule');
        const data = sheet.getDataRange().getValues();
        const headers = getColumnMap(data[0]);
        
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][headers.eventid]) === String(courseId)) {
                sheet.deleteRow(i + 1);
                return { success: true };
            }
        }
        return { success: false, message: "Class not found." };
    } catch(e) { return { success: false, message: e.message || String(e) }; }
    finally { lock.releaseLock(); }
}

function api_mst_updateCourseZoom(courseId, zoomLink) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        requireRole(['MST', 'Lead', 'Admin']);
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName('Course_Schedule');
        const data = sheet.getDataRange().getValues();
        const headers = getColumnMap(data[0]);
        
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][headers.eventid]) === String(courseId)) {
                sheet.getRange(i + 1, headers.zoomlink + 1).setValue(sanitizeInput(zoomLink));
                return { success: true };
            }
        }
        return { success: false, message: "Class not found." };
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

// --- SETTINGS & SYNC CONFIG ---

function api_saveMstExternalSettings(url, tabName) {
    try {
        requireRole(['MST', 'Lead', 'Admin']);
        saveSetting('mstExternalUrl', url);
        saveSetting('mstExternalTab', tabName);
        return { success: true, message: "Sync settings saved! The schedule will now sync automatically." };
    } catch(e) {
        return { success: false, message: e.message || String(e) };
    }
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

// --- CALENDAR SYNC (Untouched logic, mapped to new local data) ---

function api_previewMstCalendarSync(targetCalendarId) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        if (!targetCalendarId) return { success: false, message: "No Calendar Selected." };

        const localData = getLocalMstData();
        if (localData.length === 0) return { success: false, message: "No course data found." };

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

        localData.forEach(row => {
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

        const proposals = [];

        localData.forEach(row => {
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
                        guests: targetEmailLower ? [targetEmail] : []
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
                                const desiredGuests = (p.guests || []);
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