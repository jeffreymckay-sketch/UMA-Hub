/**
 * -------------------------------------------------------------------
 * CONTROLLER: TECH HUB SCHEDULING
 * Handles Shifts, Roster, Availability, and Calendar Sync
 * -------------------------------------------------------------------
 */

// --- CLIENT-CALLABLE API FUNCTIONS ---

function api_getTechHubViewData() {
    try {
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const ss = getMasterDataHub();
        const tz = ss.getSpreadsheetTimeZone(); 

        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target; 
        };

        const staffSheet = ss.getSheetByName(sheetTabs[findKey('Staff_List')]);
        const shiftsSheet = ss.getSheetByName(sheetTabs[findKey('TechHub_Shifts')]);
        const assignSheet = ss.getSheetByName(sheetTabs[findKey('Staff_Assignments')]);
        
        let availSheet = ss.getSheetByName(sheetTabs[findKey('Staff_Availability')]);
        let prefsSheet = ss.getSheetByName(sheetTabs[findKey('Staff_Preferences')]);

        if (!staffSheet || !shiftsSheet || !assignSheet) {
            throw new Error("Required sheets are missing.");
        }

        const staffData = staffSheet.getDataRange().getValues();
        const shiftsData = shiftsSheet.getDataRange().getValues();
        const assignData = assignSheet.getDataRange().getValues();
        
        const availData = availSheet ? availSheet.getDataRange().getValues() : [];
        const prefsData = prefsSheet ? prefsSheet.getDataRange().getValues() : [];

        return processTechHubData(staffData, shiftsData, assignData, availData, prefsData, tz);

    } catch (e) {
        console.error("api_getTechHubViewData failed: " + e.stack);
        return { success: false, message: e.message };
    }
}

/**
 * Internal Processor
 */
function processTechHubData(staffData, shiftsData, assignData, availData, prefsData, timezone) {
    try {
        const normalize = (h) => String(h).toLowerCase().replace(/[\s_]/g, '');
        const getColMap = (row) => {
            const map = {};
            row.forEach((cell, i) => map[normalize(cell)] = i);
            return map;
        };

        const staffHeader = getColMap(staffData[0]);
        const techHubStaff = [];
        for(let i=1; i<staffData.length; i++) {
            const row = staffData[i];
            const roles = String(row[staffHeader.roles] || '').toLowerCase();
            const active = String(row[staffHeader.isactive] || 'true').toLowerCase();
            
            if (roles.includes('tech hub') && active !== 'false') {
                techHubStaff.push({
                    id: String(row[staffHeader.staffid]),
                    name: row[staffHeader.fullname] || row[staffHeader.name]
                });
            }
        }
        techHubStaff.sort((a,b) => a.name.localeCompare(b.name));

        const assignHeader = getColMap(assignData[0]);
        const assignmentsMap = {};
        for(let i=1; i<assignData.length; i++) {
            const row = assignData[i];
            if (row[assignHeader.assignmenttype] === 'Tech Hub') {
                assignmentsMap[String(row[assignHeader.referenceid])] = String(row[assignHeader.staffid]);
            }
        }

        const availMap = {}; 
        if (availData.length > 1) {
             const h = getColMap(availData[0]);
             for(let i=1; i<availData.length; i++) {
                 const sid = String(availData[i][h.staffid]).toLowerCase();
                 if(!availMap[sid]) availMap[sid] = [];
                 
                 const startMins = parseTimeContext(availData[i][h.starttime], timezone);
                 const endMins = parseTimeContext(availData[i][h.endtime], timezone);
                 
                 availMap[sid].push({
                     day: String(availData[i][h.dayofweek] || availData[i][h.day]),
                     startMins: startMins,
                     endMins: endMins
                 });
             }
        }

        const prefsMap = {}; 
        if (prefsData.length > 1) {
            const h = getColMap(prefsData[0]);
            const staffIdx = h.staffid !== undefined ? h.staffid : 0;
            const blockIdx = h.timeblock !== undefined ? h.timeblock : 1;
            const prefIdx = h.preference !== undefined ? h.preference : 2;

            for(let i=1; i<prefsData.length; i++) {
                const sid = String(prefsData[i][staffIdx]).toLowerCase();
                if(!prefsMap[sid]) prefsMap[sid] = {};
                prefsMap[sid][prefsData[i][blockIdx]] = prefsData[i][prefIdx];
            }
        }

        const roster = [];
        const manageShifts = [];
        
        if (shiftsData.length > 1) {
            const h = getColMap(shiftsData[0]);
            
            for(let i=1; i<shiftsData.length; i++) {
                const row = shiftsData[i];
                const shiftId = String(row[h.shiftid]);
                const day = String(row[h.dayofweek] || row[h.day]);
                
                const shiftStartMins = parseTimeContext(row[h.starttime], timezone);
                const shiftEndMins = parseTimeContext(row[h.endtime], timezone);

                let timeBlock = "Morning";
                if (shiftStartMins >= 720) timeBlock = "Afternoon";
                if (shiftStartMins >= 1020) timeBlock = "Evening";
                const prefKey = `${day}_${timeBlock}`;

                let startDisplay = row[h.starttime];
                let endDisplay = row[h.endtime];
                if (startDisplay instanceof Date) startDisplay = Utilities.formatDate(startDisplay, timezone, "h:mm a");
                if (endDisplay instanceof Date) endDisplay = Utilities.formatDate(endDisplay, timezone, "h:mm a");

                const smartList = techHubStaff.map(staff => {
                    const sid = staff.id.toLowerCase();
                    let isBlocked = false; 
                    let preference = "Neutral"; 
                    
                    if (availMap[sid]) {
                        isBlocked = availMap[sid].some(slot => {
                            if (slot.day !== day) return false;
                            return (slot.startMins < shiftEndMins && slot.endMins > shiftStartMins);
                        });
                    }

                    if (prefsMap[sid] && prefsMap[sid][prefKey]) {
                        preference = prefsMap[sid][prefKey];
                    }

                    return { 
                        id: staff.id, 
                        name: staff.name, 
                        isBlocked: isBlocked,
                        preference: preference
                    };
                });

                const prefScore = { "Yes Please": 3, "Eh, Sure": 2, "Neutral": 2, "No Thanks": 1 };
                smartList.sort((a, b) => {
                    if (a.isBlocked && !b.isBlocked) return 1;
                    if (!a.isBlocked && b.isBlocked) return -1;
                    const scoreA = prefScore[a.preference] || 2;
                    const scoreB = prefScore[b.preference] || 2;
                    return scoreB - scoreA;
                });

                roster.push({
                    shiftId: shiftId,
                    description: row[h.description],
                    day: day,
                    start: startDisplay,
                    end: endDisplay,
                    assignedStaffId: assignmentsMap[shiftId] || "",
                    smartList: smartList
                });

                manageShifts.push({
                    id: shiftId,
                    desc: row[h.description],
                    day: day,
                    start: startDisplay,
                    end: endDisplay,
                    zoom: String(row[h.zoom]).toLowerCase() === 'true'
                });
            }
        }

        return { success: true, data: { roster, manageShifts } };

    } catch (e) {
        return { success: false, message: "Processing Error: " + e.message };
    }
}

// --- CALENDAR SYNC LOGIC ---

function api_previewTechHubSync(targetCalendarId, semesterStartStr, semesterEndStr) {
    try {
        if (!targetCalendarId) throw new Error("No calendar selected.");
        if (!semesterStartStr || !semesterEndStr) throw new Error("Semester dates are missing.");

        const ss = getMasterDataHub();
        const tz = ss.getSpreadsheetTimeZone();
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        
        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target; 
        };

        const shiftsSheet = ss.getSheetByName(sheetTabs[findKey('TechHub_Shifts')]);
        const assignSheet = ss.getSheetByName(sheetTabs[findKey('Staff_Assignments')]);
        const staffSheet = ss.getSheetByName(sheetTabs[findKey('Staff_List')]);

        const shiftsData = shiftsSheet.getDataRange().getValues();
        const assignData = assignSheet.getDataRange().getValues();
        const staffData = staffSheet.getDataRange().getValues();

        const assignHeader = getColumnMap(assignData[0]);
        const assignmentsMap = {}; 
        
        const staffHeader = getColumnMap(staffData[0]);
        const staffInfoMap = {}; // ID -> { email, name }
        for(let i=1; i<staffData.length; i++) {
            const id = String(staffData[i][staffHeader.staffid]);
            staffInfoMap[id] = {
                email: staffData[i][staffHeader.email] || id,
                name: staffData[i][staffHeader.fullname] || staffData[i][staffHeader.name] || id
            };
        }

        for(let i=1; i<assignData.length; i++) {
            const row = assignData[i];
            if (row[assignHeader.assignmenttype] === 'Tech Hub') {
                const shiftId = String(row[assignHeader.referenceid]);
                const staffId = String(row[assignHeader.staffid]);
                if(staffInfoMap[staffId]) {
                    assignmentsMap[shiftId] = staffInfoMap[staffId];
                }
            }
        }

        const proposals = [];
        const shiftsHeader = getColumnMap(shiftsData[0]);
        
        const semesterStart = new Date(semesterStartStr);
        const semesterEnd = new Date(semesterEndStr);
        semesterEnd.setHours(23, 59, 59);

        if (isNaN(semesterStart.getTime()) || isNaN(semesterEnd.getTime())) {
            throw new Error("Invalid semester dates provided.");
        }

        if (shiftsData.length <= 1) return { success: true, data: [] };

        for(let i=1; i<shiftsData.length; i++) {
            const row = shiftsData[i];
            const shiftId = String(row[shiftsHeader.shiftid]);
            const dayStr = String(row[shiftsHeader.dayofweek] || row[shiftsHeader.day]);
            const desc = row[shiftsHeader.description];
            const isZoom = String(row[shiftsHeader.zoom]).toLowerCase() === 'true';
            
            const startTimeMins = parseTimeContext(row[shiftsHeader.starttime], tz);
            const endTimeMins = parseTimeContext(row[shiftsHeader.endtime], tz);
            
            const firstDate = getNextDayOccurrence(semesterStart, dayStr);
            
            if (firstDate > semesterEnd) continue;

            const startDt = new Date(firstDate);
            startDt.setHours(Math.floor(startTimeMins/60), startTimeMins%60, 0, 0);
            
            const endDt = new Date(firstDate);
            endDt.setHours(Math.floor(endTimeMins/60), endTimeMins%60, 0, 0);

            const assignedInfo = assignmentsMap[shiftId];
            
            const staffDisplayName = assignedInfo ? assignedInfo.name : "Unassigned";
            const title = `Tech Hub: ${desc} (${staffDisplayName})`;
            
            const description = `Shift: ${desc}\nStaff: ${staffDisplayName}\nEmail: ${assignedInfo ? assignedInfo.email : 'Unassigned'}\nZoom: ${isZoom ? 'Yes' : 'No'}`;
            const location = isZoom ? "https://maine.zoom.us/j/2076213123" : "Tech Hub";

            proposals.push({
                shiftId: shiftId,
                title: title,
                start: startDt.getTime(),
                end: endDt.getTime(),
                recurrenceEnd: semesterEnd.getTime(),
                assignedEmail: assignedInfo ? assignedInfo.email : null,
                description: description,
                location: location
            });
        }

        const cal = CalendarApp.getCalendarById(targetCalendarId);
        if(!cal) throw new Error("Calendar not found or permission denied.");

        const scanEnd = new Date(semesterStart);
        scanEnd.setDate(scanEnd.getDate() + 14);

        let existingEvents = [];
        try {
            existingEvents = cal.getEvents(semesterStart, scanEnd);
        } catch(e) {
            throw new Error("Failed to fetch calendar events. " + e.message);
        }

        // OPTIMIZATION: Memory Cache for Series Tags
        const seriesCache = {};
        existingEvents.forEach(e => {
            let tag = e.getTag('TechHub_ShiftID');
            if (!tag) {
                const baseId = (e.getId() || "").split('_')[0];
                if (seriesCache[baseId] !== undefined) {
                    tag = seriesCache[baseId];
                } else {
                    try {
                        const series = e.getEventSeries();
                        tag = series ? series.getTag('TechHub_ShiftID') : null;
                        seriesCache[baseId] = tag;
                    } catch(err) { seriesCache[baseId] = null; }
                }
            }
            e._cachedTag = tag;
        });

        const results = [];

        proposals.forEach(prop => {
            // Instantly check memory instead of calling Google Calendar
            const match = existingEvents.find(e => e._cachedTag === prop.shiftId);

            let status = "NEW";
            let diffs = [];

            if (match) {
                status = "SYNCED";
                if (match.getTitle() !== prop.title) {
                    status = "UPDATE";
                    diffs.push(`Title: ${match.getTitle()} -> ${prop.title}`);
                }
                const guests = match.getGuestList().map(g => g.getEmail());
                if (prop.assignedEmail && !guests.includes(prop.assignedEmail)) {
                    status = "UPDATE";
                    diffs.push(`Invite: ${prop.assignedEmail}`);
                }
                if ((match.getLocation() || "") !== prop.location) {
                    status = "UPDATE";
                    diffs.push(`Location: ${match.getLocation()} -> ${prop.location}`);
                }
            }

            results.push({
                shiftId: prop.shiftId,
                status: status,
                title: prop.title,
                diffs: diffs,
                payload: prop 
            });
        });

        return { success: true, data: results };

    } catch (e) { return { success: false, message: e.message }; }
}

function api_commitTechHubSync(targetCalendarId, eventsToSync) {
    try {
        // Security: Must be Tech Hub Lead or Admin
        requireRole(['Tech Hub', 'Admin']);

        const cal = CalendarApp.getCalendarById(targetCalendarId);
        const stats = { created: 0, updated: 0, errors: 0 };

        eventsToSync.forEach(item => {
            try {
                const p = item.payload;
                const startDt = new Date(p.start);
                const endDt = new Date(p.end);
                const recurEnd = new Date(p.recurrenceEnd);

                const recurrence = CalendarApp.newRecurrence().addWeeklyRule().until(recurEnd);
                
                const scanEnd = new Date(startDt);
                scanEnd.setDate(scanEnd.getDate() + 14);
                
                const existingEvents = cal.getEvents(startDt, scanEnd);
                const seriesCache = {};
                
                const existing = existingEvents.find(e => {
                    let tag = e.getTag('TechHub_ShiftID');
                    if (!tag) {
                        const baseId = (e.getId() || "").split('_')[0];
                        if (seriesCache[baseId] !== undefined) {
                            tag = seriesCache[baseId];
                        } else {
                            try {
                                const series = e.getEventSeries();
                                tag = series ? series.getTag('TechHub_ShiftID') : null;
                                seriesCache[baseId] = tag;
                            } catch(err) { seriesCache[baseId] = null; }
                        }
                    }
                    return tag === p.shiftId;
                });

                if (existing) {
                    try { existing.getEventSeries().deleteEventSeries(); } catch(e) { existing.deleteEvent(); }
                    stats.updated++;
                } else {
                    stats.created++;
                }

                // Sanitize before creating
                const series = cal.createEventSeries(sanitizeInput(p.title), startDt, endDt, recurrence, {
                    description: sanitizeInput(p.description),
                    location: sanitizeInput(p.location)
                });
                series.setTag('TechHub_ShiftID', p.shiftId);
                
                if (p.assignedEmail) {
                    series.addGuest(p.assignedEmail);
                }
                
                Utilities.sleep(500); 

            } catch (err) {
                console.error(err);
                stats.errors++;
            }
        });

        return { success: true, stats: stats };
    } catch (e) { return { success: false, message: e.message }; }
}

// --- HELPERS ---

function parseTimeContext(val, timezone) {
    if (val instanceof Date) {
        const timeStr = Utilities.formatDate(val, timezone, "HH:mm");
        const [h, m] = timeStr.split(':').map(Number);
        return (h * 60) + m;
    }
    if (typeof val === 'string') {
        const d = new Date(`1/1/2000 ${val}`);
        if (!isNaN(d.getTime())) {
             return (d.getHours() * 60) + d.getMinutes();
        }
    }
    return 0;
}


function getNextDayOccurrence(startDate, dayName) {
    const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    const targetIdx = days.indexOf(dayName.toLowerCase());
    const startIdx = startDate.getDay();
    
    let daysToAdd = targetIdx - startIdx;
    if (daysToAdd < 0) daysToAdd += 7;
    
    const nextDate = new Date(startDate);
    nextDate.setDate(startDate.getDate() + daysToAdd);
    return nextDate;
}

// --- STANDARD ACTIONS ---

function api_saveSingleTechHubAssignment(shiftId, staffId, startStr, endStr) {
    try {
        // Security: Must be Tech Hub Lead or Admin
        requireRole(['Tech Hub', 'Admin']);

        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target;
        };
        const sheet = getSheet(findKey('Staff_Assignments'));
        const data = sheet.getDataRange().getValues(); 
        
        const normalizeHeader = (h) => String(h).toLowerCase().replace(/[\s_]/g, '');
        const headers = data[0].map(normalizeHeader);
        
        const staffIdIndex = headers.indexOf('staffid');
        const typeIndex = headers.indexOf('assignmenttype');
        const refIdIndex = headers.indexOf('referenceid');
        const startDateIndex = headers.indexOf('startdate');
        const endDateIndex = headers.indexOf('enddate');
        
        if (staffIdIndex === -1) throw new Error("Header 'StaffID' not found.");

        let rowIndex = -1;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][refIdIndex]) === String(shiftId) && data[i][typeIndex] === 'Tech Hub') {
                rowIndex = i + 1; 
                break;
            }
        }

        if (rowIndex > -1) {
            if (staffId) {
                sheet.getRange(rowIndex, staffIdIndex + 1).setValue(sanitizeInput(staffId));
                sheet.getRange(rowIndex, startDateIndex + 1).setValue(new Date(startStr));
                sheet.getRange(rowIndex, endDateIndex + 1).setValue(new Date(endStr));
            } else {
                sheet.deleteRow(rowIndex);
            }
        } else {
            if (staffId) {
                const newRow = Array(headers.length).fill('');
                newRow[0] = 'A-' + Utilities.getUuid();
                newRow[staffIdIndex] = sanitizeInput(staffId);
                newRow[typeIndex] = 'Tech Hub';
                newRow[refIdIndex] = shiftId;
                newRow[startDateIndex] = new Date(startStr);
                newRow[endDateIndex] = new Date(endStr);
                sheet.appendRow(newRow);
            }
        }

        return { success: true, message: "Assignment saved." };
    } catch (e) { return { success: false, message: e.message }; }
}

function saveAllTechHubAssignments(assignmentList, startDate, endDate) {
    try {
        // Security: Must be Tech Hub Lead or Admin
        requireRole(['Tech Hub', 'Admin']);

        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target;
        };
        const sheet = getSheet(findKey('Staff_Assignments'));
        const data = sheet.getDataRange().getValues(); 
        
        const normalizeHeader = (h) => String(h).toLowerCase().replace(/[\s_]/g, '');
        const headers = data[0].map(normalizeHeader);
        
        const staffIdIndex = headers.indexOf('staffid');
        const typeIndex = headers.indexOf('assignmenttype');
        const refIdIndex = headers.indexOf('referenceid');
        const startDateIndex = headers.indexOf('startdate');
        const endDateIndex = headers.indexOf('enddate');
        
        const updatedData = [data[0]];
        const assignmentMap = new Map(assignmentList.map(a => [a.shiftId, a.staffId]));

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const assignmentType = row[typeIndex];
            const refId = row[refIdIndex];

            if (assignmentType !== 'Tech Hub' || !assignmentMap.has(refId)) {
                updatedData.push(row);
                continue;
            }

            const newStaffId = assignmentMap.get(refId);
            if (newStaffId) {
                row[staffIdIndex] = sanitizeInput(newStaffId);
                row[startDateIndex] = startDate;
                row[endDateIndex] = endDate;
                updatedData.push(row);
            }
            assignmentMap.delete(refId);
        }

        assignmentMap.forEach((staffId, shiftId) => {
            if(staffId) {
                const newRow = Array(headers.length).fill('');
                newRow[0] = 'A-' + Utilities.getUuid();
                newRow[staffIdIndex] = sanitizeInput(staffId);
                newRow[typeIndex] = 'Tech Hub';
                newRow[refIdIndex] = shiftId;
                newRow[startDateIndex] = startDate;
                newRow[endDateIndex] = endDate;
                updatedData.push(newRow);
            }
        });

        sheet.clearContents();
        sheet.getRange(1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);

        return { success: true, message: `Assignments updated.` };
    } catch (e) { return { success: false, message: e.message }; }
}

function addTechHubShift(shiftData) {
    try {
        // Security: Must be Tech Hub Lead or Admin
        requireRole(['Tech Hub', 'Admin']);

        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target;
        };
        const sheet = getSheet(findKey('TechHub_Shifts'));
        const days = Array.isArray(shiftData.days) ? shiftData.days : [shiftData.day];
        
        days.forEach(day => {
            const zoomVal = shiftData.zoom === true ? 'TRUE' : 'FALSE';
            sheet.appendRow([
                'SH-' + Utilities.getUuid(), 
                sanitizeInput(shiftData.description), 
                sanitizeInput(day), 
                sanitizeInput(shiftData.startTime), 
                sanitizeInput(shiftData.endTime), 
                zoomVal
            ]);
        });
        return { success: true }; 
    } catch (e) { return { success: false, message: e.message }; }
}

function deleteTechHubShift(shiftId) {
    try {
        // Security: Must be Tech Hub Lead or Admin
        requireRole(['Tech Hub', 'Admin']);

        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target;
        };
        const shiftsSheet = getSheet(findKey('TechHub_Shifts'));
        const assignmentsSheet = getSheet(findKey('Staff_Assignments'));

        const shiftsData = shiftsSheet.getDataRange().getDisplayValues();
        const assignmentsData = assignmentsSheet.getDataRange().getDisplayValues();
        const normalizeHeader = (h) => String(h).toLowerCase().replace(/[\s_]/g, '');

        const shiftsIdIndex = shiftsData[0].map(normalizeHeader).indexOf('shiftid');
        const assignmentsRefIdIndex = assignmentsData[0].map(normalizeHeader).indexOf('referenceid');

        const idsToDelete = new Set([shiftId]);
        
        const remainingShifts = shiftsData.filter((row, index) => {
            if (index === 0) return true;
            return !idsToDelete.has(row[shiftsIdIndex]);
        });

        const remainingAssignments = assignmentsData.filter((row, index) => {
            if (index === 0) return true;
            return !idsToDelete.has(row[assignmentsRefIdIndex]);
        });

        shiftsSheet.clearContents();
        shiftsSheet.getRange(1, 1, remainingShifts.length, remainingShifts[0].length).setValues(remainingShifts);
        assignmentsSheet.clearContents();
        assignmentsSheet.getRange(1, 1, remainingAssignments.length, remainingAssignments[0].length).setValues(remainingAssignments);

        return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
}