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

        // Note: Calendars are now handled by the Frontend Global Cache (g.writableCalendars)

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

        // 1. Parse Staff
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

        // 2. Parse Assignments Map
        const assignHeader = getColMap(assignData[0]);
        const assignmentsMap = {};
        for(let i=1; i<assignData.length; i++) {
            const row = assignData[i];
            if (row[assignHeader.assignmenttype] === 'Tech Hub') {
                assignmentsMap[String(row[assignHeader.referenceid])] = String(row[assignHeader.staffid]);
            }
        }

        // 3. Parse Availability (Blackout Times)
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

        // 4. Parse Preferences
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

        // 5. Build Roster
        const roster = [];
        const manageShifts = [];
        
        if (shiftsData.length > 1) {
            const h = getColMap(shiftsData[0]);
            
            for(let i=1; i<shiftsData.length; i++) {
                const row = shiftsData[i];
                const shiftId = String(row[h.shiftid]);
                const day = String(row[h.dayofweek] || row[h.day]);
                
                // Parse Shift Times
                const shiftStartMins = parseTimeContext(row[h.starttime], timezone);
                const shiftEndMins = parseTimeContext(row[h.endtime], timezone);

                // Determine Time Block for Preferences
                let timeBlock = "Morning";
                if (shiftStartMins >= 720) timeBlock = "Afternoon";
                if (shiftStartMins >= 1020) timeBlock = "Evening";
                const prefKey = `${day}_${timeBlock}`;

                // Format for Display
                let startDisplay = row[h.starttime];
                let endDisplay = row[h.endtime];
                if (startDisplay instanceof Date) startDisplay = Utilities.formatDate(startDisplay, timezone, "h:mm a");
                if (endDisplay instanceof Date) endDisplay = Utilities.formatDate(endDisplay, timezone, "h:mm a");

                // Build Smart List
                const smartList = techHubStaff.map(staff => {
                    const sid = staff.id.toLowerCase();
                    let isBlocked = false; 
                    let preference = "Neutral"; 
                    
                    // Check Availability (Blackouts)
                    if (availMap[sid]) {
                        isBlocked = availMap[sid].some(slot => {
                            if (slot.day !== day) return false;
                            // Overlap Logic
                            return (slot.startMins < shiftEndMins && slot.endMins > shiftStartMins);
                        });
                    }

                    // Check Preferences
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

                // Sort Smart List
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

        // 1. Map Assignments
        const assignHeader = createHeaderMap(assignData[0]);
        const assignmentsMap = {}; 
        
        const staffHeader = createHeaderMap(staffData[0]);
        const staffEmailMap = {};
        for(let i=1; i<staffData.length; i++) {
            staffEmailMap[String(staffData[i][staffHeader.staffid])] = staffData[i][staffHeader.email] || staffData[i][staffHeader.staffid];
        }

        for(let i=1; i<assignData.length; i++) {
            const row = assignData[i];
            if (row[assignHeader.assignmenttype] === 'Tech Hub') {
                const shiftId = String(row[assignHeader.referenceid]);
                const staffId = String(row[assignHeader.staffid]);
                if(staffEmailMap[staffId]) {
                    assignmentsMap[shiftId] = staffEmailMap[staffId];
                }
            }
        }

        // 2. Process Shifts into Proposed Events
        const proposals = [];
        const shiftsHeader = createHeaderMap(shiftsData[0]);
        
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

            const assignedEmail = assignmentsMap[shiftId];
            const title = `Tech Hub: ${desc}` + (assignedEmail ? ` (${assignedEmail.split('@')[0]})` : " (Unassigned)");
            const description = `Shift: ${desc}\nStaff: ${assignedEmail || 'Unassigned'}\nZoom: ${isZoom ? 'Yes' : 'No'}`;
            
            // FIX: Hardcoded Zoom Link Logic
            const location = isZoom ? "https://maine.zoom.us/j/2076213123" : "Tech Hub";

            proposals.push({
                shiftId: shiftId,
                title: title,
                start: startDt.getTime(),
                end: endDt.getTime(),
                recurrenceEnd: semesterEnd.getTime(),
                assignedEmail: assignedEmail,
                description: description,
                location: location
            });
        }

        // 3. Compare with Calendar (OPTIMIZED)
        const cal = CalendarApp.getCalendarById(targetCalendarId);
        if(!cal) throw new Error("Calendar not found or permission denied.");

        // OPTIMIZATION: Only fetch the first 14 days of the semester.
        const scanEnd = new Date(semesterStart);
        scanEnd.setDate(scanEnd.getDate() + 14);

        let existingEvents = [];
        try {
            existingEvents = cal.getEvents(semesterStart, scanEnd);
        } catch(e) {
            throw new Error("Failed to fetch calendar events. " + e.message);
        }

        const results = [];

        proposals.forEach(prop => {
            const match = existingEvents.find(e => {
                let tag = e.getTag('TechHub_ShiftID');
                if(!tag) { try { tag = e.getEventSeries().getTag('TechHub_ShiftID'); } catch(err){} }
                return tag === prop.shiftId;
            });

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
                // Check location change
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
        const cal = CalendarApp.getCalendarById(targetCalendarId);
        const stats = { created: 0, updated: 0, errors: 0 };

        eventsToSync.forEach(item => {
            try {
                const p = item.payload;
                const startDt = new Date(p.start);
                const endDt = new Date(p.end);
                const recurEnd = new Date(p.recurrenceEnd);

                const recurrence = CalendarApp.newRecurrence().addWeeklyRule().until(recurEnd);
                
                // Find existing to update/delete
                const scanEnd = new Date(startDt);
                scanEnd.setDate(scanEnd.getDate() + 14);
                
                const existing = cal.getEvents(startDt, scanEnd).find(e => {
                    let tag = e.getTag('TechHub_ShiftID');
                    if(!tag) { try { tag = e.getEventSeries().getTag('TechHub_ShiftID'); } catch(err){} }
                    return tag === p.shiftId;
                });

                if (existing) {
                    try { existing.getEventSeries().deleteEventSeries(); } catch(e) { existing.deleteEvent(); }
                    stats.updated++;
                } else {
                    stats.created++;
                }

                const series = cal.createEventSeries(p.title, startDt, endDt, recurrence, {
                    description: p.description,
                    location: p.location
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

function createHeaderMap(row) {
    const map = {};
    row.forEach((cell, i) => map[String(cell).toLowerCase().replace(/[\s_]/g, '')] = i);
    return map;
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
                sheet.getRange(rowIndex, staffIdIndex + 1).setValue(staffId);
                sheet.getRange(rowIndex, startDateIndex + 1).setValue(new Date(startStr));
                sheet.getRange(rowIndex, endDateIndex + 1).setValue(new Date(endStr));
            } else {
                sheet.deleteRow(rowIndex);
            }
        } else {
            if (staffId) {
                const newRow = Array(headers.length).fill('');
                newRow[0] = 'A-' + Utilities.getUuid();
                newRow[staffIdIndex] = staffId;
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
                row[staffIdIndex] = newStaffId;
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
                newRow[staffIdIndex] = staffId;
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
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const findKey = (target) => {
            const match = Object.keys(sheetTabs).find(k => k.toLowerCase() === target.toLowerCase());
            return match || target;
        };
        const sheet = getSheet(findKey('TechHub_Shifts'));
        const days = Array.isArray(shiftData.days) ? shiftData.days : [shiftData.day];
        days.forEach(day => {
            const zoomVal = shiftData.zoom === true ? 'TRUE' : 'FALSE';
            sheet.appendRow(['SH-' + Utilities.getUuid(), shiftData.description, day, shiftData.startTime, shiftData.endTime, zoomVal]);
        });
        return { success: true }; 
    } catch (e) { return { success: false, message: e.message }; }
}

function deleteTechHubShift(shiftId) {
    try {
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