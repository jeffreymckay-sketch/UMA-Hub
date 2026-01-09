/**
 * -------------------------------------------------------------------
 * CONTROLLER: TECH HUB SCHEDULING
 * Handles Shifts, Roster, Availability, and Calendar Sync
 * -------------------------------------------------------------------
 */

// --- VIEW MODEL GENERATOR ---

function getTechHubViewData(master) {
    const roster = [];
    const manageShifts = [];
    
    try {
        const assignmentsMap = createLookupMap(master.assignmentsData, 'ReferenceID', 'StaffID'); 
        const allStaffObjects = Object.values(master.staffMap);
        
        const techHubStaffList = allStaffObjects
            .filter(s => (s.Roles || '').toLowerCase().includes('tech hub') && s.IsActive !== 'FALSE')
            .map(s => ({ id: s.StaffID, name: s.FullName }))
            .sort((a, b) => a.name.localeCompare(b.name));

        // 1. Build Roster (Tech Hub)
        if (master.shiftsData.length > 1 && master.shiftsData[0]) {
            const h = master.shiftsData[0].map(normalizeHeader);
            const idx = { id: h.indexOf('shiftid'), desc: h.indexOf('description'), day: h.indexOf('dayofweek'), start: h.indexOf('starttime'), end: h.indexOf('endtime') };
            if (idx.id > -1) {
                for (let i = 1; i < master.shiftsData.length; i++) {
                    const row = master.shiftsData[i];
                    const tbKey = `${row[idx.day]}_${getTimeBlock(row[idx.start])}`; 
                    const smartList = [];
                    for (const staff of techHubStaffList) {
                        const sid = (staff.id || '').toLowerCase();
                        let label = staff.name;
                        let avail = true;
                        if (master.prefsMap[sid] && master.prefsMap[sid][tbKey]) label += ` (Prefers: ${master.prefsMap[sid][tbKey]})`;
                        if (master.availMap[sid]) {
                            for (const b of master.availMap[sid]) {
                                if (normalizeHeader(b.day) === normalizeHeader(row[idx.day]) && timesOverlap(row[idx.start], row[idx.end], b.start, b.end)) {
                                    label += ` (NOT AVAILABLE)`; avail = false; break;
                                }
                            }
                        }
                        smartList.push({ id: staff.id, name: label, available: avail });
                    }
                    roster.push({ shiftId: row[idx.id], description: row[idx.desc], day: row[idx.day], start: row[idx.start], end: row[idx.end], assignedStaffId: assignmentsMap[row[idx.id]] || "", smartList: smartList });
                }
            }
        }
        
        // 2. Manage Shifts List
        if (master.shiftsData.length > 1 && master.shiftsData[0]) {
             const h = master.shiftsData[0].map(normalizeHeader);
             const idx = { 
                 id: h.indexOf('shiftid'), 
                 desc: h.indexOf('description'), 
                 day: h.indexOf('dayofweek'), 
                 start: h.indexOf('starttime'), 
                 end: h.indexOf('endtime'),
                 zoom: h.indexOf('zoom')
             };
             if (idx.id > -1) {
                for(let i=1; i<master.shiftsData.length; i++) {
                    const r = master.shiftsData[i];
                    const hasZoom = (idx.zoom > -1 && r[idx.zoom].toString().toLowerCase() === 'true');
                    manageShifts.push({ 
                        id: r[idx.id], 
                        desc: r[idx.desc], 
                        day: r[idx.day], 
                        start: r[idx.start], 
                        end: r[idx.end],
                        zoom: hasZoom
                    });
                }
            }
        }
    } catch (e) {
        // Log error but return empty arrays to prevent crash
        console.error("Tech Hub View Error: " + e.message);
    }

    return { roster, manageShifts };
}

// --- ACTIONS ---

function saveAllTechHubAssignments(assignmentList, startDate, endDate) {
    try {
        const sheet = getSheet('Staff_Assignments');
        const data = sheet.getDataRange().getValues(); 
        const headers = data[0].map(normalizeHeader);
        
        const staffIdIndex = headers.indexOf('staffid');
        const typeIndex = headers.indexOf('assignmenttype');
        const refIdIndex = headers.indexOf('referenceid');
        const startDateIndex = headers.indexOf('startdate');
        const endDateIndex = headers.indexOf('enddate');
        
        if (staffIdIndex === -1) throw new Error("Header 'StaffID' not found in Staff_Assignments.");
        if (typeIndex === -1) throw new Error("Header 'AssignmentType' not found in Staff_Assignments.");
        if (refIdIndex === -1) throw new Error("Header 'ReferenceID' not found in Staff_Assignments.");
        if (startDateIndex === -1) throw new Error("Header 'StartDate' not found in Staff_Assignments.");
        if (endDateIndex === -1) throw new Error("Header 'EndDate' not found in Staff_Assignments.");
        
        const shiftsToUpdate = new Set(assignmentList.map(a => a.shiftId));
        const newAssignmentsMap = assignmentList.reduce((map, item) => { map[item.shiftId] = item.staffId; return map; }, {});
        const rowsToDelete = [];
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][typeIndex] === 'Tech Hub' && shiftsToUpdate.has(data[i][refIdIndex])) {
                const newStaffId = newAssignmentsMap[data[i][refIdIndex]];
                if (newStaffId) {
                    if (data[i][staffIdIndex] !== newStaffId || data[i][startDateIndex] !== startDate) {
                        sheet.getRange(i + 1, staffIdIndex + 1).setValue(newStaffId);
                        sheet.getRange(i + 1, startDateIndex + 1).setValue(startDate);
                        sheet.getRange(i + 1, endDateIndex + 1).setValue(endDate);
                    }
                    shiftsToUpdate.delete(data[i][refIdIndex]); 
                } else {
                    rowsToDelete.push(i + 1);
                    shiftsToUpdate.delete(data[i][refIdIndex]);
                }
            }
        }
        rowsToDelete.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
        const rowsToAdd = [];
        assignmentList.forEach(a => {
            if (shiftsToUpdate.has(a.shiftId) && a.staffId) rowsToAdd.push(['A-' + Utilities.getUuid(), a.staffId, 'Tech Hub', a.shiftId, startDate, endDate]);
        });
        if (rowsToAdd.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        return { success: true, message: `Assignments updated.` };
    } catch (e) { return { success: false, message: e.message }; }
}

function addTechHubShift(shiftData) {
    try {
        const sheet = getSheet('TechHub_Shifts');
        
        const days = Array.isArray(shiftData.days) ? shiftData.days : [shiftData.day];
        
        days.forEach(day => {
            const zoomVal = shiftData.zoom === true ? 'TRUE' : 'FALSE';
            sheet.appendRow(['SH-' + Utilities.getUuid(), shiftData.description, day, shiftData.startTime, shiftData.endTime, zoomVal]);
        });

        return getSchedulingRosterData_refactored();
    } catch (e) { return { error: e.message }; }
}

function deleteBulkTechHubShifts(shiftIds) {
    try {
        const sheet = getSheet('TechHub_Shifts');
        const data = sheet.getDataRange().getDisplayValues();
        const idIndex = data[0].map(normalizeHeader).indexOf('shiftid');
        
        if (idIndex === -1) throw new Error("ShiftID column not found.");
        
        const idsToDelete = new Set(shiftIds);
        const rowsToDelete = [];
        
        for (let i = data.length - 1; i >= 1; i--) { 
            if (idsToDelete.has(data[i][idIndex])) {
                rowsToDelete.push(i + 1);
                deleteAssignmentsByReferenceId(data[i][idIndex]);
            }
        }
        
        rowsToDelete.forEach(r => sheet.deleteRow(r));
        
        return getSchedulingRosterData_refactored();
    } catch (e) { return { error: e.message }; }
}

function deleteTechHubShift(shiftId) {
    return deleteBulkTechHubShifts([shiftId]);
}

function deleteAssignmentsByReferenceId(refId) {
    try {
        const sheet = getSheet('Staff_Assignments');
        if (!sheet) return;
        const data = sheet.getDataRange().getDisplayValues();
        const refIdIndex = data[0].map(normalizeHeader).indexOf('referenceid');
        for (let i = data.length - 1; i >= 1; i--) { 
            if (data[i][refIdIndex] === refId) sheet.deleteRow(i + 1);
        }
    } catch (e) {}
}

function deleteAssignment(assignmentId) {
    try {
        const sheet = getSheet('Staff_Assignments');
        const data = sheet.getDataRange().getDisplayValues();
        const idIndex = data[0].map(normalizeHeader).indexOf('assignmentid');
        for (let i = data.length - 1; i >= 1; i--) { 
            if (data[i][idIndex] === assignmentId) {
                sheet.deleteRow(i + 1);
                return getSchedulingRosterData_refactored();
            }
        }
        throw new Error('ID not found.');
    } catch (e) { return { error: e.message }; }
}

function resetAllAssignments() {
    try {
        const sheet = getSheet('Staff_Assignments');
        if (!sheet) return { success: true };
        const data = sheet.getDataRange().getValues();
        const typeIndex = data[0].map(normalizeHeader).indexOf('assignmenttype');
        const rowsToDelete = [];
        for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][typeIndex] === 'Tech Hub') rowsToDelete.push(i + 1);
        }
        rowsToDelete.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
        return { success: true, message: 'Reset complete.' };
    } catch (e) { return { success: false, message: e.message }; }
}

function handleClearAllShifts() {
    try {
        const sheet = getSheet('TechHub_Shifts');
        if (sheet) {
            const lastRow = sheet.getLastRow();
            if (lastRow > 1) {
                sheet.deleteRows(2, lastRow - 1);
            }
        }
        resetAllAssignments();
        return getSchedulingRosterData_refactored();
    } catch (e) { return { error: e.message }; }
}

function validateUserEditPermission(targetEmail) {
    const currentUser = Session.getActiveUser().getEmail().toLowerCase();
    const target = (targetEmail || currentUser).toLowerCase();
    if (target === currentUser) return target; 
    const access = api_getAccessControlData(); 
    if (access.userRole === 'Admin' || access.userRole === 'Lead') return target;
    throw new Error("Permission denied: You can only edit your own data.");
}

function api_getMyAvailability(targetEmail) {
    try {
        const email = validateUserEditPermission(targetEmail);
        const sheet = getSheet('Staff_Availability');
        const data = sheet.getDataRange().getDisplayValues();
        const list = [];
        if (data.length > 1) {
            const h = data[0].map(normalizeHeader);
            const idIdx = h.indexOf('availabilityid'), staffIdx = h.indexOf('staffid'), dayIdx = h.indexOf('dayofweek'), startIdx = h.indexOf('starttime'), endIdx = h.indexOf('endtime');
            for (let i = 1; i < data.length; i++) {
                if (normalizeHeader(data[i][staffIdx]) === email) list.push({ id: data[i][idIdx], day: data[i][dayIdx], start: data[i][startIdx], end: data[i][endIdx] });
            }
        }
        return { success: true, data: list };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_addNotAvailable(day, start, end, targetEmail) {
    try {
        const email = validateUserEditPermission(targetEmail);
        const sheet = getSheet('Staff_Availability');
        sheet.appendRow(['AV-' + Utilities.getUuid(), email, day, start, end]);
        return api_getMyAvailability(email);
    } catch (e) { return { success: false, message: e.message }; }
}

function api_deleteAvailability(id, targetEmail) {
    try {
        const email = validateUserEditPermission(targetEmail);
        const sheet = getSheet('Staff_Availability');
        const data = sheet.getDataRange().getDisplayValues();
        const idIdx = data[0].map(normalizeHeader).indexOf('availabilityid');
        for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][idIdx] === id) { sheet.deleteRow(i + 1); return api_getMyAvailability(email); }
        }
        return { success: false, message: 'Item not found.' };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_getMyPreferences(targetEmail) {
    try {
        const email = validateUserEditPermission(targetEmail);
        const sheet = getSheet('Staff_Preferences');
        const data = sheet.getDataRange().getDisplayValues();
        const prefs = {};
        if (data.length > 1) {
            const h = data[0].map(normalizeHeader);
            const staffIdx = h.indexOf('staffid'), blockIdx = h.indexOf('timeblock'), prefIdx = h.indexOf('preference');
            for (let i = 1; i < data.length; i++) {
                if (normalizeHeader(data[i][staffIdx]) === email) prefs[data[i][blockIdx]] = data[i][prefIdx];
            }
        }
        return { success: true, data: prefs };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_savePreference(timeBlock, value, targetEmail) {
    try {
        const email = validateUserEditPermission(targetEmail);
        const sheet = getSheet('Staff_Preferences');
        const data = sheet.getDataRange().getValues();
        const h = data[0].map(normalizeHeader);
        const staffIdx = h.indexOf('staffid'), blockIdx = h.indexOf('timeblock'), prefIdx = h.indexOf('preference');
        let found = false;
        for (let i = 1; i < data.length; i++) {
            if (normalizeHeader(data[i][staffIdx]) === email && data[i][blockIdx] === timeBlock) {
                sheet.getRange(i + 1, prefIdx + 1).setValue(value); found = true; break;
            }
        }
        if (!found) sheet.appendRow([email, timeBlock, value]);
        return { success: true, message: 'Saved' };
    } catch (e) { return { success: false, message: e.message }; }
}
