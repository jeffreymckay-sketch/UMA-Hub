/**
 * -------------------------------------------------------------------
 * CONTROLLER: TECH HUB SCHEDULING
 * Handles Shifts, Roster, Availability, and Calendar Sync
 * -------------------------------------------------------------------
 */

// --- CLIENT-CALLABLE API FUNCTIONS ---

/**
 * API endpoint for the client to fetch all data required for the Tech Hub view.
 * This function acts as a bridge, calling the main data model and then the specific view generator.
 * @returns {object} A success/fail object with the generated view data.
 */
function api_getTechHubViewData() {
    try {
        // 1. Get the master data model & settings
        const masterData = getMasterDataModel();
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');

        // 2. Pass the data to the view generator
        return getTechHubViewData(masterData, sheetTabs);

    } catch (e) {
        // Log the error for debugging and return a user-friendly message
        console.error("api_getTechHubViewData failed: " + e.stack);
        return { success: false, message: "An error occurred while loading the Tech Hub data. Details: " + e.message };
    }
}


// --- VIEW MODEL GENERATOR ---

function getTechHubViewData(master, sheetTabs) {
    try {
        const roster = [];
        const manageShifts = [];
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
        
        return { success: true, data: { roster, manageShifts } };

    } catch (e) {
        console.error("Tech Hub View Error: " + e.message);
        return { success: false, message: "Tech Hub View Error: " + e.message };
    }
}

// --- ACTIONS ---

function saveAllTechHubAssignments(assignmentList, startDate, endDate) {
    try {
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const sheet = getSheet(sheetTabs.Staff_Assignments);
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
        const sheet = getSheet(sheetTabs.TechHub_Shifts);
        
        const days = Array.isArray(shiftData.days) ? shiftData.days : [shiftData.day];
        
        days.forEach(day => {
            const zoomVal = shiftData.zoom === true ? 'TRUE' : 'FALSE';
            sheet.appendRow(['SH-' + Utilities.getUuid(), shiftData.description, day, shiftData.startTime, shiftData.endTime, zoomVal]);
        });

        return { success: true }; // Let the client re-fetch data
    } catch (e) { return { success: false, message: e.message }; }
}

function handleClearAllShifts() {
    try {
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const shiftsSheet = getSheet(sheetTabs.TechHub_Shifts);
        const assignmentsSheet = getSheet(sheetTabs.Staff_Assignments);

        // Clear shifts but keep headers
        if (shiftsSheet) {
             const headers = shiftsSheet.getRange(1, 1, 1, shiftsSheet.getLastColumn()).getValues();
            shiftsSheet.clearContents();
            shiftsSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        }

        // Remove only 'Tech Hub' assignments
        if (assignmentsSheet) {
            const data = assignmentsSheet.getDataRange().getValues();
            const typeIndex = data[0].map(normalizeHeader).indexOf('assignmenttype');
            const remainingData = data.filter((row, index) => {
                if (index === 0) return true; // Keep header
                return row[typeIndex] !== 'Tech Hub';
            });
            assignmentsSheet.clearContents();
            if(remainingData.length > 0) {
                assignmentsSheet.getRange(1, 1, remainingData.length, remainingData[0].length).setValues(remainingData);
            }
        }

        return { success: true };
    } catch (e) { return { success: false, message: e.message }; }
}

function deleteTechHubShift(shiftId) {
    try {
        const sheetTabs = JSON.parse(getSettings().sheetTabs || '{}');
        const shiftsSheet = getSheet(sheetTabs.TechHub_Shifts);
        const assignmentsSheet = getSheet(sheetTabs.Staff_Assignments);

        const shiftsData = shiftsSheet.getDataRange().getDisplayValues();
        const assignmentsData = assignmentsSheet.getDataRange().getDisplayValues();

        const shiftsIdIndex = shiftsData[0].map(normalizeHeader).indexOf('shiftid');
        if (shiftsIdIndex === -1) throw new Error("ShiftID column not found in TechHub_Shifts.");

        const assignmentsRefIdIndex = assignmentsData[0].map(normalizeHeader).indexOf('referenceid');
        if (assignmentsRefIdIndex === -1) throw new Error("ReferenceID column not found in Staff_Assignments.");

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