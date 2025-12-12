/**
 * -------------------------------------------------------------------
 * CONTROLLER: SCHEDULING (AGGREGATOR)
 * Main Entry Point for Scheduling UI
 * -------------------------------------------------------------------
 */

function loadAllSchedulingData(ss) {
    const dataHubData = {};
    
    // Helper to safely get data even if tab name varies
    const getSafeData = (key, defaultName) => {
        let sheet = ss.getSheetByName(defaultName);
        if (!sheet) sheet = ss.getSheetByName(defaultName.replace('_', ' '));
        if (!sheet) return [];
        const vals = sheet.getDataRange().getDisplayValues();
        return (vals.length > 0) ? vals : [];
    };

    dataHubData[CONFIG.TABS.STAFF_LIST] = getSafeData(CONFIG.TABS.STAFF_LIST, 'Staff_List');
    dataHubData[CONFIG.TABS.TECH_HUB_SHIFTS] = getSafeData(CONFIG.TABS.TECH_HUB_SHIFTS, 'TechHub_Shifts');
    dataHubData[CONFIG.TABS.STAFF_ASSIGNMENTS] = getSafeData(CONFIG.TABS.STAFF_ASSIGNMENTS, 'Staff_Assignments');
    dataHubData[CONFIG.TABS.STAFF_AVAILABILITY] = getSafeData(CONFIG.TABS.STAFF_AVAILABILITY, 'Staff_Availability');
    dataHubData[CONFIG.TABS.STAFF_PREFERENCES] = getSafeData(CONFIG.TABS.STAFF_PREFERENCES, 'Staff_Preferences');
    dataHubData[CONFIG.TABS.COURSE_SCHEDULE] = getSafeData(CONFIG.TABS.COURSE_SCHEDULE, 'Course_Schedule');

    const output = {
        staffData: dataHubData[CONFIG.TABS.STAFF_LIST],
        shiftsData: dataHubData[CONFIG.TABS.TECH_HUB_SHIFTS],
        courseData: dataHubData[CONFIG.TABS.COURSE_SCHEDULE],
        assignmentsData: dataHubData[CONFIG.TABS.STAFF_ASSIGNMENTS],
        staffMap: {}, availMap: {}, prefsMap: {}
    };

    if (output.staffData.length > 1) output.staffMap = createDataMap(output.staffData, 'StaffID');

    // Availability Map
    const availData = dataHubData[CONFIG.TABS.STAFF_AVAILABILITY];
    if (availData.length > 1 && availData[0]) {
        const h = availData[0].map(normalizeHeader);
        const sIdx = h.indexOf('staffid');
        const dIdx = h.indexOf('dayofweek');
        const stIdx = h.indexOf('starttime');
        const eIdx = h.indexOf('endtime');
        if (sIdx > -1) {
            for (let i = 1; i < availData.length; i++) {
                const sid = (availData[i][sIdx] || '').toLowerCase();
                if (!output.availMap[sid]) output.availMap[sid] = [];
                output.availMap[sid].push({ day: availData[i][dIdx], start: availData[i][stIdx], end: availData[i][eIdx] });
            }
        }
    }

    // Preferences Map
    const prefData = dataHubData[CONFIG.TABS.STAFF_PREFERENCES];
    if (prefData.length > 1 && prefData[0]) {
        const h = prefData[0].map(normalizeHeader);
        const sIdx = h.indexOf('staffid');
        const bIdx = h.indexOf('timeblock');
        const pIdx = h.indexOf('preference');
        if (sIdx > -1) {
            for (let i = 1; i < prefData.length; i++) {
                const sid = (prefData[i][sIdx] || '').toLowerCase();
                if (!output.prefsMap[sid]) output.prefsMap[sid] = {};
                output.prefsMap[sid][prefData[i][bIdx]] = prefData[i][pIdx];
            }
        }
    }
    return output;
}

function getSchedulingRosterData() {
    try {
        const ss = getMasterDataHub();
        const master = loadAllSchedulingData(ss);

        // 1. Get Tech Hub Data
        const techHubView = getTechHubViewData(master);

        // 2. Get MST Data
        const mstView = getMSTViewData(master);

        // 3. Combine and Return
        return { 
            data: { 
                roster: techHubView.roster, 
                mstStaffList: mstView.mstStaffList, 
                manageShifts: techHubView.manageShifts, 
                courseItems: mstView.courseItems, 
                courseAssignments: mstView.courseAssignments,
                debug: mstView.debugMsg 
            } 
        };
    } catch (e) { return { error: e.message }; }
}