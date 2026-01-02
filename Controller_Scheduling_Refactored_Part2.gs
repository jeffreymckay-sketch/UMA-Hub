/**
 * -------------------------------------------------------------------
 * NEW REFACTORED SCHEDULING CONTROLLER - PART 2
 * -------------------------------------------------------------------
 */

// --- UTILITY ---

function getColumnMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const normalizedHeader = String(header).toLowerCase().replace(/[\s_]/g, '');
    map[normalizedHeader] = index;
  });
  return map;
}

// --- REFACTORED IMPORT/SYNC FUNCTIONS ---

/**
 * Imports Zoom Links from an external sheet using a more robust method.
 * Matches courses by Course, Day, and Time, then fuzzy matches by instructor.
 */
function importCourseLinks_refactored(sheetUrl, tabName) {
  const debugLog = [];
  try {
    if (!sheetUrl || !tabName) throw new Error("Missing Sheet URL or Tab Name.");

    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    if (!localSheet) throw new Error("Local Course Schedule sheet not found.");

    // 1. Prepare Local Data
    const localData = localSheet.getDataRange().getValues();
    const { headerIdx: localHeaderIdx, headers: lHeaders } = _findHeaderRow(localData, ['course', 'faculty']);
    if (localHeaderIdx === -1) throw new Error("Local header row not found.");
    debugLog.push(`Local Headers Found: ${lHeaders.join(', ')}`);
    const lMap = getColumnMap(lHeaders);

    // 2. Fetch External Data
    const sourceSS = SpreadsheetApp.openByUrl(sheetUrl);
    const sourceSheet = sourceSS.getSheetByName(tabName);
    if (!sourceSheet) throw new Error(`External tab "${tabName}" not found.`);
    
    const sourceValues = sourceSheet.getDataRange().getValues();
    const sourceRichText = sourceSheet.getDataRange().getRichTextValues();
    const { headerIdx: sourceHeaderIdx, headers: sHeaders } = _findHeaderRow(sourceValues, ['coursenumber', 'instructorname']);
    if (sourceHeaderIdx === -1) throw new Error("External header row not found.");
    debugLog.push(`External Headers Found: ${sHeaders.join(', ')}`);
    const sMap = getColumnMap(sHeaders);
    
    if (sMap.zoomlink === undefined) throw new Error("'Zoom Link' column not found in source data.");

    // 3. Build Source Map (Bucket by Course|Day|Time)
    const sourceMap = _buildSourceLinkMap(sourceValues, sourceRichText, sourceHeaderIdx, sMap);
    debugLog.push(`External Map Size: ${sourceMap.size}`);

    // 4. Match and Prepare Updates
    const { updates, outputColumn } = _matchAndPrepareUpdates(localData, localHeaderIdx, lMap, sourceMap);
    debugLog.push(`Found ${updates} links to update.`);

    // 5. Write Back if updates are found
    if (updates > 0) {
        const zoomLinkCol = (lMap.zoomlink || lHeaders.length) + 1;
        if (lMap.zoomlink === undefined) {
            localSheet.getRange(localHeaderIdx + 1, zoomLinkCol).setValue("Zoom Link");
            debugLog.push("Created new 'Zoom Link' column.");
        }
        localSheet.getRange(localHeaderIdx + 2, zoomLinkCol, outputColumn.length, 1).setValues(outputColumn);
    }

    // 6. Save settings
    PropertiesService.getScriptProperties().setProperty('mst_links_settings', JSON.stringify({ url: sheetUrl, tab: tabName }));

    return { success: true, message: `Import Complete. Matched: ${updates} links.` };

  } catch (e) {
    return { success: false, message: e.message + "\n\nLog:\n" + debugLog.join('\n') };
  }
}

function _findHeaderRow(data, keywords) {
    for (let r = 0; r < Math.min(data.length, 10); r++) {
        const rowStr = data[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (keywords.every(k => rowStr.includes(k))) {
            return { headerIdx: r, headers: data[r] };
        }
    }
    return { headerIdx: -1, headers: [] };
}

function _buildSourceLinkMap(values, richText, headerIdx, sMap) {
    const sourceMap = new Map();
    let lastCourse = "", lastInstructor = "";

    for (let i = headerIdx + 1; i < values.length; i++) {
        const row = values[i];
        const richRow = richText[i];

        let courseVal = row[sMap.coursenumber] || lastCourse;
        let instrVal = row[sMap.instructorname] || lastInstructor;
        lastCourse = courseVal; lastInstructor = instrVal;

        if (!courseVal) continue;

        const c = String(courseVal).toLowerCase().replace(/[^a-z0-9]/g, '');
        const d = sched_normalizeDay(row[sMap.day]);
        const t = sched_normalizeTime(row[sMap.starttime], null);
        const bucketKey = `${c}|${d}|${t}`;
        
        let link = String(row[sMap.zoomlink] || '').trim();
        if (!link.startsWith('http')) {
            link = richRow[sMap.zoomlink].getLinkUrl();
        }

        if (link && link.startsWith('http')) {
            if (!sourceMap.has(bucketKey)) sourceMap.set(bucketKey, []);
            sourceMap.get(bucketKey).push({
                instructor: String(instrVal).toLowerCase().replace(/[^a-z]/g, ''),
                link: link
            });
        }
    }
    return sourceMap;
}

function _matchAndPrepareUpdates(localData, headerIdx, lMap, sourceMap) {
    let updates = 0;
    const outputColumn = [];

    for (let i = headerIdx + 1; i < localData.length; i++) {
        const row = localData[i];
        
        let rawTime = String(row[lMap.runtime] || '');
        if (rawTime.includes('-')) rawTime = rawTime.split('-')[0].trim();
        
        const c = String(row[lMap.course] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const d = sched_normalizeDay(row[lMap.day]);
        const t = sched_normalizeTime(rawTime, row[lMap.timeofday]);
        const bucketKey = `${c}|${d}|${t}`;

        const candidates = sourceMap.get(bucketKey);
        let link = "";

        if (candidates) {
            const localInstr = String(row[lMap.faculty] || '').toLowerCase().replace(/[^a-z]/g, '');
            // NOTE: Fuzzy matching logic is simple and may have errors with similar names.
            const match = candidates.find(cand => cand.instructor.includes(localInstr) || localInstr.includes(cand.instructor));
            if (match) link = match.link;
        }
        
        if (link) {
            outputColumn.push([link]);
            updates++;
        } else {
            outputColumn.push([row[lMap.zoomlink] || ""]); // Preserve existing link if no new one is found
        }
    }
    return { updates, outputColumn };
}


/**
 * Performs a safe sync from an external data source to the local course schedule.
 * This version preserves event IDs and performs an in-memory merge before clearing the sheet.
 */
function syncExternalCourseData_refactored() {
  try {
    const settings = getSettings('courseImportSettings');
    if (!settings || !settings.sheetUrl || !settings.tabName) {
        return { success: true, message: "No sync configured." }; 
    }

    // 1. Read all necessary data
    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    const localData = localSheet.getDataRange().getValues();
    const sourceId = extractFileIdFromUrl(settings.sheetUrl);
    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName(settings.tabName);
    const sourceValues = sourceSheet.getDataRange().getValues();

    // 2. Prepare local ID map to preserve IDs
    const { headerIdx: localHeaderIdx, headers: localHeaders } = _findHeaderRow(localData, ['eventid', 'course']);
    const localIdMap = _buildLocalIdMap(localData, localHeaderIdx, getColumnMap(localHeaders));

    // 3. Process and merge external data
    const { headerIdx: sourceHeaderIdx, headers: sourceHeaders } = _findHeaderRow(sourceValues, ['course']);
    const mergedData = _mergeSourceData(sourceValues, sourceHeaderIdx, getColumnMap(sourceHeaders), localIdMap);

    // 4. Overwrite local sheet with merged data
    localSheet.clear();
    localSheet.getRange(1, 1, mergedData.length, mergedData[0].length).setValues(mergedData);

    // TODO: Assignment sync logic could also be improved here.

    return { success: true, message: "Sync successful" };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _buildLocalIdMap(data, headerIdx, lMap) {
    const idMap = new Map();
    if (headerIdx > -1 && lMap.eventid !== undefined) {
        for (let i = headerIdx + 1; i < data.length; i++) {
            const row = data[i];
            const id = String(row[lMap.eventid]).trim();
            if (id) {
                const key = `${row[lMap.course]}|${row[lMap.day]}|${row[lMap.runtime]}`.toLowerCase().replace(/\s/g, '');
                idMap.set(key, id);
            }
        }
    }
    return idMap;
}

function _mergeSourceData(sourceData, headerIdx, sMap, localIdMap) {
    const merged = [sourceData[headerIdx]]; // Start with headers
    let idCol = sMap.eventid !== undefined ? sMap.eventid : sourceData[headerIdx].length;
    if (sMap.eventid === undefined) merged[0][idCol] = 'eventID'; // Add header if new column

    for (let i = headerIdx + 1; i < sourceData.length; i++) {
        const row = sourceData[i];
        while (row.length <= idCol) row.push(''); // Ensure row is long enough

        const key = `${row[sMap.course]}|${row[sMap.day]}|${row[sMap.runtime]}`.toLowerCase().replace(/\s/g, '');
        
        const existingId = String(row[idCol] || '').trim();
        const preservedId = localIdMap.get(key);

        row[idCol] = existingId || preservedId || Utilities.getUuid();
        
        // TODO: Data Sanitization (like duration calculation) should be done here.

        merged.push(row);
    }
    return merged;
}

/**
 * Syncs Tech Hub shifts to Google Calendar with improved safety.
 * Creates new events first before deleting old ones on overwrite.
 */
function api_syncTechHubToCalendar_refactored(startStr, endStr, calendarId, overwrite) {
  try {
    if (!calendarId) throw new Error("Calendar ID missing.");
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error("Calendar not found.");

    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    endDate.setHours(23, 59, 59);

    // 1. Get data from sheets
    const ss = getMasterDataHub();
    const shiftsData = getRequiredSheetData(ss, CONFIG.TABS.TECH_HUB_SHIFTS);
    const assignmentsData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);

    // 2. Create maps for efficient lookup
    const staffMap = {}; staffData.slice(1).forEach(r => { if(r[1]) staffMap[String(r[1]).trim()] = r[0]; });
    const shiftsMap = {}; shiftsData.slice(1).forEach(r => { if(r[0]) shiftsMap[String(r[0]).trim()] = { id: r[0], desc: r[1], day: r[2], start: r[3], end: r[4], zoom: (r[5] === true || r[5] === "TRUE") }; });

    const thAssignments = assignmentsData.slice(1).filter(r => r[2] === 'Tech Hub');
    const settings = getSettings('schedulingSettings');
    const masterZoom = settings.zoomUrl || "";

    // 3. Create new events
    const newEvents = [];
    let count = 0;
    for (const assign of thAssignments) {
        // ... [logic to create event details] ...
        const staffId = String(assign[1]).trim();
        const shiftId = String(assign[3]).trim();
        const shift = shiftsMap[shiftId];
        const staffName = staffMap[staffId];

        if (!shift || !staffName) continue;

        // ... [Date and time calculation logic from original function] ...
        const firstOccurrence = sched_getNextDayOfWeek(startDate, shift.day);
        if (firstOccurrence > endDate) continue;

        const startDateTime = new Date(firstOccurrence);
        const endDateTime = new Date(firstOccurrence);
        const startParts = String(shift.start).split(':');
        const endParts = String(shift.end).split(':');
        if (startParts.length < 2 || endParts.length < 2) continue;

        startDateTime.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0);
        endDateTime.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0);
        if (endDateTime <= startDateTime) endDateTime.setHours(endDateTime.getHours() + 12);
        if (endDateTime.getTime() === startDateTime.getTime()) endDateTime.setMinutes(endDateTime.getMinutes() + 60);


        const recurrence = CalendarApp.newRecurrence().addWeeklyRule().until(endDate);
        const title = `Tech Hub: ${staffName}`;
        const location = shift.zoom ? "Zoom" : "Tech Hub";
        let desc = `Shift: ${shift.desc}\nStaff: ${staffName}`;
        if (shift.zoom && masterZoom) desc += `\nZoom: ${masterZoom}`;
        
        try {
            const series = cal.createEventSeries(title, startDateTime, endDateTime, recurrence, { location: location, description: desc });
            series.setTag('AppSource', 'StaffHub');
            newEvents.push(series);
            count++;
        } catch(e) {
            // If one event fails, roll back any created in this run
            newEvents.forEach(evt => { try { evt.deleteEventSeries(); } catch(err){} });
            throw new Error(`Failed to create event for ${staffName}: ${e.message}`);
        }
    }

    // 4. If successful and overwrite is true, delete old events
    if (overwrite) {
      const existingEvents = cal.getEvents(startDate, endDate);
      existingEvents.forEach(e => {
        if (e.getTag('AppSource') === 'StaffHub' && !newEvents.some(ne => ne.getId() === e.getId())) {
            try { e.getEventSeries().deleteEventSeries(); } catch(err) { e.deleteEvent(); }
        }
      });
    }

    return { success: true, message: `Created ${count} recurring shift series.` };
  } catch (e) { 
      return { success: false, message: e.message }; 
  }
}