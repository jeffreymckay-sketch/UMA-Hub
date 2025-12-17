/**
 * -------------------------------------------------------------------
 * CALENDAR SYNC CONTROLLER (MST & Generic)
 * -------------------------------------------------------------------
 */

/**
 * Returns a list of calendars the user has write access to.
 */
function getWritableCalendars() {
  try {
    const calendars = CalendarApp.getAllOwnedCalendars();
    const writable = calendars.map(cal => ({
      id: cal.getId(),
      name: cal.getName(),
      description: cal.getDescription()
    }));
    
    writable.sort((a, b) => a.name.localeCompare(b.name));
    
    return { success: true, data: writable };
  } catch (e) {
    console.error("Error fetching calendars: " + e.message);
    return { success: false, message: "Error fetching calendars: " + e.message };
  }
}

/**
 * Returns a list of ALL calendars the user can view.
 */
function getAllViewableCalendars() {
  try {
    const calendars = CalendarApp.getAllCalendars();
    const viewable = calendars.map(cal => ({
      id: cal.getId(),
      name: cal.getName(),
      description: cal.getDescription()
    }));
    
    viewable.sort((a, b) => a.name.localeCompare(b.name));
    
    return { success: true, data: viewable };
  } catch (e) {
    console.error("Error fetching calendars: " + e.message);
    return { success: false, message: "Error fetching calendars: " + e.message };
  }
}

/**
 * Returns the name of the currently configured target calendar.
 */
function api_getCalendarTargetName() {
  try {
    const settings = getSettings('calendarSettings');
    if (!settings || !settings.targetCalendarId) {
      return { success: false, message: "No calendar configured." };
    }
    
    const cal = CalendarApp.getCalendarById(settings.targetCalendarId);
    if (!cal) {
      return { success: false, message: "Calendar not found (ID invalid)." };
    }
    
    return { success: true, name: cal.getName() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Syncs data from the Source Sheet to the Target Calendar.
 */
function api_syncSheetToCalendar(templateName) {
  return core_syncLogic(templateName, false); // false = execute
}

/**
 * Previews changes without applying them.
 */
function api_previewSheetToCalendar(templateName) {
  return core_syncLogic(templateName, true); // true = preview
}

/**
 * Core Sync Logic (Shared by Sync and Preview)
 * UPDATED: Added Throttling (Utilities.sleep) to prevent Rate Limit errors.
 */
function core_syncLogic(templateName, isPreview) {
  try {
    // 1. Load Settings
    const settings = getSettings('calendarSettings');
    if (!settings.targetCalendarId || !settings.sourceTabName) {
      return { success: false, message: "Calendar or Source Tab not configured." };
    }

    // 2. Get Data
    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName(settings.sourceTabName);
    if (!sheet) return { success: false, message: `Tab "${settings.sourceTabName}" not found.` };
    
    const data = sheet.getDataRange().getValues();
    
    // Handle Header Row (Row 2 detection logic)
    let headerRowIdx = 0;
    for(let r=0; r<Math.min(data.length, 5); r++) {
        const rowStr = data[r].join(' ').toLowerCase();
        if(rowStr.includes('start date') && (rowStr.includes('time of day') || rowStr.includes('run time'))) {
            headerRowIdx = r;
            break;
        }
    }

    const headers = data[headerRowIdx];
    const rows = data.slice(headerRowIdx + 1);

    // 3. Get Template Pattern
    let pattern = "{{Course}} - {{Faculty}}"; // Default
    if (settings.savedTemplates) {
      const t = settings.savedTemplates.find(x => x.name === templateName);
      if (t) pattern = t.pattern;
    }

    // 4. Prepare Calendar
    const calendar = CalendarApp.getCalendarById(settings.targetCalendarId);
    const stats = { created: 0, updated: 0, skipped: 0, errors: 0 };
    const changes = [];

    // 5. Identify Columns
    const hMap = headers.map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    
    const colIdx = {
        id: hMap.indexOf('eventid'),
        course: hMap.indexOf('course'),
        date: hMap.indexOf('startdate'),
        startTime: hMap.indexOf('runtime'), 
        ampm: hMap.indexOf('timeofday'),    
        location: hMap.indexOf('bxlocation')
    };
    
    // Fallbacks
    if (colIdx.date === -1) colIdx.date = hMap.findIndex(h => h.includes('date'));
    if (colIdx.startTime === -1) colIdx.startTime = hMap.findIndex(h => h.includes('time'));
    if (colIdx.location === -1) colIdx.location = hMap.findIndex(h => h.includes('location') || h.includes('room'));

    // Determine Date Range
    let minDate = new Date(8640000000000000);
    let maxDate = new Date(-8640000000000000);

    rows.forEach(row => {
        if (colIdx.date > -1 && row[colIdx.date]) {
            const d = new Date(row[colIdx.date]);
            if (!isNaN(d.getTime())) {
                if (d < minDate) minDate = d;
                if (d > maxDate) maxDate = d;
            }
        }
    });

    // Fetch Existing Events
    const eventIdMap = new Map(); 
    const legacyMap = new Map();  

    if (minDate < maxDate) {
        minDate.setHours(0,0,0,0);
        maxDate.setHours(23,59,59,999);
        const existingEvents = calendar.getEvents(minDate, maxDate);
        
        existingEvents.forEach(e => {
            const tagId = e.getTag('StaffHub_EventID');
            if (tagId) {
                eventIdMap.set(tagId, e);
            } else {
                const key = e.getStartTime().toISOString();
                if (!legacyMap.has(key)) legacyMap.set(key, []);
                legacyMap.get(key).push(e);
            }
        });
    }

    // 6. Process Rows
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowId = (colIdx.id > -1) ? String(row[colIdx.id]).trim() : null;
        const courseName = (colIdx.course > -1) ? String(row[colIdx.course]).trim() : "";
        
        // Construct Title
        let title = pattern;
        headers.forEach((h, idx) => {
            title = title.replace(new RegExp(`{{${h}}}`, 'gi'), row[idx]);
        });
        title = title.replace(/{{.*?}}/g, '').trim();
        if (!title || title === '-') title = "Untitled Event";

        // Determine Time
        let startDt = null;
        let endDt = null;

        if (colIdx.date > -1 && row[colIdx.date]) {
             const dVal = row[colIdx.date];
             let tStr = (colIdx.startTime > -1) ? String(row[colIdx.startTime]) : "09:00";
             const ampmVal = (colIdx.ampm > -1) ? String(row[colIdx.ampm]).trim() : "";

             if (tStr.includes('-')) tStr = tStr.split('-')[0].trim();
             
             startDt = new Date(dVal);
             const timeMatch = tStr.match(/(\d+):(\d+)/);
             if (timeMatch) {
                 let h = parseInt(timeMatch[1]);
                 const m = parseInt(timeMatch[2]);
                 if (ampmVal.toLowerCase() === 'pm' && h < 12) h += 12;
                 if (ampmVal.toLowerCase() === 'am' && h === 12) h = 0;
                 if (!ampmVal && tStr.toLowerCase().includes('pm') && h < 12) h += 12;
                 startDt.setHours(h, m, 0, 0);
             } else {
                 startDt.setHours(9, 0, 0, 0);
             }
             
             let durationMinutes = 60;
             if (colIdx.duration > -1 && row[colIdx.duration]) {
                 const durVal = row[colIdx.duration];
                 if (String(durVal).includes(':')) {
                     const parts = String(durVal).split(':');
                     durationMinutes = (parseInt(parts[0]) * 60) + parseInt(parts[1]);
                 } else if (typeof durVal === 'number') {
                     if (durVal < 1) durationMinutes = durVal * 24 * 60;
                     else durationMinutes = durVal;
                 }
             }
             endDt = new Date(startDt.getTime() + (durationMinutes * 60000));
        }

        if (startDt && !isNaN(startDt.getTime())) {
            const location = (colIdx.location > -1) ? row[colIdx.location] : "";
            let eventToUpdate = null;
            let matchType = "NEW";

            // A. Check ID Match
            if (rowId && eventIdMap.has(rowId)) {
                eventToUpdate = eventIdMap.get(rowId);
                matchType = "ID";
            } 
            // B. Check Legacy Match
            else {
                const timeKey = startDt.toISOString();
                if (legacyMap.has(timeKey)) {
                    const candidates = legacyMap.get(timeKey);
                    eventToUpdate = candidates.find(e => e.getTitle().includes(courseName));
                    if (eventToUpdate) matchType = "LEGACY";
                }
            }

            if (eventToUpdate) {
                // UPDATE EXISTING
                if (!isPreview) {
                    let needsUpdate = false;
                    
                    // Only call API if data actually changed (Saves Quota)
                    if (eventToUpdate.getTitle() !== title) {
                        eventToUpdate.setTitle(title);
                        needsUpdate = true;
                    }
                    if (eventToUpdate.getLocation() !== location) {
                        eventToUpdate.setLocation(location);
                        needsUpdate = true;
                    }
                    
                    // Always ensure tags are set for future
                    if (rowId && eventToUpdate.getTag('StaffHub_EventID') !== rowId) {
                        eventToUpdate.setTag('StaffHub_EventID', rowId);
                        needsUpdate = true;
                    }
                    if (eventToUpdate.getTag('AppSource') !== 'StaffHub') {
                        eventToUpdate.setTag('AppSource', 'StaffHub');
                        needsUpdate = true;
                    }

                    if (needsUpdate) {
                        stats.updated++;
                        // THROTTLE: Sleep 800ms after every update to prevent "Too Many Requests"
                        Utilities.sleep(800); 
                    } else {
                        stats.skipped++;
                    }
                } else {
                    changes.push({ row: i + headerRowIdx + 2, action: "UPDATE (" + matchType + ")", title: title, details: "Updating existing event" });
                    stats.updated++;
                }
            } else {
                // CREATE NEW
                if (!isPreview) {
                    const evt = calendar.createEvent(title, startDt, endDt, { location: location });
                    if (rowId) evt.setTag('StaffHub_EventID', rowId);
                    evt.setTag('AppSource', 'StaffHub');
                    stats.created++;
                    
                    // THROTTLE: Sleep 1000ms after creation (Creation is heavier than update)
                    Utilities.sleep(1000); 
                } else {
                    changes.push({ row: i + headerRowIdx + 2, action: "CREATE", title: title, details: startDt.toLocaleString() });
                    stats.created++;
                }
            }
        }
    }
    
    return { 
      success: true, 
      isPreview: isPreview, 
      stats: stats, 
      changes: changes,
      message: isPreview ? `Preview: ${stats.created} Create, ${stats.updated} Update.` : `Sync Complete. Created ${stats.created}, Updated ${stats.updated}, Skipped ${stats.skipped}.` 
    };

  } catch (e) {
    return { success: false, message: "Sync Error: " + e.message };
  }
}

/**
 * Sends invites to staff based on the synced events.
 * Checks existing guest list to avoid duplicate invites.
 */
function api_syncStaffToCalendar() {
  try {
    // 1. Setup
    const settings = getSettings('calendarSettings');
    if (!settings.targetCalendarId) return { success: false, message: "No calendar configured." };
    
    const cal = CalendarApp.getCalendarById(settings.targetCalendarId);
    if (!cal) return { success: false, message: "Calendar not found." };

    const ss = getMasterDataHub();
    const assignData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);

    // 2. Build Maps
    const staffEmailMap = new Map();
    for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][1]) staffEmailMap.set(String(staffData[i][1]).trim(), String(staffData[i][1]).trim());
    }

    const assignmentMap = new Map();
    for (let i = 1; i < assignData.length; i++) {
        if (assignData[i][2] === 'Course') {
            const staffId = String(assignData[i][1]).trim();
            const eventId = String(assignData[i][3]).trim(); 
            
            if (staffId && eventId && staffEmailMap.has(staffId)) {
                assignmentMap.set(eventId, staffEmailMap.get(staffId));
            }
        }
    }

    // 3. Scan Future Events
    const now = new Date();
    const future = new Date();
    future.setDate(now.getDate() + 120);
    
    const events = cal.getEvents(now, future);
    let invitesSent = 0;
    let skipped = 0;

    for (const event of events) {
        const eventId = event.getTag('StaffHub_EventID');
        
        if (eventId && assignmentMap.has(eventId)) {
            const targetEmail = assignmentMap.get(eventId);
            const currentGuests = event.getGuestList().map(g => g.getEmail());
            
            if (!currentGuests.includes(targetEmail)) {
                event.addGuest(targetEmail);
                invitesSent++;
                // THROTTLE: Sleep 500ms after adding guest
                Utilities.sleep(500);
            } else {
                skipped++;
            }
        }
    }

    return { 
        success: true, 
        message: `Process Complete.\nSent ${invitesSent} new invites.\nSkipped ${skipped} existing guests.` 
    };

  } catch (e) {
    return { success: false, message: "Invite Error: " + e.message };
  }
}

/**
 * Inspects the calendar for events in a date range.
 */
function api_getCalendarEvents(startStr, endStr, calendarId) {
  try {
    let targetId = calendarId;
    
    if (!targetId) {
        const settings = getSettings('calendarSettings');
        targetId = settings.targetCalendarId;
    }
    
    if (!targetId) return { success: false, message: "No calendar selected or configured." };
    
    const cal = CalendarApp.getCalendarById(targetId);
    if (!cal) return { success: false, message: "Calendar not found." };

    const start = new Date(startStr);
    const end = new Date(endStr);
    end.setHours(23, 59, 59); 

    const events = cal.getEvents(start, end);
    const result = events.map(evt => ({
      title: evt.getTitle(),
      start: evt.getStartTime().toLocaleString(),
      end: evt.getEndTime().toLocaleString(),
      location: evt.getLocation(),
      guests: evt.getGuestList().map(g => g.getEmail()).join(", ")
    }));

    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Exports inspected events to a new sheet.
 */
function api_exportCalendarEvents(events) {
  try {
    const ss = SpreadsheetApp.create("Calendar Export " + new Date().toISOString().split('T')[0]);
    const sheet = ss.getActiveSheet();
    
    if (!events || events.length === 0) return { success: false, message: "No data." };
    
    const headers = ["Title", "Start", "End", "Location", "Guests"];
    const rows = events.map(e => [e.title, e.start, e.end, e.location, e.guests]);
    
    sheet.appendRow(headers);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    return { success: true, url: ss.getUrl() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}