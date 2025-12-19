/**
 * -------------------------------------------------------------------
 * CALENDAR SYNC CONTROLLER (MST & Generic)
 * -------------------------------------------------------------------
 */

function getWritableCalendars() {
  try {
    const calendars = CalendarApp.getAllOwnedCalendars();
    const writable = calendars.map(cal => ({ id: cal.getId(), name: cal.getName() }));
    writable.sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, data: writable };
  } catch (e) { return { success: false, message: e.message }; }
}

function getAllViewableCalendars() {
  try {
    const calendars = CalendarApp.getAllCalendars();
    const viewable = calendars.map(cal => ({ id: cal.getId(), name: cal.getName() }));
    viewable.sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, data: viewable };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_getCalendarTargetName() {
  try {
    const settings = getSettings('calendarSettings');
    if (!settings || !settings.targetCalendarId) return { success: false, message: "No calendar configured." };
    const cal = CalendarApp.getCalendarById(settings.targetCalendarId);
    return cal ? { success: true, name: cal.getName() } : { success: false, message: "Calendar not found." };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_syncSheetToCalendar(templateName) { return core_syncLogic(templateName, false); }
function api_previewSheetToCalendar(templateName) { return core_syncLogic(templateName, true); }

function core_syncLogic(templateName, isPreview) {
  try {
    const settings = getSettings('calendarSettings');
    if (!settings.targetCalendarId || !settings.sourceTabName) return { success: false, message: "Config missing." };

    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName(settings.sourceTabName);
    if (!sheet) return { success: false, message: "Tab not found." };
    
    const data = sheet.getDataRange().getValues();
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
    let pattern = "{{Course}} - {{Faculty}}";
    if (settings.savedTemplates) {
      const t = settings.savedTemplates.find(x => x.name === templateName);
      if (t) pattern = t.pattern;
    }

    const calendar = CalendarApp.getCalendarById(settings.targetCalendarId);
    const stats = { created: 0, updated: 0, skipped: 0, errors: 0 };
    const changes = [];
    const warnings = []; 

    const hMap = headers.map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    const colIdx = {
        id: hMap.indexOf('eventid'),
        course: hMap.indexOf('course'),
        startDate: hMap.findIndex(h => h.includes('startdate')),
        endDate: hMap.findIndex(h => h.includes('enddate')),
        day: hMap.indexOf('day'),
        startTime: hMap.indexOf('runtime'), 
        ampm: hMap.indexOf('timeofday'),    
        location: hMap.indexOf('bxlocation'),
        duration: hMap.indexOf('coveragehrs')
    };
    
    if (colIdx.startTime === -1) colIdx.startTime = hMap.findIndex(h => h.includes('time'));
    if (colIdx.location === -1) colIdx.location = hMap.findIndex(h => h.includes('location') || h.includes('room'));

    // Cache Existing Events
    let minDate = new Date(8640000000000000);
    let maxDate = new Date(-8640000000000000);

    rows.forEach(row => {
        if (colIdx.startDate > -1 && row[colIdx.startDate]) {
            const d = new Date(row[colIdx.startDate]);
            if (!isNaN(d.getTime())) {
                if (d < minDate) minDate = d;
                if (colIdx.endDate > -1 && row[colIdx.endDate]) {
                    const ed = new Date(row[colIdx.endDate]);
                    if (ed > maxDate) maxDate = ed;
                } else {
                    if (d > maxDate) maxDate = d;
                }
            }
        }
    });

    const eventIdMap = new Map(); 
    const allExistingEvents = []; 

    if (minDate < maxDate) {
        minDate.setHours(0,0,0,0);
        maxDate.setHours(23,59,59,999);
        const existingEvents = calendar.getEvents(minDate, maxDate);
        existingEvents.forEach(e => {
            const tagId = e.getTag('StaffHub_EventID');
            if (tagId) eventIdMap.set(tagId, e);
            allExistingEvents.push(e); 
        });
    }

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowNum = i + headerRowIdx + 2;
        
        const rowId = (colIdx.id > -1) ? String(row[colIdx.id]).trim() : null;
        if (!rowId) {
            if (row[colIdx.course]) warnings.push({ row: rowNum, message: "Skipped: Missing Event ID" });
            continue; 
        }

        let title = pattern;
        headers.forEach((h, idx) => { title = title.replace(new RegExp(`{{${h}}}`, 'gi'), row[idx]); });
        title = title.replace(/{{.*?}}/g, '').trim();
        if (!title || title === '-') title = "Untitled Event";

        // --- TIME PARSING ---
        let startDt = null;
        let endDt = null;
        let seriesEndDate = null;

        if (colIdx.startDate > -1 && row[colIdx.startDate]) {
             const dVal = row[colIdx.startDate];
             let tStr = (colIdx.startTime > -1) ? String(row[colIdx.startTime]) : "09:00";
             const ampmVal = (colIdx.ampm > -1) ? String(row[colIdx.ampm]).trim() : "";

             startDt = new Date(dVal);
             if (isNaN(startDt.getTime())) {
                 warnings.push({ row: rowNum, message: "Skipped: Invalid Start Date" });
                 continue;
             }

             // Robust Split: Handle "9:00-11:00", "9:00 - 11:00", "9:00 to 11:00"
             let startPart = tStr;
             let endPart = null;
             
             if (tStr.includes('-')) {
                 const parts = tStr.split('-');
                 startPart = parts[0].trim();
                 endPart = parts[1].trim();
             } else if (tStr.toLowerCase().includes(' to ')) {
                 const parts = tStr.toLowerCase().split(' to ');
                 startPart = parts[0].trim();
                 endPart = parts[1].trim();
             }

             const timeMatch = startPart.match(/(\d+):(\d+)/);
             if (timeMatch) {
                 let h = parseInt(timeMatch[1]);
                 const m = parseInt(timeMatch[2]);
                 
                 const isPM = ampmVal.toLowerCase() === 'pm' || (!ampmVal && startPart.toLowerCase().includes('pm'));
                 const isAM = ampmVal.toLowerCase() === 'am' || (!ampmVal && startPart.toLowerCase().includes('am'));
                 
                 if (isPM && h < 12) h += 12;
                 if (isAM && h === 12) h = 0;
                 
                 startDt.setHours(h, m, 0, 0);
             } else {
                 startDt.setHours(9, 0, 0, 0);
             }
             
             // END TIME LOGIC
             let explicitEndFound = false;

             // 1. Try Explicit End Time from String (Highest Priority)
             if (endPart) {
                 const endMatch = endPart.match(/(\d+):(\d+)/);
                 if (endMatch) {
                     let eh = parseInt(endMatch[1]);
                     const em = parseInt(endMatch[2]);
                     
                     // AM/PM Inference
                     // If End < Start (e.g. 11:00 - 1:00), assume PM wrap
                     if (eh < startDt.getHours()) eh += 12;
                     // If Start is PM (e.g. 1:00 PM), End must be PM
                     if (startDt.getHours() >= 12 && eh < 12) eh += 12;

                     endDt = new Date(startDt);
                     endDt.setHours(eh, em, 0, 0);
                     explicitEndFound = true;
                 }
             }
             
             // 2. Fallback to Duration Column
             if (!explicitEndFound) {
                 let durationMinutes = 60;
                 if (colIdx.duration > -1 && row[colIdx.duration]) {
                     const durVal = row[colIdx.duration];
                     if (typeof durVal === 'number') {
                         if (durVal < 1) durationMinutes = durVal * 24 * 60;
                         else if (durVal <= 12) durationMinutes = durVal * 60;
                         else durationMinutes = durVal;
                     }
                 }
                 if (durationMinutes < 15) durationMinutes = 60;
                 endDt = new Date(startDt.getTime() + (durationMinutes * 60000));
             }

             if (colIdx.endDate > -1 && row[colIdx.endDate]) {
                 seriesEndDate = new Date(row[colIdx.endDate]);
                 seriesEndDate.setHours(23, 59, 59);
             }
        } else {
            warnings.push({ row: rowNum, message: "Skipped: No Start Date column found" });
            continue;
        }

        if (startDt && !isNaN(startDt.getTime())) {
            const location = (colIdx.location > -1) ? row[colIdx.location] : "";
            const dayStr = (colIdx.day > -1) ? String(row[colIdx.day]) : "";
            
            // --- MATCHING LOGIC ---
            let eventToUpdate = null;
            let matchType = "NEW";

            if (eventIdMap.has(rowId)) {
                eventToUpdate = eventIdMap.get(rowId);
                matchType = "ID";
            } else {
                const targetDay = startDt.getDay();
                const candidates = allExistingEvents.filter(e => e.getStartTime().getDay() === targetDay);
                const targetTitle = title.toLowerCase().replace(/\s+/g, '');
                
                eventToUpdate = candidates.find(e => {
                    const eTitle = e.getTitle().toLowerCase().replace(/\s+/g, '');
                    const eStart = e.getStartTime();
                    const timeDiff = Math.abs(eStart.getHours()*60 + eStart.getMinutes() - (startDt.getHours()*60 + startDt.getMinutes()));
                    const isTimeClose = timeDiff <= 45;
                    const isTitleMatch = eTitle.includes(targetTitle) || targetTitle.includes(eTitle);
                    return isTimeClose && isTitleMatch;
                });

                if (eventToUpdate) matchType = "LEGACY";
            }
            
            if (eventToUpdate) {
                // --- DIFF LOGIC ---
                const diffs = [];
                
                if (String(eventToUpdate.getTitle()).trim() !== String(title).trim()) {
                    diffs.push(`Title: "${eventToUpdate.getTitle()}" → "${title}"`);
                }
                
                const loc1 = String(eventToUpdate.getLocation() || "").trim();
                const loc2 = String(location || "").trim();
                if (loc1 !== loc2) {
                    diffs.push(`Loc: "${loc1}" → "${loc2}"`);
                }
                
                const existStart = eventToUpdate.getStartTime();
                const existEnd = eventToUpdate.getEndTime();
                
                const time1 = `${formatTime(existStart)} - ${formatTime(existEnd)}`;
                const time2 = `${formatTime(startDt)} - ${formatTime(endDt)}`;
                
                // Compare minutes to avoid second-level diffs
                const isTimeDiff = Math.abs(existStart.getTime() - startDt.getTime()) > 60000 || 
                                   Math.abs(existEnd.getTime() - endDt.getTime()) > 60000;

                if (isTimeDiff) {
                     diffs.push(`Time: ${time1} → ${time2}`);
                }

                if (diffs.length > 0 || matchType === "LEGACY") {
                    if (!isPreview) {
                        if (matchType === "LEGACY") eventToUpdate.setTag('StaffHub_EventID', rowId);
                        if (diffs.length > 0) {
                            eventToUpdate.setTitle(title);
                            eventToUpdate.setLocation(location);
                            // Only update time if it's a single event or we are confident
                            // For series, we usually can't update time easily without recreating
                        }
                        stats.updated++;
                    } else {
                        const actionLabel = matchType === "LEGACY" ? "LINK & UPDATE" : "UPDATE";
                        const details = diffs.length > 0 ? diffs.join("<br>") : "Restoring Link (ID Mismatch)";
                        changes.push({ row: rowNum, action: actionLabel, title: title, details: details });
                        stats.updated++;
                    }
                } else {
                    stats.skipped++;
                }

            } else {
                // CREATE NEW
                if (!isPreview) {
                    if (seriesEndDate && seriesEndDate > startDt) {
                        const weekday = parseDayOfWeek(dayStr);
                        if (weekday) {
                            const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(weekday).until(seriesEndDate);
                            const series = calendar.createEventSeries(title, startDt, endDt, recurrence, { location: location });
                            series.setTag('StaffHub_EventID', rowId);
                            series.setTag('AppSource', 'StaffHub');
                            stats.created++;
                            Utilities.sleep(1500); 
                        } else {
                            const evt = calendar.createEvent(title, startDt, endDt, { location: location });
                            evt.setTag('StaffHub_EventID', rowId);
                            evt.setTag('AppSource', 'StaffHub');
                            stats.created++;
                            Utilities.sleep(800);
                        }
                    } else {
                        const evt = calendar.createEvent(title, startDt, endDt, { location: location });
                        evt.setTag('StaffHub_EventID', rowId);
                        evt.setTag('AppSource', 'StaffHub');
                        stats.created++;
                        Utilities.sleep(800);
                    }
                } else {
                    const type = (seriesEndDate && seriesEndDate > startDt) ? "SERIES" : "SINGLE";
                    changes.push({ row: rowNum, action: "CREATE", title: title, details: `${type} starting ${startDt.toLocaleString()}` });
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
        warnings: warnings, 
        message: isPreview ? `Preview: ${stats.created} Create, ${stats.updated} Update.` : `Sync Complete. Created ${stats.created}.` 
    };

  } catch (e) { return { success: false, message: "Sync Error: " + e.message }; }
}

function parseDayOfWeek(dayStr) {
    if (!dayStr) return null;
    const s = dayStr.toLowerCase().trim();
    if (s.includes('mon')) return CalendarApp.Weekday.MONDAY;
    if (s.includes('tue')) return CalendarApp.Weekday.TUESDAY;
    if (s.includes('wed')) return CalendarApp.Weekday.WEDNESDAY;
    if (s.includes('thu')) return CalendarApp.Weekday.THURSDAY;
    if (s.includes('fri')) return CalendarApp.Weekday.FRIDAY;
    if (s.includes('sat')) return CalendarApp.Weekday.SATURDAY;
    if (s.includes('sun')) return CalendarApp.Weekday.SUNDAY;
    return null;
}

function formatTime(date) {
    let h = date.getHours();
    const m = date.getMinutes();
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12;
    h = h ? h : 12; 
    const mStr = m < 10 ? '0'+m : m;
    return `${h}:${mStr} ${ampm}`;
}

function api_syncStaffToCalendar() {
  try {
    const settings = getSettings('calendarSettings');
    if (!settings.targetCalendarId) return { success: false, message: "No calendar configured." };
    
    const cal = CalendarApp.getCalendarById(settings.targetCalendarId);
    if (!cal) return { success: false, message: "Calendar not found." };

    const ss = getMasterDataHub();
    const assignData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);

    const allStaffEmails = new Set();
    for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][1]) allStaffEmails.add(String(staffData[i][1]).trim().toLowerCase());
    }

    const assignmentMap = new Map();
    for (let i = 1; i < assignData.length; i++) {
        if (assignData[i][2] === 'Course') {
            const staffId = String(assignData[i][1]).trim().toLowerCase();
            const eventId = String(assignData[i][3]).trim(); 
            if (staffId && eventId) assignmentMap.set(eventId, staffId);
        }
    }

    const now = new Date();
    const future = new Date();
    future.setDate(now.getDate() + 120);
    
    const events = cal.getEvents(now, future);
    let invitesSent = 0;
    let invitesRemoved = 0;

    for (const event of events) {
        let eventId = event.getTag('StaffHub_EventID');
        if (!eventId) {
            try {
                const series = event.getEventSeries();
                if (series) eventId = series.getTag('StaffHub_EventID');
            } catch(e) { }
        }
        
        if (eventId && assignmentMap.has(eventId)) {
            const targetEmail = assignmentMap.get(eventId);
            const currentGuests = event.getGuestList().map(g => g.getEmail().toLowerCase());
            
            if (!currentGuests.includes(targetEmail)) {
                event.addGuest(targetEmail);
                invitesSent++;
                Utilities.sleep(500);
            }

            for (const guestEmail of currentGuests) {
                if (allStaffEmails.has(guestEmail) && guestEmail !== targetEmail) {
                    event.removeGuest(guestEmail);
                    invitesRemoved++;
                    Utilities.sleep(200);
                }
            }
        }
    }

    return { 
        success: true, 
        message: `Process Complete.\nSent ${invitesSent} new invites.\nRemoved ${invitesRemoved} outdated staff guests.` 
    };

  } catch (e) { return { success: false, message: "Invite Error: " + e.message }; }
}

function api_getCalendarEvents(startStr, endStr, calendarId) {
  try {
    let targetId = calendarId || getSettings('calendarSettings').targetCalendarId;
    if (!targetId) return { success: false, message: "No calendar." };
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
  } catch (e) { return { success: false, message: e.message }; }
}

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
  } catch (e) { return { success: false, message: e.message }; }
}