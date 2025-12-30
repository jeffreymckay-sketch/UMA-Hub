/**
 * -------------------------------------------------------------------
 * CALENDAR SYNC CONTROLLER (Surgical, Transparent & Invite-Aware)
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
    const settings = getSettings(CONFIG.SETTINGS_KEYS.CALENDAR);
    if (!settings || !settings.targetCalendarId) return { success: false, message: "No calendar configured." };
    const cal = CalendarApp.getCalendarById(settings.targetCalendarId);
    return cal ? { success: true, name: cal.getName() } : { success: false, message: "Calendar not found." };
  } catch (e) { return { success: false, message: e.message }; }
}

function api_previewSheetToCalendar(templateName, targetCalendarId) {
  return core_syncLogic(templateName, true, null, targetCalendarId); 
}

function api_syncSheetToCalendar(templateName, selectedRowIndices, targetCalendarId) {
  return core_syncLogic(templateName, false, selectedRowIndices, targetCalendarId); 
}

/**
 * CORE LOGIC: Handles both Preview and Execution.
 * Now includes ZOOM LINK in description and diff logic.
 */
function core_syncLogic(templateName, isPreview, selectedRows, targetCalendarId) {
  try {
    const settings = getSettings(CONFIG.SETTINGS_KEYS.CALENDAR);
    const calId = targetCalendarId || settings.targetCalendarId;
    if (!calId) return { success: false, message: "No Calendar Selected." };

    if (!settings.sourceTabName) return { success: false, message: "Source Tab not configured." };

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

    const assignData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS);
    const staffData = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);
    
    const assignmentMap = new Map();
    for (let i = 1; i < assignData.length; i++) {
        if (assignData[i][2] === 'Course') {
            const staffId = String(assignData[i][1]).trim().toLowerCase();
            const eventId = String(assignData[i][3]).trim(); 
            if (staffId && eventId) assignmentMap.set(eventId, staffId);
        }
    }

    const allStaffEmails = new Set();
    for (let i = 1; i < staffData.length; i++) {
        if (staffData[i][1]) allStaffEmails.add(String(staffData[i][1]).trim().toLowerCase());
    }

    const calendar = CalendarApp.getCalendarById(calId);
    if (!calendar) return { success: false, message: "Target Calendar not found." };
    const calendarName = calendar.getName(); 

    const stats = { created: 0, updated: 0, recreated: 0, skipped: 0, errors: 0 };
    const changes = []; 
    const results = []; 
    const warnings = []; 

    const hMap = headers.map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    const colIdx = {
        id: hMap.indexOf('eventid'),
        course: hMap.indexOf('course'),
        faculty: hMap.indexOf('faculty'),
        startDate: hMap.findIndex(h => h.includes('startdate')),
        endDate: hMap.findIndex(h => h.includes('enddate')),
        day: hMap.indexOf('day'),
        startTime: hMap.indexOf('runtime'), 
        ampm: hMap.indexOf('timeofday'),    
        location: hMap.indexOf('bxlocation'),
        duration: hMap.indexOf('coveragehrs'),
        zoomLink: hMap.indexOf('zoomlink') // NEW
    };
    
    if (colIdx.startTime === -1) colIdx.startTime = hMap.findIndex(h => h.includes('time'));
    if (colIdx.location === -1) colIdx.location = hMap.findIndex(h => h.includes('location') || h.includes('room'));

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
        
        if (!isPreview && selectedRows && !selectedRows.includes(rowNum)) {
            continue;
        }

        const rowId = (colIdx.id > -1) ? String(row[colIdx.id]).trim() : null;
        if (!rowId) {
            if (row[colIdx.course]) warnings.push({ row: rowNum, message: "Skipped: Missing Event ID" });
            continue; 
        }

        let title = pattern;
        headers.forEach((h, idx) => { title = title.replace(new RegExp(`{{${h}}}`, 'gi'), row[idx]); });
        title = title.replace(/{{.*?}}/g, '').trim();
        if (!title || title === '-') title = "Untitled Event";

        // Build Description with Link
        const courseName = (colIdx.course > -1) ? row[colIdx.course] : "";
        const faculty = (colIdx.faculty > -1) ? row[colIdx.faculty] : "";
        const zoomLink = (colIdx.zoomLink > -1) ? String(row[colIdx.zoomLink]).trim() : "";
        
        let description = `Course: ${courseName}\nFaculty: ${faculty}`;
        if (zoomLink) {
            description += `\n\n--- RESOURCES ---\nZoom Link: ${zoomLink}`;
        }

        let startDt = null;
        let endDt = null;
        let seriesEndDate = null;

        if (colIdx.startDate > -1 && row[colIdx.startDate]) {
             const dVal = row[colIdx.startDate];
             let tStr = (colIdx.startTime > -1) ? String(row[colIdx.startTime]) : "09:00";
             const ampmVal = (colIdx.ampm > -1) ? String(row[colIdx.ampm]).trim().toUpperCase() : "";

             startDt = new Date(dVal);
             if (isNaN(startDt.getTime())) {
                 warnings.push({ row: rowNum, message: "Skipped: Invalid Start Date" });
                 continue;
             }

             let startPart = tStr;
             let endPart = null;
             if (tStr.includes('-')) {
                 const parts = tStr.split('-');
                 startPart = parts[0].trim();
                 endPart = parts[1].trim();
             }

             const timeMatch = startPart.match(/(\d+):(\d+)/);
             if (timeMatch) {
                 let h = parseInt(timeMatch[1]);
                 const m = parseInt(timeMatch[2]);
                 
                 const isPm = ampmVal === 'PM';
                 if (isPm && h !== 12) h += 12;
                 else if (!isPm && h === 12) h = 0;
                 
                 startDt.setHours(h, m, 0, 0);
             } else {
                 startDt.setHours(9, 0, 0, 0);
             }
             
             let durationMinutes = 60;
             
             if (endPart) {
                 const endMatch = endPart.match(/(\d+):(\d+)/);
                 if (endMatch) {
                     let eh = parseInt(endMatch[1]);
                     const em = parseInt(endMatch[2]);
                     
                     const isPm = ampmVal === 'PM';
                     let eh_abs = eh;
                     if (isPm && eh !== 12) eh_abs += 12;
                     else if (!isPm && eh === 12) eh_abs = 0;
                     
                     if (eh_abs < startDt.getHours()) eh_abs += 12;

                     const tempEnd = new Date(startDt);
                     tempEnd.setHours(eh_abs, em, 0, 0);
                     
                     const diff = (tempEnd - startDt) / 60000;
                     if (diff > 0) durationMinutes = diff;
                 }
             } else if (colIdx.duration > -1 && row[colIdx.duration]) {
                 const durVal = row[colIdx.duration];
                 if (typeof durVal === 'number') {
                     durationMinutes = durVal * 60;
                 }
             }
             
             if (durationMinutes < 15) durationMinutes = 60;
             endDt = new Date(startDt.getTime() + (durationMinutes * 60000));

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
            
            const timeSignature = `${startDt.toISOString()}_${endDt.toISOString()}_${dayStr}`;

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
                    const existingTag = e.getTag('StaffHub_EventID');
                    if (existingTag && existingTag !== rowId) return false;

                    const eStart = e.getStartTime();
                    if (eStart < startDt) return false; 
                    if (seriesEndDate && eStart > seriesEndDate) return false;

                    const eTitle = e.getTitle().toLowerCase().replace(/\s+/g, '');
                    const isTitleMatch = eTitle.includes(targetTitle) || targetTitle.includes(eTitle);
                    if (!isTitleMatch) return false;

                    const normE = normalizeDateToEpoch(eStart);
                    const normS = normalizeDateToEpoch(startDt);
                    const timeDiff = Math.abs(normE - normS);
                    const isTimeClose = timeDiff <= (45 * 60000);

                    return isTimeClose;
                });

                if (eventToUpdate) matchType = "LEGACY";
            }
            
            if (eventToUpdate) {
                const diffs = [];
                let needsRecreate = false;
                
                if (String(eventToUpdate.getTitle()).trim() !== String(title).trim()) {
                    diffs.push({ field: "Title", old: eventToUpdate.getTitle(), new: title });
                }
                
                const loc1 = String(eventToUpdate.getLocation() || "").trim();
                const loc2 = String(location || "").trim();
                if (loc1 !== loc2) {
                    diffs.push({ field: "Location", old: loc1, new: loc2 });
                }

                // NEW: Check Description (Link)
                const desc1 = String(eventToUpdate.getDescription() || "").trim();
                const desc2 = String(description || "").trim();
                // Simple check: if new description has link and old doesn't, or if they differ
                if (desc1 !== desc2) {
                    diffs.push({ field: "Description", old: "Old Text", new: "Updated (Link)" });
                }
                
                let series = null;
                try { series = eventToUpdate.getEventSeries(); } catch(e) {}
                
                const storedSig = series ? series.getTag('StaffHub_TimeSignature') : eventToUpdate.getTag('StaffHub_TimeSignature');
                
                if (storedSig !== timeSignature) {
                     if (!storedSig) {
                         const existStart = eventToUpdate.getStartTime();
                         const existEnd = eventToUpdate.getEndTime();
                         const normExistStart = normalizeDateToEpoch(existStart);
                         const normStart = normalizeDateToEpoch(startDt);
                         const normExistEnd = normalizeDateToEpoch(existEnd);
                         const normEnd = normalizeDateToEpoch(endDt);
                         const timeDiff = Math.abs(normExistStart - normStart) > 60000 || 
                                          Math.abs(normExistEnd - normEnd) > 60000;
                         if (timeDiff) {
                             diffs.push({ field: "Time", old: formatTime(existStart), new: formatTime(startDt) });
                             needsRecreate = true;
                         }
                     } else {
                         diffs.push({ field: "Time", old: "Old Schedule", new: "New Schedule" });
                         needsRecreate = true;
                     }
                }

                const targetEmail = assignmentMap.get(rowId);
                const currentGuests = eventToUpdate.getGuestList().map(g => g.getEmail().toLowerCase());
                
                if (targetEmail && !currentGuests.includes(targetEmail)) {
                    diffs.push({ field: "Invite", old: "(Missing)", new: targetEmail });
                }
                
                currentGuests.forEach(email => {
                    if (allStaffEmails.has(email) && email !== targetEmail) {
                        diffs.push({ field: "Uninvite", old: email, new: "(Remove)" });
                    }
                });

                if (diffs.length > 0 || matchType === "LEGACY") {
                    const actionType = needsRecreate ? "RECREATE" : "UPDATE";
                    
                    if (!isPreview) {
                        try {
                            if (matchType === "LEGACY") {
                                if(series) series.setTag('StaffHub_EventID', rowId);
                                else eventToUpdate.setTag('StaffHub_EventID', rowId);
                            }
                            
                            if (needsRecreate) {
                                if(series) series.deleteEventSeries();
                                else eventToUpdate.deleteEvent();
                                
                                const weekday = parseDayOfWeek(dayStr);
                                if (weekday && seriesEndDate) {
                                    const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(weekday).until(seriesEndDate);
                                    const newSeries = calendar.createEventSeries(title, startDt, endDt, recurrence, { location: location, description: description });
                                    newSeries.setTag('StaffHub_EventID', rowId);
                                    newSeries.setTag('StaffHub_TimeSignature', timeSignature);
                                    newSeries.setTag('AppSource', 'StaffHub');
                                    if (targetEmail) newSeries.addGuest(targetEmail);
                                    stats.recreated++;
                                    results.push({ title: title, status: "✅ Recreated", details: "Time changed, series rebuilt." });
                                } else {
                                    stats.errors++;
                                    results.push({ title: title, status: "❌ Error", details: `Could not recreate: Invalid Day (${dayStr}) or End Date.` });
                                }
                            } else {
                                const targetObj = series || eventToUpdate;
                                targetObj.setTitle(title);
                                targetObj.setLocation(location);
                                targetObj.setDescription(description); // Update Description
                                if (!storedSig) targetObj.setTag('StaffHub_TimeSignature', timeSignature);
                                
                                stats.updated++;
                                results.push({ title: title, status: "✅ Updated", details: "Metadata updated." });
                            }
                            Utilities.sleep(500); 
                        } catch (err) {
                            stats.errors++;
                            results.push({ title: title, status: "❌ Error", details: err.message });
                        }
                    } else {
                        changes.push({ 
                            row: rowNum, 
                            action: actionType, 
                            title: title, 
                            diffs: diffs,
                            matchType: matchType
                        });
                    }
                } else {
                    stats.skipped++;
                }

            } else {
                // --- CREATE NEW ---
                const targetEmail = assignmentMap.get(rowId);
                
                if (!isPreview) {
                    try {
                        if (seriesEndDate && seriesEndDate > startDt) {
                            const weekday = parseDayOfWeek(dayStr);
                            if (weekday) {
                                const recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(weekday).until(seriesEndDate);
                                const series = calendar.createEventSeries(title, startDt, endDt, recurrence, { location: location, description: description });
                                series.setTag('StaffHub_EventID', rowId);
                                series.setTag('StaffHub_TimeSignature', timeSignature);
                                series.setTag('AppSource', 'StaffHub');
                                if (targetEmail) series.addGuest(targetEmail);
                                stats.created++;
                                results.push({ title: title, status: "✅ Created", details: "New Series created." });
                                Utilities.sleep(1000); 
                            } else {
                                stats.errors++;
                                results.push({ title: title, status: "❌ Skipped", details: `Invalid Day of Week: "${dayStr}"` });
                            }
                        } else {
                            stats.errors++;
                            results.push({ title: title, status: "❌ Skipped", details: "Missing or Invalid End Date for Series." });
                        }
                    } catch (err) {
                        stats.errors++;
                        results.push({ title: title, status: "❌ Error", details: err.message });
                    }
                } else {
                    const inviteDiff = targetEmail ? [{ field: "Invite", old: "-", new: targetEmail }] : [];
                    changes.push({ 
                        row: rowNum, 
                        action: "CREATE", 
                        title: title, 
                        diffs: [{ field: "New Event", old: "-", new: "Create Series" }, ...inviteDiff] 
                    });
                }
            }
        }
    }
    
    const msg = isPreview 
        ? `Found ${changes.length} proposed changes for calendar: <strong>${calendarName}</strong>` 
        : `Sync Complete for <strong>${calendarName}</strong>. Created ${stats.created}, Updated ${stats.updated}, Recreated ${stats.recreated}.`;

    return { 
        success: true, 
        isPreview: isPreview, 
        stats: stats, 
        changes: changes, 
        results: results, 
        warnings: warnings, 
        message: msg
    };

  } catch (e) { return { success: false, message: "Sync Error: " + e.message }; }
}

function normalizeDateToEpoch(d) {
    const n = new Date(d);
    n.setFullYear(2000, 0, 1); 
    n.setSeconds(0);
    n.setMilliseconds(0);
    return n.getTime();
}

function parseDayOfWeek(dayStr) {
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

function formatTime(date) {
    let h = date.getHours();
    const m = date.getMinutes();
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12;
    h = h ? h : 12; 
    const mStr = m < 10 ? '0'+m : m;
    return `${h}:${mStr} ${ampm}`;
}

function api_syncStaffToCalendar(targetCalendarId) {
  try {
    if (typeof Calendar === 'undefined') {
        return { success: false, message: "Error: 'Google Calendar API' service is not enabled in the script editor." };
    }

    const settings = getSettings(CONFIG.SETTINGS_KEYS.CALENDAR);
    const calId = targetCalendarId || settings.targetCalendarId;
    if (!calId) return { success: false, message: "No Calendar Selected." };

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
    
    const calendar = CalendarApp.getCalendarById(calId);
    const events = calendar.getEvents(now, future);
    
    let invitesSent = 0;
    let invitesRemoved = 0;
    let errors = 0;

    const processedSeries = new Set();

    for (const event of events) {
        let eventId = event.getTag('StaffHub_EventID');
        let rawId = event.getId();
        let cleanId = rawId.split('@')[0].split('_')[0]; 

        if (!eventId) {
            try {
                const series = event.getEventSeries();
                if (series) eventId = series.getTag('StaffHub_EventID');
            } catch(e) { }
        }
        
        if (eventId && assignmentMap.has(eventId)) {
            if (processedSeries.has(cleanId)) continue;
            processedSeries.add(cleanId);

            const targetEmail = assignmentMap.get(eventId);
            const currentGuests = event.getGuestList().map(g => g.getEmail().toLowerCase());
            
            let needsUpdate = false;
            let attendees = event.getGuestList().map(g => ({ email: g.getEmail() }));

            if (!currentGuests.includes(targetEmail)) {
                attendees.push({ email: targetEmail });
                needsUpdate = true;
                invitesSent++;
            }

            const filteredAttendees = [];
            attendees.forEach(att => {
                const email = att.email.toLowerCase();
                if (allStaffEmails.has(email) && email !== targetEmail) {
                    needsUpdate = true;
                    invitesRemoved++;
                } else {
                    filteredAttendees.push(att);
                }
            });

            if (needsUpdate) {
                try {
                    const resource = { attendees: filteredAttendees };
                    Calendar.Events.patch(resource, calId, cleanId, { sendUpdates: 'all' });
                    Utilities.sleep(500);
                } catch (err) {
                    console.error("API Error: " + err.message);
                    errors++;
                }
            }
        }
    }

    return { 
        success: true, 
        message: `Invites Sent: ${invitesSent}\nInvites Removed: ${invitesRemoved}\nErrors: ${errors}` 
    };

  } catch (e) { return { success: false, message: "Invite Error: " + e.message }; }
}

function api_getCalendarEvents(startStr, endStr, calendarId) {
  try {
    let targetId = calendarId || getSettings(CONFIG.SETTINGS_KEYS.CALENDAR).targetCalendarId;
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