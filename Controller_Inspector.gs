/**
 * -------------------------------------------------------------------
 * INSPECTOR & ANALYTICS CONTROLLER
 * -------------------------------------------------------------------
 */

/**
 * MAIN INSPECTOR FUNCTION
 * Scans multiple calendars and returns a merged list of events.
 * Used by the "Calendar Inspector" UI.
 */
function api_inspector_getEvents(calIds, startStr, endStr) {
  try {
    if (!calIds || !Array.isArray(calIds) || calIds.length === 0) {
        return { success: false, message: "No calendars selected." };
    }

    const start = new Date(startStr);
    const end = new Date(endStr);
    end.setHours(23, 59, 59);
    
    let allEvents = [];
    
    calIds.forEach(id => {
      try {
        const cal = CalendarApp.getCalendarById(id);
        if(cal) {
          const calName = cal.getName();
          const events = cal.getEvents(start, end);
          
          const mapped = events.map(e => ({
            calendarName: calName,
            title: e.getTitle(),
            start: e.getStartTime().toLocaleString(),
            end: e.getEndTime().toLocaleString(),
            location: e.getLocation() || "",
            guests: e.getGuestList().map(g => g.getEmail()).join(", ")
          }));
          
          allEvents = allEvents.concat(mapped);
        }
      } catch(err) {
        console.error("Error reading cal " + id + ": " + err.message);
      }
    });
    
    // Sort by Start Time
    allEvents.sort((a, b) => new Date(a.start) - new Date(b.start));
    
    return { success: true, data: allEvents };
    
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Exports the inspected events to a new Google Sheet.
 */
function api_inspector_exportEvents(events) {
  try {
    if (!events || events.length === 0) return { success: false, message: "No data to export." };
    
    const ss = SpreadsheetApp.create("Inspector Export " + new Date().toISOString().split('T')[0]);
    const sheet = ss.getActiveSheet();
    
    const headers = ["Calendar", "Title", "Start", "End", "Location", "Guests"];
    const rows = events.map(e => [
        e.calendarName, 
        e.title, 
        e.start, 
        e.end, 
        e.location, 
        e.guests
    ]);
    
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e0e0");
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    
    return { success: true, url: ss.getUrl() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- DISCOVERY & EVENT TYPES ---

/**
 * Fetches defined Event Types from the "Event_Types" tab.
 */
function api_getEventTypes() {
  try {
    const sheet = getSheet('Event_Types');
    const data = sheet.getDataRange().getValues();
    
    const types = [];
    // Skip header
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        types.push({
          name: data[i][0],
          keywords: data[i][1] || ""
        });
      }
    }
    return { success: true, data: types };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
