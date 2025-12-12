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
    const ss = getMasterDataHub();
    const sheet = getOrCreateSheet(ss, CONFIG.TABS.EVENT_TYPES, ['Category Name', 'Keywords']);
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

/**
 * Scans a calendar for frequent event titles to help discover categories.
 */
function api_discovery_scanCalendar(calId, daysBack) {
  try {
    if (!calId) throw new Error("No calendar ID provided.");
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) throw new Error("Calendar not found.");

    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - (daysBack || 30));

    const events = cal.getEvents(start, end);
    const frequencyMap = {};

    events.forEach(e => {
      let title = e.getTitle().toLowerCase().trim();
      if (title) {
        frequencyMap[title] = (frequencyMap[title] || 0) + 1;
      }
    });

    const sorted = Object.keys(frequencyMap)
      .map(key => ({ keyword: key, count: frequencyMap[key] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 50); 

    return { success: true, count: events.length, data: sorted };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Saves new Event Type rules to the "Event_Types" tab.
 */
function api_discovery_saveRules(rules) {
  try {
    const ss = getMasterDataHub();
    const sheet = getOrCreateSheet(ss, CONFIG.TABS.EVENT_TYPES, ['Category Name', 'Keywords']);
    const data = sheet.getDataRange().getValues();
    
    const categoryMap = {}; 
    for (let i = 1; i < data.length; i++) {
      categoryMap[data[i][0]] = i + 1;
    }

    rules.forEach(rule => {
      const catName = rule.category;
      const keyword = rule.keyword;

      if (categoryMap[catName]) {
        const rowIdx = categoryMap[catName];
        const currentKeywords = sheet.getRange(rowIdx, 2).getValue();
        if (!currentKeywords.includes(keyword)) {
          const newKeywords = currentKeywords ? (currentKeywords + ", " + keyword) : keyword;
          sheet.getRange(rowIdx, 2).setValue(newKeywords);
        }
      } else {
        sheet.appendRow([catName, keyword]);
        categoryMap[catName] = sheet.getLastRow();
      }
    });

    return { success: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}