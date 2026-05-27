/**
 * -------------------------------------------------------------------
 * CALENDAR SYNC CONTROLLER
 * Handles fetching, scanning, and syncing with Google Calendar.
 * -------------------------------------------------------------------
 */

/**
 * Fetches all calendars the user has WRITE access to (Owner or Writer).
 * Used by: MST Sync, Nursing Sync, Tech Hub Sync
 */
function getWritableCalendars() {
  try {
    // We try to use the Advanced Calendar Service first (Better results)
    if (typeof Calendar !== 'undefined') {
      const response = Calendar.CalendarList.list({
        showDeleted: false, 
        minAccessRole: 'writer' 
      });
      
      if (response.items) {
        const calendars = response.items.map(cal => ({
          id: cal.id,
          name: cal.summaryOverride || cal.summary
        }));
        // Sort alphabetically
        calendars.sort((a, b) => a.name.localeCompare(b.name));
        console.log("Found calendars: " + JSON.stringify(calendars));
        return { success: true, data: calendars };
      }
    }

    // Fallback to basic app class if Advanced Service fails
    const calendars = CalendarApp.getAllOwnedCalendars();
    const writable = calendars.map(cal => ({ id: cal.getId(), name: cal.getName() }));
    writable.sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, data: writable };

  } catch (e) { 
    return { success: false, message: "Calendar Error: " + e.message };
  }
}

/**
 * Fetches ALL calendars visible to the user (including Read-Only).
 * Used by: Reporting / Inspector Tool
 */
function api_getAllCalendars() {
  try {
    const calendars = CalendarApp.getAllCalendars();
    const viewable = calendars.map(cal => ({ 
      id: cal.getId(), 
      name: cal.getName(),
      isOwned: cal.isOwnedByMe()
    }));
    
    // Sort: Owned first, then alphabetical
    viewable.sort((a, b) => {
        if (a.isOwned && !b.isOwned) return -1;
        if (!a.isOwned && b.isOwned) return 1;
        return a.name.localeCompare(b.name);
    });
    
    return { success: true, data: viewable };
  } catch (e) { 
    return { success: false, message: e.message };
  }
}

/**
 * Fetches a specific calendar's name by ID.
 */
function api_getCalendarTargetName() {
  try {
    const settings = getSettings();
    if (!settings || !settings.targetCalendarId) return { success: false, message: "No calendar configured." };
    const cal = CalendarApp.getCalendarById(settings.targetCalendarId);
    return cal ? { success: true, name: cal.getName() } : { success: false, message: "Calendar not found." };
  } catch (e) { 
    return { success: false, message: e.message };
  }
}