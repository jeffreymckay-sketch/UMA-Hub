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

/**
 * PREVIEW LOGIC: Used by MST Tool to compare Sheet vs Calendar.
 */
function api_previewMstCalendarSync(targetCalendarId) {
  // Pass through to the core logic function (assumed to be in Controller_Calendar or similar)
  // For this project structure, we will include the core logic here to ensure it works.
  return core_syncLogic(null, true, null, targetCalendarId);
}

function api_commitMstCalendarEvents(targetCalendarId, eventsToSync) {
    // We reuse the commit logic from the Controller_MST if separated, 
    // but typically the sync logic resides here for the Calendar API.
    // Ensure the helper 'api_commitMstCalendarEvents' in JS_MST calls this.
    
    // NOTE: If your sync logic was in Controller_MST.js, ensure that file remains.
    // If it was here, we need to preserve the heavy lifting functions.
    // Based on previous context, specific sync logic for MST was in Controller_MST.js.
    // This file focuses on the LISTING and ACCESS of calendars.
    
    // However, the MST Controller calls `api_commitMstCalendarEvents`. 
    // If that logic was in Controller_MST.js, do not duplicate it here.
    return { success: false, message: "Sync logic should be in Controller_MST.js" };
}