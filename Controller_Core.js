/**
 * -------------------------------------------------------------------
 * CORE CONTROLLER
 * Handles essential, app-wide data and actions.
 * -------------------------------------------------------------------
 */

function api_getInitialAppData() {
    var payload = {
        success: false,
        message: "Uninitialized",
        data: {}
    };

    try {
        // 1. User Info
        var userInfo = { success: false, message: "Loading..." };
        try {
            userInfo = getUserInfo();
        } catch (e) { userInfo = { success: false, message: e.message }; }

        // 2. Settings
        var allSettings = { success: false, message: "Loading..." };
        try {
            allSettings = getAllSettings_();
        } catch (e) { allSettings = { success: false, message: e.message }; }

        // 3. Nursing Data
        var nursingData = { success: false, message: "Loading..." };
        try {
            if (typeof api_getNursingData === 'function') {
                nursingData = api_getNursingData();
            } else {
                nursingData = { success: false, message: "Nursing function missing." };
            }
        } catch (e) { nursingData = { success: false, message: "Nursing Error: " + e.message }; }

        // 4. Placeholders for other tools (Disabled to prevent crashes)
        var mstViewData = { success: false, error: "MST Disabled" };
        var calendars = { success: true, data: [] };

        // Construct Payload
        payload = {
            success: true,
            data: {
                userInfo: userInfo.success ? userInfo.data : { error: userInfo.message },
                mstData: mstViewData,
                settings: allSettings.success ? allSettings.data : { error: allSettings.message },
                writableCalendars: calendars.data,
                nursingData: nursingData
            }
        };

        // SAFETY CHECK: Ensure payload is serializable
        // This prevents the "null response" error by catching it server-side
        try {
            JSON.stringify(payload);
        } catch (jsonError) {
            throw new Error("Data Serialization Failed (Circular reference or invalid object): " + jsonError.message);
        }

        return payload;

    } catch (e) {
        console.error("api_getInitialAppData Critical Failure: " + e.stack);
        return {
            success: false,
            message: "Critical Server Error: " + e.message
        };
    }
}

function api_saveSettings(key, settingsObject) {
    try {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty(key, JSON.stringify(settingsObject));
        const allSettings = getAllSettings_();
        return { success: true, message: 'Settings saved!', data: allSettings.data };
    } catch (e) {
        return { success: false, message: 'Failed to save settings: ' + e.message };
    }
}

// --- INTERNAL HELPERS ---

function getAllSettings_() {
    try {
        const props = PropertiesService.getUserProperties().getProperties();
        const parsed = {};
        for (const key in props) {
            try { parsed[key] = JSON.parse(props[key]); } 
            catch (e) { parsed[key] = props[key]; }
        }
        return { success: true, data: parsed };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function getUserInfo() {
    try {
        return { success: true, data: { email: Session.getActiveUser().getEmail(), photoUrl: "" } };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * TIME NORMALIZER
 * Used by Nursing Tool to clean up time inputs.
 */
function normalizeTime(input) {
  if (input === null || input === undefined || input === '') return '';

  // Handle Date Objects
  if (input instanceof Date) {
    if (isNaN(input.getTime())) return 'Invalid Time';
    return Utilities.formatDate(input, Session.getScriptTimeZone(), "h:mm a");
  }

  // Handle Numbers
  if (typeof input === 'number') {
    let num = input;
    // Excel serial time
    if (num < 1 && num > 0) {
      const totalMinutes = Math.round(num * 24 * 60);
      const h = Math.floor(totalMinutes / 60);
      const m = totalMinutes % 60;
      const ampm = h >= 12 ? 'PM' : 'AM';
      const h12 = h % 12 || 12;
      return `${h12}:${m.toString().padStart(2, '0')} ${ampm}`;
    }
    // Integer time (900, 1300)
    let str = num.toString();
    if (num < 24) str = num + "00";
    if (str.length === 3) str = "0" + str;
    if (str.length === 4) {
      const h = parseInt(str.substring(0, 2));
      const m = str.substring(2);
      const ampm = h >= 12 ? 'PM' : 'AM';
      const h12 = h % 12 || 12;
      return `${h12}:${m} ${ampm}`;
    }
  }

  // Handle Strings
  const text = String(input).trim().toLowerCase();
  const match = text.match(/(\d{1,2})[:.]?(\d{2})?\s*(a|p|am|pm)?/);
  if (match) {
    let h = parseInt(match[1]);
    let m = match[2] || "00";
    let period = match[3]; 

    if (!period) {
      if (h >= 7 && h <= 11) period = 'am';
      else if (h === 12) period = 'pm';
      else if (h >= 1 && h <= 6) period = 'pm'; 
      else if (h > 12) period = 'pm'; 
    }
    if (h > 12) { h = h - 12; period = 'pm'; }
    
    const cleanPeriod = (period && period.startsWith('p')) ? 'PM' : 'AM';
    return `${h}:${m} ${cleanPeriod}`;
  }
  return String(input); 
}