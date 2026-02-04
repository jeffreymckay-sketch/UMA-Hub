
/**
 * ------------------------------------------------------------------
 * DATA PARSERS
 * ------------------------------------------------------------------
 * This file contains functions for parsing raw data from Google Sheets
 * into the clean data models defined in Model.gs.
 * ------------------------------------------------------------------
 */

/**
 * Parses a raw data row from the Staff_List tab into a Staff object.
 * @param {Array} row - The raw data row from the spreadsheet.
 * @param {Object} headers - A map of header names to column indices.
 * @returns {Staff|null} A Staff object or null if the row is invalid.
 */
function parseStaff(row, headers) {
    const id = getString(row, headers, 'staffid', true); // Aggressively clean ID
    if (!id) return null;

    const name = getString(row, headers, 'fullname');
    const role = getString(row, headers, 'roles');
    const isActive = getBoolean(row, headers, 'isactive');

    return new Staff(id, name, role, isActive);
}

/**
 * Parses a raw data row from the Course_Schedule tab into a Course object.
 * @param {Array} row - The raw data row from the spreadsheet.
 * @param {Object} headers - A map of header names to column indices.
 * @returns {Course|null} A Course object or null if the row is invalid.
 */
function parseCourse(row, headers) {
    let id = getString(row, headers, 'eventid', true); // Aggressively clean ID
    const courseName = getString(row, headers, 'course');

    if (!courseName) return null; // Skip rows without a course name
    if (!id) id = Utilities.getUuid(); // Generate a UUID if no ID is present
    
    const faculty = getString(row, headers, 'faculty');
    const location = getString(row, headers, 'bxlocation');
    const zoomLink = getString(row, headers, 'zoomlink');
    const days = getString(row, headers, 'day');

    const { startDate, endDate } = parseCourseTimes(row, headers);
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        return null;
    }

    const daysOfWeek = days ? days.split('/').map(d => d.trim()) : [];

    return new Course(id, courseName, startDate, endDate, faculty, location, zoomLink, daysOfWeek);
}

/**
 * Parses a raw data row from the Tech_Hub_Shifts tab into a Shift object.
 * @param {Array} row - The raw data row from the spreadsheet.
 * @param {Object} headers - A map of header names to column indices.
 * @returns {Shift|null} A Shift object or null if the row is invalid.
 */
function parseShift(row, headers) {
    const id = getString(row, headers, 'shiftid', true); // Aggressively clean ID
    if (!id) return null;

    const description = getString(row, headers, 'description');
    const day = getString(row, headers, 'day');
    const startTimeStr = getString(row, headers, 'starttime'); 
    const endTimeStr = getString(row, headers, 'endtime');     
    const isZoom = getBoolean(row, headers, 'iszoom');
    const location = isZoom ? 'Zoom' : 'Tech Hub Desk';

    const startDate = new Date(`1970-01-01T${startTimeStr}`);
    const endDate = new Date(`1970-01-01T${endTimeStr}`);

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        return null;
    }

    return new Shift(id, `${description} (${day})`, startDate, endDate, location);
}

/**
 * Parses a raw data row from the Staff_Assignments tab into an Assignment object.
 * @param {Array} row - The raw data row from the spreadsheet.
 * @param {Object} headers - A map of header names to column indices.
 * @returns {Assignment|null} An Assignment object or null if the row is invalid.
 */
function parseAssignment(row, headers) {
    const recordId = getString(row, headers, 'assignmentid', true);
    const staffId = getString(row, headers, 'staffid', true);
    const itemId = getString(row, headers, 'referenceid', true);
    const type = getString(row, headers, 'assignmenttype');

    if (!recordId || !staffId || !itemId || !type) return null;

    const eventType = (type.toLowerCase() === 'tech hub') ? 'SHIFT' : 'COURSE';

    return new Assignment(recordId, staffId, itemId, eventType);
}


// --- UTILITY HELPERS ---

/**
 * Robustly gets and cleans a string value from a sheet row.
 * Can perform extra-aggressive cleaning for ID fields.
 * @param {Array} row The sheet row
 * @param {Object} headers The column map
 * @param {string} key The header key to look for
 * @param {boolean} isId If true, performs aggressive cleaning for ID fields.
 * @returns {string} The cleaned string value.
 */
function getString(row, headers, key, isId = false) {
    const index = headers[key.toLowerCase().replace(/[\s_]/g, '')];
    if (index === undefined || row[index] === null || row[index] === undefined || row[index] === '') {
        return '';
    }

    let value = String(row[index]);

    if (isId) {
      value = value.replace(/[^a-zA-Z0-9-_.@+]/g, '');
    } 

    return value.replace(/\u00A0/g, ' ').trim();
}

function getBoolean(row, headers, key) {
    const index = headers[key.toLowerCase().replace(/[\s_]/g, '')];
    if (index === undefined || row[index] === '') return false;
    
    const value = getString(row, headers, key);
    return value.toUpperCase() === 'TRUE';
}

/**
 * Parses the complex time format from the course schedule spreadsheet.
 * @param {Array} row - The raw data row.
 * @param {Object} headers - The header map.
 * @returns {{startDate: Date, endDate: Date}} An object with Date objects.
 */
function parseCourseTimes(row, headers) {
    const runtime = getString(row, headers, 'runtime');
    const timeOfDay = getString(row, headers, 'timeofday').toUpperCase();

    const invalidDate = new Date(NaN);
    if (!runtime || !timeOfDay || (timeOfDay !== 'AM' && timeOfDay !== 'PM')) {
        return { startDate: invalidDate, endDate: invalidDate };
    }

    const times = runtime.split('-');
    if (times.length !== 2) {
        return { startDate: invalidDate, endDate: invalidDate };
    }

    try {
        const [startTimeStr, endTimeStr] = times.map(t => t.trim());

        const [startHourStr, startMinuteStr] = startTimeStr.split(':');
        const [endHourStr, endMinuteStr] = endTimeStr.split(':');

        let startHour = parseInt(startHourStr, 10);
        const startMinute = parseInt(startMinuteStr, 10);
        let endHour = parseInt(endHourStr, 10);
        const endMinute = parseInt(endMinuteStr, 10);
        
        if (isNaN(startHour) || isNaN(startMinute) || isNaN(endHour) || isNaN(endMinute)) {
            return { startDate: invalidDate, endDate: invalidDate };
        }

        if (timeOfDay === 'PM' && startHour < 12) {
            startHour += 12;
        } else if (timeOfDay === 'AM' && startHour === 12) { 
            startHour = 0;
        }

        if (endHour < startHour || (endHour === 12 && startHour < 12)) {
            endHour += 12;
        }

        const startDate = new Date();
        startDate.setHours(startHour, startMinute, 0, 0);

        const endDate = new Date();
        endDate.setHours(endHour, endMinute, 0, 0);

        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            return { startDate: invalidDate, endDate: invalidDate };
        }

        return { startDate, endDate };

    } catch (e) {
        console.error(`Failed to parse time: '${runtime} ${timeOfDay}'. Error: ${e.message}`);
        return { startDate: invalidDate, endDate: invalidDate };
    }
}
