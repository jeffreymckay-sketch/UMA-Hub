/**
 * @file Utilities.gs
 * @description This file contains shared utility functions used throughout the application.
 */

/**
 * Normalizes a header string by converting it to lowercase, trimming whitespace,
 * and removing all non-alphanumeric characters.
 * 
 * @param {string} header The header string to normalize.
 * @returns {string} The normalized header string.
 */
function normalizeHeader(header) {
  if (!header) return '';
  return header.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');
}

/**
 * Converts a time value (Date object, number, or string) into a standardized
 * "h:mm a" format.
 * 
 * @param {Date|number|string} input The time value to normalize.
 * @returns {string} The normalized time string.
 */
function normalizeTime(input) {
  if (input === null || input === undefined || input === '') return '';
  if (input instanceof Date) {
    if (isNaN(input.getTime())) return 'Invalid Time';
    return Utilities.formatDate(input, Session.getScriptTimeZone(), "h:mm a");
  }
  if (typeof input === 'number') {
    let num = input;
    if (num < 1 && num > 0) {
      const totalMinutes = Math.round(num * 24 * 60);
      const h = Math.floor(totalMinutes / 60);
      const m = totalMinutes % 60;
      const ampm = h >= 12 ? 'PM' : 'AM';
      const h12 = h % 12 || 12;
      return `${h12}:${m.toString().padStart(2, '0')} ${ampm}`;
    }
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

/**
 * Creates a map of column names to their index from a header row.
 * 
 * @param {Array<string>} headerRow The header row of a spreadsheet.
 * @returns {object} A map where keys are normalized column names and values are their indices.
 */
function getColumnMap(headerRow) {
    const map = {};
    headerRow.forEach((col, index) => {
        if (col) map[String(col).trim().toLowerCase().replace(/\s+/g, '')] = index;
    });
    return map;
}

/**
 * Creates a lookup map from a 2D array of data.
 *
 * @param {Array<Array<string>>} data The 2D array of data.
 * @param {string} keyColumnName The name of the column to use as the key.
 * @param {string} valueColumnName The name of the column to use as the value.
 * @returns {object} A map where keys are values from the key column and values are from the value column.
 */
function createLookupMap(data, keyColumnName, valueColumnName) {
    if (!data || data.length < 2) return {};
    const headers = data[0].map(normalizeHeader);
    const keyIndex = headers.indexOf(normalizeHeader(keyColumnName));
    const valueIndex = headers.indexOf(normalizeHeader(valueColumnName));
    if (keyIndex === -1 || valueIndex === -1) return {};

    const lookupMap = {};
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const key = row[keyIndex];
        if (key) {
            lookupMap[key] = row[valueIndex];
        }
    }
    return lookupMap;
}

/**
 * Generates a standardized time block key from a time value.
 * This is used for preference and availability lookups.
 * @param {Date|string} time The time to convert.
 * @returns {string} A string in 'HHMM' format.
 */
function getTimeBlock(time) {
    const d = (time instanceof Date) ? time : new Date('1970-01-01T' + time);
    if (isNaN(d.getTime())) return null;
    return d.getHours().toString().padStart(2, '0') + d.getMinutes().toString().padStart(2, '0');
}

/**
 * Checks if two time ranges overlap.
 * @param {Date|string} start1 Start of the first time range.
 * @param {Date|string} end1 End of the first time range.
 * @param {Date|string} start2 Start of the second time range.
 * @param {Date|string} end2 End of the second time range.
 * @returns {boolean} True if the time ranges overlap.
 */
function timesOverlap(start1, end1, start2, end2) {
    const s1 = (start1 instanceof Date) ? start1 : new Date('1970-01-01T' + start1);
    const e1 = (end1 instanceof Date) ? end1 : new Date('1970-01-01T' + end1);
    const s2 = (start2 instanceof Date) ? start2 : new Date('1970-01-01T' + start2);
    const e2 = (end2 instanceof Date) ? end2 : new Date('1970-01-01T' + end2);

    if (isNaN(s1.getTime()) || isNaN(e1.getTime()) || isNaN(s2.getTime()) || isNaN(e2.getTime())) {
        return false;
    }

    return (s1 < e2 && s2 < e1);
}