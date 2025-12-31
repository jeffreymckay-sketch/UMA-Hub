/**
 * @file Controller_Zoom.gs
 * @description Controller for the Zoom Link Management feature.
 */

/**
 * Previews the Zoom link import process without making any changes.
 *
 * @param {string} sheetUrl The URL of the Google Sheet containing the Zoom links.
 * @param {string} tabName The name of the tab containing the link data.
 * @returns {object} A result object with a success flag, message, and data for the preview.
 */
function previewZoomImport(sheetUrl, tabName) {
  try {
    if (!sheetUrl || !tabName) {
      return { success: false, message: "Missing URL or Tab Name." };
    }

    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    if (!localSheet) {
      return { success: false, message: "Local Course Schedule not found." };
    }

    // 1. Prepare Local Data
    const localData = localSheet.getDataRange().getValues();
    let localHeaderIdx = -1;
    for (let r = 0; r < Math.min(localData.length, 5); r++) {
      const rowStr = localData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
      if (rowStr.includes('course') || rowStr.includes('startdate')) {
        localHeaderIdx = r;
        break;
      }
    }
    if (localHeaderIdx === -1) {
      return { success: false, message: "Local headers not found." };
    }

    const lHeaders = localData[localHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    const lIdx = {
      course: lHeaders.indexOf('course'),
      faculty: lHeaders.indexOf('faculty'),
      day: lHeaders.indexOf('day'),
      runTime: lHeaders.indexOf('runtime'),
      timeOfDay: lHeaders.indexOf('timeofday'),
      zoomLink: lHeaders.indexOf('zoomlink')
    };

    // 2. Fetch External Data
    const sourceId = extractFileIdFromUrl(sheetUrl);
    const sourceSS = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceSS.getSheetByName(tabName);
    if (!sourceSheet) {
      return { success: false, message: "External tab not found." };
    }

    const sourceValues = sourceSheet.getDataRange().getValues();
    const sourceRichText = sourceSheet.getDataRange().getRichTextValues();

    let sourceHeaderIdx = -1;
    for (let r = 0; r < Math.min(sourceValues.length, 5); r++) {
      const rowStr = sourceValues[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
      if (rowStr.includes('coursenumber') && rowStr.includes('yourname')) {
        sourceHeaderIdx = r;
        break;
      }
    }
    if (sourceHeaderIdx === -1) {
      return { success: false, message: "External headers not found." };
    }

    const sHeaders = sourceValues[sourceHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    const sIdx = {
      purpose: sHeaders.indexOf('whatisthepurposeofthissubmission?'),
      course: sHeaders.indexOf('coursenumber(e.g.mat115)'),
      instructor: sHeaders.indexOf('yourname'),
      day: sHeaders.indexOf('day(s)(ifapplicable)'),
      startTime: sHeaders.indexOf('starttime(ifapplicable)'),
      link: sHeaders.indexOf('whatisyourzoomlinkforyourhyflexclass?'),
      passcode: sHeaders.indexOf('whatisthepasscodeforthiszoomlink?')
    };

    if (sIdx.link === -1) {
      return { success: false, message: "'Zoom Link' column not found in source." };
    }

    // 3. Build Source Map (Bucket by Course|Day|Time)
    const sourceMap = new Map();
    for (let i = sourceHeaderIdx + 1; i < sourceValues.length; i++) {
      const row = sourceValues[i];
      const richRow = sourceRichText[i];

      if (row[sIdx.purpose] !== "Submit your course Zoom link") {
        continue;
      }

      const courseVal = row[sIdx.course];
      if (!courseVal) {
        continue;
      }

      const c = String(courseVal).toLowerCase().replace(/[^a-z0-9]/g, '');
      const d = sched_normalizeDay(row[sIdx.day]);
      const t = sched_normalizeTime(row[sIdx.startTime], null);
      const bucketKey = `${c}|${d}|${t}`;

      let link = String(row[sIdx.link]).trim();
      if (!link.startsWith('http')) {
        const richCell = richRow[sIdx.link];
        const url = richCell.getLinkUrl();
        if (url) {
          link = url;
        }
      }

      if (link && link.startsWith('http')) {
        const passcode = row[sIdx.passcode];
        if (passcode) {
          link += ` (Passcode: ${passcode})`;
        }

        if (!sourceMap.has(bucketKey)) {
          sourceMap.set(bucketKey, []);
        }
        sourceMap.get(bucketKey).push({
          instructor: String(row[sIdx.instructor]).toLowerCase().replace(/[^a-z]/g, ''),
          link: link
        });
      }
    }

    // 4. Find changes
    const changes = [];
    for (let i = localHeaderIdx + 1; i < localData.length; i++) {
      const row = localData[i];

      let rawTime = String(row[lIdx.runTime]);
      if (rawTime.includes('-')) {
        rawTime = rawTime.split('-')[0].trim();
      }

      const c = String(row[lIdx.course]).toLowerCase().replace(/[^a-z0-9]/g, '');
      const d = sched_normalizeDay(row[lIdx.day]);
      const t = sched_normalizeTime(rawTime, row[lIdx.timeOfDay]);
      const bucketKey = `${c}|${d}|${t}`;

      const candidates = sourceMap.get(bucketKey);
      let newLink = "";

      if (candidates) {
        const localInstr = String(row[lIdx.faculty]).toLowerCase().replace(/[^a-z]/g, '');
        const match = candidates.find(cand => cand.instructor.includes(localInstr) || localInstr.includes(cand.instructor));
        if (match) {
          newLink = match.link;
        }
      }
      
      const oldLink = lIdx.zoomLink > -1 ? row[lIdx.zoomLink] : "";

      if (newLink && newLink !== oldLink) {
        changes.push({
          rowIndex: i + 1, // 1-based for clarity
          courseName: row[lIdx.course],
          faculty: row[lIdx.faculty],
          day: row[lIdx.day],
          time: rawTime,
          oldLink: oldLink,
          newLink: newLink
        });
      }
    }

    return { success: true, data: changes };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Executes the Zoom link import based on the provided changes.
 *
 * @param {Array<object>} changes The array of changes to apply.
 * @returns {object} A result object with a success flag and a message.
 */
function executeZoomImport(changes) {
  try {
    if (!changes || changes.length === 0) {
      return { success: false, message: "No changes to apply." };
    }

    const ss = getMasterDataHub();
    const localSheet = ss.getSheetByName(CONFIG.TABS.COURSE_SCHEDULE);
    if (!localSheet) {
      return { success: false, message: "Local Course Schedule not found." };
    }

    const localData = localSheet.getDataRange().getValues();
    let localHeaderIdx = -1;
    for (let r = 0; r < Math.min(localData.length, 5); r++) {
        const rowStr = localData[r].join(' ').toLowerCase().replace(/[\s_]/g, '');
        if (rowStr.includes('course') || rowStr.includes('startdate')) {
            localHeaderIdx = r;
            break;
        }
    }
    if (localHeaderIdx === -1) return { success: false, message: "Local headers not found." };

    const lHeaders = localData[localHeaderIdx].map(h => String(h).toLowerCase().replace(/[\s_]/g, ''));
    let zoomLinkCol = lHeaders.indexOf('zoomlink');

    if (zoomLinkCol === -1) {
      zoomLinkCol = localData[localHeaderIdx].length;
      localSheet.getRange(localHeaderIdx + 1, zoomLinkCol + 1).setValue("Zoom Link");
    }

    changes.forEach(change => {
      localSheet.getRange(change.rowIndex, zoomLinkCol + 1).setValue(change.newLink);
    });

    return { success: true, message: `Successfully updated ${changes.length} Zoom links.` };

  } catch (e) {
    return { success: false, message: e.message };
  }
}


/**
 * Normalizes a day string to a number (0-6).
 *
 * @param {string} dayStr The day string.
 * @returns {number} The normalized day number.
 */
function sched_normalizeDay(dayStr) {
  if (!dayStr) return 0;
  const s = String(dayStr).toLowerCase().trim();
  if (s.includes('mon') || s === 'm') return 1;
  if (s.includes('tue') || s === 'tu' || s === 't') return 2;
  if (s.includes('wed') || s === 'w') return 3;
  if (s.includes('thu') || s === 'th' || s === 'r') return 4;
  if (s.includes('fri') || s === 'f') return 5;
  if (s.includes('sat') || s === 'sa') return 6;
  if (s.includes('sun') || s === 'su') return 0;
  return 0;
}

/**
 * Normalizes a time string to minutes from midnight.
 *
 * @param {string} timeVal The time string.
 * @param {string} amPmVal The AM/PM string.
 * @returns {number} The normalized time in minutes.
 */
function sched_normalizeTime(timeVal, amPmVal) {
  if (!timeVal) return 0;

  let h = 0, m = 0;

  if (timeVal instanceof Date) {
    h = timeVal.getHours();
    m = timeVal.getMinutes();
  } else {
    const str = String(timeVal).trim();
    const match = str.match(/(\d+):(\d+)/);
    if (match) {
      h = parseInt(match[1]);
      m = parseInt(match[2]);
    }

    if (!amPmVal) {
      if (str.toUpperCase().includes('PM')) amPmVal = 'PM';
      if (str.toUpperCase().includes('AM')) amPmVal = 'AM';
    }
  }

  if (amPmVal) {
    const isPm = String(amPmVal).trim().toUpperCase().includes('PM');
    const isAm = String(amPmVal).trim().toUpperCase().includes('AM');

    if (isPm && h < 12) h += 12;
    if (isAm && h === 12) h = 0;
  }

  return (h * 60) + m;
}

/**
 * Extracts a file ID from a Google Sheet URL.
 *
 * @param {string} url The URL to extract the ID from.
 * @returns {string|null} The extracted file ID.
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Gets the Master Spreadsheet instance.
 * RELIABLE METHOD: Defaults to ActiveSpreadsheet (Bound Script).
 */
function getMasterDataHub() {
  try {
    // 1. Primary: Get the sheet this script is attached to
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;

    // 2. Fallback: Check properties if not bound
    const props = PropertiesService.getScriptProperties();
    const settingsStr = props.getProperty('adminSettings');
    if (settingsStr) {
      const settings = JSON.parse(settingsStr);
      if (settings.dataHubUrl) {
        return settings.dataHubUrl.includes('http')
          ? SpreadsheetApp.openByUrl(settings.dataHubUrl)
          : SpreadsheetApp.openById(settings.dataHubUrl);
      }
    }

    throw new Error("Script is not bound to a sheet and no URL is saved in settings.");
  } catch (e) {
    console.error("Connection Error: " + e.message);
    throw new Error("System Error: Could not connect to Data Hub.");
  }
}
