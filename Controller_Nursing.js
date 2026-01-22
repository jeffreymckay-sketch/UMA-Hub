/**
 * ----------------------------------------------------------------------------------------
 * Controller for the Nursing Proctoring Tool
 * 
 * Contains all server-side logic specific to the Nursing feature, including data retrieval,
 * document creation/updating, and accommodation management.
 * ----------------------------------------------------------------------------------------
 */

/**
 * ----------------------------------------------------------------------------------------
 * API Endpoints - Nursing Proctoring Tool
 * ----------------------------------------------------------------------------------------
 */

function api_getNursingData() {
  try {
    const settings = getNursingSettings();
    
    // Validate IDs
    if (!settings.nursingSheetId) throw new Error("Invalid Sheet ID. Please check your settings.");
    if (!settings.nursingFolderId) throw new Error("Invalid Folder ID. Please check your settings.");
    
    const spreadsheet = SpreadsheetApp.openById(settings.nursingSheetId);
    
    const allSheetData = spreadsheet.getSheets()
      .map(sheet => {
        const sheetName = sheet.getName();
        // Skip templates/masters
        if (sheetName.toLowerCase().match(/(template|master)/)) return null;

        // 1. Parse the Sheet
        const parsed = parseNursingSheet(sheet);
        
        if (!parsed || parsed.exams.length === 0) {
            console.log(`Skipping sheet "${sheetName}" - No valid data found.`);
            return null;
        }
        
        return {
          sheetName: sheetName,
          course: parsed.course,
          exams: parsed.exams
        };
      })
      .filter(s => s !== null);

    // 2. Find Document URLs (Subfolder Logic) & Check Calendar Status
    const mainFolder = DriveApp.getFolderById(settings.nursingFolderId);
    
    // Attempt to load calendar for status checks
    let calendar = null;
    if (settings.nursingCalendarId) {
        try { calendar = CalendarApp.getCalendarById(settings.nursingCalendarId); } catch(e) { console.warn("Calendar check failed: " + e.message); }
    }
    
    allSheetData.forEach(sheetData => {
        // Look for the subfolder matching the Course Code (e.g., "NUR 220")
        const courseCode = sheetData.course.code;
        const subfolders = mainFolder.getFoldersByName(courseCode);
        
        // Resolve Folder for Docs
        let courseFolder = null;
        if (subfolders.hasNext()) {
            courseFolder = subfolders.next();
        }

        sheetData.exams.forEach(exam => {
            // A. Check for Document
            if (courseFolder) {
                // Construct filename: "NUR 220 Concepts - Exam 1"
                const docName = `${sheetData.course.code} ${sheetData.course.name} - ${exam.name}`;
                const files = courseFolder.getFilesByName(docName);
                if (files.hasNext()) {
                    exam.docUrl = files.next().getUrl();
                }
            }

            // B. Check for Calendar Event (Visual Indication)
            exam.isOnCalendar = false;
            if (calendar) {
                exam.isOnCalendar = checkExamOnCalendar(calendar, sheetData.course, exam);
            }
        });
    });

    return { success: true, data: { sheets: allSheetData, settings: settings } };

  } catch (e) {
    console.error(`Error in api_getNursingData: ${e.toString()}`);
    return { success: false, message: e.message };
  }
}

function api_saveNursingAccommodations(sheetName, examName, accommodationsText) {
  try {
    const settings = getNursingSettings();
    const spreadsheet = SpreadsheetApp.openById(settings.nursingSheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet '${sheetName}' not found.`);
    
    const data = sheet.getDataRange().getValues();
    
    let headerRowIndex = -1;
    for (let i = 0; i < 15; i++) {
        const rowStr = data[i].join(' ').toLowerCase();
        if (rowStr.includes('exam') && rowStr.includes('date')) {
            headerRowIndex = i;
            break;
        }
    }
    if (headerRowIndex === -1) throw new Error("Could not locate data table in sheet.");

    const headers = data[headerRowIndex].map(h => String(h).trim().toLowerCase());
    const nameCol = headers.findIndex(h => h.includes('exam'));
    const accommCol = headers.findIndex(h => h.includes('accommodations'));

    if (nameCol === -1) throw new Error("Exam Name column not found.");
    if (accommCol === -1) throw new Error("Accommodations column not found.");

    for (let i = headerRowIndex + 1; i < data.length; i++) {
      if (String(data[i][nameCol]).trim() === examName) {
        sheet.getRange(i + 1, accommCol + 1).setValue(accommodationsText);
        return { success: true, message: 'Accommodations saved!' }; 
      }
    }
    throw new Error(`Exam '${examName}' not found in sheet '${sheetName}'.`);
  } catch (e) {
    return { success: false, message: e.message }; 
  }
}

/**
 * Syncs nursing exams to a Google Calendar
 * @param {Object} data - The payload containing settings and exam data
 */
function api_syncNursingCalendar(data) {
  try {
    const settings = data.settings || getNursingSettings();
    if (!settings.nursingCalendarId) throw new Error("No Target Calendar ID found in settings.");
    
    const calendar = CalendarApp.getCalendarById(settings.nursingCalendarId);
    if (!calendar) throw new Error("Could not access the specified Calendar. Check the ID and permissions.");

    let count = 0;

    data.sheets.forEach(sheet => {
      sheet.exams.forEach(exam => {
        const eventCreated = createOrUpdateExamEvent(calendar, sheet.course, exam);
        if (eventCreated) count++;
        
        // Prevent API rate limiting errors
        Utilities.sleep(1500);
      });
    });

    return { success: true, message: `Successfully synced ${count} exam(s) to the calendar.`, count: count };
  } catch (e) {
    console.error(`Error in api_syncNursingCalendar: ${e.toString()}`);
    return { success: false, message: e.message };
  }
}

/**
 * Helper to check if an event already exists (read-only)
 */
function checkExamOnCalendar(calendar, course, exam) {
    if (!exam.date || !exam.siteTime) return false;

    const startDateTime = new Date(`${exam.date} ${exam.siteTime}`);
    if (isNaN(startDateTime.getTime())) return false;

    // Calculate End Time (Default to 2 hours if duration is missing)
    let durationMinutes = 120; 
    if (exam.duration) {
        const durStr = String(exam.duration); // Force string
        const match = durStr.match(/\d+/);
        if (match) {
            const num = parseInt(match[0]);
            if (!isNaN(num)) {
                durationMinutes = durStr.toLowerCase().includes('hour') ? num * 60 : num;
            }
        }
    }
    const endDateTime = new Date(startDateTime.getTime() + durationMinutes * 60000);
    const eventTitle = `${course.code}: ${exam.name}`;

    // Search window (+/- 1 hour)
    const existingEvents = calendar.getEvents(
        new Date(startDateTime.getTime() - 3600000), 
        new Date(endDateTime.getTime() + 3600000)
    );

    return existingEvents.some(e => e.getTitle() === eventTitle);
}

/**
 * Helper to create or update a single calendar event
 */
function createOrUpdateExamEvent(calendar, course, exam) {
  if (!exam.date || !exam.siteTime) return false;

  // 1. Parse Start Time
  // Assumes exam.date is "YYYY-MM-DD" and siteTime is "HH:mm AM/PM"
  const startDateTime = new Date(`${exam.date} ${exam.siteTime}`);
  if (isNaN(startDateTime.getTime())) return false;

  // 2. Calculate End Time (Default to 2 hours if duration is missing)
  let durationMinutes = 120; 
  if (exam.duration) {
    const durStr = String(exam.duration); // Force string
    const match = durStr.match(/\d+/);
    if (match) {
        const num = parseInt(match[0]);
        if (!isNaN(num)) {
            durationMinutes = durStr.toLowerCase().includes('hour') ? num * 60 : num;
        }
    }
  }
  const endDateTime = new Date(startDateTime.getTime() + durationMinutes * 60000);

  // 3. Construct Event Details
  const eventTitle = `${course.code}: ${exam.name}`;
  const description = `Course: ${course.name}\nPassword: ${exam.password || 'N/A'}\nZoom Time: ${exam.zoomTime || 'N/A'}\nAccommodations: ${exam.accommodations || 'None'}`;
  const location = exam.room || "Nursing Dept";

  // 4. Duplicate Prevention
  // Search for existing events with the same title on that specific day (+/- 1 hour window)
  const existingEvents = calendar.getEvents(
    new Date(startDateTime.getTime() - 3600000), 
    new Date(endDateTime.getTime() + 3600000)
  );

  const duplicate = existingEvents.find(e => e.getTitle() === eventTitle);

  if (duplicate) {
    // Update existing
    duplicate.setTime(startDateTime, endDateTime);
    duplicate.setDescription(description);
    duplicate.setLocation(location);
  } else {
    // Create new
    calendar.createEvent(eventTitle, startDateTime, endDateTime, {
      description: description,
      location: location
    });
  }
  return true;
}

/**
 * ----------------------------------------------------------------------------------------
 * Document Creation & Updating Logic
 * ----------------------------------------------------------------------------------------
 */

function createNursingProctoringDocuments(data) {
  return genericDocAction(data, false);
}

function updateAllNursingDocuments(data) {
  return genericDocAction(data, true);
}

function genericDocAction(data, isUpdateOnly) {
  try {
    const settings = data.settings || getNursingSettings();
    const mainFolder = DriveApp.getFolderById(settings.nursingFolderId);
    let count = 0;
    
    data.sheets.forEach(sheet => {
      const courseCode = sheet.course.code;
      let courseFolder = null;

      // 1. Resolve Subfolder
      const subfolders = mainFolder.getFoldersByName(courseCode);
      if (subfolders.hasNext()) {
          courseFolder = subfolders.next();
      } else {
          // If folder doesn't exist:
          if (isUpdateOnly) {
              return; // Skip if updating, as no docs can exist without a folder
          } else {
              // Create folder if generating
              courseFolder = mainFolder.createFolder(courseCode);
          }
      }

      // 2. Process Exams in Subfolder
      sheet.exams.forEach(exam => {
        const docName = `${sheet.course.code} ${sheet.course.name} - ${exam.name}`;
        
        const files = courseFolder.getFilesByName(docName);
        let doc = null;

        if (files.hasNext()) {
            doc = DocumentApp.openById(files.next().getId());
        }

        if (doc) {
          // Update existing
          updateDocContent(doc, exam, sheet.course, settings);
          count++;
        } else if (!isUpdateOnly) {
          // Create new
          const newDoc = DocumentApp.create(docName);
          const file = DriveApp.getFileById(newDoc.getId());
          
          // Move to subfolder
          courseFolder.addFile(file);
          DriveApp.getRootFolder().removeFile(file); 
          
          updateDocContent(newDoc, exam, sheet.course, settings);
          count++;
        }
      });
    });
    return { success: true, message: `${count} document(s) processed.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateDocContent(doc, exam, course, settings) {
    const body = doc.getBody().clear();
    const FONT_FAMILY = 'Calibri';
    const FONT_SIZE = 11;
    
    const docAttributes = {};
    docAttributes[DocumentApp.Attribute.FONT_FAMILY] = FONT_FAMILY;
    docAttributes[DocumentApp.Attribute.FONT_SIZE] = FONT_SIZE;
    body.setAttributes(docAttributes);
    body.setMarginTop(72).setMarginBottom(72).setMarginLeft(72).setMarginRight(72);

    // --- Title ---
    const titleText = `${course.code} ${course.name} - ${exam.name}`;
    const title = body.appendParagraph(titleText);
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1)
         .setBold(true)
         .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    body.appendParagraph(''); 

    // --- Exam Details Section ---
    body.appendParagraph('Exam Details').setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Extract Faculty Name
    let facultyName = course.name;
    const profMatch = course.name.match(/Professor\s+(.+)/i);
    if (profMatch) facultyName = profMatch[1];
    else {
        const colonSplit = course.name.split(':');
        if (colonSplit.length > 1) facultyName = colonSplit[1].trim();
    }

    // Format Date with Ordinals
    let dateStr = exam.date;
    if (dateStr) { 
        const parts = dateStr.split('-');
        if (parts.length === 3) {
            const year = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1; 
            const day = parseInt(parts[2]);
            const d = new Date(year, month, day);
            
            const monthNames = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"
            ];
            
            const getOrdinal = (n) => {
                const s = ["th", "st", "nd", "rd"];
                const v = n % 100;
                return n + (s[(v - 20) % 10] || s[v] || s[0]);
            };

            dateStr = `${monthNames[month]} ${getOrdinal(day)}, ${year}`;
        }
    } else {
        dateStr = 'N/A';
    }

    // Build the list
    body.appendParagraph(`1. Faculty: ${facultyName}`);
    body.appendParagraph(`2. Date: ${dateStr}`);
    body.appendParagraph(`3. Start Time (On Site): ${exam.siteTime || 'N/A'}`);
    body.appendParagraph(`4. Start Time (Zoom): ${exam.zoomTime || 'N/A'}`);
    body.appendParagraph(`5. Duration: ${exam.duration || 'N/A'}`);
    // Room removed as requested
    body.appendParagraph(`6. Password: ${exam.password || 'N/A'}`);

    body.appendParagraph(''); 

    // --- Important Links ---
    body.appendParagraph('Important Links').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Red Flag Reporting Form');
    body.appendParagraph('Nursing Protocol');
    
    body.appendParagraph(''); 

    // --- Location Rosters ---
    body.appendParagraph('Location Rosters').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // Define the standard order of locations
    const locationOrder = [
        'Augusta', 'UMAAL', 'UMF Testing Ctr', 'Bangor', 'Brunswick', 
        'East Millinocket', 'Ellsworth', 'Lewiston', 'Rockland', 'Rumford', 'Saco'
    ];

    locationOrder.forEach(locName => {
        // Create Header for every location (even if empty)
        const header = body.appendParagraph(locName);
        header.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        
        // Check if we have students
        const students = (exam.rosters && exam.rosters[locName]) ? exam.rosters[locName] : [];
        
        if (students.length > 0) {
            students.forEach(student => {
                body.appendListItem(student).setGlyphType(DocumentApp.GlyphType.BULLET);
            });
        } else {
            // Placeholder for empty locations
            body.appendParagraph("(No students assigned)").setAttributes({
                [DocumentApp.Attribute.ITALIC]: true,
                [DocumentApp.Attribute.FOREGROUND_COLOR]: '#666666'
            });
        }
    });

    // --- Accommodations ---
    if (exam.accommodations) {
        body.appendParagraph('');
        body.appendParagraph('Accommodations').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(exam.accommodations);
    }

    doc.saveAndClose();
}

/**
 * ----------------------------------------------------------------------------------------
 * Helper Functions
 * ----------------------------------------------------------------------------------------
 */

function getNursingSettings() {
  const props = PropertiesService.getUserProperties().getProperty('nursing_settings');
  if (!props) throw new Error("Nursing settings not found. Please save your settings first.");
  const settings = JSON.parse(props);
  
  settings.nursingSheetId = extractIdFromUrl(settings.nursingSheetId);
  settings.nursingFolderId = extractIdFromUrl(settings.nursingFolderId);
  
  if (!settings.nursingSheetId || !settings.nursingFolderId) {
    throw new Error("Nursing Sheet URL or Folder URL is missing or invalid in settings.");
  }
  return settings;
}

function extractIdFromUrl(url) {
    if (!url) return null;
    const folderMatch = url.match(/\/folders\/([a-zA-Z0-9-_]+)/);
    if (folderMatch) return folderMatch[1];
    const docMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (docMatch) return docMatch[1];
    const idParamMatch = url.match(/id=([a-zA-Z0-9-_]+)/);
    if (idParamMatch) return idParamMatch[1];
    return url; 
}

function parseFlexibleDate(input) {
    if (!input) return null;
    if (input instanceof Date) {
        return Utilities.formatDate(input, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    let str = String(input).trim();
    str = str.replace(/(\d+)(st|nd|rd|th)/ig, "$1");
    const d = new Date(str);
    if (isNaN(d.getTime())) return null;
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Parses the sheet using a Two-Zone strategy.
 * Zone A: Exam Table (Top)
 * Zone B: Roster Table (Bottom)
 */
function parseNursingSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 5) return null; 

  // --- Parse Course Info (A1) ---
  const a1 = String(data[0][0]).trim(); 
  let courseCode = "Unknown";
  let courseName = a1;

  const splitMatch = a1.match(/^([^:-]+)[:\s-](.+)/);
  if (splitMatch) {
      courseCode = splitMatch[1].trim();
      courseName = splitMatch[2].trim();
  } else {
      const parts = a1.split(' ');
      if (parts.length > 1) {
          courseCode = parts.slice(0, 2).join(' ');
          courseName = parts.slice(2).join(' ');
      }
  }

  // --- ZONE A: Find Exam Table ---
  let examHeaderRowIndex = -1;
  for (let i = 0; i < 20; i++) { 
      if (!data[i]) continue;
      const rowStr = data[i].join(' ').toLowerCase();
      if (rowStr.includes('exam') && rowStr.includes('date')) {
          examHeaderRowIndex = i;
          break;
      }
  }

  if (examHeaderRowIndex === -1) return null;

  // Map Exam Columns
  const examHeaders = data[examHeaderRowIndex].map(h => String(h).trim().toLowerCase());
  const colMap = {
      name: examHeaders.findIndex(h => h.includes('exam')),
      date: examHeaders.findIndex(h => h === 'date'), 
      timeSite: examHeaders.findIndex(h => h.includes('time') && !h.includes('zoom')),
      timeZoom: examHeaders.findIndex(h => h.includes('time') && h.includes('zoom')),
      duration: examHeaders.findIndex(h => h.includes('duration')),
      room: examHeaders.findIndex(h => h.includes('room') || h.includes('location')),
      password: examHeaders.findIndex(h => h.includes('password')),
      accommodations: examHeaders.findIndex(h => h.includes('accommodations'))
  };

  if (colMap.name === -1 || colMap.date === -1) return null;

  // --- ZONE B: Find Roster Table ---
  // We look for a row containing ANY of our specific location names
  const locationKeywords = [
      'Augusta', 'UMAAL', 'UMF Testing Ctr', 'Bangor', 'Brunswick', 
      'East Millinocket', 'Ellsworth', 'Lewiston', 'Rockland', 'Rumford', 'Saco'
  ];
  
  let rosterHeaderRowIndex = -1;
  
  // Start searching *after* the exam header
  for (let i = examHeaderRowIndex + 1; i < data.length; i++) {
      const rowValues = data[i].map(v => String(v).trim());
      // Check if this row contains at least one known location header
      const match = rowValues.some(v => locationKeywords.includes(v));
      
      if (match) {
          rosterHeaderRowIndex = i;
          break;
      }
  }

  // Parse Rosters if found
  let rosters = {};
  if (rosterHeaderRowIndex > -1) {
      rosters = parseRosterData(data, rosterHeaderRowIndex);
  }

  // --- Extract Exams ---
  const exams = [];
  const safeNormalize = (typeof normalizeTime === 'function') ? normalizeTime : String;

  // Iterate from Exam Header down to Roster Header (or end of sheet)
  const endRow = (rosterHeaderRowIndex > -1) ? rosterHeaderRowIndex : data.length;

  for (let i = examHeaderRowIndex + 1; i < endRow; i++) {
      const row = data[i];
      const examName = row[colMap.name];

      // Stop if empty or looks like a new header
      if (!examName || String(examName).trim() === '') continue;
      
      // Safety check: if we accidentally hit the roster header
      if (locationKeywords.some(kw => String(examName).includes(kw))) break;

      const dateVal = parseFlexibleDate(row[colMap.date]);

      exams.push({
          name: String(examName).trim(),
          date: dateVal, 
          siteTime: safeNormalize(row[colMap.timeSite]), 
          zoomTime: safeNormalize(row[colMap.timeZoom]), 
          duration: colMap.duration > -1 ? row[colMap.duration] : '',
          room: colMap.room > -1 ? row[colMap.room] : '',
          password: colMap.password > -1 ? row[colMap.password] : '',
          accommodations: colMap.accommodations > -1 ? row[colMap.accommodations] : '',
          rosters: rosters // Attach the parsed rosters to every exam
      });
  }

  return {
      course: { code: courseCode, name: courseName },
      exams: exams
  };
}

/**
 * Helper to parse the Roster Zone.
 * Hard-coded to look for the "Students with accommodations..." text block.
 */
function parseRosterData(data, headerRowIndex) {
    const headers = data[headerRowIndex].map(h => String(h).trim());
    const rosters = {};
    
    // 1. Determine where the actual data starts
    // We expect: Header -> Address -> Merged Note -> DATA
    let dataStartIndex = headerRowIndex + 3; // Default fallback (skip 2 rows after header)

    // Scan the next 5 rows to find the specific "Students with accommodations" text
    for (let i = headerRowIndex + 1; i < Math.min(data.length, headerRowIndex + 6); i++) {
        const rowStr = data[i].join(' ').toLowerCase();
        if (rowStr.includes('students with accommodations')) {
            dataStartIndex = i + 1; // Data starts immediately after this row
            break;
        }
    }

    const targetLocations = [
        'Augusta', 'UMAAL', 'UMF Testing Ctr', 'Bangor', 'Brunswick', 
        'East Millinocket', 'Ellsworth', 'Lewiston', 'Rockland', 'Rumford', 'Saco'
    ];

    targetLocations.forEach(loc => {
        const colIndex = headers.findIndex(h => h === loc); // Exact match preferred
        
        if (colIndex > -1) {
            const students = [];
            
            // Read down from the calculated start index
            for (let i = dataStartIndex; i < data.length; i++) {
                const cell = String(data[i][colIndex]).trim();
                
                if (!cell) continue; // Skip empty
                if (cell.length < 2) continue; // Skip junk

                students.push(cell);
            }
            rosters[loc] = students;
        } else {
            rosters[loc] = [];
        }
    });

    return rosters;
}