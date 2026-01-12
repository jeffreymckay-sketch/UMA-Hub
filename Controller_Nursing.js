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

    // 2. Find Document URLs (Subfolder Logic)
    const mainFolder = DriveApp.getFolderById(settings.nursingFolderId);
    
    allSheetData.forEach(sheetData => {
        // Look for the subfolder matching the Course Code (e.g., "NUR 220")
        const courseCode = sheetData.course.code;
        const subfolders = mainFolder.getFoldersByName(courseCode);
        
        if (subfolders.hasNext()) {
            const courseFolder = subfolders.next();
            
            sheetData.exams.forEach(exam => {
                // Construct filename: "NUR 220 Concepts - Exam 1"
                const docName = `${sheetData.course.code} ${sheetData.course.name} - ${exam.name}`;
                
                // Look for file inside the course subfolder
                const files = courseFolder.getFilesByName(docName);
                if (files.hasNext()) {
                    exam.docUrl = files.next().getUrl();
                }
            });
        }
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

    // Format Date with Ordinals (e.g., February 24th, 2026)
    let dateStr = exam.date;
    if (dateStr) { 
        // We expect exam.date to be yyyy-MM-dd from our parser
        // We need to parse it back to a Date object to format it nicely
        // Note: new Date("yyyy-MM-dd") treats it as UTC, which can shift the day.
        // We split manually to avoid timezone shifts.
        const parts = dateStr.split('-');
        if (parts.length === 3) {
            const year = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1; // JS months are 0-11
            const day = parseInt(parts[2]);
            const d = new Date(year, month, day);
            
            const monthNames = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"
            ];
            
            // Ordinal suffix logic
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
    body.appendParagraph(`6. Room: ${exam.room || 'N/A'}`);
    body.appendParagraph(`7. Password: ${exam.password || 'N/A'}`);

    body.appendParagraph(''); 

    // --- Important Links ---
    body.appendParagraph('Important Links').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Red Flag Reporting Form');
    body.appendParagraph('Nursing Protocol');
    
    body.appendParagraph(''); 

    // --- Location Rosters ---
    body.appendParagraph('Location Rosters').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const locations = [
        'UMAAL', 'Augusta', 'Bangor', 'Brunswick', 'East Millinocket',
        'Ellsworth', 'Lewiston', 'Rockland', 'Rumford', 'Saco', 'UMF Testing Ctr'
    ];
    locations.forEach(loc => body.appendParagraph(loc));

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

/**
 * Parses a date input flexibly.
 * Handles: Date objects, "2/15/26", "February 15th, 2026", etc.
 * Returns: "yyyy-MM-dd" string or null.
 */
function parseFlexibleDate(input) {
    if (!input) return null;
    
    // 1. Already a Date object
    if (input instanceof Date) {
        return Utilities.formatDate(input, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    let str = String(input).trim();
    
    // 2. Remove ordinal suffixes (st, nd, rd, th) to make it parsable by JS
    // Regex: look for digits followed immediately by st/nd/rd/th, case insensitive
    str = str.replace(/(\d+)(st|nd|rd|th)/ig, "$1");

    // 3. Try parsing
    const d = new Date(str);
    if (isNaN(d.getTime())) {
        return null; // Failed to parse
    }

    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseNursingSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 5) return null; 

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

  let headerRowIndex = -1;
  for (let i = 0; i < 15; i++) { 
      if (!data[i]) continue;
      const rowStr = data[i].join(' ').toLowerCase();
      if (rowStr.includes('exam') && rowStr.includes('date')) {
          headerRowIndex = i;
          break;
      }
  }

  if (headerRowIndex === -1) return null;

  const headers = data[headerRowIndex].map(h => String(h).trim().toLowerCase());
  
  const colMap = {
      name: headers.findIndex(h => h.includes('exam')),
      date: headers.findIndex(h => h === 'date'), 
      timeSite: headers.findIndex(h => h.includes('time') && !h.includes('zoom')),
      timeZoom: headers.findIndex(h => h.includes('time') && h.includes('zoom')),
      duration: headers.findIndex(h => h.includes('duration')),
      room: headers.findIndex(h => h.includes('room') || h.includes('location')),
      password: headers.findIndex(h => h.includes('password')),
      accommodations: headers.findIndex(h => h.includes('accommodations'))
  };

  if (colMap.name === -1 || colMap.date === -1) return null;

  const exams = [];
  for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      const examName = row[colMap.name];

      if (!examName || String(examName).trim() === '') break;
      if (String(examName).toLowerCase().includes('exam') && String(row[colMap.date]).toLowerCase().includes('date')) break;

      const safeNormalize = (typeof normalizeTime === 'function') ? normalizeTime : String;

      // Use new flexible date parser
      const dateVal = parseFlexibleDate(row[colMap.date]);

      exams.push({
          name: String(examName).trim(),
          date: dateVal, // Now consistently yyyy-MM-dd or null
          siteTime: safeNormalize(row[colMap.timeSite]), 
          zoomTime: safeNormalize(row[colMap.timeZoom]), 
          duration: colMap.duration > -1 ? row[colMap.duration] : '',
          room: colMap.room > -1 ? row[colMap.room] : '',
          password: colMap.password > -1 ? row[colMap.password] : '',
          accommodations: colMap.accommodations > -1 ? row[colMap.accommodations] : ''
      });
  }

  return {
      course: { code: courseCode, name: courseName },
      exams: exams
  };
}