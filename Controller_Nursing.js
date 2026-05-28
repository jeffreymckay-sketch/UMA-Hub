function _getNursingSettings() {
  const allProps = getSettings();
  let raw = allProps.nursing_settings;
  let settings = {};
  if (raw && typeof raw === 'string') {
    try { settings = JSON.parse(raw);
    } catch (e) { console.error("JSON Parse Error", e); }
  } else if (typeof raw === 'object') {
    settings = raw;
  }
  const extractId = (input) => {
    if (!input) return null;
    const match = input.match(/[-\w]{25,}/);
    return match ? match[0] : input;
  };
  return {
    sheetId: extractId(settings.nursingSheetId),
    folderId: extractId(settings.nursingFolderId),
    calendarId: settings.nursingCalendarId, 
    customNotes: settings.customNotes
  };
}

function _normalizeStr(str) {
  if (!str) return "";
  return str.toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function _toTitleCase(str) {
  if (!str) return "";
  return str.replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

/**
 * Parses Cell A1 to separate Course info from Faculty info.
 * Input Format: "Course Code Course Name : Faculty Name"
 */
function _parseCellA1(cellValue) {
  if (!cellValue) return { 
      folderKey: "Unknown", 
      courseTitle: "Unknown", 
      facultyName: "Faculty Unassigned", 
      fullTitlePrefix: "Unknown : Faculty Unassigned" 
  };

  const str = String(cellValue).trim();
  
  // We assume a colon separator based on requirements.
  // If missing, we fallback to hyphen or treat whole string as course.
  let separator = ':';
  if (!str.includes(':')) separator = '-';

  const parts = str.split(separator);
  
  // Part 1: Course (Everything before the first separator)
  const courseInfo = parts[0].trim();
  
  // Part 2: Faculty (Everything after the first separator)
  let facultyInfo = "";
  if (parts.length > 1) {
      facultyInfo = parts.slice(1).join(separator).trim();
  }

  // Fallback if faculty name is empty
  if (!facultyInfo) facultyInfo = "Faculty Unassigned";

  // Reconstruct the canonical Title Prefix for filenames
  // Format: "Course : Faculty"
  const fullTitlePrefix = `${courseInfo} : ${facultyInfo}`;

  return { 
      folderKey: courseInfo,        // Used to find/create the folder
      courseTitle: courseInfo,      // Used for Document Title inside the doc
      facultyName: facultyInfo,     // Used for Heading 1 inside the doc
      fullTitlePrefix: fullTitlePrefix // Used for the File Name in Drive
  };
}

function _findSubFolder(parentFolder, targetName) {
  const targetNorm = _normalizeStr(targetName);
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (_normalizeStr(folder.getName()) === targetNorm) return folder;
  }
  return null;
}

function formatDateToPlainLanguage(input) {
  if (!input) return "";
  let dateObj;
  if (input instanceof Date) {
    dateObj = input;
  } else {
    let str = String(input).replace(/(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|st|nd|rd|th),?/gi, "").trim();
    dateObj = new Date(str);
  }
  if (isNaN(dateObj.getTime())) return String(input); 
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMMM d, yyyy");
}

function api_getNursingData() {
  try {
    const config = _getNursingSettings();
    if (!config.sheetId || !config.folderId) {
      return { success: false, message: "Nursing settings missing. Please save Sheet/Folder IDs in Settings."
      };
    }
    const ss = SpreadsheetApp.openById(config.sheetId);
    const sheets = ss.getSheets();
    const rootFolder = DriveApp.getFolderById(config.folderId);
    const dbMap = getAccommodationsDBMap();
    const courseData =[];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("_")) return; 
      // NEW: Safety Switch Toggle
      const isBetaEnabled = getSettings().enableBetaNursing === 'true' || getSettings().enableBetaNursing === true;
      let parsed = null;
      if (isBetaEnabled) {
          parsed = parseNursingSheet_V2(sheet, dbMap);
      } else {
          parsed = parseNursingSheet(sheet, dbMap);
      }
      if (!parsed) return;
      
      // Meta now contains the split variables
      const meta = _parseCellA1(parsed.rawA1);
      
      let targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      const existingFiles =[];
      if (targetFolder) {
        const files = targetFolder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          existingFiles.push({ name: f.getName(), nameNorm: _normalizeStr(f.getName()), url: f.getUrl() });
        }
      }
      
      parsed.exams.forEach(exam => {
        const examNameNorm = _normalizeStr(exam.name);
        const expectedTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        const expectedTitleNorm = _normalizeStr(expectedTitle);
        
        let match = existingFiles.find(f => f.nameNorm === expectedTitleNorm);
        if (!match) match = existingFiles.find(f => f.nameNorm.includes(examNameNorm));
        
        exam.docUrl = match ? match.url : null;
      });
      
      if (parsed.exams.length > 0) {
        courseData.push({ sheetName: sheetName, course: parsed.course, exams: parsed.exams, _meta: meta });
      }
    });
    return { success: true, data: { sheets: courseData, settings: config } };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * V2 Parser: Uses Utility_SheetReader.js to extract Nursing sheet data heuristically.
 * Outputs JSON exactly matching the shape of the original parseNursingSheet function.
 */
function parseNursingSheet_V2(sheet, dbMap) {
  const range = sheet.getDataRange();
  const data = range.getValues();
  const fontColors = range.getFontColors();
  const fontLines = range.getFontLines();
  
  if (!data || data.length < 5) return null;
  
  const a1 = String(data[0][0]).trim();
  
  let courseCode = "Unknown";
  let courseName = a1;
  const splitMatch = a1.match(/^([^:-]+)[:\s-](.+)/);
  if (splitMatch) {
      courseCode = splitMatch[1].trim();
      courseName = splitMatch[2].trim();
  }

  const headerKeywords = ['exam', 'date'];
  const headerRowIndex = SheetReader.findHeaderRowHeuristic(data, headerKeywords, 15);
  
  if (headerRowIndex === -1) return null;
  
  const synonymConfig = {
      name: ['exam'],
      date: ['date'],
      timeSite: ['time', 'onsite time', 'on site time', 'site time'],
      timeZoom: ['zoom time', 'zoom'],
      duration: ['duration'],
      room: ['room', 'location'],
      password: ['password'],
      accommodations: ['accommodations', 'notes']
  };
  
  const colMap = SheetReader.mapColumnsBySynonyms(data[headerRowIndex], synonymConfig);
  
  const headers = data[headerRowIndex].map(h => String(h).trim().toLowerCase());
  colMap.timeSite = headers.findIndex(h => h.includes('time') && !h.includes('zoom'));
  colMap.timeZoom = headers.findIndex(h => h.includes('time') && h.includes('zoom'));
  colMap.name = headers.findIndex(h => h.includes('exam'));

  if (colMap.name === -1 || colMap.date === -1) return null;

  const rosterHeaderIdx = SheetReader.findAnchorRow(data, ['augusta'], headerRowIndex + 1);
  const scanEnd = (rosterHeaderIdx !== -1) ? rosterHeaderIdx : data.length;

  const sheetRoster = {};
  if (rosterHeaderIdx !== -1) {
      const locHeaders = data[rosterHeaderIdx];
      locHeaders.forEach(h => {
          const locName = String(h).trim();
          if (locName && !sheetRoster[locName]) sheetRoster[locName] = [];
      });
      const rosterStartRow = rosterHeaderIdx + 3;
      for (let r = rosterStartRow; r < data.length; r++) {
          for (let c = 0; c < locHeaders.length; c++) {
              const loc = String(locHeaders[c]).trim();
              const name = String(data[r][c]).trim();
              const color = fontColors[r][c];
              if (loc && name) {
                  if (!sheetRoster[loc]) sheetRoster[loc] = [];
                  sheetRoster[loc].push({ name: name, color: color });
              }
          }
      }
  }

  const exams = [];
  const rawExams = SheetReader.readDynamicRoster(data.slice(0, scanEnd), headerRowIndex + 1, colMap.name, ['total'], 3);
  
  rawExams.forEach(parsedRow => {
      const rIndex = parsedRow.rowIndex;
      const rowData = parsedRow.data;
      const examName = parsedRow.name;
      
      let isDone = false;
      if (fontLines[rIndex][colMap.date] === 'line-through' || fontLines[rIndex][colMap.name] === 'line-through') {
          isDone = true;
      }
      
      const rawDate = rowData[colMap.date];
      let dateObj;
      if (rawDate instanceof Date) { dateObj = rawDate; } 
      else if (rawDate && rawDate !== 'TBD') {
          const str = String(rawDate).replace(/(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|st|nd|rd|th),?/gi, "").trim();
          dateObj = new Date(str);
      }
      
      if (dateObj && !isNaN(dateObj.getTime())) {
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          if (dateObj < today) isDone = true;
      }
      
      const dateVal = formatDateToPlainLanguage(rawDate !== 'TBD' ? rawDate : "");
      
      const siteTimeStr = (colMap.timeSite > -1 && rowData[colMap.timeSite] !== 'TBD') ? rowData[colMap.timeSite] : '';
      const zoomTimeStr = (colMap.timeZoom > -1 && rowData[colMap.timeZoom] !== 'TBD') ? rowData[colMap.timeZoom] : '';
      
      const siteTime = normalizeTime(siteTimeStr);
      const zoomTime = normalizeTime(zoomTimeStr);
      
      const dbKey = `${courseCode}|${examName}`;
      const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };
      
      const durationVal = (colMap.duration > -1 && rowData[colMap.duration] !== 'TBD') ? rowData[colMap.duration] : '';
      const roomVal = (colMap.room > -1 && rowData[colMap.room] !== 'TBD') ? rowData[colMap.room] : '';
      const passVal = (colMap.password > -1 && rowData[colMap.password] !== 'TBD') ? rowData[colMap.password] : '';
      const accVal = (colMap.accommodations > -1 && rowData[colMap.accommodations] !== 'TBD') ? rowData[colMap.accommodations] : '';

      exams.push({
          name: examName,
          date: dateVal,
          siteTime: siteTime,
          zoomTime: zoomTime,
          duration: durationVal,
          room: roomVal,
          password: passVal,
          generalNotes: dbEntry.generalNotes || accVal,
          studentTags: dbEntry.studentTags,
          rosters: sheetRoster,
          isDone: isDone
      });
  });

  return { rawA1: a1, course: { code: courseCode, name: courseName }, exams: exams };
}

function parseNursingSheet(sheet, dbMap) {
  const range = sheet.getDataRange();
  const data = range.getValues();
  const fontColors = range.getFontColors();
  const fontLines = range.getFontLines();
  if (data.length < 5) return null; 
  const a1 = String(data[0][0]).trim(); 
  
  // Note: This regex split is for the JSON return object, distinct from _parseCellA1
  let courseCode = "Unknown";
  let courseName = a1;
  const splitMatch = a1.match(/^([^:-]+)[:\s-](.+)/);
  if (splitMatch) {
      courseCode = splitMatch[1].trim();
      courseName = splitMatch[2].trim();
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
  let rosterHeaderIdx = -1;
  for(let i = headerRowIndex + 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === 'augusta') {
          rosterHeaderIdx = i;
          break;
      }
  }
  const sheetRoster = {}; 
  if (rosterHeaderIdx !== -1) {
      const locHeaders = data[rosterHeaderIdx];
      locHeaders.forEach(h => {
          const locName = String(h).trim();
          if (locName && !sheetRoster[locName]) sheetRoster[locName] =[];
      });
      const rosterStartRow = rosterHeaderIdx + 3; 
      for (let r = rosterStartRow; r < data.length; r++) {
          for (let c = 0; c < locHeaders.length; c++) {
              const loc = String(locHeaders[c]).trim();
              const name = String(data[r][c]).trim();
              const color = fontColors[r][c];
              if (loc && name) {
                  if (!sheetRoster[loc]) sheetRoster[loc] = [];
                  sheetRoster[loc].push({ name: name, color: color });
              }
          }
      }
  }
  const scanEnd = (rosterHeaderIdx !== -1) ?
  rosterHeaderIdx : data.length;
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
  const exams =[];
  
  for (let i = headerRowIndex + 1; i < scanEnd; i++) {
      const row = data[i];
      const examName = row[colMap.name];
      if (!examName || String(examName).trim() === '') continue;
      if (String(examName).toLowerCase().includes('total')) continue;
      
      let isDone = false;
      
      // 1. Check for manual strikethrough (Standard way users mark things as cancelled/done)
      if (fontLines[i][colMap.date] === 'line-through' || fontLines[i][colMap.name] === 'line-through') {
          isDone = true;
      }
      
      // 2. Automatically flag as done if the date has passed
      const rawDate = row[colMap.date];
      let dateObj;
      if (rawDate instanceof Date) {
          dateObj = rawDate;
      } else {
          const str = String(rawDate).replace(/(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|st|nd|rd|th),?/gi, "").trim();
          dateObj = new Date(str);
      }
      
      if (!isNaN(dateObj.getTime())) {
          const today = new Date();
          today.setHours(0, 0, 0, 0); // Reset to midnight for an accurate day comparison
          if (dateObj < today) {
              isDone = true;
          }
      }
      
      const dateVal = formatDateToPlainLanguage(row[colMap.date]);
      const siteTime = normalizeTime(row[colMap.timeSite]);
      const zoomTime = normalizeTime(row[colMap.timeZoom]);

      const dbKey = `${courseCode}|${String(examName).trim()}`; 
      const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };
      
      exams.push({
          name: String(examName).trim(),
          date: dateVal, 
          siteTime: siteTime, 
          zoomTime: zoomTime, 
          duration: colMap.duration > -1 ? row[colMap.duration] : '',
          room: colMap.room > -1 ? row[colMap.room] : '',
          password: colMap.password > -1 ? row[colMap.password] : '',
          generalNotes: dbEntry.generalNotes || (colMap.accommodations > -1 ? row[colMap.accommodations] : ''),
          studentTags: dbEntry.studentTags, 
          rosters: sheetRoster,
          isDone: isDone // <--- NEW PROPERTY
      });
  }
  return { rawA1: a1, course: { code: courseCode, name: courseName }, exams: exams };
}

function createNursingProctoringDocuments(payload) {
  try {
    const config = _getNursingSettings();
    const rootFolder = DriveApp.getFolderById(config.folderId);
    let createdCount = 0;
    const createdUrls = {}; 

    payload.sheets.forEach(sheetData => {
      const meta = sheetData._meta || _parseCellA1(sheetData.course.name); 
      let targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      if (!targetFolder) targetFolder = rootFolder.createFolder(meta.folderKey);
      
      sheetData.exams.forEach(exam => {
        // Ideal Title: "Course : Faculty - Exam"
        const docTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        let targetFile = null;

        // --- SEARCH STRATEGY ---
        // 1. Try finding by exact Ideal Name
        const filesByName = targetFolder.getFilesByName(docTitle);
        if (filesByName.hasNext()) {
            targetFile = filesByName.next();
        }

        // 2. If not found, fuzzy search in folder
        if (!targetFile) {
            const suffix = ` - ${exam.name}`;
            const files = targetFolder.getFiles();
            while (files.hasNext()) {
                const f = files.next();
                if (f.getName().includes(suffix) || f.getName().trim() === exam.name) {
                    targetFile = f;
                    // Auto-Rename existing file to match new standard
                    if (targetFile.getName() !== docTitle) {
                        targetFile.setName(docTitle);
                    }
                    break;
                }
            }
        }

        if (targetFile) {
            createdUrls[exam.name] = targetFile.getUrl();
            return; // Skip creation
        }

        // Create New
        const doc = DocumentApp.create(docTitle);
        // Pass distinct variables to the population function
        _populateNursingDoc(doc, meta.courseTitle, meta.facultyName, exam, config.customNotes);
        
        const file = DriveApp.getFileById(doc.getId());
        targetFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
        createdUrls[exam.name] = file.getUrl();
        if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Created Doc", docTitle, doc.getId(), `Date: ${exam.date}`);
        createdCount++;
      });
    });
    return { success: true, message: `Created ${createdCount} documents.`, createdUrls: createdUrls };
  } catch (e) { return { success: false, message: e.message };
  }
}

function updateAllNursingDocuments(payload) {
  try {
    const config = _getNursingSettings();
    const rootFolder = DriveApp.getFolderById(config.folderId);
    let updatedCount = 0;

    payload.sheets.forEach(sheetData => {
      const meta = sheetData._meta || _parseCellA1(sheetData.course.name);
      const targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      if (!targetFolder) return; 
      
      sheetData.exams.forEach(exam => {
        const docTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        let targetFile = null;

        // Use URL if available
        if (exam.docUrl) {
          try { targetFile = DriveApp.getFileByUrl(exam.docUrl); } catch(e) {}
        }
        
        // --- SEARCH STRATEGY (If no URL) ---
        if (!targetFile) {
            // 1. Exact Name Search
            const filesByName = targetFolder.getFilesByName(docTitle);
            if (filesByName.hasNext()) {
                targetFile = filesByName.next();
            }

            // 2. Fuzzy Search & Auto-Rename
            if (!targetFile) {
                const suffix = ` - ${exam.name}`;
                const files = targetFolder.getFiles();
                while (files.hasNext()) {
                    const f = files.next();
                    // Check if file contains the exam suffix
                    if (f.getName().includes(suffix) || f.getName().trim() === exam.name) {
                        targetFile = f;
                        // Auto-Correction: Rename file to standard format
                        if (targetFile.getName() !== docTitle) {
                            targetFile.setName(docTitle);
                        }
                        break;
                    }
                }
            }
        }

        if (targetFile) {
          const doc = DocumentApp.openById(targetFile.getId());
          // Pass distinct variables: Course Title, Faculty Name, Exam Object
          _populateNursingDoc(doc, meta.courseTitle, meta.facultyName, exam, config.customNotes);
          
          if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Updated Doc", targetFile.getName(), targetFile.getId(), `Date: ${exam.date}`);
          updatedCount++;
        }
      });
    });
    return { success: true, message: `Updated ${updatedCount} documents.` };
  } catch (e) { return { success: false, message: e.message };
  }
}

// === SURGICAL UPDATE FUNCTIONS ===

function _populateNursingDoc(doc, courseTitle, facultyName, exam, customNotes) {
    const body = doc.getBody();
    
    // --- 1. Header Parsing & Styling ---
    // NO Parsing needed. We use the variables passed directly from the sheet logic.
    
    const displayTitle = `${courseTitle} - ${exam.name}`;
    const displayFaculty = facultyName || "Faculty Unassigned";

    // Update Title (Course - Exam)
    const existingTitle = body.getChild(0);
    if (existingTitle.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const titlePara = existingTitle.asParagraph();
        if (titlePara.getText() !== displayTitle) {
            titlePara.setText(displayTitle);
        }
        titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
    }

    // Update Heading 1 (Faculty)
    let facultyPara = null;
    if (body.getNumChildren() > 1) {
        const child = body.getChild(1);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            facultyPara = child.asParagraph();
        }
    }
    
    if (!facultyPara) {
        facultyPara = body.insertParagraph(1, displayFaculty);
    } else {
        if (facultyPara.getText() !== displayFaculty) {
            facultyPara.setText(displayFaculty);
        }
    }
    facultyPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // --- 2. Exam Details (Heading 2) ---
    const examDetailsMap = {
        'Date': exam.date,
        'Start Time (On Site)': exam.siteTime || "N/A",
        'Start Time (Zoom)': exam.zoomTime || "N/A",
        'Duration': exam.duration || "-",
        'Password': exam.password || "-"
    };
    // Use HEADING2 for section headers
    _updateKeyValueSection(body, "Exam Details", examDetailsMap, DocumentApp.ParagraphHeading.HEADING2);
    
    // --- 3. General Instructions (Heading 2) ---
    if (customNotes) {
        _updateOrCreateTextSection(body, "General Instructions", customNotes, true);
    }

    // --- 4. Important Links (Heading 2) ---
    _updateImportantLinks(body);
    
    // --- 5. Accommodations (Heading 2) ---
    if (exam.generalNotes) {
        _updateOrCreateHighlightedSection(body, "Exam Accommodations", exam.generalNotes);
    } else {
        _removeSectionIfExists(body, "Exam Accommodations");
    }
    
    // --- 6. Roster (Heading 2 for Main, Heading 3 for Locs) ---
    _updateLocationRosters(body, exam);
    
    doc.saveAndClose();
}

// === HELPER FUNCTIONS FOR SURGICAL UPDATES ===

function _findHeadingIndex(body, headingText) {
  const numChildren = body.getNumChildren();
  // Normalize target text for comparison
  const targetNorm = headingText.toLowerCase().replace(/[^a-z0-9]/g, '');
  
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      const text = para.getText();
      // Fuzzy match: check if document header contains our target
      if (text && text.toLowerCase().replace(/[^a-z0-9]/g, '').includes(targetNorm)) {
        // Ensure it's actually a Heading (H1, H2, or H3) to avoid false positives
        const h = para.getHeading();
        if (h !== DocumentApp.ParagraphHeading.NORMAL && h !== DocumentApp.ParagraphHeading.TITLE) {
          return i;
        }
      }
    }
  }
  return -1;
}

function _updateImportantLinks(body) {
  const sectionTitle = "Important Links";
  const links =[
    { text: "Red Flag Reporting Form", url: "https://docs.google.com/forms/d/e/1FAIpQLSfORKCKol8SsRldNKfvsDy3ILNs9HcFv3gKb8TuxrNrlqxijw/viewform" },
    { text: "Nursing Protocol", url: "https://docs.google.com/document/d/1TgKtmoDFqXLK0lBFPNirOAz_TW4S3E_BFhS934VcjOo/edit" }
  ];

  let sectionIndex = _findHeadingIndex(body, sectionTitle);
  
  // If not found, find insertion point
  if (sectionIndex === -1) {
    let insertAt = _findHeadingIndex(body, "Exam Accommodations");
    if (insertAt === -1) insertAt = _findHeadingIndex(body, "Location Rosters");
    if (insertAt === -1) insertAt = body.getNumChildren();
    
    body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    sectionIndex = insertAt;
  } else {
    // Update Heading Level just in case
    body.getChild(sectionIndex).asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // Clear existing content below it
    let scanIndex = sectionIndex + 1;
    while (scanIndex < body.getNumChildren()) {
      const child = body.getChild(scanIndex);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const h = child.asParagraph().getHeading();
        if (h !== DocumentApp.ParagraphHeading.NORMAL) break;
      }
      body.removeChild(child);
    }
  }

  // Write fresh links
  links.forEach((link, i) => {
    const li = body.insertListItem(sectionIndex + 1 + i, link.text);
    li.setLinkUrl(link.url);
    li.setGlyphType(DocumentApp.GlyphType.BULLET);
  });
}

function _updateKeyValueSection(body, sectionTitle, dataMap, headingLevel) {
  let sectionIndex = _findHeadingIndex(body, sectionTitle);
  
  if (sectionIndex === -1) {
    const rostersIndex = _findHeadingIndex(body, "Location Rosters");
    const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
    body.insertParagraph(insertAt, sectionTitle).setHeading(headingLevel);
    sectionIndex = insertAt;
  } else {
    body.getChild(sectionIndex).asParagraph().setHeading(headingLevel);
    
    let scanIndex = sectionIndex + 1;
    while (scanIndex < body.getNumChildren()) {
      const child = body.getChild(scanIndex);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const h = child.asParagraph().getHeading();
        if (h !== DocumentApp.ParagraphHeading.NORMAL) break;
      }
      body.removeChild(child);
    }
  }
  
  let insertCount = 1;
  for (const [key, value] of Object.entries(dataMap)) {
    const text = `${key}: ${value}`;
    const li = body.insertListItem(sectionIndex + insertCount, text);
    li.setGlyphType(DocumentApp.GlyphType.BULLET);

    const keyCheck = key.toLowerCase();
    if (keyCheck.includes("date") || keyCheck.includes("time") || keyCheck.includes("password")) {
      li.setBackgroundColor('#ffff00');
    }
    insertCount++;
  }
}

function _updateOrCreateTextSection(body, sectionTitle, content, isItalic) {
  let sectionIndex = _findHeadingIndex(body, sectionTitle);
  if (sectionIndex === -1) {
    const rostersIndex = _findHeadingIndex(body, "Location Rosters");
    const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
    body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const para = body.insertParagraph(insertAt + 1, content);
    if (isItalic) para.setItalic(true);
  } else {
    body.getChild(sectionIndex).asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    let scanIndex = sectionIndex + 1;
    if (scanIndex < body.getNumChildren()) {
      const nextChild = body.getChild(scanIndex);
      if (nextChild.getType() === DocumentApp.ElementType.PARAGRAPH && 
        nextChild.asParagraph().getHeading() === DocumentApp.ParagraphHeading.NORMAL) {
        nextChild.asParagraph().setText(content);
        if (isItalic) nextChild.asParagraph().setItalic(true);
        scanIndex++;
      } else {
        const para = body.insertParagraph(scanIndex, content);
        if (isItalic) para.setItalic(true);
        scanIndex++;
      }
    } else {
       const para = body.insertParagraph(scanIndex, content);
       if (isItalic) para.setItalic(true);
       scanIndex++;
    }

    while (scanIndex < body.getNumChildren()) {
      const child = body.getChild(scanIndex);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        if (child.asParagraph().getHeading() !== DocumentApp.ParagraphHeading.NORMAL) break;
      }
      body.removeChild(child);
    }
  }
}

function _updateOrCreateHighlightedSection(body, sectionTitle, content) {
  let sectionIndex = _findHeadingIndex(body, sectionTitle);
  if (sectionIndex === -1) {
    const rostersIndex = _findHeadingIndex(body, "Location Rosters");
    const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
    body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const p = body.insertParagraph(insertAt + 1, content);
    p.setBackgroundColor('#e8f5e9');
    p.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(10).setPaddingRight(10);
  } else {
    body.getChild(sectionIndex).asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const nextChild = body.getChild(sectionIndex + 1);
    if (nextChild && nextChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = nextChild.asParagraph();
      if (para.getText() !== content) {
        para.setText(content);
        para.setBackgroundColor('#e8f5e9');
        para.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(10).setPaddingRight(10);
      }
    }
  }
}

function _removeSectionIfExists(body, sectionTitle) {
  const sectionIndex = _findHeadingIndex(body, sectionTitle);
  if (sectionIndex > -1) {
    body.removeChild(body.getChild(sectionIndex));
    if (sectionIndex < body.getNumChildren()) {
      const nextChild = body.getChild(sectionIndex);
      if (nextChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
         if (nextChild.asParagraph().getHeading() === DocumentApp.ParagraphHeading.NORMAL) {
           body.removeChild(nextChild);
         }
      }
    }
  }
}

function _updateLocationRosters(body, exam) {
  if (!exam.rosters || Object.keys(exam.rosters).length === 0) {
    _removeSectionIfExists(body, "Location Rosters");
    return;
  }
  
  let rostersIndex = _findHeadingIndex(body, "Location Rosters");
  if (rostersIndex === -1) {
    rostersIndex = body.getNumChildren();
    body.insertParagraph(rostersIndex, "Location Rosters").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  } else {
    body.getChild(rostersIndex).asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
  }
  
  const sortedLocations = Object.keys(exam.rosters).sort((a, b) => {
    if (a === 'UMAAL') return -1;
    if (b === 'UMAAL') return 1;
    return a.localeCompare(b);
  });
  
  let currentIndex = rostersIndex + 1;
  
  sortedLocations.forEach(location => {
    const students = exam.rosters[location];
    
    const activeStudents = students.filter(s => {
      const tagData = (exam.studentTags && exam.studentTags[s.name]) ? exam.studentTags[s.name] : null;
      return !(tagData && typeof tagData === 'object' && tagData.excluded);
    });
    
    let locationIndex = -1;
    const locNorm = location.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    for (let i = currentIndex; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const para = child.asParagraph();
        const h = para.getHeading();
        const t = para.getText().toLowerCase().replace(/[^a-z0-9]/g, '');
        
        if (t.includes(locNorm) && (h === DocumentApp.ParagraphHeading.HEADING3 || h === DocumentApp.ParagraphHeading.HEADING2)) {
          locationIndex = i;
          break;
        }
        if (h === DocumentApp.ParagraphHeading.HEADING2 || h === DocumentApp.ParagraphHeading.HEADING1) break;
      }
    }
    
    const displayLocation = _toTitleCase(location);

    if (locationIndex === -1) {
      body.insertParagraph(currentIndex, displayLocation).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      locationIndex = currentIndex;
      currentIndex++;
    } else {
      const para = body.getChild(locationIndex).asParagraph();
      para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (para.getText() !== displayLocation) para.setText(displayLocation);
      currentIndex = locationIndex + 1;
    }
    
    const existingLines =[];
    let scanIndex = currentIndex;
    while (scanIndex < body.getNumChildren()) {
      const child = body.getChild(scanIndex);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const para = child.asParagraph();
        if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) break;
      }
      if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
        existingLines.push({
          index: scanIndex,
          listItem: child.asListItem(),
          text: child.asListItem().getText()
        });
      }
      scanIndex++;
    }

    const processedDocIndices = new Set();
    let insertPosition = currentIndex;

    activeStudents.forEach(studentObj => {
      const studentName = studentObj.name;
      let note = "";
      let isHighlighted = false;
      let isLocked = false;
      let color = studentObj.color;
      
      if (exam.studentTags && exam.studentTags[studentName]) {
        const tag = exam.studentTags[studentName];
        if (typeof tag === 'object') {
          note = tag.note || "";
          isHighlighted = tag.highlighted || false;
          isLocked = tag.locked || false;
        }
      }
      if (!color || color === '#000000') color = null;

      const foundLine = existingLines.find(line => {
        return !processedDocIndices.has(line.index) && line.text.includes(studentName);
      });

      if (foundLine) {
        processedDocIndices.add(foundLine.index);
        const li = foundLine.listItem;

        if (!isLocked) {
          const expectedText = note ? `${studentName} [${note}]` : studentName;
          if (li.getText() !== expectedText) {
            li.clear();
            li.setText(studentName);
            if (note) {
              const t = li.appendText(` [${note}]`);
              t.setBold(true).setForegroundColor('#000000');
              if (isHighlighted) t.setBackgroundColor('#ffff00');
              else t.setBackgroundColor('#fff59d');
            }
          }
          if (color) li.setForegroundColor(color);
          else li.setForegroundColor('#000000');
        }
        
        insertPosition = Math.max(insertPosition, foundLine.index + 1);

      } else {
        const li = body.insertListItem(insertPosition, studentName);
        if (note) {
          const t = li.appendText(` [${note}]`);
          t.setBold(true).setForegroundColor('#000000');
          if (isHighlighted) t.setBackgroundColor('#ffff00');
          else t.setBackgroundColor('#fff59d');
        }
        if (color) li.setForegroundColor(color);
        insertPosition++;
      }
    });

    for (let i = existingLines.length - 1; i >= 0; i--) {
      const line = existingLines[i];
      if (!processedDocIndices.has(line.index)) {
        const elementIndex = body.getChildIndex(line.listItem);
        if (elementIndex === body.getNumChildren() - 1) {
          body.appendParagraph(" "); 
        }
        body.removeChild(line.listItem);
        if (line.index < insertPosition) insertPosition--;
      }
    }
    
    currentIndex = insertPosition;
  });
}

function api_syncNursingCalendar(payload) {
  try {
    const config = _getNursingSettings();
    if (!config.calendarId) return { success: false, message: "No Calendar ID." };
    const cal = CalendarApp.getCalendarById(config.calendarId);
    if (!cal) return { success: false, message: "Calendar not found." };
    let count = 0;
    const sheetsToProcess = Array.isArray(payload.sheets) ?
    payload.sheets : [payload.sheets];
    sheetsToProcess.forEach(sheetData => {
      sheetData.exams.forEach(exam => {
        if (!exam.date) return;
        
        let dateObj = new Date(exam.date.replace(/(st|nd|rd|th)/gi, ""));
        if (isNaN(dateObj.getTime())) return;

        const start = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), 8, 0, 0); 
        const end = new Date(start);
        end.setHours(17, 0, 0); 
     
        const titlePrefix = sheetData._meta ? sheetData._meta.fullTitlePrefix : sheetData.sheetName;
        const title = `Proctor: ${titlePrefix} - ${exam.name}`;
        const events = cal.getEvents(start, end);
        const exists = events.some(e => e.getTitle() === title);
        if (!exists) {
          cal.createEvent(title, start, end, {
            description: `Password: ${exam.password}\nZoom: ${exam.zoomTime}\nSite: ${exam.siteTime}\nNotes: ${exam.generalNotes || ''}`
       
          });
          count++;
        }
      });
    });
    if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Calendar Sync", "Batch", config.calendarId, `Synced ${count}`);
    return { success: true, message: `Synced ${count} events.` };
  } catch (e) { return { success: false, message: e.message };
  }
}

function getAccommodationsDBMap() {
    const map = {};
    try {
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName("_DB_ACCOMMODATIONS");
        if (!sheet) return map; 
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const id = String(row[0]); 
            if (!id) continue;
            let tags = {};
            try { tags = JSON.parse(row[4]);
            } catch (e) { } 
            map[id] = { generalNotes: row[3], studentTags: tags };
        }
    } catch (e) { console.warn("DB Error: " + e.message); }
    return map;
}

function api_saveNursingAccommodations(payload) {
  try {
    const ss = getMasterDataHub(); 
    let sheet = ss.getSheetByName("_DB_ACCOMMODATIONS");
    if (!sheet) {
        sheet = ss.insertSheet("_DB_ACCOMMODATIONS");
        sheet.appendRow(["Unique_ID", "Course_Code", "Exam_Name", "General_Notes", "Student_Data"]);
    }
    const uniqueId = `${payload.courseCode}|${payload.examName}`;
    const studentJson = JSON.stringify(payload.studentTags || {});
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === uniqueId) { rowIndex = i + 1;
        break; }
    }
    if (rowIndex > -1) sheet.getRange(rowIndex, 4, 1, 2).setValues([[payload.generalNotes, studentJson]]);
    else sheet.appendRow([uniqueId, payload.courseCode, payload.examName, payload.generalNotes, studentJson]);
    return { success: true, message: 'Saved to Database!' };
  } catch (e) { return { success: false, message: e.message };
  }
}