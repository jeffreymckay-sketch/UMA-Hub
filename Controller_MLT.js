/**
 * -------------------------------------------------------------------
 * MLT PROCTORING CONTROLLER (REFACTORED)
 * Features: API-Driven UI, Deep Folder Search, Database Persistence,
 * Gap-tolerant Roster Scanning (30 rows), D1-based Folders.
 * UPDATED: Tracks "Unassigned" students for Sidebar visibility only.
 * -------------------------------------------------------------------
 */

// --- LOCAL CONFIGURATION DEFAULTS ---
const MLT_CONSTANTS = {
  SETTINGS_KEY: 'MLT_SETTINGS_V2', 
  DEFAULTS: {
    ROSTER_KEYWORD: 'Students',
    KEYWORDS: {
      EXAM: 'Exam',
      DATE: 'Date',
      START_TIME: 'Start Time',
      START_SITE: 'Site',
      DURATION: 'Duration',
      ROOM: 'Room',
      PASSWORD: 'Password'
    }
  }
};

// --- SHARED HELPERS ---

function _mlt_normalizeStr(str) {
  if (!str) return "";
  return str.toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function _mlt_findSubFolder(parentFolder, targetName) {
  const targetNorm = _mlt_normalizeStr(targetName);
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (_mlt_normalizeStr(folder.getName()) === targetNorm) {
      return folder;
    }
  }
  return null;
}

/**
 * 1. GETTER
 */
function _getMLTSettings() {
  let allProps = {};
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(MLT_CONSTANTS.SETTINGS_KEY);
    if (raw) allProps = JSON.parse(raw);
  } catch (e) {
    console.error("Error reading MLT settings:", e);
  }

  if (!allProps.config) allProps.config = MLT_CONSTANTS.DEFAULTS.KEYWORDS;
  if (!allProps.rosterKeyword) allProps.rosterKeyword = MLT_CONSTANTS.DEFAULTS.ROSTER_KEYWORD;

  const extractId = (input) => {
    if (!input) return null;
    const match = input.match(/[-\w]{25,}/); 
    return match ? match[0] : input;
  };

  return {
    sheetId: extractId(allProps.spreadsheetUrl),
    folderId: extractId(allProps.targetFolderId),
    calendarId: allProps.calendarId || "", 
    customNotes: allProps.customNotes || "",
    config: allProps.config,
    rosterKeyword: allProps.rosterKeyword
  };
}

/**
 * 2. SAVER
 */
function mlt_saveSettings(settingsObj) {
    try {
        let current = {};
        const raw = PropertiesService.getScriptProperties().getProperty(MLT_CONSTANTS.SETTINGS_KEY);
        if (raw) { try { current = JSON.parse(raw); } catch(e) {} }

        const newSettings = { ...current, ...settingsObj };
        if (!newSettings.config) newSettings.config = MLT_CONSTANTS.DEFAULTS.KEYWORDS;

        PropertiesService.getScriptProperties().setProperty(MLT_CONSTANTS.SETTINGS_KEY, JSON.stringify(newSettings));
        return { success: true, message: "MLT Settings Saved." };
    } catch (e) { 
        return { success: false, message: "Save Failed: " + e.message }; 
    }
}

/**
 * 3. API
 */
function api_getMLTData() {
  try {
    const settings = _getMLTSettings();
    if (!settings.sheetId || !settings.folderId) {
      return { success: false, message: "MLT settings incomplete. Please check Settings tab." };
    }

    const ss = SpreadsheetApp.openById(settings.sheetId);
    const sheets = ss.getSheets();
    const rootFolder = DriveApp.getFolderById(settings.folderId);
    const dbMap = getAccommodationsDBMap(); 

    const courseData = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("_")) return; 

      const parsed = parseMLTSheet(sheet, settings, dbMap);
      if (!parsed) return;

      const folderKey = parsed.fullTitle; 
      let targetFolder = _mlt_findSubFolder(rootFolder, folderKey);
      
      const existingFiles = [];
      if (targetFolder) {
        const files = targetFolder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          existingFiles.push({
            name: f.getName(),
            nameNorm: _mlt_normalizeStr(f.getName()),
            url: f.getUrl()
          });
        }
      }

      parsed.exams.forEach(exam => {
        const examNameNorm = _mlt_normalizeStr(exam.name);
        const expectedTitle = `${parsed.fullTitle} - ${exam.name}`;
        const expectedTitleNorm = _mlt_normalizeStr(expectedTitle);
        
        let match = existingFiles.find(f => f.nameNorm === expectedTitleNorm);
        if (!match) match = existingFiles.find(f => f.nameNorm.includes(examNameNorm));
        
        exam.docUrl = match ? match.url : null;
      });

      if (parsed.exams.length > 0) {
        courseData.push({
          sheetName: sheetName,
          course: parsed.course, 
          fullTitle: parsed.fullTitle,
          exams: parsed.exams
        });
      }
    });

    return { success: true, data: { sheets: courseData, settings: settings } };

  } catch (e) {
    console.error("MLT API Error", e);
    return { success: false, message: e.message };
  }
}

/**
 * PARSER
 */
function parseMLTSheet(sheet, settings, dbMap) {
  const data = sheet.getDataRange().getValues();
  const fontLines = sheet.getDataRange().getFontLines();
  const fontColors = sheet.getDataRange().getFontColors();

  if (data.length < 5) return null;

  // 1. Course Info
  const d1 = (data[0] && data[0][3]) ? String(data[0][3]).trim() : sheet.getName();
  let courseCode = "MLT";
  let courseName = d1;
  const splitMatch = d1.match(/^([A-Z]{2,4}\s?\d{3}[A-Z]?)\s?[:\-]?\s?(.*)/i);
  if (splitMatch) {
      courseCode = splitMatch[1].trim();
      courseName = splitMatch[2].trim();
  }

  // 2. Header Row
  const config = settings.config;
  let headerRowIndex = -1;
  for(let i=0; i<20; i++) { 
     if (!data[i]) continue;
     const rowStr = data[i].join(' ').toLowerCase();
     if(rowStr.includes(config.EXAM.toLowerCase()) && rowStr.includes(config.DATE.toLowerCase())) { 
         headerRowIndex = i; 
         break; 
     }
  }
  if (headerRowIndex === -1) return null;

  // 3. Roster Anchor
  const rosterKeyword = (settings.rosterKeyword || "students").toLowerCase();
  let rosterHeaderIdx = -1;
  for(let i = headerRowIndex + 1; i < data.length; i++) {
     if (String(data[i][0]).trim().toLowerCase().includes(rosterKeyword)) {
         rosterHeaderIdx = i;
         break;
     }
  }

  // 4. Parse Rosters (UPDATED FOR UNASSIGNED LOGIC)
  const sheetRoster = {};
  const allAssigned = new Set(); // Track students who have a location
  const masterList = [];         // Track all students in Column 0

  if (rosterHeaderIdx !== -1) {
    const locHeaders = data[rosterHeaderIdx]; 
    locHeaders.forEach((h, idx) => {
        if (idx === 0) return; 
        const loc = String(h).trim();
        if (loc && !sheetRoster[loc]) sheetRoster[loc] = [];
    });

    const startRow = rosterHeaderIdx + 2; 
    const maxScan = Math.min(startRow + 30, data.length); 

    for (let r = startRow; r < maxScan; r++) {
       // A. Capture Master List (Col 0)
       const masterName = String(data[r][0]).trim();
       if (masterName) masterList.push({ name: masterName, row: r });

       // B. Capture Assigned (Cols 1+)
       for (let c = 1; c < locHeaders.length; c++) { 
           const loc = String(locHeaders[c]).trim();
           if (!loc) continue;
           const rawVal = data[r][c];
           if (rawVal) {
               const name = String(rawVal).trim();
               if (name) {
                  if (!sheetRoster[loc]) sheetRoster[loc] = [];
                  sheetRoster[loc].push({ name: name, color: fontColors[r][c] });
                  allAssigned.add(name.toLowerCase()); // Mark as assigned
               }
           }
       }
    }

    // C. Calculate Unassigned
    const unassigned = [];
    masterList.forEach(student => {
        if (!allAssigned.has(student.name.toLowerCase())) {
            unassigned.push({ name: student.name, color: '#000000' });
        }
    });

    if (unassigned.length > 0) {
        sheetRoster['Unassigned'] = unassigned;
    }
  }

  // 5. Parse Exams
  const headers = data[headerRowIndex].map(h => String(h).trim().toLowerCase());
  const colMap = {};
  for (const [key, val] of Object.entries(config)) {
      colMap[key] = headers.findIndex(h => h.includes(val.toLowerCase()));
  }
  
  const colIdx = {
      name: colMap.EXAM,
      date: colMap.DATE,
      startTime: (colMap.START_TIME > -1) ? colMap.START_TIME : colMap.START_SITE,
      duration: colMap.DURATION,
      room: colMap.ROOM,
      password: colMap.PASSWORD
  };

  if (colIdx.name === -1 || colIdx.date === -1) return null;

  const exams = [];
  const scanEnd = (rosterHeaderIdx !== -1) ? rosterHeaderIdx : data.length;

  for (let i = headerRowIndex + 1; i < scanEnd; i++) {
      const row = data[i];
      const examName = String(row[colIdx.name]).trim();
      
      if (!examName) continue;
      if (examName.toLowerCase().includes('total')) continue;
      if (fontLines[i][colIdx.date] === 'line-through') continue; 

      const dbKey = `${d1}|${examName}`; 
      const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };
      
      let dateVal = row[colIdx.date];
      if (dateVal instanceof Date) dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");

      let timeVal = "";
      if (colIdx.startTime > -1) {
          const rawT = row[colIdx.startTime];
          if (rawT instanceof Date) timeVal = Utilities.formatDate(rawT, Session.getScriptTimeZone(), "hh:mm a");
          else timeVal = String(rawT);
      }

      exams.push({
          name: examName,
          date: dateVal,
          siteTime: timeVal,
          duration: (colIdx.duration > -1) ? row[colIdx.duration] : "",
          room: (colIdx.room > -1) ? row[colIdx.room] : "",
          password: (colIdx.password > -1) ? row[colIdx.password] : "",
          generalNotes: dbEntry.generalNotes, 
          studentTags: dbEntry.studentTags,   
          rosters: sheetRoster
      });
  }

  return { 
      course: { code: courseCode, name: courseName },
      fullTitle: d1,
      exams 
  };
}

/**
 * GENERATOR
 */
function createMLTProctoringDocuments(payload) {
  try {
    const settings = _getMLTSettings();
    const rootFolder = DriveApp.getFolderById(settings.folderId);
    let createdCount = 0;
    const createdUrls = {};

    payload.sheets.forEach(sheetData => {
        const folderName = sheetData.fullTitle;
        let targetFolder = _mlt_findSubFolder(rootFolder, folderName);
        if (!targetFolder) {
            targetFolder = rootFolder.createFolder(folderName);
        }

        sheetData.exams.forEach(exam => {
            const docTitle = `${folderName} - ${exam.name}`;
            
            const existing = targetFolder.getFilesByName(docTitle);
            if (existing.hasNext()) {
                createdUrls[exam.name] = existing.next().getUrl();
                return;
            }

            const doc = DocumentApp.create(docTitle);
            _populateMLTDoc(doc, docTitle, exam, settings.customNotes);
            
            const file = DriveApp.getFileById(doc.getId());
            targetFolder.addFile(file);
            DriveApp.getRootFolder().removeFile(file);

            createdUrls[exam.name] = file.getUrl();
            createdCount++;

            try {
                if (typeof logSystemAction === 'function') {
                    logSystemAction("MLT", "Created Doc", docTitle, doc.getId(), `Date: ${exam.date}`);
                }
            } catch (e) { console.warn("Log failed: " + e.message); }
        });
    });

    return { success: true, message: `Created ${createdCount} documents.`, docUrl: null, createdUrls: createdUrls };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * UPDATER
 */
function updateAllMLTDocuments(payload) {
  try {
    const settings = _getMLTSettings();
    const rootFolder = DriveApp.getFolderById(settings.folderId);
    let updatedCount = 0;

    payload.sheets.forEach(sheetData => {
        const folderName = sheetData.fullTitle;
        const targetFolder = _mlt_findSubFolder(rootFolder, folderName);
        
        if (!targetFolder) return; 

        sheetData.exams.forEach(exam => {
            const expectedTitle = `${folderName} - ${exam.name}`;
            const expectedNorm = _mlt_normalizeStr(expectedTitle);
            const examNorm = _mlt_normalizeStr(exam.name);

            let targetFile = null;
            if (exam.docUrl) {
                try { targetFile = DriveApp.getFileByUrl(exam.docUrl); } catch(e){}
            }

            if (!targetFile) {
                const files = targetFolder.getFiles();
                while (files.hasNext()) {
                    const f = files.next();
                    const fNorm = _mlt_normalizeStr(f.getName());
                    if (fNorm === expectedNorm || fNorm.includes(examNorm)) {
                        targetFile = f;
                        break;
                    }
                }
            }

            if (targetFile) {
                const doc = DocumentApp.openById(targetFile.getId());
                doc.getBody().clear();
                _populateMLTDoc(doc, targetFile.getName(), exam, settings.customNotes);
                updatedCount++;
                
                try {
                    if (typeof logSystemAction === 'function') {
                        logSystemAction("MLT", "Updated Doc", targetFile.getName(), targetFile.getId(), `Date: ${exam.date}`);
                    }
                } catch (e) { console.warn("Log failed: " + e.message); }
            }
        });
    });

    return { success: true, message: `Updated ${updatedCount} documents.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _populateMLTDoc(doc, title, exam, customNotes) {
    const body = doc.getBody();
    
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    body.appendParagraph("Exam Details").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendListItem(`Date: ${exam.date || "TBD"}`);
    body.appendListItem(`Start Time: ${exam.siteTime || "TBD"}`);
    body.appendListItem(`Duration: ${exam.duration || "-"}`);
    body.appendListItem(`Room: ${exam.room || "TBD"}`);
    body.appendListItem(`Password: ${exam.password || "-"}`);

    if (customNotes) {
        body.appendParagraph("General Instructions").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        body.appendParagraph(customNotes).setItalic(true);
    }

    if (exam.generalNotes) {
        body.appendParagraph("Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        const p = body.appendParagraph(exam.generalNotes);
        p.setBackgroundColor('#e8f5e9');
        p.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(10).setPaddingRight(10);
    }

    if (exam.rosters && Object.keys(exam.rosters).length > 0) {
        body.appendParagraph("Rosters").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        
        for (const [location, students] of Object.entries(exam.rosters)) {
            // SKIP UNASSIGNED IN THE DOC (Sidebar will still see it)
            if (location === 'Unassigned') continue; 

            body.appendParagraph(location).setHeading(DocumentApp.ParagraphHeading.HEADING2);
            
            const active = students.filter(s => {
                const tags = (exam.studentTags && exam.studentTags[s.name]) ? exam.studentTags[s.name] : null;
                return !(tags && tags.excluded);
            });

            if (active.length === 0) {
                body.appendParagraph("(No students)");
            } else {
                active.forEach(s => {
                    const li = body.appendListItem(s.name);
                    
                    if (exam.studentTags && exam.studentTags[s.name]) {
                        const tag = exam.studentTags[s.name];
                        if (tag.note) {
                           const t = li.appendText(` [${tag.note}]`);
                           t.setBold(true);
                           t.setBackgroundColor(tag.highlighted ? '#ffff00' : '#fff59d');
                        }
                    }
                    if (s.color && s.color !== '#000000') li.setForegroundColor(s.color);
                });
            }
        }
    }
    
    body.appendHorizontalRule();
    body.appendParagraph("Generated by Staff Hub (MLT)").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#888888').setFontSize(8);
    doc.saveAndClose();
}

function api_saveMLTAccommodations(payload) {
    return api_saveNursingAccommodations(payload);
}