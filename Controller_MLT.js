const MLT_CONSTANTS = {
  SETTINGS_KEY: 'MLT_SETTINGS_V2', 
  DEFAULTS: {
    ROSTER_KEYWORD: 'Students',
    KEYWORDS: { EXAM: 'Exam', DATE: 'Date', START_TIME: 'Start Time', START_SITE: 'Site', DURATION: 'Duration', ROOM: 'Room', PASSWORD: 'Password' }
  }
};

function _mlt_normalizeStr(str) {
  if (!str) return "";
  return str.toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function _mlt_findSubFolder(parentFolder, targetName) {
  const targetNorm = _mlt_normalizeStr(targetName);
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (_mlt_normalizeStr(folder.getName()) === targetNorm) return folder;
  }
  return null;
}

function _getMLTSettings() {
  let allProps = {};
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(MLT_CONSTANTS.SETTINGS_KEY);
    if (raw) allProps = JSON.parse(raw);
  } catch (e) {}
  if (!allProps.config) allProps.config = MLT_CONSTANTS.DEFAULTS.KEYWORDS;
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
    rosterKeyword: allProps.rosterKeyword || MLT_CONSTANTS.DEFAULTS.ROSTER_KEYWORD
  };
}

function mlt_saveSettings(settingsObj) {
    try {
        let current = {};
        const raw = PropertiesService.getScriptProperties().getProperty(MLT_CONSTANTS.SETTINGS_KEY);
        if (raw) { try { current = JSON.parse(raw); } catch(e) {} }
        const newSettings = { ...current, ...settingsObj };
        PropertiesService.getScriptProperties().setProperty(MLT_CONSTANTS.SETTINGS_KEY, JSON.stringify(newSettings));
        return { success: true, message: "MLT Settings Saved." };
    } catch (e) { return { success: false, message: e.message }; }
}

function api_getMLTData() {
  try {
    const settings = _getMLTSettings();
    if (!settings.sheetId || !settings.folderId) return { success: false, message: "MLT settings incomplete." };
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
          existingFiles.push({ name: f.getName(), nameNorm: _mlt_normalizeStr(f.getName()), url: f.getUrl() });
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
        courseData.push({ sheetName: sheetName, course: parsed.course, fullTitle: parsed.fullTitle, exams: parsed.exams });
      }
    });
    return { success: true, data: { sheets: courseData, settings: settings } };
  } catch (e) { return { success: false, message: e.message }; }
}

function parseMLTSheet(sheet, settings, dbMap) {
  const data = sheet.getDataRange().getValues();
  const fontLines = sheet.getDataRange().getFontLines();
  const fontColors = sheet.getDataRange().getFontColors();
  if (data.length < 5) return null;
  const d1 = (data[0] && data[0][3]) ? String(data[0][3]).trim() : sheet.getName();
  let courseCode = "MLT";
  let courseName = d1;
  const splitMatch = d1.match(/^([A-Z]{2,4}\s?\d{3}[A-Z]?)\s?[:\-]?\s?(.*)/i);
  if (splitMatch) { courseCode = splitMatch[1].trim(); courseName = splitMatch[2].trim(); }
  const config = settings.config;
  let headerRowIndex = -1;
  for(let i=0; i<20; i++) { 
     if (!data[i]) continue;
     const rowStr = data[i].join(' ').toLowerCase();
     if(rowStr.includes(config.EXAM.toLowerCase()) && rowStr.includes(config.DATE.toLowerCase())) { headerRowIndex = i; break; }
  }
  if (headerRowIndex === -1) return null;
  const rosterKeyword = (settings.rosterKeyword || "students").toLowerCase();
  let rosterHeaderIdx = -1;
  for(let i = headerRowIndex + 1; i < data.length; i++) {
     if (String(data[i][0]).trim().toLowerCase().includes(rosterKeyword)) { rosterHeaderIdx = i; break; }
  }
  const sheetRoster = {};
  const allAssigned = new Set();
  const masterList = [];         
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
       const masterName = String(data[r][0]).trim();
       if (masterName) masterList.push({ name: masterName, row: r });
       for (let c = 1; c < locHeaders.length; c++) { 
           const loc = String(locHeaders[c]).trim();
           if (!loc) continue;
           const rawVal = data[r][c];
           if (rawVal) {
               const name = String(rawVal).trim();
               if (name) {
                  if (!sheetRoster[loc]) sheetRoster[loc] = [];
                  sheetRoster[loc].push({ name: name, color: fontColors[r][c] });
                  allAssigned.add(name.toLowerCase()); 
               }
           }
       }
    }
    const unassigned = [];
    masterList.forEach(student => {
        if (!allAssigned.has(student.name.toLowerCase())) unassigned.push({ name: student.name, color: '#000000' });
    });
    if (unassigned.length > 0) sheetRoster['Unassigned'] = unassigned;
  }
  const headers = data[headerRowIndex].map(h => String(h).trim().toLowerCase());
  const colMap = {};
  for (const [key, val] of Object.entries(config)) { colMap[key] = headers.findIndex(h => h.includes(val.toLowerCase())); }
  const colIdx = { name: colMap.EXAM, date: colMap.DATE, startTime: (colMap.START_TIME > -1) ? colMap.START_TIME : colMap.START_SITE, duration: colMap.DURATION, room: colMap.ROOM, password: colMap.PASSWORD };
  if (colIdx.name === -1 || colIdx.date === -1) return null;
  const exams = [];
  const scanEnd = (rosterHeaderIdx !== -1) ? rosterHeaderIdx : data.length;
  for (let i = headerRowIndex + 1; i < scanEnd; i++) {
      const row = data[i];
      const examName = String(row[colIdx.name]).trim();
      if (!examName || String(examName).toLowerCase().includes('total')) continue;
      if (fontLines[i][colIdx.date] === 'line-through') continue; 
      const dbKey = `${d1}|${examName}`; 
      const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };
      
      // FLEXIBLE DATE & TIME
      const dateVal = formatDateToPlainLanguage(row[colIdx.date]);
      const siteTime = normalizeTime(row[colIdx.startTime]);

      exams.push({ name: examName, date: dateVal, siteTime: siteTime, duration: (colIdx.duration > -1) ? row[colIdx.duration] : "", room: (colIdx.room > -1) ? row[colIdx.room] : "", password: (colIdx.password > -1) ? row[colIdx.password] : "", generalNotes: dbEntry.generalNotes, studentTags: dbEntry.studentTags, rosters: sheetRoster });
  }
  return { course: { code: courseCode, name: courseName }, fullTitle: d1, exams };
}

function createMLTProctoringDocuments(payload) {
  try {
    const settings = _getMLTSettings();
    const rootFolder = DriveApp.getFolderById(settings.folderId);
    let createdCount = 0;
    const createdUrls = {};
    payload.sheets.forEach(sheetData => {
        const folderName = sheetData.fullTitle;
        let targetFolder = _mlt_findSubFolder(rootFolder, folderName);
        if (!targetFolder) targetFolder = rootFolder.createFolder(folderName);
        sheetData.exams.forEach(exam => {
            const docTitle = `${folderName} - ${exam.name}`;
            const existing = targetFolder.getFilesByName(docTitle);
            if (existing.hasNext()) { createdUrls[exam.name] = existing.next().getUrl(); return; }
            const doc = DocumentApp.create(docTitle);
            _populateMLTDoc(doc, docTitle, exam, settings.customNotes);
            const file = DriveApp.getFileById(doc.getId());
            targetFolder.addFile(file);
            DriveApp.getRootFolder().removeFile(file);
            createdUrls[exam.name] = file.getUrl();
            if (typeof logSystemAction === 'function') logSystemAction("MLT", "Created Doc", docTitle, doc.getId(), `Date: ${exam.date}`);
            createdCount++;
        });
    });
    return { success: true, message: `Created ${createdCount} documents.`, createdUrls: createdUrls };
  } catch (e) { return { success: false, message: e.message }; }
}

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
            const examNameNorm = _mlt_normalizeStr(exam.name);
            const expectedTitle = `${folderName} - ${exam.name}`;
            const expectedNorm = _mlt_normalizeStr(expectedTitle);
            let targetFile = null;
            if (exam.docUrl) { try { targetFile = DriveApp.getFileByUrl(exam.docUrl); } catch(e){} }
            if (!targetFile) {
                const files = targetFolder.getFiles();
                while (files.hasNext()) {
                    const f = files.next();
                    const fNorm = _mlt_normalizeStr(f.getName());
                    if (fNorm === expectedNorm || fNorm.includes(examNameNorm)) { targetFile = f; break; }
                }
            }
            if (targetFile) {
                const doc = DocumentApp.openById(targetFile.getId());
                _populateMLTDoc(doc, targetFile.getName(), exam, settings.customNotes);
                if (typeof logSystemAction === 'function') logSystemAction("MLT", "Updated Doc", targetFile.getName(), targetFile.getId(), `Date: ${exam.date}`);
                updatedCount++;
            }
        });
    });
    return { success: true, message: `Updated ${updatedCount} documents.` };
  } catch (e) { return { success: false, message: e.message }; }
}

// === SURGICAL UPDATE FUNCTIONS FOR MLT ===

function _populateMLTDoc(doc, title, exam, customNotes) {
    const body = doc.getBody();
    
    // === SURGICAL UPDATE ===
    
    // 1. UPDATE TITLE
    const existingTitle = body.getChild(0);
    if (existingTitle.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const titlePara = existingTitle.asParagraph();
        if (titlePara.getText() !== title) {
            titlePara.setText(title);
            titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
        }
    }
    
    // 2. UPDATE EXAM DETAILS
    const examDetailsMap = {
        'Date': exam.date || "TBD",
        'Start Time': exam.siteTime || "TBD",
        'Duration': exam.duration || "-",
        'Room': exam.room || "TBD",
        'Password': exam.password || "-"
    };
    _updateKeyValueSection(body, "Exam Details", examDetailsMap, DocumentApp.ParagraphHeading.HEADING1);
    
    // 3. UPDATE GENERAL INSTRUCTIONS
    if (customNotes) {
        _updateOrCreateTextSection(body, "General Instructions", customNotes, true);
    }
    
    // 4. UPDATE ACCOMMODATIONS
    if (exam.generalNotes) {
        _updateOrCreateHighlightedSection(body, "Accommodations", exam.generalNotes);
    } else {
        _removeSectionIfExists(body, "Accommodations");
    }
    
    // 5. UPDATE ROSTERS
    _updateMLTRosters(body, exam);
    
    doc.saveAndClose();
}

function _updateMLTRosters(body, exam) {
    if (!exam.rosters || Object.keys(exam.rosters).length === 0) {
        _removeSectionIfExists(body, "Rosters");
        return;
    }
    
    let rostersIndex = _findHeadingIndex(body, "Rosters");
    
    if (rostersIndex === -1) {
        rostersIndex = body.getNumChildren();
        body.insertParagraph(rostersIndex, "Rosters").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    }
    
    // UMAAL FIRST
    const sortedLocations = Object.keys(exam.rosters).sort((a, b) => {
        if (a === 'UMAAL') return -1;
        if (b === 'UMAAL') return 1;
        return a.localeCompare(b);
    }).filter(loc => loc !== 'Unassigned');
    
    let currentIndex = rostersIndex + 1;
    
    sortedLocations.forEach(location => {
        const active = exam.rosters[location].filter(s => {
            const tags = (exam.studentTags && exam.studentTags[s.name]) ? exam.studentTags[s.name] : null;
            return !(tags && tags.excluded);
        });
        
        // Find or create location
        let locationIndex = -1;
        for (let i = currentIndex; i < body.getNumChildren(); i++) {
            const child = body.getChild(i);
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                const para = child.asParagraph();
                if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING2 && para.getText() === location) {
                    locationIndex = i;
                    break;
                }
                if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) break;
            }
        }
        
        if (locationIndex === -1) {
            body.insertParagraph(currentIndex, location).setHeading(DocumentApp.ParagraphHeading.HEADING2);
            locationIndex = currentIndex;
            currentIndex++;
        } else {
            currentIndex = locationIndex + 1;
        }
        
        // Build target
        const targetStudents = new Map();
        if (active.length === 0) {
            targetStudents.set('(No students)', { color: null, note: null, isHighlighted: false });
        } else {
            active.forEach(s => {
                let note = null;
                let isHighlighted = false;
                
                if (exam.studentTags && exam.studentTags[s.name]) {
                    const tag = exam.studentTags[s.name];
                    note = tag.note || null;
                    isHighlighted = tag.highlighted || false;
                }
                
                targetStudents.set(s.name, { color: s.color, note, isHighlighted });
            });
        }
        
        // Remove old students
        const existingStudents = new Map();
        let scanIndex = currentIndex;
        while (scanIndex < body.getNumChildren()) {
            const child = body.getChild(scanIndex);
            
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                const para = child.asParagraph();
                if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) break;
            }
            
            if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
                const listItem = child.asListItem();
                const fullText = listItem.getText();
                const nameMatch = fullText.match(/^([^\[]+)/);
                const studentName = nameMatch ? nameMatch[1].trim() : fullText;
                existingStudents.set(studentName, scanIndex);
            }
            
            scanIndex++;
        }
        
        existingStudents.forEach((index, studentName) => {
            if (!targetStudents.has(studentName)) {
                body.removeChild(body.getChild(index));
                existingStudents.forEach((idx, name) => {
                    if (idx > index) existingStudents.set(name, idx - 1);
                });
            }
        });
        
        // Add or update students
        let insertPosition = currentIndex;
        targetStudents.forEach((data, studentName) => {
            const existingIndex = existingStudents.get(studentName);
            
            if (existingIndex !== undefined) {
                const listItem = body.getChild(existingIndex).asListItem();
                const expectedText = data.note ? `${studentName} [${data.note}]` : studentName;
                
                if (listItem.getText() !== expectedText) {
                    listItem.clear();
                    listItem.setText(studentName);
                    
                    if (data.note) {
                        const t = listItem.appendText(` [${data.note}]`);
                        t.setBold(true).setBackgroundColor(data.isHighlighted ? '#ffff00' : '#fff59d');
                    }
                    
                    if (data.color && data.color !== '#000000') listItem.setForegroundColor(data.color);
                }
                
                insertPosition = Math.max(insertPosition, existingIndex + 1);
            } else {
                const li = body.insertListItem(insertPosition, studentName);
                
                if (data.note) {
                    const t = li.appendText(` [${data.note}]`);
                    t.setBold(true).setBackgroundColor(data.isHighlighted ? '#ffff00' : '#fff59d');
                }
                
                if (data.color && data.color !== '#000000') li.setForegroundColor(data.color);
                
                insertPosition++;
            }
        });
        
        currentIndex = insertPosition;
    });
}