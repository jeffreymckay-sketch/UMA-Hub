/**
 * -------------------------------------------------------------------
 * CONTROLLER: MLT PROCTORING (Aligned with Nursing Tool)
 * -------------------------------------------------------------------
 */

const MLT_CONSTANTS = {
    SETTINGS_KEY: 'mlt_settings', 
    DEFAULTS: {
      ROSTER_KEYWORD: 'Students',
      KEYWORDS: { EXAM: 'Exam', DATE: 'Date', START_TIME: 'Start Time', START_SITE: 'Site', DURATION: 'Duration', ROOM: 'Room', PASSWORD: 'Password', ITEMS: 'Items Allowed' }
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
      const dbMap = _mlt_getAccommodationsDBMap(); 
      
      const courseData = [];
  
      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.startsWith("_")) return; 
        
        // NEW: Safety Switch Toggle
        const isBetaEnabled = getSettings().enableBetaNursing === 'true' || getSettings().enableBetaNursing === true;
        let parsed = null;
        if (isBetaEnabled) {
            parsed = parseMLTSheet_V2(sheet, settings, dbMap);
        } else {
            parsed = parseMLTSheet(sheet, settings, dbMap);
        }
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
          courseData.push({ 
              sheetName: sheetName, 
              course: parsed.course, 
              fullTitle: parsed.fullTitle, 
              faculty: parsed.faculty,
              courseDisplay: parsed.courseDisplay,
              exams: parsed.exams 
          });
        }
      });
      
      return { success: true, data: { sheets: courseData, settings: settings } };
    } catch (e) { return { success: false, message: e.message }; }
  }
  
  /**
   * V2 Parser: Uses Utility_SheetReader.js to extract MLT sheet data heuristically.
   * Preserves exact JSON output shape and specific overrides.
   */
  function parseMLTSheet_V2(sheet, settings, dbMap) {
    const range = sheet.getDataRange();
    const data = range.getValues();
    const fontLines = range.getFontLines();
    const fontColors = range.getFontColors();
    
    if (!data || data.length < 5) return null;
    
    // --- Title & Course Code Parsing (Identical to V1) ---
    let rawTitle = "";
    try {
        const cellD1 = sheet.getRange("D1").getValue();
        if (cellD1 && String(cellD1).trim() !== "") rawTitle = String(cellD1).trim();
    } catch(e) {}
  
    if (!rawTitle) {
        try {
            const cellA1 = sheet.getRange("A1").getValue();
            if (cellA1 && String(cellA1).trim() !== "") rawTitle = String(cellA1).trim();
        } catch(e) {}
    }
    if (!rawTitle) rawTitle = sheet.getName();
    
    let courseDisplay = rawTitle; 
    let facultyName = "Faculty Unassigned";
  
    if (rawTitle.includes(':')) {
        const parts = rawTitle.split(':');
        courseDisplay = parts[0].trim(); 
        facultyName = parts.slice(1).join(':').trim(); 
    } else if (rawTitle.includes(' - ')) {
        const parts = rawTitle.split(' - ');
        courseDisplay = parts[0].trim();
        facultyName = parts.slice(1).join(' - ').trim();
    }
  
    const codeMatch = courseDisplay.match(/([A-Z]{2,4}\s?\d{3}[A-Z]?)/i);
    let courseCode = "MLT";
    if (codeMatch) courseCode = codeMatch[1].toUpperCase();
    
    // --- Header & Column Mapping (V2) ---
    const config = settings.config;
    const headerKeywords = [config.EXAM, config.DATE];
    const headerRowIndex = SheetReader.findHeaderRowHeuristic(data, headerKeywords, 20);
    
    if (headerRowIndex === -1) return null;

    // Create synonym map from user config
    const synonymConfig = {
        name: [config.EXAM],
        date: [config.DATE],
        startTime: [config.START_TIME, config.START_SITE], // Will match either
        duration: [config.DURATION],
        room: [config.ROOM],
        password: [config.PASSWORD],
        items: [config.ITEMS]
    };

    const colMap = SheetReader.mapColumnsBySynonyms(data[headerRowIndex], synonymConfig);
    if (colMap.name === -1 || colMap.date === -1) return null;

    // --- Roster Anchor & Mapping (V2) ---
    const rosterKeyword = settings.rosterKeyword || "students";
    const rosterHeaderIdx = SheetReader.findAnchorRow(data, [rosterKeyword], headerRowIndex + 1);
    const scanEnd = (rosterHeaderIdx !== -1) ? rosterHeaderIdx : data.length;

    // Roster Logic (Preserved exactly as V1 to maintain "Unassigned" logic)
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

    // --- Data Reading & Overrides (V2) ---
    const exams = [];
    const rawExams = SheetReader.readDynamicRoster(data.slice(0, scanEnd), headerRowIndex + 1, colMap.name, ['total'], 3);

    rawExams.forEach(parsedRow => {
        const rIndex = parsedRow.rowIndex;
        const rowData = parsedRow.data;
        const examName = parsedRow.name;

        if (fontLines[rIndex][colMap.date] === 'line-through') return; // Skip cancelled
        
        const dbKey = `${courseCode}|${examName}`;
        const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };
        
        // === MOVEMENT OVERRIDE LOGIC (CRITICAL) ===
        const examRoster = JSON.parse(JSON.stringify(sheetRoster));
        const studentTags = dbEntry.studentTags;
  
        if (studentTags) {
            for (const [sName, sData] of Object.entries(studentTags)) {
                if (sData.overrideLocation) {
                    const targetLoc = sData.overrideLocation;
                    let studentObj = null;
                    for (const loc in examRoster) {
                        const idx = examRoster[loc].findIndex(s => s.name === sName);
                        if (idx > -1) {
                            studentObj = examRoster[loc][idx];
                            examRoster[loc].splice(idx, 1);
                            break;
                        }
                    }
                    if (studentObj) {
                        if (!examRoster[targetLoc]) examRoster[targetLoc] = [];
                        examRoster[targetLoc].push(studentObj);
                    }
                }
            }
        }
        
        const rawDate = rowData[colMap.date];
        const dateVal = formatDateToPlainLanguage(rawDate !== 'TBD' ? rawDate : "");
        const siteTimeStr = (colMap.startTime > -1 && rowData[colMap.startTime] !== 'TBD') ? rowData[colMap.startTime] : '';
        const itemsVal = (colMap.items > -1 && rowData[colMap.items] !== 'TBD') ? String(rowData[colMap.items]).trim() : "";

        exams.push({ 
            name: examName, date: dateVal, siteTime: normalizeTime(siteTimeStr), 
            duration: (colMap.duration > -1 && rowData[colMap.duration] !== 'TBD') ? rowData[colMap.duration] : "", 
            room: (colMap.room > -1 && rowData[colMap.room] !== 'TBD') ? rowData[colMap.room] : "", 
            password: (colMap.password > -1 && rowData[colMap.password] !== 'TBD') ? rowData[colMap.password] : "", 
            itemsAllowed: itemsVal, generalNotes: dbEntry.generalNotes, studentTags: dbEntry.studentTags, rosters: examRoster 
        });
    });
    
    return { course: { code: courseCode, name: courseDisplay }, fullTitle: rawTitle, courseDisplay: courseDisplay, faculty: facultyName, exams: exams };
  }


  function parseMLTSheet(sheet, settings, dbMap) {
    const data = sheet.getDataRange().getValues();
    const fontLines = sheet.getDataRange().getFontLines();
    const fontColors = sheet.getDataRange().getFontColors();
    
    if (data.length < 5) return null;
    
    let rawTitle = "";
    try {
        const cellD1 = sheet.getRange("D1").getValue();
        if (cellD1 && String(cellD1).trim() !== "") {
            rawTitle = String(cellD1).trim();
        }
    } catch(e) {}
  
    if (!rawTitle) {
        try {
            const cellA1 = sheet.getRange("A1").getValue();
            if (cellA1 && String(cellA1).trim() !== "") {
                rawTitle = String(cellA1).trim();
            }
        } catch(e) {}
    }
    
    if (!rawTitle) rawTitle = sheet.getName();
    
    let courseDisplay = rawTitle; 
    let facultyName = "Faculty Unassigned";
  
    if (rawTitle.includes(':')) {
        const parts = rawTitle.split(':');
        courseDisplay = parts[0].trim(); 
        facultyName = parts.slice(1).join(':').trim(); 
    } else if (rawTitle.includes(' - ')) {
        const parts = rawTitle.split(' - ');
        courseDisplay = parts[0].trim();
        facultyName = parts.slice(1).join(' - ').trim();
    }
  
    const codeMatch = courseDisplay.match(/([A-Z]{2,4}\s?\d{3}[A-Z]?)/i);
    let courseCode = "MLT";
    
    if (codeMatch) {
        courseCode = codeMatch[1].toUpperCase();
    }
    
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
    
    const colIdx = { 
        name: colMap.EXAM, 
        date: colMap.DATE, 
        startTime: (colMap.START_TIME > -1) ? colMap.START_TIME : colMap.START_SITE, 
        duration: colMap.DURATION, 
        room: colMap.ROOM, 
        password: colMap.PASSWORD,
        items: colMap.ITEMS 
    };
    
    if (colIdx.name === -1 || colIdx.date === -1) return null;
    
    const exams = [];
    const scanEnd = (rosterHeaderIdx !== -1) ? rosterHeaderIdx : data.length;
    
    for (let i = headerRowIndex + 1; i < scanEnd; i++) {
        const row = data[i];
        const examName = String(row[colIdx.name]).trim();
        
        if (!examName || String(examName).toLowerCase().includes('total')) continue;
        if (fontLines[i][colIdx.date] === 'line-through') continue; 
        
        const dbKey = `${courseCode}|${examName}`;
        const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };
        
        // === MOVEMENT OVERRIDE LOGIC ===
        // Deep clone the roster so we don't affect other exams sharing this roster
        const examRoster = JSON.parse(JSON.stringify(sheetRoster));
        const studentTags = dbEntry.studentTags;
  
        if (studentTags) {
            // Iterate all students in the tags to find overrides
            for (const [sName, sData] of Object.entries(studentTags)) {
                if (sData.overrideLocation) {
                    const targetLoc = sData.overrideLocation;
                    let studentObj = null;
                    let originalLoc = null;
  
                    // 1. Find and Remove Student from their original location
                    for (const loc in examRoster) {
                        const idx = examRoster[loc].findIndex(s => s.name === sName);
                        if (idx > -1) {
                            studentObj = examRoster[loc][idx];
                            examRoster[loc].splice(idx, 1); // Remove
                            originalLoc = loc;
                            break;
                        }
                    }
  
                    // 2. Add to New Location (if found)
                    if (studentObj) {
                        if (!examRoster[targetLoc]) examRoster[targetLoc] = [];
                        examRoster[targetLoc].push(studentObj);
                    }
                }
            }
        }
        
        const dateVal = formatDateToPlainLanguage(row[colIdx.date]);
        const siteTime = normalizeTime(row[colIdx.startTime]);
        const itemsVal = (colIdx.items > -1) ? String(row[colIdx.items]).trim() : "";
  
        exams.push({ 
            name: examName, 
            date: dateVal, 
            siteTime: siteTime, 
            duration: (colIdx.duration > -1) ? row[colIdx.duration] : "", 
            room: (colIdx.room > -1) ? row[colIdx.room] : "", 
            password: (colIdx.password > -1) ? row[colIdx.password] : "", 
            itemsAllowed: itemsVal,
            generalNotes: dbEntry.generalNotes, 
            studentTags: dbEntry.studentTags, 
            rosters: examRoster // Use the modified roster
        });
    }
    
    return { 
        course: { code: courseCode, name: courseDisplay }, 
        fullTitle: rawTitle, 
        courseDisplay: courseDisplay,
        faculty: facultyName,
        exams 
    };
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
              
              if (existing.hasNext()) { 
                  createdUrls[exam.name] = existing.next().getUrl(); 
                  return; 
              }
              
              const doc = DocumentApp.create(docTitle);
              _populateMLTDoc(doc, docTitle, sheetData.courseDisplay, sheetData.faculty, exam, settings.customNotes);
              
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
                      if (fNorm === expectedNorm || fNorm.includes(examNameNorm)) { 
                          targetFile = f; 
                          if (f.getName() !== expectedTitle) f.setName(expectedTitle);
                          break; 
                      }
                  }
              }
              
              if (targetFile) {
                  const doc = DocumentApp.openById(targetFile.getId());
                  _populateMLTDoc(doc, targetFile.getName(), sheetData.courseDisplay, sheetData.faculty, exam, settings.customNotes);
                  if (typeof logSystemAction === 'function') logSystemAction("MLT", "Updated Doc", targetFile.getName(), targetFile.getId(), `Date: ${exam.date}`);
                  updatedCount++;
              }
          });
      });
      
      return { success: true, message: `Updated ${updatedCount} documents.` };
    } catch (e) { return { success: false, message: e.message }; }
  }
  
  // === SURGICAL UPDATE FUNCTIONS ===
  
  function _populateMLTDoc(doc, filename, courseTitle, facultyName, exam, customNotes) {
      const body = doc.getBody();
      
      const displayCourse = courseTitle || "MLT Exam";
      const displayFaculty = facultyName || "Faculty Unassigned";
      
      const docDisplayTitle = `${displayCourse}: ${exam.name}`;
  
      // 1. UPDATE TITLE
      const existingTitle = body.getChild(0);
      if (existingTitle.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const titlePara = existingTitle.asParagraph();
          if (titlePara.getText() !== docDisplayTitle) {
              titlePara.setText(docDisplayTitle);
          }
          titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
      }
      
      // 2. UPDATE HEADING 1 (Faculty Name)
      let subHeader = null;
      if (body.getNumChildren() > 1) {
          const child = body.getChild(1);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
              subHeader = child.asParagraph();
          }
      }
      
      if (!subHeader) {
          subHeader = body.insertParagraph(1, displayFaculty);
      } else {
          if (subHeader.getText() !== displayFaculty) subHeader.setText(displayFaculty);
      }
      subHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
      // 3. UPDATE EXAM DETAILS (Heading 2)
      const examDetailsMap = {
          'Date': exam.date || "TBD",
          'Start Time': exam.siteTime || "TBD",
          'Duration': exam.duration || "-",
          'Room': exam.room || "TBD",
          'Password': exam.password || "-",
          'Items Allowed': exam.itemsAllowed || "None"
      };
      
      _mlt_updateKeyValueSection(body, "Exam Details", examDetailsMap, DocumentApp.ParagraphHeading.HEADING2);
      
      // 4. UPDATE GENERAL INSTRUCTIONS (Heading 2)
      if (customNotes) {
          _mlt_updateOrCreateTextSection(body, "General Instructions", customNotes, true);
      }
      
      // 5. UPDATE IMPORTANT LINKS
      _mlt_updateImportantLinks(body);
      
      // 6. UPDATE ACCOMMODATIONS (Green Box)
      if (exam.generalNotes) {
          _mlt_updateOrCreateHighlightedSection(body, "Exam Accommodations", exam.generalNotes);
      } else {
          _mlt_removeSectionIfExists(body, "Exam Accommodations");
      }
      
      // 7. UPDATE ROSTERS
      _mlt_updateLocationRosters(body, exam);
      
      doc.saveAndClose();
  }
  
  // === LOCAL HELPERS ===
  
  function _mlt_findHeadingIndex(body, headingText) {
    const numChildren = body.getNumChildren();
    const targetNorm = headingText.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (let i = 0; i < numChildren; i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const para = child.asParagraph();
        const text = para.getText();
        if (text && text.toLowerCase().replace(/[^a-z0-9]/g, '').includes(targetNorm)) {
          const h = para.getHeading();
          if (h !== DocumentApp.ParagraphHeading.NORMAL && h !== DocumentApp.ParagraphHeading.TITLE) {
            return i;
          }
        }
      }
    }
    return -1;
  }
  
  function _mlt_updateKeyValueSection(body, sectionTitle, dataMap, headingLevel) {
    let sectionIndex = _mlt_findHeadingIndex(body, sectionTitle);
    if (sectionIndex === -1) {
      const rostersIndex = _mlt_findHeadingIndex(body, "Location Rosters");
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
  
  function _mlt_updateOrCreateTextSection(body, sectionTitle, content, isItalic) {
    let sectionIndex = _mlt_findHeadingIndex(body, sectionTitle);
    if (sectionIndex === -1) {
      const rostersIndex = _mlt_findHeadingIndex(body, "Location Rosters");
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
  
  function _mlt_updateOrCreateHighlightedSection(body, sectionTitle, content) {
    let sectionIndex = _mlt_findHeadingIndex(body, sectionTitle);
    if (sectionIndex === -1) {
      const rostersIndex = _mlt_findHeadingIndex(body, "Location Rosters");
      const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
      body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const p = body.insertParagraph(insertAt + 1, content);
      p.setBackgroundColor('#e8f5e9'); // Light Green
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
  
  function _mlt_updateImportantLinks(body) {
    const sectionTitle = "Important Links";
    // Default links; in future could be pulled from settings
    const links = [
      { text: "Red Flag Reporting Form", url: "https://docs.google.com/forms/d/e/1FAIpQLSfORKCKol8SsRldNKfvsDy3ILNs9HcFv3gKb8TuxrNrlqxijw/viewform" }
    ];
    
    let sectionIndex = _mlt_findHeadingIndex(body, sectionTitle);
    if (sectionIndex === -1) {
      let insertAt = _mlt_findHeadingIndex(body, "Exam Accommodations");
      if (insertAt === -1) insertAt = _mlt_findHeadingIndex(body, "Location Rosters");
      if (insertAt === -1) insertAt = body.getNumChildren();
      body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      sectionIndex = insertAt;
    } else {
      body.getChild(sectionIndex).asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
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
    
    links.forEach((link, i) => {
      const li = body.insertListItem(sectionIndex + 1 + i, link.text);
      li.setLinkUrl(link.url);
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
    });
  }
  
  function _mlt_removeSectionIfExists(body, sectionTitle) {
    const sectionIndex = _mlt_findHeadingIndex(body, sectionTitle);
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
  
  function _mlt_updateLocationRosters(body, exam) {
    if (!exam.rosters || Object.keys(exam.rosters).length === 0) {
      _mlt_removeSectionIfExists(body, "Location Rosters");
      return;
    }
    
    let rostersIndex = _mlt_findHeadingIndex(body, "Location Rosters");
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
        return !(tagData && tagData.excluded);
      });
      
      // Find or Create Location Heading (Heading 3)
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
      
      // Use Title Case for display
      const displayLocation = location.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())));
      
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
      
      // Update Students List
      const existingLines = [];
      let scanIndex = currentIndex;
      while (scanIndex < body.getNumChildren()) {
        const child = body.getChild(scanIndex);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          if (child.asParagraph().getHeading() !== DocumentApp.ParagraphHeading.NORMAL) break;
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
          if (tag) {
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
  
      // Cleanup removed students
      for (let i = existingLines.length - 1; i >= 0; i--) {
        const line = existingLines[i];
        if (!processedDocIndices.has(line.index)) {
          body.removeChild(line.listItem);
          if (line.index < insertPosition) insertPosition--;
        }
      }
      
      currentIndex = insertPosition;
    });
  }
  
  function _mlt_getAccommodationsDBMap() {
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
              try { tags = JSON.parse(row[4]); } catch (e) { } 
              map[id] = { generalNotes: row[3], studentTags: tags };
          }
      } catch (e) { console.warn("DB Error: " + e.message); }
      return map;
  }
  
  function api_saveMLTAccommodations(payload) {
    try {
      const ss = getMasterDataHub(); 
      let sheet = ss.getSheetByName("_DB_ACCOMMODATIONS");
      if (!sheet) {
          sheet = ss.insertSheet("_DB_ACCOMMODATIONS");
          sheet.appendRow(["Unique_ID", "Course_Code", "Exam_Name", "General_Notes", "Student_Data"]);
      }
      
      // Unique ID format: CourseCode|ExamName
      const uniqueId = `${payload.courseCode}|${payload.examName}`;
      const studentJson = JSON.stringify(payload.studentTags || {});
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === uniqueId) { rowIndex = i + 1; break; }
      }
      
      if (rowIndex > -1) {
          sheet.getRange(rowIndex, 4, 1, 2).setValues([[payload.generalNotes, studentJson]]);
      } else {
          sheet.appendRow([uniqueId, payload.courseCode, payload.examName, payload.generalNotes, studentJson]);
      }
      return { success: true, message: 'Saved to Database!' };
    } catch (e) { return { success: false, message: e.message }; }
  }