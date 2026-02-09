function _getNursingSettings() {
  const allProps = getSettings();
  let raw = allProps.nursing_settings;
  let settings = {};
  if (raw && typeof raw === 'string') {
    try { settings = JSON.parse(raw); } catch (e) { console.error("JSON Parse Error", e); }
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

function _parseCellA1(cellValue) {
  if (!cellValue) return { folderKey: "Unknown", fullTitlePrefix: "Unknown" };
  const str = String(cellValue).trim();
  let separator = ':';
  let parts = str.split(':');
  if (parts.length === 1) {
    separator = '-';
    parts = str.split('-');
  }
  if (parts.length > 1) {
    const courseInfo = parts[0].trim();
    const facultyInfo = parts.slice(1).join(separator).trim();
    const fullTitlePrefix = `${courseInfo} ${facultyInfo}`;
    return { folderKey: courseInfo, fullTitlePrefix: fullTitlePrefix };
  }
  return { folderKey: str, fullTitlePrefix: str };
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
      return { success: false, message: "Nursing settings missing. Please save Sheet/Folder IDs in Settings." };
    }
    const ss = SpreadsheetApp.openById(config.sheetId);
    const sheets = ss.getSheets();
    const rootFolder = DriveApp.getFolderById(config.folderId);
    const dbMap = getAccommodationsDBMap();
    const courseData = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("_")) return; 
      const parsed = parseNursingSheet(sheet, dbMap);
      if (!parsed) return;
      const meta = _parseCellA1(parsed.rawA1);
      let targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      const existingFiles = [];
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

function parseNursingSheet(sheet, dbMap) {
  const range = sheet.getDataRange();
  const data = range.getValues();
  const fontColors = range.getFontColors();
  const fontLines = range.getFontLines();
  if (data.length < 5) return null; 
  const a1 = String(data[0][0]).trim(); 
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
  const scanEnd = (rosterHeaderIdx !== -1) ? rosterHeaderIdx : data.length;
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
  for (let i = headerRowIndex + 1; i < scanEnd; i++) {
      const row = data[i];
      const examName = row[colMap.name];
      if (!examName || String(examName).trim() === '') continue;
      if (String(examName).toLowerCase().includes('total')) continue;
      if (fontLines[i][colMap.date] === 'line-through' || fontLines[i][colMap.name] === 'line-through') continue;
      
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
          rosters: sheetRoster 
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
        const docTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        const existing = targetFolder.getFilesByName(docTitle);
        if (existing.hasNext()) {
          createdUrls[exam.name] = existing.next().getUrl();
          return;
        }
        const doc = DocumentApp.create(docTitle);
        _populateNursingDoc(doc, docTitle, exam, config.customNotes);
        const file = DriveApp.getFileById(doc.getId());
        targetFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
        createdUrls[exam.name] = file.getUrl();
        if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Created Doc", docTitle, doc.getId(), `Date: ${exam.date}`);
        createdCount++;
      });
    });
    return { success: true, message: `Created ${createdCount} documents.`, createdUrls: createdUrls };
  } catch (e) { return { success: false, message: e.message }; }
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
        const examNameNorm = _normalizeStr(exam.name);
        const expectedTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        const expectedTitleNorm = _normalizeStr(expectedTitle);
        let targetFile = null;
        if (exam.docUrl) {
          try { targetFile = DriveApp.getFileByUrl(exam.docUrl); } catch(e) {}
        }
        if (!targetFile) {
          const files = targetFolder.getFiles();
          while (files.hasNext()) {
            const f = files.next();
            if (_normalizeStr(f.getName()) === expectedTitleNorm) { targetFile = f; break; }
          }
        }
        if (targetFile) {
          const doc = DocumentApp.openById(targetFile.getId());
          _populateNursingDoc(doc, targetFile.getName(), exam, config.customNotes);
          if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Updated Doc", targetFile.getName(), targetFile.getId(), `Date: ${exam.date}`);
          updatedCount++;
        }
      });
    });
    return { success: true, message: `Updated ${updatedCount} documents.` };
  } catch (e) { return { success: false, message: e.message }; }
}

// === SURGICAL UPDATE FUNCTIONS ===

function _populateNursingDoc(doc, title, exam, customNotes) {
    const body = doc.getBody();
    
    // === SURGICAL UPDATE: Only change what's different ===
    
    // 1. UPDATE TITLE (if different)
    const existingTitle = body.getChild(0);
    if (existingTitle.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const titlePara = existingTitle.asParagraph();
        if (titlePara.getText() !== title) {
            titlePara.setText(title);
            titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
        }
    }
    
    // 2. FIND OR CREATE EXAM DETAILS SECTION
    const examDetailsMap = {
        'Faculty': title.split(' - ')[0] || title,
        'Date': exam.date,
        'Start Time (On Site)': exam.siteTime || "N/A",
        'Start Time (Zoom)': exam.zoomTime || "N/A",
        'Duration': exam.duration || "-",
        'Password': exam.password || "-"
    };
    _updateKeyValueSection(body, "Exam Details", examDetailsMap, DocumentApp.ParagraphHeading.HEADING1);
    
    // 3. UPDATE IMPORTANT LINKS (leave as-is, they rarely change)
    // Skip for now to avoid unnecessary updates
    
    // 4. UPDATE GENERAL INSTRUCTIONS
    if (customNotes) {
        _updateOrCreateTextSection(body, "General Instructions", customNotes, true);
    }
    
    // 5. UPDATE ACCOMMODATIONS
    if (exam.generalNotes) {
        _updateOrCreateHighlightedSection(body, "Exam Accommodations", exam.generalNotes);
    } else {
        _removeSectionIfExists(body, "Exam Accommodations");
    }
    
    // 6. SURGICAL ROSTER UPDATE (the complex part)
    _updateLocationRosters(body, exam);
    
    doc.saveAndClose();
}

// === HELPER FUNCTIONS FOR SURGICAL UPDATES ===

function _findHeadingIndex(body, headingText) {
    const numChildren = body.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
        const child = body.getChild(i);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const para = child.asParagraph();
            if (para.getText().trim() === headingText) {
                return i;
            }
        }
    }
    return -1;
}

function _updateKeyValueSection(body, sectionTitle, dataMap, headingLevel) {
    let sectionIndex = _findHeadingIndex(body, sectionTitle);
    
    // If section doesn't exist, create it at the end (before rosters)
    if (sectionIndex === -1) {
        const rostersIndex = _findHeadingIndex(body, "Location Rosters");
        const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
        body.insertParagraph(insertAt, sectionTitle).setHeading(headingLevel);
        sectionIndex = insertAt;
    }
    
    // Update or create each list item
    let currentIndex = sectionIndex + 1;
    
    for (const [key, value] of Object.entries(dataMap)) {
        const expectedText = `${key}: ${value}`;
        let found = false;
        
        // Look ahead a few items to find this key
        for (let i = currentIndex; i < Math.min(currentIndex + 10, body.getNumChildren()); i++) {
            const child = body.getChild(i);
            
            // Stop if we hit another heading
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                const para = child.asParagraph();
                const heading = para.getHeading();
                if (heading !== DocumentApp.ParagraphHeading.NORMAL) break;
            }
            
            // Check if this is our item
            if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
                const listItem = child.asListItem();
                const text = listItem.getText();
                
                if (text.startsWith(key + ':')) {
                    // Update if different
                    if (text !== expectedText) {
                        listItem.setText(expectedText);
                    }
                    found = true;
                    currentIndex = i + 1;
                    break;
                }
            }
        }
        
        // If not found, insert it
        if (!found) {
            body.insertListItem(currentIndex, expectedText);
            currentIndex++;
        }
    }
}

function _updateOrCreateTextSection(body, sectionTitle, content, isItalic) {
    let sectionIndex = _findHeadingIndex(body, sectionTitle);
    
    if (sectionIndex === -1) {
        // Create section before Location Rosters
        const rostersIndex = _findHeadingIndex(body, "Location Rosters");
        const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
        body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        const para = body.insertParagraph(insertAt + 1, content);
        if (isItalic) para.setItalic(true);
    } else {
        // Update existing content
        const nextChild = body.getChild(sectionIndex + 1);
        if (nextChild && nextChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const para = nextChild.asParagraph();
            if (para.getText() !== content) {
                para.setText(content);
                if (isItalic) para.setItalic(true);
            }
        }
    }
}

function _updateOrCreateHighlightedSection(body, sectionTitle, content) {
    let sectionIndex = _findHeadingIndex(body, sectionTitle);
    
    if (sectionIndex === -1) {
        const rostersIndex = _findHeadingIndex(body, "Location Rosters");
        const insertAt = rostersIndex > -1 ? rostersIndex : body.getNumChildren();
        body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        const p = body.insertParagraph(insertAt + 1, content);
        p.setBackgroundColor('#e8f5e9');
        p.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(10).setPaddingRight(10);
    } else {
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
        // Remove heading and next paragraph
        body.removeChild(body.getChild(sectionIndex));
        if (sectionIndex < body.getNumChildren()) {
            const nextChild = body.getChild(sectionIndex);
            if (nextChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
                body.removeChild(nextChild);
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
    
    // Create section if it doesn't exist
    if (rostersIndex === -1) {
        rostersIndex = body.getNumChildren();
        body.insertParagraph(rostersIndex, "Location Rosters").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    }
    
    // UMAAL FIRST, then alphabetical
    const sortedLocations = Object.keys(exam.rosters).sort((a, b) => {
        if (a === 'UMAAL') return -1;
        if (b === 'UMAAL') return 1;
        return a.localeCompare(b);
    });
    
    let currentIndex = rostersIndex + 1;
    
    sortedLocations.forEach(location => {
        const students = exam.rosters[location];
        
        // Filter active students (respecting sidebar exclusions)
        const activeStudents = students.filter(s => {
            const tagData = (exam.studentTags && exam.studentTags[s.name]) ? exam.studentTags[s.name] : null;
            return !(tagData && typeof tagData === 'object' && tagData.excluded);
        });
        
        // Find or create location heading
        let locationIndex = -1;
        for (let i = currentIndex; i < body.getNumChildren(); i++) {
            const child = body.getChild(i);
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                const para = child.asParagraph();
                if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING2 && para.getText() === location) {
                    locationIndex = i;
                    break;
                }
                // Stop if we hit another H1
                if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) break;
            }
        }
        
        if (locationIndex === -1) {
            // Create new location section
            body.insertParagraph(currentIndex, location).setHeading(DocumentApp.ParagraphHeading.HEADING2);
            locationIndex = currentIndex;
            currentIndex++;
        } else {
            currentIndex = locationIndex + 1;
        }
        
        // Build target student list
        const targetStudents = new Map();
        if (activeStudents.length === 0) {
            targetStudents.set('(No students assigned)', { color: null, note: null });
        } else {
            activeStudents.forEach(studentObj => {
                let note = "";
                let isHighlighted = false;
                let isLocked = false;
                let color = studentObj.color;
                
                if (exam.studentTags && exam.studentTags[studentObj.name]) {
                    const tag = exam.studentTags[studentObj.name];
                    if (typeof tag === 'object') {
                        note = tag.note || "";
                        isHighlighted = tag.highlighted || false;
                        isLocked = tag.locked || false;
                    }
                }
                
                if (isLocked) color = '#000000';
                else if (!color || color === '#000000') color = null;
                
                targetStudents.set(studentObj.name, { color, note, isHighlighted });
            });
        }
        
        // Get existing students in this location
        const existingStudents = new Map();
        let scanIndex = currentIndex;
        while (scanIndex < body.getNumChildren()) {
            const child = body.getChild(scanIndex);
            
            // Stop at next heading
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
        
        // Remove students that are no longer in target
        existingStudents.forEach((index, studentName) => {
            if (!targetStudents.has(studentName)) {
                body.removeChild(body.getChild(index));
                // Adjust indices
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
                // Update existing student
                const listItem = body.getChild(existingIndex).asListItem();
                const expectedText = data.note ? `${studentName} [${data.note}]` : studentName;
                
                if (listItem.getText() !== expectedText) {
                    listItem.clear();
                    listItem.setText(studentName);
                    
                    if (data.color) listItem.setForegroundColor(data.color);
                    
                    if (data.note) {
                        const text = listItem.appendText(` [${data.note}]`);
                        text.setBold(true).setForegroundColor('#000000');
                        if (data.isHighlighted) text.setBackgroundColor('#ffff00');
                        else text.setBackgroundColor('#fff59d');
                    }
                }
                
                insertPosition = Math.max(insertPosition, existingIndex + 1);
            } else {
                // Add new student
                const li = body.insertListItem(insertPosition, studentName);
                
                if (data.color) li.setForegroundColor(data.color);
                
                if (data.note) {
                    const text = li.appendText(` [${data.note}]`);
                    text.setBold(true).setForegroundColor('#000000');
                    if (data.isHighlighted) text.setBackgroundColor('#ffff00');
                    else text.setBackgroundColor('#fff59d');
                }
                
                insertPosition++;
            }
        });
        
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
    const sheetsToProcess = Array.isArray(payload.sheets) ? payload.sheets : [payload.sheets];
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
  } catch (e) { return { success: false, message: e.message }; }
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
            try { tags = JSON.parse(row[4]); } catch (e) { } 
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
        if (String(data[i][0]) === uniqueId) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) sheet.getRange(rowIndex, 4, 1, 2).setValues([[payload.generalNotes, studentJson]]);
    else sheet.appendRow([uniqueId, payload.courseCode, payload.examName, payload.generalNotes, studentJson]);
    return { success: true, message: 'Saved to Database!' };
  } catch (e) { return { success: false, message: e.message }; }
}