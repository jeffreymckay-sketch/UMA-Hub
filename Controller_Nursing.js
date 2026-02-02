/**
 * -------------------------------------------------------------------
 * NURSING PROCTORING CONTROLLER
 * Handles: Data Fetching, Document Generation, Calendar Sync, and Logging
 * Features: Deep-Search Folder Logic, Fuzzy Matching, Roster Management
 * -------------------------------------------------------------------
 */

/**
 * HELPER: Centralized Settings Parser
 */
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

/**
 * HELPER: Normalizes strings for fuzzy comparison
 * Removes special characters and converts to lowercase.
 */
function _normalizeStr(str) {
  if (!str) return "";
  return str.toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * HELPER: Parses Cell A1 to extract Folder Name and Full Title prefix
 * Input: "NUR 220 Concepts: Professor Erin Voisine" or "NUR 220 Concepts - Professor Erin Voisine"
 * Returns: { folderKey: "NUR 220 Concepts", fullTitlePrefix: "NUR 220 Concepts Professor Erin Voisine" }
 */
function _parseCellA1(cellValue) {
  if (!cellValue) return { folderKey: "Unknown", fullTitlePrefix: "Unknown" };
  
  const str = String(cellValue).trim();
  
  // Try to split by colon first (most common), then hyphen
  let separator = ':';
  let parts = str.split(':');
  
  if (parts.length === 1) {
    // No colon found, try hyphen
    separator = '-';
    parts = str.split('-');
  }
  
  if (parts.length > 1) {
    const courseInfo = parts[0].trim(); // "NUR 220 Concepts"
    const facultyInfo = parts.slice(1).join(separator).trim(); // "Professor Erin Voisine"
    
    // Build title prefix matching the old document naming convention (space separator, no colon/hyphen)
    const fullTitlePrefix = `${courseInfo} ${facultyInfo}`;
    
    return {
      folderKey: courseInfo,
      fullTitlePrefix: fullTitlePrefix
    };
  }
  
  // Fallback if no separator found
  return { folderKey: str, fullTitlePrefix: str };
}

/**
 * HELPER: Recursively finds a subfolder by fuzzy name match
 */
function _findSubFolder(parentFolder, targetName) {
  const targetNorm = _normalizeStr(targetName);
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (_normalizeStr(folder.getName()) === targetNorm) {
      return folder;
    }
  }
  return null;
}

/**
 * Main API called by the frontend.
 * Fetches sheet data and resolves existing document URLs using deep search.
 */
function api_getNursingData() {
  try {
    const config = _getNursingSettings();
    if (!config.sheetId || !config.folderId) {
      return { success: false, message: "Nursing settings missing. Please save Sheet/Folder IDs in Settings." };
    }

    const ss = SpreadsheetApp.openById(config.sheetId);
    const sheets = ss.getSheets();
    const rootFolder = DriveApp.getFolderById(config.folderId);
    
    // Cache Accommodations DB
    const dbMap = getAccommodationsDBMap();

    const courseData = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("_")) return; 

      const parsed = parseNursingSheet(sheet, dbMap);
      if (!parsed) return;

      // --- DEEP SEARCH LOGIC ---
      // 1. Determine expected folder name from A1
      const meta = _parseCellA1(parsed.rawA1);
      
      console.log("=== Processing Sheet:", sheetName, "===");
      console.log("📋 Raw A1 cell:", parsed.rawA1);
      console.log("📁 Extracted folder key:", meta.folderKey);
      console.log("📄 Full title prefix:", meta.fullTitlePrefix);
      
      // 2. Try to find the specific sub-folder for this class
      let targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      
      // 3. Scan files in that folder (if it exists)
      const existingFiles = [];
      if (targetFolder) {
        console.log("📁 Found target folder:", targetFolder.getName());
        const files = targetFolder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          existingFiles.push({
            name: f.getName(),
            nameNorm: _normalizeStr(f.getName()),
            url: f.getUrl()
          });
        }
        console.log("📄 Files in folder:", existingFiles.length);
        existingFiles.forEach(f => console.log("  -", f.name, "→", f.nameNorm));
      } else {
        console.log("⚠️ Target folder NOT FOUND:", meta.folderKey);
      }

      // 4. Attach Doc URLs based on improved fuzzy match
      parsed.exams.forEach(exam => {
        const examNameNorm = _normalizeStr(exam.name);
        
        // Build the expected full document title
        const expectedTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        const expectedTitleNorm = _normalizeStr(expectedTitle);
        
        console.log("🔍 Looking for exam:", exam.name);
        console.log("   Expected title:", expectedTitle);
        console.log("   Normalized:", expectedTitleNorm);
        
        // Priority 1: Try exact match on full expected title
        let match = existingFiles.find(f => f.nameNorm === expectedTitleNorm);
        
        // Priority 2: Fallback to partial match (file contains exam name)
        if (!match) {
          match = existingFiles.find(f => f.nameNorm.includes(examNameNorm));
          if (match) {
            console.log("   ✓ Found via partial match:", match.name);
          }
        } else {
          console.log("   ✓ Found via exact match:", match.name);
        }
        
        if (!match) {
          console.log("   ✗ No match found for:", exam.name);
        }
        
        exam.docUrl = match ? match.url : null;
      });

      if (parsed.exams.length > 0) {
        courseData.push({
          sheetName: sheetName,
          course: parsed.course,
          exams: parsed.exams,
          // Pass metadata for creation/updates
          _meta: meta 
        });
      }
    });

    return { success: true, data: { sheets: courseData, settings: config } };

  } catch (e) {
    console.error("ERROR in api_getNursingData:", e.message, e.stack);
    return { success: false, message: e.message + " (Stack: " + e.stack + ")" };
  }
}

/**
 * CORE PARSER: Handles the "Augusta Anchor" and 2-row skip
 */
function parseNursingSheet(sheet, dbMap) {
  const range = sheet.getDataRange();
  const data = range.getValues();
  const fontColors = range.getFontColors();
  const fontLines = range.getFontLines();
  
  if (data.length < 5) return null; 

  // 1. Course Info
  const a1 = String(data[0][0]).trim(); 
  let courseCode = "Unknown";
  let courseName = a1;
  const splitMatch = a1.match(/^([^:-]+)[:\s-](.+)/);
  if (splitMatch) {
      courseCode = splitMatch[1].trim();
      courseName = splitMatch[2].trim();
  }

  // 2. Find Exam Header Row
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

  // 3. Find Roster Anchor ("Augusta")
  let rosterHeaderIdx = -1;
  for(let i = headerRowIndex + 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === 'augusta') {
          rosterHeaderIdx = i;
          break;
      }
  }

  // 4. Parse Rosters (Bottom Section)
  const sheetRoster = {}; // Raw roster from sheet
  if (rosterHeaderIdx !== -1) {
      const locHeaders = data[rosterHeaderIdx];
      const rosterStartRow = rosterHeaderIdx + 3; // Skip 2 rows
      
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

  // 5. Parse Exams
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
      
      // Check Strikethrough
      if (fontLines[i][colMap.date] === 'line-through' || fontLines[i][colMap.name] === 'line-through') continue;

      const dateVal = parseFlexibleDate(row[colMap.date]);
      const safeNormalize = (typeof normalizeTime === 'function') ? normalizeTime : String;

      // DB Lookup
      const dbKey = `${courseCode}|${String(examName).trim()}`; 
      const dbEntry = dbMap[dbKey] || { generalNotes: "", studentTags: {} };

      exams.push({
          name: String(examName).trim(),
          date: dateVal, 
          siteTime: safeNormalize(row[colMap.timeSite]), 
          zoomTime: safeNormalize(row[colMap.timeZoom]), 
          duration: colMap.duration > -1 ? row[colMap.duration] : '',
          room: colMap.room > -1 ? row[colMap.room] : '',
          password: colMap.password > -1 ? row[colMap.password] : '',
          generalNotes: dbEntry.generalNotes || (colMap.accommodations > -1 ? row[colMap.accommodations] : ''),
          studentTags: dbEntry.studentTags, // This now contains objects {note, excluded, highlighted, locked}
          rosters: sheetRoster // Attach raw roster, filtering happens in UI/DocGen
      });
  }

  return {
      rawA1: a1,
      course: { code: courseCode, name: courseName },
      exams: exams
  };
}

/**
 * Creates Google Docs for selected exams using intelligent folder management.
 */
function createNursingProctoringDocuments(payload) {
  try {
    const config = _getNursingSettings();
    if (!config.folderId) return { success: false, message: "Folder ID missing." };
    
    const rootFolder = DriveApp.getFolderById(config.folderId);
    let createdCount = 0;
    const createdUrls = {}; // Track URLs for response

    payload.sheets.forEach(sheetData => {
      // 1. Resolve Target Folder (Create if missing)
      const meta = sheetData._meta || _parseCellA1(sheetData.course.name); // Fallback
      
      let targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      if (!targetFolder) {
        targetFolder = rootFolder.createFolder(meta.folderKey);
        console.log("📁 Created new folder:", meta.folderKey);
      }

      // 2. Create Docs in Target Folder
      sheetData.exams.forEach(exam => {
        // Build Title: [Course Info] [Faculty] - [Exam Name]
        // Matches existing naming: "NUR 220 Concepts Professor Erin Voisine - Exam 1"
        const docTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        
        // Double check existence inside target folder to prevent duplicates
        const existing = targetFolder.getFilesByName(docTitle);
        if (existing.hasNext()) {
          console.log("⚠️ Doc already exists, skipping:", docTitle);
          return;
        }

        const doc = DocumentApp.create(docTitle);
        _populateNursingDoc(doc, docTitle, exam, config.customNotes);

        const file = DriveApp.getFileById(doc.getId());
        targetFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
        
        // Track the new URL
        const newUrl = file.getUrl();
        createdUrls[exam.name] = newUrl;
        
        console.log("✓ Created doc:", docTitle);
        
        if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Created Doc", docTitle, doc.getId(), `Date: ${exam.date}`);
        createdCount++;
      });
    });

    return { success: true, message: `Created ${createdCount} documents.`, createdUrls: createdUrls };
  } catch (e) { 
    console.error("ERROR in createNursingProctoringDocuments:", e.message, e.stack);
    return { success: false, message: e.message }; 
  }
}

/**
 * Updates existing Google Docs.
 */
function updateAllNursingDocuments(payload) {
  try {
    const config = _getNursingSettings();
    if (!config.folderId) return { success: false, message: "Folder ID missing." };

    const rootFolder = DriveApp.getFolderById(config.folderId);
    let updatedCount = 0;

    payload.sheets.forEach(sheetData => {
      // 1. Resolve Target Folder
      const meta = sheetData._meta || _parseCellA1(sheetData.course.name);
      const targetFolder = _findSubFolder(rootFolder, meta.folderKey);
      
      if (!targetFolder) {
        console.log("⚠️ Cannot update - folder not found:", meta.folderKey);
        return; // Cannot update if folder doesn't exist
      }

      sheetData.exams.forEach(exam => {
        // 2. Find File using improved fuzzy match
        const examNameNorm = _normalizeStr(exam.name);
        const expectedTitle = `${meta.fullTitlePrefix} - ${exam.name}`;
        const expectedTitleNorm = _normalizeStr(expectedTitle);
        
        let targetFile = null;

        // Priority 1: Exact URL match (if provided by frontend)
        if (exam.docUrl) {
          try { 
            targetFile = DriveApp.getFileByUrl(exam.docUrl);
            console.log("✓ Found via URL:", exam.docUrl);
          } catch(e) {
            console.log("⚠️ URL lookup failed:", e.message);
          }
        }

        // Priority 2: Exact title match
        if (!targetFile) {
          const files = targetFolder.getFiles();
          while (files.hasNext()) {
            const f = files.next();
            if (_normalizeStr(f.getName()) === expectedTitleNorm) {
              targetFile = f;
              console.log("✓ Found via exact title match:", f.getName());
              break;
            }
          }
        }

        // Priority 3: Partial match (contains exam name)
        if (!targetFile) {
          const files = targetFolder.getFiles();
          while (files.hasNext()) {
            const f = files.next();
            if (_normalizeStr(f.getName()).includes(examNameNorm)) {
              targetFile = f;
              console.log("✓ Found via partial match:", f.getName());
              break;
            }
          }
        }
        
        if (targetFile) {
          const doc = DocumentApp.openById(targetFile.getId());
          doc.getBody().clear();
          // Use current file name to preserve any manual renaming user might have done
          _populateNursingDoc(doc, targetFile.getName(), exam, config.customNotes);
          
          console.log("✓ Updated doc:", targetFile.getName());
          
          if (typeof logSystemAction === 'function') logSystemAction("Nursing", "Updated Doc", targetFile.getName(), targetFile.getId(), `Date: ${exam.date}`);
          updatedCount++;
        } else {
          console.log("✗ No doc found to update for exam:", exam.name);
        }
      });
    });

    return { success: true, message: `Updated ${updatedCount} documents.` };
  } catch (e) { 
    console.error("ERROR in updateAllNursingDocuments:", e.message, e.stack);
    return { success: false, message: e.message }; 
  }
}

/**
 * SHARED DOC POPULATOR (The Magic Happens Here)
 */
function _populateNursingDoc(doc, title, exam, customNotes) {
    const body = doc.getBody();
    
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(`Date: ${exam.date}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // --- DETAILS TABLE ---
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell("Detail").setBackgroundColor('#f3f3f3').setBold(true);
    headerRow.appendTableCell("Value").setBackgroundColor('#f3f3f3').setBold(true);
    
    table.appendTableRow().appendTableCell("Site Time").getParent().appendTableCell(exam.siteTime || "-");
    table.appendTableRow().appendTableCell("Zoom Time").getParent().appendTableCell(exam.zoomTime || "-");
    const pwCell = table.appendTableRow().appendTableCell("Password").getParent().appendTableCell(exam.password || "-");
    pwCell.setBackgroundColor('#fff8e1'); 

    if (customNotes) {
        body.appendParagraph("General Instructions").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        body.appendParagraph(customNotes).setItalic(true);
    }

    if (exam.generalNotes) {
        body.appendParagraph("Exam Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        const p = body.appendParagraph(exam.generalNotes);
        p.setBackgroundColor('#e8f5e9'); 
        p.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(10).setPaddingRight(10);
    }

    // --- ROSTERS WITH EXCLUSION & HIGHLIGHTING ---
    if (exam.rosters && Object.keys(exam.rosters).length > 0) {
        body.appendParagraph("Rosters & Specific Needs").setHeading(DocumentApp.ParagraphHeading.HEADING1);
        
        for (const [location, students] of Object.entries(exam.rosters)) {
            // Filter students first
            const activeStudents = students.filter(s => {
                const tagData = (exam.studentTags && exam.studentTags[s.name]) ? exam.studentTags[s.name] : null;
                // If tagData is an object, check excluded. If string (old data), assume included.
                if (tagData && typeof tagData === 'object' && tagData.excluded) return false;
                return true;
            });

            if(activeStudents.length > 0) {
                body.appendParagraph(location).setHeading(DocumentApp.ParagraphHeading.HEADING2);
                
                activeStudents.forEach(studentObj => {
                    const li = body.appendListItem(studentObj.name);
                    
                    // Get Saved Data
                    let note = "";
                    let isHighlighted = false;
                    let isLocked = false;

                    if (exam.studentTags && exam.studentTags[studentObj.name]) {
                        const tag = exam.studentTags[studentObj.name];
                        if (typeof tag === 'object') {
                            note = tag.note || "";
                            isHighlighted = tag.highlighted || false;
                            isLocked = tag.locked || false;
                        } else {
                            note = tag; // Backward compatibility
                        }
                    }

                    // Apply Color (Sheet Color vs Lock)
                    if (isLocked) {
                        // If locked, we ignore sheet color (default black)
                        li.setForegroundColor('#000000'); 
                    } else if (studentObj.color && studentObj.color !== '#000000') {
                        li.setForegroundColor(studentObj.color);
                    }

                    // Append Note
                    if (note) {
                        const tagText = ` [${note}]`;
                        const text = li.appendText(tagText);
                        text.setBold(true);
                        text.setForegroundColor('#000000'); // Reset color for the tag text
                        
                        if (isHighlighted) {
                            text.setBackgroundColor('#ffff00'); // Yellow Highlight
                        } else {
                            text.setBackgroundColor('#fff59d'); // Default subtle yellow
                        }
                    }
                });
            }
        }
    }

    body.appendHorizontalRule();
    body.appendParagraph("Generated by Staff Hub").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#888888').setFontSize(8);
    
    doc.saveAndClose();
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
        const parts = exam.date.split('/');
        if (parts.length !== 3) return;
        const start = new Date(parts[2], parts[0]-1, parts[1], 8, 0, 0); 
        const end = new Date(start);
        end.setHours(17, 0, 0); 

        // Use meta full title if available for better calendar naming
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

// --- DATABASE HELPERS ---

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

function parseFlexibleDate(input) {
    if (!input) return null;
    if (input instanceof Date) return Utilities.formatDate(input, Session.getScriptTimeZone(), "yyyy-MM-dd");
    let str = String(input).trim();
    str = str.replace(/(\d+)(st|nd|rd|th)/ig, "$1");
    const d = new Date(str);
    if (isNaN(d.getTime())) return null;
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function extractIdFromUrl(url) {
    if (!url) return null;
    const match = url.match(/[-\w]{25,}/);
    return match ? match[0] : url;
}