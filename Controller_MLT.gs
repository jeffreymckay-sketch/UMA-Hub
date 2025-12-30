/**
 * -------------------------------------------------------------------
 * MLT PROCTORING CONTROLLER
 * Features: Persistent ID Updates, Accommodation Preservation, Logging
 *           Time & Duration Parsing for Calendar Sync
 * -------------------------------------------------------------------
 */

function mlt_getSettingsAndIDs() {
  const settings = getSettings(CONFIG.SETTINGS_KEYS.MLT);
  
  if (!settings.config) {
    settings.config = CONFIG.MLT.DEFAULTS.KEYWORDS;
    settings.rosterKeyword = CONFIG.MLT.DEFAULTS.ROSTER_KEYWORD;
  }
  
  if (!settings.spreadsheetUrl || !settings.targetFolderId) {
    throw new Error('MLT Settings are incomplete. Please go to MLT Settings to save URLs.');
  }

  const spreadsheetId = extractFileIdFromUrl(settings.spreadsheetUrl);
  const targetFolderId = extractFileIdFromUrl(settings.targetFolderId);
  
  if (!spreadsheetId) throw new Error('Invalid MLT Sheet URL.');
  if (!targetFolderId) throw new Error('Invalid Target Folder URL.');
  
  return { settings, spreadsheetId, targetFolderId };
}

function mlt_generateAllDocuments() {
  try {
    const { settings, spreadsheetId, targetFolderId } = mlt_getSettingsAndIDs();
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let docsGenerated = 0;
    let docsUpdated = 0;
    let debugLog = [];

    const sheets = spreadsheet.getSheets();
    for (const sheet of sheets) {
      const allValues = sheet.getDataRange().getValues();
      if (!allValues || allValues.length === 0) continue;

      const discoveredExams = mlt_discoverExamsOnSheet(sheet, allValues, settings);
      
      if (discoveredExams.length === 0) continue;

      // MLT Title Logic: Cell D1 (index 3) or Sheet Name
      const sheetTitle = (allValues[0] && allValues[0][3]) ? allValues[0][3] : sheet.getName();
      
      for (const exam of discoveredExams) {
        const docTitle = `${sheetTitle} - ${exam.name}`;
        const existingFiles = targetFolder.getFilesByName(docTitle);
        
        let doc;
        let actionType = "";
        let preservedAccommodations = "";

        if (existingFiles.hasNext()) {
          // UPDATE EXISTING
          const file = existingFiles.next();
          doc = DocumentApp.openById(file.getId());
          const body = doc.getBody();
          
          // 1. Preserve Accommodations
          preservedAccommodations = mlt_findSectionText(body, "Accommodations");
          
          // 2. Clear Body (Safe Method)
          body.setText('');
          
          actionType = "Updated";
          docsUpdated++;
        } else {
          // CREATE NEW
          doc = DocumentApp.create(docTitle);
          const newDocFile = DriveApp.getFileById(doc.getId());
          targetFolder.addFile(newDocFile);
          DriveApp.getRootFolder().removeFile(newDocFile);
          
          actionType = "Created";
          docsGenerated++;
        }
        
        // 3. Populate Content
        mlt_populateDocContent(doc.getBody(), docTitle, exam.data, sheet, settings);
        
        // 4. Restore Accommodations
        if (preservedAccommodations) {
           const b = doc.getBody();
           let i = mlt_findInsertionIndex(b, ["Notes", "General Notes", "Rosters", "Location Rosters"]);
           b.insertParagraph(i, "Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1);
           b.insertParagraph(i+1, preservedAccommodations);
        }

        doc.saveAndClose();
        
        // 5. Log Action
        logSystemAction("MLT", actionType, docTitle, doc.getId(), `Exam Date: ${exam.data.date}`);
      }
      debugLog.push(`Processed ${sheet.getName()}: ${discoveredExams.length} exams.`);
    }
    return { data: `Success! Created ${docsGenerated}, Updated ${docsUpdated}.\n${debugLog.join('\n')}` };
  } catch (e) { return { error: e.message }; }
}

function mlt_discoverExamsOnSheet(sheet, allValues, settings) {
    const exams = [];
    const config = settings.config; 
    const rosterKeyword = (settings.rosterKeyword || CONFIG.MLT.DEFAULTS.ROSTER_KEYWORD).toLowerCase();
    
    let headerRowIndex = -1;
    
    for(let i=0; i<allValues.length; i++) {
        const rowStr = allValues[i].join(' ').toLowerCase();
        if(rowStr.includes(config.EXAM) && rowStr.includes(config.DATE)) { 
            headerRowIndex = i; 
            break; 
        }
    }
    if (headerRowIndex === -1) return [];

    const headers = allValues[headerRowIndex].map(h => String(h).toLowerCase());
    
    const colMap = {};
    for (const [key, keyword] of Object.entries(config)) {
        colMap[key] = headers.findIndex(h => h.includes(keyword.toLowerCase()));
    }
    
    if (colMap.EXAM === -1 || colMap.DATE === -1) return [];

    const allFontLines = sheet.getDataRange().getFontLines();

    for (let i = headerRowIndex + 1; i < allValues.length; i++) {
        const row = allValues[i];
        if (String(row[0]).toLowerCase().trim().includes(rosterKeyword)) break;
        
        const examName = String(row[colMap.EXAM]).trim();
        if (!examName) continue; 
        
        if (allFontLines[i][colMap.DATE] === 'line-through') continue;
        
        let startTimeVal = '';
        if (colMap.START_TIME > -1) startTimeVal = row[colMap.START_TIME];
        else if (colMap.START_SITE > -1) startTimeVal = row[colMap.START_SITE];

        exams.push({
            name: examName,
            data: {
                date: row[colMap.DATE],
                startTime: startTimeVal,
                duration: (colMap.DURATION > -1) ? row[colMap.DURATION] : '',
                room: (colMap.ROOM > -1) ? row[colMap.ROOM] : '',
                password: (colMap.PASSWORD > -1) ? row[colMap.PASSWORD] : ''
            }
        });
    }
    return exams;
}

function mlt_populateDocContent(body, docTitle, data, sheet, settings) {
  const allValues = sheet.getDataRange().getValues();
  const allFontColors = sheet.getDataRange().getFontColors();
  const rosterKeyword = (settings.rosterKeyword || CONFIG.MLT.DEFAULTS.ROSTER_KEYWORD).toLowerCase();
  
  let rosterRow = -1;
  for(let i=0; i<allValues.length; i++) {
      if(String(allValues[i][0]).toLowerCase().trim().includes(rosterKeyword)) { rosterRow = i; break; }
  }

  body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Exam Details').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const addDetail = (label, val, highlight) => {
      const item = body.appendListItem('');
      item.appendText(`${label}: `);
      let displayVal = val;
      if (!displayVal || displayVal === '') { displayVal = (label === 'Date') ? "TBD" : " "; }
      
      if (displayVal instanceof Date) {
          if (label === 'Start Time') {
             displayVal = displayVal.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
          } else {
             displayVal = displayVal.toLocaleDateString();
          }
      }
      
      const text = item.appendText(String(displayVal));
      if(highlight && String(displayVal).trim() !== "") text.setBackgroundColor('#FFFF00');
  };

  let dateStr = "";
  if (data.date) { try { dateStr = new Date(data.date).toLocaleDateString("en-US", {weekday:'long', month:'long', day:'numeric'}); } catch(e){ dateStr = String(data.date); } }

  addDetail('Date', dateStr, true);
  addDetail('Start Time', data.startTime, true);
  addDetail('Duration', data.duration, true);
  addDetail('Room', data.room, false);
  addDetail('Password', data.password, true);

  if (settings.customNotes) {
      body.appendParagraph('');
      body.appendParagraph('Notes').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(settings.customNotes);
  }

  if (rosterRow !== -1) {
      body.appendParagraph('Rosters').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      const headers = allValues[rosterRow];
      for (let c = 0; c < headers.length; c++) {
          const header = String(headers[c]).trim();
          if (!header) continue;
          body.appendParagraph(header).setHeading(DocumentApp.ParagraphHeading.HEADING2);
          for (let r = rosterRow + 3; r < allValues.length; r++) {
              const student = String(allValues[r][c]).trim();
              if (student) {
                   const li = body.appendListItem(student);
                   li.setGlyphType(DocumentApp.GlyphType.BULLET);
                   if(allFontColors[r][c] !== '#000000') li.setForegroundColor(allFontColors[r][c]);
              }
          }
      }
  }
}

function mlt_syncExamsToCalendar(calendarId, startStr, endStr, overwrite) {
    try {
        if (!calendarId) throw new Error("Calendar ID missing.");
        const cal = CalendarApp.getCalendarById(calendarId);
        if (!cal) throw new Error("Calendar not found. Check permissions or ID.");
        
        const startDate = new Date(startStr);
        const endDate = new Date(endStr);
        endDate.setHours(23, 59, 59);

        if (overwrite) {
            const events = cal.getEvents(startDate, endDate);
            events.forEach(e => {
                if (e.getTag('AppSource') === 'StaffHub') {
                    e.deleteEvent();
                }
            });
        }

        const { settings, spreadsheetId } = mlt_getSettingsAndIDs();
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheets = spreadsheet.getSheets();
        
        let count = 0;

        for (const sheet of sheets) {
            const allValues = sheet.getDataRange().getValues();
            const discovered = mlt_discoverExamsOnSheet(sheet, allValues, settings);
            
            const sheetTitle = sheet.getName();

            for (const exam of discovered) {
                if (!exam.data.date) continue;
                const examDate = new Date(exam.data.date);
                if (examDate < startDate || examDate > endDate) continue;

                let hour = 9, min = 0;
                const rawTime = exam.data.startTime;
                
                if (rawTime instanceof Date) {
                    hour = rawTime.getHours();
                    min = rawTime.getMinutes();
                } else if (rawTime) {
                    const timeStr = String(rawTime);
                    const timeMatch = timeStr.match(/(\d+):(\d+)/);
                    if (timeMatch) {
                        hour = parseInt(timeMatch[1]);
                        min = parseInt(timeMatch[2]);
                        if (timeStr.toLowerCase().includes('pm') && hour < 12) hour += 12;
                    }
                }

                let durationMinutes = 120;
                const rawDur = exam.data.duration;
                if (rawDur) {
                    const durStr = String(rawDur).trim();
                    if (durStr.includes(':')) {
                        const parts = durStr.split(':');
                        durationMinutes = (parseInt(parts[0]) * 60) + parseInt(parts[1]);
                    } else if (durStr.toLowerCase().includes('h')) {
                        durationMinutes = parseFloat(durStr) * 60;
                    } else {
                        const val = parseFloat(durStr);
                        if (!isNaN(val)) {
                            durationMinutes = (val <= 8) ? val * 60 : val;
                        }
                    }
                }

                const eventStart = new Date(examDate);
                eventStart.setHours(hour, min);
                
                const eventEnd = new Date(eventStart.getTime() + (durationMinutes * 60000));

                const title = `${sheetTitle} - ${exam.name}`;
                const location = exam.data.room || "TBD";
                const desc = `Password: ${exam.data.password}\nDuration: ${exam.data.duration}`;

                const event = cal.createEvent(title, eventStart, eventEnd, { location: location, description: desc });
                event.setTag('AppSource', 'StaffHub');
                
                count++;
            }
        }
        
        logSystemAction("MLT", "Calendar Sync", "N/A", calendarId, `Synced ${count} events.`);
        return { success: true, message: `Synced ${count} exams to calendar.` };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function mlt_getActiveDocsList() { 
    try { 
        const { targetFolderId } = mlt_getSettingsAndIDs(); 
        const list = []; 
        const files = DriveApp.getFolderById(targetFolderId).getFiles(); 
        while(files.hasNext()) { 
            const f = files.next(); 
            if(f.getMimeType() === MimeType.GOOGLE_DOCS) list.push({name: f.getName(), id: f.getId(), url: f.getUrl()}); 
        } 
        return { data: list }; 
    } catch(e) { 
        return { error: e.message }; 
    } 
}

function mlt_refreshAllActiveDocs() { 
  return mlt_generateAllDocuments(); 
}

function mlt_getDocsFromFolder() { 
    return mlt_getActiveDocsList(); 
}

function mlt_saveSettings(settingsObj) {
    try {
        const current = getSettings(CONFIG.SETTINGS_KEYS.MLT);
        const newSettings = { ...current, ...settingsObj };
        saveSettings(CONFIG.SETTINGS_KEYS.MLT, newSettings);
        return { success: true, message: "MLT Settings Saved." };
    } catch (e) { return { success: false, message: e.message }; }
}

function mlt_getSettingsData() {
    try {
        const s = getSettings(CONFIG.SETTINGS_KEYS.MLT);
        if (!s.config) s.config = CONFIG.MLT.DEFAULTS.KEYWORDS;
        if (!s.rosterKeyword) s.rosterKeyword = CONFIG.MLT.DEFAULTS.ROSTER_KEYWORD;
        return { success: true, data: s };
    } catch (e) { return { success: false, message: e.message }; }
}

function mlt_regenerateSingleDocById(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const targetTitle = doc.getName();
    
    const { settings, spreadsheetId } = mlt_getSettingsAndIDs();
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    let foundMatch = false;

    for (const sheet of sheets) {
      const allValues = sheet.getDataRange().getValues();
      if (!allValues || allValues.length === 0) continue;
      
      const discoveredExams = mlt_discoverExamsOnSheet(sheet, allValues, settings);
      const sheetTitle = (allValues[0] && allValues[0][3]) ? allValues[0][3] : sheet.getName();

      for (const exam of discoveredExams) {
        const generatedTitle = `${sheetTitle} - ${exam.name}`;
        
        if (generatedTitle === targetTitle) {
          foundMatch = true;
          
          const body = doc.getBody();
          const preservedAccommodations = mlt_findSectionText(body, "Accommodations");
          
          body.setText('');
          mlt_populateDocContent(body, targetTitle, exam.data, sheet, settings);
          
          if (preservedAccommodations) {
             let i = mlt_findInsertionIndex(body, ["Notes", "General Notes", "Rosters", "Location Rosters"]);
             body.insertParagraph(i, "Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1);
             body.insertParagraph(i+1, preservedAccommodations);
          }
          
          doc.saveAndClose();
          logSystemAction("MLT", "Regenerated Single", targetTitle, docId, `Exam Date: ${exam.data.date}`);
          break; 
        }
      }
      if (foundMatch) break;
    }

    if (!foundMatch) {
      return { error: "Could not find matching data in the spreadsheet. Has the Exam Name or Sheet Title changed?" };
    }

    return { data: "Document updated successfully." };

  } catch (e) {
    return { error: e.message };
  }
}

function mlt_getAccommodations(id) { 
    try { 
        return { data: mlt_findSectionText(DocumentApp.openById(id).getBody(), "Accommodations") }; 
    } catch (e) { return { error: e.message }; } 
}

function mlt_saveAccommodations(id, t) { 
    try { 
        const d = DocumentApp.openById(id); 
        const b = d.getBody(); 
        mlt_removeSection(b, "Accommodations"); 
        if(t){ 
            let i = mlt_findInsertionIndex(b, ["Notes", "General Notes", "Rosters", "Location Rosters"]); 
            b.insertParagraph(i, "Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1); 
            b.insertParagraph(i+1, t); 
        } 
        d.saveAndClose(); 
        return { success: true }; 
    } catch (e) { return { error: e.message }; } 
}

function mlt_removeSection(b, h) {
  const p = b.getParagraphs();
  for (let i = 0; i < p.length; i++) {
    if (p[i].getHeading() == DocumentApp.ParagraphHeading.HEADING1 && p[i].getText() == h) {
      let el = p[i];
      while (el && (el.getType() != DocumentApp.ElementType.PARAGRAPH || el.asParagraph().getHeading() != DocumentApp.ParagraphHeading.HEADING1)) {
        let next = el.getNextSibling();
        if (b.getNumChildren() === 1) b.appendParagraph(""); 
        b.removeChild(el);
        el = next;
      }
      return;
    }
  }
}

function mlt_findSectionText(body, heading) {
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
        if (paragraphs[i].getHeading() == DocumentApp.ParagraphHeading.HEADING1 && paragraphs[i].getText() == heading) {
            return (i + 1 < paragraphs.length) ? paragraphs[i + 1].getText() : "";
        }
    }
    return "";
}

function mlt_findInsertionIndex(body, possibleHeaders) {
    for (const t of possibleHeaders) {
        for (let i = 0; i < body.getNumChildren(); i++) {
            const e = body.getChild(i);
            if (e.getType() == DocumentApp.ElementType.PARAGRAPH && e.asParagraph().getHeading() == DocumentApp.ParagraphHeading.HEADING1 && e.getText() == t) {
                return i;
            }
        }
    }
    return body.getNumChildren();
}