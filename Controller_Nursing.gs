/**
 * -------------------------------------------------------------------
 * NURSING EXAM CONTROLLER
 * Features: Persistent ID Updates, Accommodation Preservation, Logging
 * -------------------------------------------------------------------
 */

function nursing_getSettingsAndIDs() {
  const settings = getSettings(CONFIG.SETTINGS_KEYS.NURSING);
  if (!settings.spreadsheetUrl || !settings.targetFolderId) {
    throw new Error('Nursing Settings are incomplete. Please go to Settings > Nursing Exams and save the URLs.');
  }
  
  const spreadsheetId = extractFileIdFromUrl(settings.spreadsheetUrl);
  const targetFolderId = extractFileIdFromUrl(settings.targetFolderId);
  
  if (!spreadsheetId) throw new Error('Invalid Nursing Sheet URL. Could not extract ID.');
  if (!targetFolderId) throw new Error('Invalid Target Folder URL. Could not extract ID.');
  
  return { settings, spreadsheetId, targetFolderId };
}

function nursing_generateAllDocuments() {
  try {
    const { settings, spreadsheetId, targetFolderId } = nursing_getSettingsAndIDs();
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let docsGenerated = 0;
    let docsUpdated = 0;
    let debugLog = [];

    const sheets = spreadsheet.getSheets();
    for (const sheet of sheets) {
      const allValues = sheet.getDataRange().getValues();
      if (!allValues || allValues.length === 0) continue;
      
      const discoveredExams = nursing_discoverExamsOnSheet(sheet, allValues);
      if (discoveredExams.length === 0) continue;

      const sheetTitle = allValues[0][0] || `Report for ${sheet.getName()}`;
      
      for (const exam of discoveredExams) {
        const docTitle = `${sheetTitle} - ${exam.name}`;
        const existingFiles = targetFolder.getFilesByName(docTitle);
        
        let doc;
        let actionType = "";
        let preservedAccommodations = "";

        if (existingFiles.hasNext()) {
          // UPDATE EXISTING (Persistent ID)
          const file = existingFiles.next();
          doc = DocumentApp.openById(file.getId());
          const body = doc.getBody();
          
          preservedAccommodations = nursing_findSectionText(body, "Accommodations");
          
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
        
        nursing_populateDocContent(doc.getBody(), docTitle, exam.data, sheet, settings.customNotes);
        
        if (preservedAccommodations) {
           const b = doc.getBody();
           let i = nursing_findInsertionIndex(b, ["General Notes", "Location Rosters"]);
           b.insertParagraph(i, "Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1);
           b.insertParagraph(i+1, preservedAccommodations);
        }

        doc.saveAndClose();
        
        logSystemAction("Nursing", actionType, docTitle, doc.getId(), `Exam Date: ${exam.data.date}`);
      }
      debugLog.push(`Processed ${sheet.getName()}: ${discoveredExams.length} exams.`);
    }
    return { data: `Success! Created ${docsGenerated}, Updated ${docsUpdated}.\n${debugLog.join('\n')}` };
  } catch (e) { return { error: e.message }; }
}

function nursing_discoverExamsOnSheet(sheet, allValues) { 
    const exams = [];
    let headerRowIndex = -1;
    for(let i=0; i<allValues.length; i++) {
        const rowStr = allValues[i].join(' ').toLowerCase();
        if(rowStr.includes('exam') && rowStr.includes('date')) { headerRowIndex = i; break; }
    }
    if (headerRowIndex === -1) return [];

    const headers = allValues[headerRowIndex].map(h => String(h).toLowerCase());
    const colMap = {
        name: headers.findIndex(h => h.includes('exam')),
        date: headers.findIndex(h => h.includes('date')),
        site: headers.findIndex(h => h.includes('site')), 
        zoom: headers.findIndex(h => h.includes('zoom')), 
        duration: headers.findIndex(h => h.includes('duration')),
        room: headers.findIndex(h => h.includes('room')),
        password: headers.findIndex(h => h.includes('password'))
    };
    if (colMap.name === -1 || colMap.date === -1) return [];

    const allFontLines = sheet.getDataRange().getFontLines();
    for (let i = headerRowIndex + 1; i < allValues.length; i++) {
        const row = allValues[i];
        if (String(row[0]).toLowerCase().trim() === CONFIG.NURSING.ROSTER_KEYWORD) break;
        const examName = String(row[colMap.name]).trim();
        if (!examName) continue; 
        if (allFontLines[i][colMap.date] === 'line-through') continue;
        
        exams.push({
            name: examName,
            data: {
                date: row[colMap.date],
                siteTime: (colMap.site > -1) ? row[colMap.site] : '',
                zoomTime: (colMap.zoom > -1) ? row[colMap.zoom] : '',
                duration: (colMap.duration > -1) ? row[colMap.duration] : '',
                room: (colMap.room > -1) ? row[colMap.room] : '',
                password: (colMap.password > -1) ? row[colMap.password] : ''
            }
        });
    }
    return exams;
}

function nursing_populateDocContent(body, docTitle, data, sheet, customNotes) {
  const allValues = sheet.getDataRange().getValues();
  const allFontColors = sheet.getDataRange().getFontColors();
  let rosterRow = -1;
  for(let i=0; i<allValues.length; i++) {
      if(String(allValues[i][0]).toLowerCase().trim() === CONFIG.NURSING.ROSTER_KEYWORD) { rosterRow = i; break; }
  }

  body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Exam Details').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const addDetail = (label, val, highlight) => {
      const item = body.appendListItem('');
      item.appendText(`${label}: `);
      let displayVal = val;
      if (!displayVal || displayVal === '') { if (label === 'Date') displayVal = "TBD"; else displayVal = " "; }
      const text = item.appendText(String(displayVal));
      if(highlight && displayVal.trim() !== "") text.setBackgroundColor('#FFFF00');
  };

  let dateStr = "";
  if (data.date) { try { dateStr = new Date(data.date).toLocaleDateString("en-US", {weekday:'long', month:'long', day:'numeric'}); } catch(e){ dateStr = String(data.date); } }

  addDetail('Date', dateStr, true);
  addDetail('Start Time (On Site)', data.siteTime, true);
  addDetail('Start Time (Zoom)', data.zoomTime, true);
  addDetail('Duration', data.duration, true);
  addDetail('Room', data.room, false);
  addDetail('Password', data.password, true);

  body.appendParagraph('');
  body.appendParagraph('Important Links').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('').appendText('Red Flag Reporting Form').setLinkUrl(CONFIG.NURSING.URLS.RED_FLAG_REPORT);
  body.appendParagraph('').appendText('Nursing Protocol').setLinkUrl(CONFIG.NURSING.URLS.PROTOCOL_DOC);
  body.appendParagraph('');

  if (rosterRow !== -1) {
      body.appendParagraph('Location Rosters').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      const headers = allValues[rosterRow];
      for (let c = 0; c < headers.length; c++) {
          const header = String(headers[c]).trim();
          if (!header) continue;
          body.appendParagraph(header).setHeading(DocumentApp.ParagraphHeading.HEADING2);
          for (let r = rosterRow + 3; r < allValues.length; r++) {
              const student = String(allValues[r][c]).trim();
              if (student) {
                  if (student.toLowerCase().includes('test time')) body.appendParagraph(student).setBackgroundColor('#DDEEFF');
                  else {
                      const li = body.appendListItem(student);
                      li.setGlyphType(DocumentApp.GlyphType.BULLET);
                      if(allFontColors[r][c] !== '#000000') li.setForegroundColor(allFontColors[r][c]);
                  }
              }
          }
      }
  }
  if (customNotes) {
      body.appendParagraph('');
      body.appendParagraph('General Notes').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(customNotes);
  }
}

function nursing_syncExamsToCalendar(calendarId, startStr, endStr, overwrite) {
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

        const { settings, spreadsheetId } = nursing_getSettingsAndIDs();
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheets = spreadsheet.getSheets();
        
        let count = 0;

        for (const sheet of sheets) {
            const allValues = sheet.getDataRange().getValues();
            const discovered = nursing_discoverExamsOnSheet(sheet, allValues);
            
            const sheetTitle = sheet.getName();

            for (const exam of discovered) {
                if (!exam.data.date) continue;
                const examDate = new Date(exam.data.date);
                
                if (examDate < startDate || examDate > endDate) continue;

                let timeStr = exam.data.siteTime || exam.data.zoomTime || "09:00";
                let hour = 9, min = 0;
                const timeMatch = timeStr.toString().match(/(\d+):(\d+)/);
                if (timeMatch) {
                     hour = parseInt(timeMatch[1]);
                     min = parseInt(timeMatch[2]);
                     if (timeStr.toLowerCase().includes('pm') && hour < 12) hour += 12;
                }

                const eventStart = new Date(examDate);
                eventStart.setHours(hour, min);
                
                const eventEnd = new Date(eventStart);
                eventEnd.setHours(hour + 2); 

                const title = `${sheetTitle} - ${exam.name}`;
                const location = exam.data.room || "TBD";
                const desc = `Password: ${exam.data.password}\nZoom: ${exam.data.zoomTime}`;

                const event = cal.createEvent(title, eventStart, eventEnd, { location: location, description: desc });
                event.setTag('AppSource', 'StaffHub');
                
                count++;
            }
        }
        
        logSystemAction("Nursing", "Calendar Sync", "N/A", calendarId, `Synced ${count} events.`);
        return { success: true, message: `Synced ${count} exams to calendar.` };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function nursing_getNotes() { 
    try { 
        return { data: getSettings(CONFIG.SETTINGS_KEYS.NURSING).customNotes || '' }; 
    } catch (e) { 
        return { error: e.message }; 
    } 
}

function nursing_saveNotes(t) { 
    try { 
        const s = getSettings(CONFIG.SETTINGS_KEYS.NURSING); 
        s.customNotes = t; 
        saveSettings(CONFIG.SETTINGS_KEYS.NURSING, s); 
        return { success: true }; 
    } catch (e) { 
        return { error: e.message }; 
    } 
}

function nursing_getAccommodations(id) { 
    try { 
        return { data: nursing_findSectionText(DocumentApp.openById(id).getBody(), "Accommodations") }; 
    } catch (e) { 
        return { error: e.message }; 
    } 
}

function nursing_saveAccommodations(id, t) { 
    try { 
        const d = DocumentApp.openById(id); 
        const b = d.getBody(); 
        nursing_removeSection(b, "Accommodations"); 
        if(t){ 
            const i = nursing_findInsertionIndex(b, ["Important Links"]); 
            b.insertParagraph(i, "Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1); 
            b.insertParagraph(i+1, t); 
        } 
        d.saveAndClose(); 
        return { success: true }; 
    } catch (e) { 
        return { error: e.message }; 
    } 
}

function nursing_getActiveDocsList() { 
    try { 
        const { targetFolderId } = nursing_getSettingsAndIDs(); 
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

function nursing_refreshAllActiveDocs() { 
    return { data: "Please use 'Generate All' to refresh content." }; 
}

function nursing_getDocsFromFolder() { 
    return nursing_getActiveDocsList(); 
}

function nursing_getEmailSettings() { 
    try { 
        const s = getSettings(CONFIG.SETTINGS_KEYS.NURSING); 
        return { 
            data: { 
                email: s.reportEmail || '', 
                isAuto: false, 
                reportDays: s.reportDaysForward || '7', 
                autoFrequency: s.autoTriggerFrequency || 'WEEKLY', 
                autoDay: s.autoTriggerDay || 'MONDAY', 
                autoHour: s.autoTriggerHour || '8', 
                reportResolvedComments: s.reportResolvedComments === 'true' 
            } 
        }; 
    } catch (e) { 
        return { error: e.message }; 
    } 
}

function nursing_saveEmailSettings(s) { 
    try { 
        const set = getSettings(CONFIG.SETTINGS_KEYS.NURSING); 
        set.reportEmail = s.email; 
        set.reportDaysForward = s.reportDays; 
        set.autoTriggerFrequency = s.autoFrequency; 
        set.autoTriggerDay = s.autoDay; 
        set.autoTriggerHour = s.autoHour; 
        set.reportResolvedComments = s.reportResolvedComments; 
        saveSettings(CONFIG.SETTINGS_KEYS.NURSING, set); 
        return { data: "Saved." }; 
    } catch (e) { 
        return { error: e.message }; 
    } 
}

function nursing_regenerateSingleDocById(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const targetTitle = doc.getName();
    
    const { settings, spreadsheetId } = nursing_getSettingsAndIDs();
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    let foundMatch = false;

    for (const sheet of sheets) {
      const allValues = sheet.getDataRange().getValues();
      if (!allValues || allValues.length === 0) continue;
      
      const discoveredExams = nursing_discoverExamsOnSheet(sheet, allValues);
      const sheetTitle = allValues[0][0] || `Report for ${sheet.getName()}`;

      for (const exam of discoveredExams) {
        const generatedTitle = `${sheetTitle} - ${exam.name}`;
        
        if (generatedTitle === targetTitle) {
          foundMatch = true;
          
          const body = doc.getBody();
          const preservedAccommodations = nursing_findSectionText(body, "Accommodations");
          
          body.setText('');
          nursing_populateDocContent(body, targetTitle, exam.data, sheet, settings.customNotes);
          
          if (preservedAccommodations) {
             let i = nursing_findInsertionIndex(body, ["General Notes", "Location Rosters"]);
             body.insertParagraph(i, "Accommodations").setHeading(DocumentApp.ParagraphHeading.HEADING1);
             body.insertParagraph(i+1, preservedAccommodations);
          }
          
          doc.saveAndClose();
          logSystemAction("Nursing", "Regenerated Single", targetTitle, docId, `Exam Date: ${exam.data.date}`);
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

function nursing_removeSection(b, h) {
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

function nursing_findSectionText(body, heading) {
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
        if (paragraphs[i].getHeading() == DocumentApp.ParagraphHeading.HEADING1 && paragraphs[i].getText() == heading) {
            return (i + 1 < paragraphs.length) ? paragraphs[i + 1].getText() : "";
        }
    }
    return "";
}

function nursing_findInsertionIndex(body, possibleHeaders) {
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