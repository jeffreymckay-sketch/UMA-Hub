/**
 * -------------------------------------------------------------------
 * NURSING EXAM CONTROLLER (REFACTORED for Multi-Sheet Support)
 * Features: Interactive analysis, robust parsing, and controlled document generation.
 * -------------------------------------------------------------------
 */

function analyzeNursingSheet(sheetUrl) {
    try {
        if (!sheetUrl) {
            const settings = getSettings(CONFIG.SETTINGS_KEYS.NURSING);
            sheetUrl = settings.nursingSheetId;
        }
        if (!sheetUrl) throw new Error("No sheet URL was provided, and none is saved in Settings.");

        const spreadsheetId = extractFileIdFromUrl(sheetUrl);
        if (!spreadsheetId) throw new Error("The provided URL is not a valid Google Sheet URL.");

        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheets = spreadsheet.getSheets();
        const allSheetsData = [];
        const analysisErrors = [];

        for (const sheet of sheets) {
            try {
                if (sheet.getLastRow() < 5 || sheet.getName().toLowerCase().includes('template') || sheet.getName().toLowerCase().includes('config')) continue;
                const parsedData = parseSheetForAnalysis_(sheet);
                if (parsedData && parsedData.exams.length > 0) allSheetsData.push(parsedData);
            } catch (e) {
                analysisErrors.push(`Sheet '${sheet.getName()}': ${e.message}`);
            }
        }

        if (allSheetsData.length === 0 && analysisErrors.length > 0) throw new Error(`Analysis failed on all sheets. Errors: ${analysisErrors.join("; ")}`)
        return { success: true, data: { spreadsheetId, sheets: allSheetsData, analysisErrors } };
    } catch (e) {
        return { success: false, message: "ANALYSIS FAILED: " + e.message };
    }
}

function createNursingProctoringDocuments(approvedData) {
    return processNursingDocuments_(approvedData, false); // false = not update-only
}

function updateAllNursingDocuments(approvedData) {
    return processNursingDocuments_(approvedData, true); // true = update-only mode
}

/**
 * [UPGRADED] Core document processing engine with a mode for "update-only".
 * @param {object} approvedData The data from the frontend.
 * @param {boolean} updateOnly If true, will only update existing docs and skip creating new ones.
 */
function processNursingDocuments_(approvedData, updateOnly) {
    try {
        const { settings, targetFolderId } = nursing_getSettingsAndIDs();
        if (!targetFolderId) throw new Error("Output folder is not defined in Nursing Settings.");

        const mainOutputFolder = DriveApp.getFolderById(targetFolderId);
        const templateId = settings.nursingTemplateId ? extractFileIdFromUrl(settings.nursingTemplateId) : null;
        let docsCreated = 0;
        let docsUpdated = 0;
        let docsSkipped = 0;

        for (const sheetData of approvedData.sheets) {
            const { course, exams, rosters, sheetName } = sheetData;
            if (!exams || exams.length === 0) continue;

            const folderName = sheetName.trim();
            const existingFolders = mainOutputFolder.getFoldersByName(folderName);
            const targetFolder = existingFolders.hasNext() ? existingFolders.next() : mainOutputFolder.createFolder(folderName);

            for (const exam of exams) {
                const docTitle = `${course.code} ${course.name} - ${exam.name}`.replace(/[\/]/g, '-');
                const existingFiles = targetFolder.getFilesByName(docTitle);

                if (existingFiles.hasNext()) {
                    const file = existingFiles.next();
                    const doc = DocumentApp.openById(file.getId());
                    const body = doc.getBody();
                    const preservedAccommodations = findSectionText_(body, "Accommodations");
                    body.clear();
                    populateDocContent_(body, docTitle, course, exam, rosters, settings.customNotes, preservedAccommodations);
                    doc.saveAndClose();
                    logSystemAction("Nursing", `Updated in folder '${folderName}'`, docTitle, doc.getId(), `Exam Date: ${exam.date}`);
                    docsUpdated++;
                } else {
                    if (updateOnly) {
                        docsSkipped++;
                        continue; // Skip creating new file in update-only mode
                    }
                    // Create new document
                    let doc;
                    if (templateId) {
                        const newFile = DriveApp.getFileById(templateId).makeCopy(docTitle, targetFolder);
                        doc = DocumentApp.openById(newFile.getId());
                    } else {
                        doc = DocumentApp.create(docTitle);
                        const newFile = DriveApp.getFileById(doc.getId());
                        targetFolder.addFile(newFile);
                        DriveApp.getRootFolder().removeFile(newFile);
                    }
                    populateDocContent_(doc.getBody(), docTitle, course, exam, rosters, settings.customNotes, "");
                    doc.saveAndClose();
                    logSystemAction("Nursing", `Created in folder '${folderName}'`, docTitle, doc.getId(), `Exam Date: ${exam.date}`);
                    docsCreated++;
                }
            }
        }

        let message = `Process Complete! `;
        if (updateOnly) {
            message += `Updated ${docsUpdated} document(s). Skipped ${docsSkipped} non-existent document(s).`;
        } else {
            message += `Created ${docsCreated} and updated ${docsUpdated} document(s).`;
        }
        if (docsCreated === 0 && docsUpdated === 0 && !updateOnly) {
             message = "Analysis complete. No documents needed to be created or updated.";
        }
        
        return { success: true, message };

    } catch (e) {
        console.error("processNursingDocuments_ Error: " + e.stack);
        return { success: false, message: `DOCUMENT PROCESSING FAILED: ${e.message}` };
    }
}


function populateDocContent_(body, docTitle, course, exam, rosters, customNotes, preservedAccommodations) {
    body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('Exam Details').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const addDetail = (label, val, highlight) => {
        if (!val || String(val).trim() === '') return;
        const item = body.appendListItem('');
        item.appendText(`${label}: `).setBold(true);
        const text = item.appendText(String(val));
        if (highlight) text.setBackgroundColor('#FFFF00');
    };
    addDetail('Date', exam.date, true);
    addDetail('Start Time (On Site)', exam.siteTime, true);
    addDetail('Start Time (Zoom)', exam.zoomTime, true);
    addDetail('Duration', exam.duration, true);
    addDetail('Room', exam.room, false);
    addDetail('Password', exam.password, true);
    body.appendParagraph('');
    if (preservedAccommodations) {
        body.appendParagraph('Accommodations').setHeading(DocumentApp.ParagraphHeading.HEADING1);
        body.appendParagraph(preservedAccommodations);
        body.appendParagraph('');
    }
    body.appendParagraph('Important Links').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('').appendText('Red Flag Reporting Form').setLinkUrl(CONFIG.NURSING.URLS.RED_FLAG_REPORT);
    body.appendParagraph('').appendText('Nursing Protocol').setLinkUrl(CONFIG.NURSING.URLS.PROTOCOL_DOC);
    body.appendParagraph('');
    body.appendParagraph('Location Rosters').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const locations = Object.keys(rosters).sort();
    if (locations.length > 0) {
        for (const location of locations) {
            const students = rosters[location];
            if (students && students.length > 0) {
                body.appendParagraph(location).setHeading(DocumentApp.ParagraphHeading.HEADING2);
                for (const student of students) {
                    const listItem = body.appendListItem(student.name);
                    if (student.color && student.color !== '#000000') {
                        listItem.setForegroundColor(student.color);
                    }
                    listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
                }
            }
        }
    } else {
        body.appendParagraph("No student rosters were found in the analyzed data.");
    }
    body.appendParagraph('');
    if (customNotes) {
        body.appendParagraph('General Notes').setHeading(DocumentApp.ParagraphHeading.HEADING1);
        body.appendParagraph(customNotes);
    }
}

function findSectionText_(body, heading) {
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
        if (paragraphs[i].getHeading() == DocumentApp.ParagraphHeading.HEADING1 && paragraphs[i].getText() == heading) {
            if (i + 1 < paragraphs.length && paragraphs[i + 1].getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
                return paragraphs[i + 1].getText();
            }
        }
    }
    return "";
}

function parseSheetForAnalysis_(sheet) {
    const allValues = sheet.getDataRange().getValues();
    const richTextValues = sheet.getDataRange().getRichTextValues();
    let warnings = [];
    let examHeaderRowIndex = -1;
    let rosterHeaderRowIndex = -1;
    for (let i = 0; i < allValues.length; i++) {
        const rowString = allValues[i].join(' ').toLowerCase();
        if (examHeaderRowIndex === -1 && rowString.includes('exam') && rowString.includes('date')) examHeaderRowIndex = i;
        if (rosterHeaderRowIndex === -1 && String(allValues[i][0]).toLowerCase().trim() === 'augusta') rosterHeaderRowIndex = i;
    }
    if (examHeaderRowIndex === -1) return null;
    const course = parseCourseInfo_(allValues[0][0]);
    const exams = parseExamBlock_(allValues, examHeaderRowIndex, rosterHeaderRowIndex, warnings);
    let rosters = {};
    if (rosterHeaderRowIndex !== -1) {
        rosters = parseRosterBlock_(richTextValues, rosterHeaderRowIndex, warnings);
    } else {
        warnings.push("Could not find the student roster section.");
    }
    return { sheetName: sheet.getName(), course, exams, rosters, warnings };
}

function parseCourseInfo_(rawString) {
    if (!rawString) return { code: 'N/A', name: 'N/A', faculty: 'N/A' };
    const parts = rawString.split(/[-–—|:]/);
    let coursePart = parts[0] || "";
    let faculty = (parts.length > 1) ? parts.slice(1).join(' ').replace(/Professor/i, '').trim() : "N/A";
    const codeMatch = coursePart.match(/[A-Z]{2,4}\s?\d{3,4}/);
    let code = "N/A";
    let name = coursePart.trim();
    if (codeMatch) {
        code = codeMatch[0];
        name = coursePart.replace(codeMatch[0], '').trim();
    }
    return { code, name, faculty };
}

function parseExamBlock_(allValues, headerRowIndex, endRowIndex, warnings) {
    const exams = [];
    const headers = allValues[headerRowIndex].map(h => String(h).toLowerCase().trim());
    const stopIndex = (endRowIndex !== -1) ? endRowIndex : allValues.length;
    const colMap = { name: headers.indexOf('exam'), date: headers.indexOf('date'), siteTime: headers.indexOf('start time (on site)'), zoomTime: headers.indexOf('start time (zoom)'), duration: headers.indexOf('duration'), room: headers.indexOf('room: on campus'), password: headers.indexOf('password') };
    if (colMap.name === -1 || colMap.date === -1) throw new Error("Crucial 'Exam' or 'Date' column is missing.");
    for (let i = headerRowIndex + 1; i < stopIndex; i++) {
        const row = allValues[i];
        const examName = row[colMap.name];
        if (!examName || String(examName).trim().length < 2) continue;
        const dateValue = row[colMap.date];
        let displayDate;
        if (dateValue) {
            try {
                const d = new Date(dateValue);
                if (d && !isNaN(d.getTime())) {
                    displayDate = d.toLocaleDateString("en-US", { weekday: 'long', month: 'long', day: 'numeric' });
                } else {
                    displayDate = String(dateValue).trim();
                    if (warnings) warnings.push(`Could not parse date for "${examName}". Using original text.`);
                }
            } catch (e) {
                displayDate = String(dateValue).trim();
                 if (warnings) warnings.push(`Could not parse date for "${examName}". Using original text.`);
            }
        } else {
            displayDate = "Not Specified";
        }
        exams.push({ name: String(examName).trim(), date: displayDate, siteTime: colMap.siteTime > -1 ? String(row[colMap.siteTime]).trim() : 'N/A', zoomTime: colMap.zoomTime > -1 ? String(row[colMap.zoomTime]).trim() : 'N/A', duration: colMap.duration > -1 ? String(row[colMap.duration]).trim() : 'N/A', room: colMap.room > -1 ? String(row[colMap.room]).trim() : 'N/A', password: colMap.password > -1 ? String(row[colMap.password]).trim() : 'N/A' });
    }
    return exams;
}

function parseRosterBlock_(richTextValues, headerRowIndex, warnings) {
    const rosters = {};
    const headers = richTextValues[headerRowIndex].map(rtv => rtv.getText().trim());
    let addressRowIndex = headerRowIndex + 1;
    while(addressRowIndex < richTextValues.length && richTextValues[addressRowIndex].every(c => c.getText().trim() === '')) {
        addressRowIndex++;
    }
    const studentDataStartIndex = addressRowIndex + 1;
    for (let c = 0; c < headers.length; c++) {
        const locationName = headers[c];
        if (!locationName) continue;
        rosters[locationName] = [];
        for (let r = studentDataStartIndex; r < richTextValues.length; r++) {
            if (!richTextValues[r] || !richTextValues[r][c]) continue;
            const richTextCell = richTextValues[r][c];
            const studentName = richTextCell.getText().trim();
            if (!studentName || studentName.toLowerCase().includes('students with accommodations')) continue;
            const color = richTextCell.getRuns()[0].getTextStyle().getForegroundColor();
            rosters[locationName].push({ name: studentName, color: color });
        }
    }
    return rosters;
}

function nursing_getSettingsAndIDs() {
    const settings = getSettings(CONFIG.SETTINGS_KEYS.NURSING);
    const spreadsheetId = extractFileIdFromUrl(settings.nursingSheetId);
    const targetFolderId = extractFileIdFromUrl(settings.nursingFolderId);
    return { settings, spreadsheetId, targetFolderId };
}

function nursing_syncExamsToCalendar(calendarId, startStr, endStr, overwrite) {
    return { success: false, message: "Calendar Sync is temporarily disabled." };
}