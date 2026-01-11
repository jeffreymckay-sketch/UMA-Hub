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
    const spreadsheet = SpreadsheetApp.openById(settings.nursingSheetId);
    const allSheetData = spreadsheet.getSheets()
      .map(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.toLowerCase().includes('template') || sheetName.toLowerCase().includes('master')) {
          return null;
        }

        const data = sheet.getDataRange().getValues();
        // --- FIX: Make header row detection more robust by looking for multiple key headers. ---
        const headerRowIndex = data.findIndex(row => {
            const joinedRow = row.join('').toLowerCase();
            return joinedRow.includes('exam name') && joinedRow.includes('course code');
        });

        if (headerRowIndex === -1) {
            console.log(`Skipping sheet "${sheetName}" - Header row not found.`);
            return null; // If no header, skip sheet.
        }
        
        const headers = data[headerRowIndex].map(h => String(h).trim().toLowerCase());
        const examData = parseExamData(data, headers, headerRowIndex);
        
        // If, after parsing, we have no valid course or no exams, discard this sheet's data.
        if (!examData.course.code || !examData.course.name || examData.exams.length === 0) {
            console.log(`Skipping sheet "${sheetName}" - No valid course or exam data found after parsing.`);
            return null;
        }
        
        return {
          sheetName: sheetName,
          course: examData.course,
          exams: examData.exams
        };
      })
      .filter(s => s !== null); // Filter out nulls from skipped sheets

    // Find document URLs for each valid exam
    allSheetData.forEach(sheetData => {
        sheetData.exams.forEach(exam => {
            exam.docUrl = findDocUrlByName(exam.name, settings.nursingFolderId);
        });
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
    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found.`);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.trim().toLowerCase());
    const nameCol = headers.indexOf('exam name');
    const accommCol = headers.indexOf('accommodations');

    if (nameCol === -1 || accommCol === -1) {
      throw new Error('Could not find required columns (Exam Name, Accommodations) in sheet: ' + sheetName);
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === examName) {
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
  try {
    const settings = data.settings || getNursingSettings();
    const folder = DriveApp.getFolderById(settings.nursingFolderId);
    let count = 0;
    
    data.sheets.forEach(sheet => {
      sheet.exams.forEach(exam => {
        const docName = exam.name;
        const existingDoc = findDocByName(docName, folder);

        if (existingDoc) {
          updateDocContent(existingDoc, exam, sheet.course, settings);
        } else {
          const newDoc = DocumentApp.create(docName);
          const file = DriveApp.getFileById(newDoc.getId());
          folder.addFile(file);
          DriveApp.getRootFolder().removeFile(file); 
          updateDocContent(newDoc, exam, sheet.course, settings);
        }
        count++;
      });
    });
    
    const docUrl = data.sheets[0]?.exams[0] ? findDocUrlByName(data.sheets[0].exams[0].name, settings.nursingFolderId) : null;
    return { success: true, message: `${count} document(s) created/updated.`, docUrl: docUrl };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateAllNursingDocuments(data) {
  try {
    const settings = data.settings || getNursingSettings();
    const folder = DriveApp.getFolderById(settings.nursingFolderId);
    let count = 0;

    data.sheets.forEach(sheet => {
      sheet.exams.forEach(exam => {
        const doc = findDocByName(exam.name, folder);
        if (doc) {
          updateDocContent(doc, exam, sheet.course, settings);
          count++;
        }
      });
    });

    if (count === 0) {
        return { success: false, message: "No matching documents found to update." };
    }
    const docUrl = data.sheets[0]?.exams[0] ? findDocUrlByName(data.sheets[0].exams[0].name, settings.nursingFolderId) : null;
    return { success: true, message: `${count} document(s) updated successfully.`, docUrl: docUrl };

  } catch(e) {
    return { success: false, message: e.message };
  }
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
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : url; 
}

function parseExamData(data, headers, headerRowIndex) {
  const dataRows = data.slice(headerRowIndex + 1);
  const courseCodeIndex = headers.indexOf('course code');
  const courseNameIndex = headers.indexOf('course name');
  const examNameIndex = headers.indexOf('exam name'); // Make sure we have this for filtering

  // Find the first row that actually has course data. Skips blank rows after the header.
  const firstDataRow = dataRows.find(row => row[courseCodeIndex] && row[courseNameIndex]);

  const course = {
    code: firstDataRow ? firstDataRow[courseCodeIndex] : null,
    name: firstDataRow ? firstDataRow[courseNameIndex] : null,
  };

  // If no course info could be found at all, we can't proceed.
  if (!course.code || !course.name) {
      return { course: {}, exams: [] };
  }

  const exams = dataRows
    .filter(row => {
      // A row is valid if it's not completely empty and has an exam name.
      const hasExamName = row[examNameIndex] && String(row[examNameIndex]).trim() !== '';
      const isNotEmpty = row.some(cell => String(cell).trim() !== '');
      return hasExamName && isNotEmpty;
    })
    .map(row => {
      const examObj = {};
      headers.forEach((header, i) => {
        if (!header) return; // Skip empty header cells
        // Robust camelCase conversion
        const camelCaseHeader = header.replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase());
        examObj[camelCaseHeader] = row[i];
      });

      // Standardize the 'name' property
      if (examObj.examName) {
        examObj.name = examObj.examName;
        // delete examObj.examName; // Keep original property for now if needed, or delete.
      }

      // Format date if it's a Date object
      if (examObj.date && examObj.date instanceof Date) {
        examObj.date = examObj.date.toLocaleDateString();
      }

      return examObj;
    });

  return { course, exams };
}


function findDocByName(name, folder) {
  const files = folder.getFilesByName(name);
  return files.hasNext() ? DocumentApp.openById(files.next().getId()) : null;
}

function findDocUrlByName(name, folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(name);
    return files.hasNext() ? files.next().getUrl() : null;
}

function updateDocContent(doc, exam, course, settings) {
    const body = doc.getBody().clear();
    const FONT_FAMILY = 'Calibri';
    const FONT_SIZE = 11;
    const HEADER_BG = '#f3f3f3';

    const docAttributes = {};
    docAttributes[DocumentApp.Attribute.FONT_FAMILY] = FONT_FAMILY;
    docAttributes[DocumentApp.Attribute.FONT_SIZE] = FONT_SIZE;
    body.setAttributes(docAttributes);
    body.setMarginTop(72).setMarginBottom(72).setMarginLeft(72).setMarginRight(72);

    const header = doc.getHeader() || doc.addHeader();
    header.clear();
    const headerTable = header.appendTable([[' ','']]);
    headerTable.setBorderWidth(0);
    const courseCell = headerTable.getCell(0,0).setWidth(300);
    const examCell = headerTable.getCell(0,1).setWidth(150);

    courseCell.getChild(0).asParagraph().setIndentFirstLine(0).setIndentStart(0).appendText(`${course.code}\n${course.name}`);
    examCell.getChild(0).asParagraph().setIndentFirstLine(0).setIndentStart(0).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).appendText(exam.name);
    header.appendHorizontalRule();

    const title = body.appendParagraph(`${course.code} ${course.name}`);
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');

    const tableData = [
        ['Exam Name', exam.name || 'N/A'],
        ['Exam Password', exam.password || 'N/A'],
        ['Date', exam.date ? new Date(exam.date).toLocaleDateString() : 'N/A'],
        ['Time', exam.siteTime || 'N/A'],
        ['Testing Site', exam.testingSite || 'N/A']
    ];

    if (tableData.length > 0) {
        const table = body.appendTable(tableData);
        const headerStyle = { [DocumentApp.Attribute.BACKGROUND_COLOR]: HEADER_BG, [DocumentApp.Attribute.BOLD]: true };
        table.getRow(0).setAttributes(headerStyle);
    }
    body.appendParagraph('');

    if (exam.accommodations) {
        body.appendParagraph('Accommodations').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(exam.accommodations);
        body.appendParagraph('');
    }

    if (settings.customNotes) {
        body.appendParagraph('General Instructions').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(settings.customNotes);
    }

    doc.saveAndClose();
}
