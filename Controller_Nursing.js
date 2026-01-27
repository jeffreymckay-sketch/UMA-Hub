/**
 * NEW HELPER: Reads the _DB_ACCOMMODATIONS tab into a fast lookup Map
 */
function getAccommodationsDBMap() {
    const map = {};
    try {
        // FIX: Use getMasterDataHub() instead of getActiveSpreadsheet()
        const ss = getMasterDataHub();
        const sheet = ss.getSheetByName("_DB_ACCOMMODATIONS");
        if (!sheet) return map; // Tab doesn't exist yet, return empty

        const data = sheet.getDataRange().getValues();
        // Skip header (Row 1)
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const id = String(row[0]); // Unique_ID
            if (!id) continue;

            let tags = {};
            try { 
                tags = JSON.parse(row[4]); // Col E is Student_Data
            } catch (e) { 
                // ignore parsing error
            } 

            map[id] = {
                generalNotes: row[3], // Col D is General_Notes
                studentTags: tags
            };
        }
    } catch (e) {
        console.warn("Error reading DB Map: " + e.message);
    }
    return map;
}

function api_saveNursingAccommodations(payload) {
  if (arguments.length === 3) {
      return { success: false, message: "Please refresh the page. The saving mechanism has been upgraded." };
  }

  try {
    // FIX: Use getMasterDataHub()
    const ss = getMasterDataHub(); 
    let sheet = ss.getSheetByName("_DB_ACCOMMODATIONS");
    
    // Auto-create if missing
    if (!sheet) {
        sheet = ss.insertSheet("_DB_ACCOMMODATIONS");
        sheet.appendRow(["Unique_ID", "Course_Code", "Exam_Name", "General_Notes", "Student_Data"]);
    }

    const uniqueId = `${payload.courseCode}|${payload.examName}`;
    const studentJson = JSON.stringify(payload.studentTags || {});
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === uniqueId) {
            rowIndex = i + 1; 
            break;
        }
    }

    if (rowIndex > -1) {
        sheet.getRange(rowIndex, 4, 1, 2).setValues([[payload.generalNotes, studentJson]]);
    } else {
        sheet.appendRow([uniqueId, payload.courseCode, payload.examName, payload.generalNotes, studentJson]);
    }

    return { success: true, message: 'Saved to Database!' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}