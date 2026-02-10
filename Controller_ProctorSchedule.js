/**
 * -------------------------------------------------------------------
 * FILE: Controller_ProctorSchedule.js
 * DESCRIPTION: Aggregates Nursing and MLT data into a master schedule.
 * -------------------------------------------------------------------
 */

function api_generateMasterExamSchedule() {
    try {
      // 1. Gather the Harvest (Fetch Data)
      // We use the existing functions from your other controllers.
      // Note: If you see red lines under these function names, fear not! 
      // They exist in Controller_Nursing.js and Controller_MLT.js and will work when run.
      var nursingRes = api_getNursingData(); 
      var mltRes = api_getMLTData();
  
      if (!nursingRes.success) throw new Error("Nursing Data Error: " + nursingRes.message);
      if (!mltRes.success) throw new Error("MLT Data Error: " + mltRes.message);
  
      var rows = [];
      var richTextLinks = []; 
  
      // 2. The Sorting Logic (Helper Function)
      // We define this to handle both lists exactly the same way.
      function processSheetData(programName, sheetObj) {
        if (!sheetObj || !sheetObj.exams) return;
  
        var exams = sheetObj.exams;
        // Handle course name safely (Nursing has code/name object, MLT might differ slightly)
        var courseTitle = "Unknown Course";
        if (sheetObj.course && sheetObj.course.code) {
          courseTitle = sheetObj.course.code + " " + sheetObj.course.name;
        } else if (sheetObj.fullTitle) {
          courseTitle = sheetObj.fullTitle;
        }
  
        for (var i = 0; i < exams.length; i++) {
          var exam = exams[i];
          
          // Build the "Smart Link"
          var linkVal;
          if (exam.docUrl) {
            linkVal = SpreadsheetApp.newRichTextValue()
              .setText("Open Document")
              .setLinkUrl(exam.docUrl)
              .build();
          } else {
            linkVal = SpreadsheetApp.newRichTextValue()
              .setText("-")
              .build();
          }
  
          // Add the row data
          // Columns: [Program, Course, Exam, Date, Start, Location, Password, Link, Notes]
          rows.push([
            programName,
            courseTitle,
            exam.name || "Unnamed Exam",
            exam.date || "-",
            exam.siteTime || "-",
            exam.room || "-",
            exam.password || "-",
            "", // Placeholder for the link column
            exam.generalNotes || ""
          ]);
  
          // Save the link for the separate styling step
          richTextLinks.push([linkVal]);
        }
      }
  
      // 3. Process Nursing Exams
      if (nursingRes.data && nursingRes.data.sheets) {
        nursingRes.data.sheets.forEach(function(sheet) {
          processSheetData("Nursing", sheet);
        });
      }
  
      // 4. Process MLT Exams
      if (mltRes.data && mltRes.data.sheets) {
        mltRes.data.sheets.forEach(function(sheet) {
          processSheetData("MLT", sheet);
        });
      }
  
      if (rows.length === 0) {
        return { success: false, message: "No exams found in either tool." };
      }
  
      // 5. Create the Spreadsheet
      var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      var ss = SpreadsheetApp.create("Master Exam Schedule - " + dateStr);
      var sheet = ss.getActiveSheet();
  
      // 6. Write Headers
      var headers = ["Program", "Course", "Exam Name", "Date", "Start Time", "Location", "Password", "Document Link", "Notes"];
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight("bold")
                 .setBackground("#003057") // University Navy
                 .setFontColor("white");
      
      // 7. Write Data Rows
      var dataRange = sheet.getRange(2, 1, rows.length, headers.length);
      dataRange.setValues(rows);
  
      // 8. Apply the "Smart Links"
      // We apply these specifically to Column H (index 8)
      sheet.getRange(2, 8, richTextLinks.length, 1).setRichTextValues(richTextLinks);
  
      // 9. Polish (Formatting)
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
  
      return { success: true, url: ss.getUrl() };
  
    } catch (e) {
      console.error("Master Schedule Error: " + e.stack);
      return { success: false, message: e.message };
    }
  }