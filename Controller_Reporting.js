/**
 * -------------------------------------------------------------------
 * REPORTING & INSPECTOR CONTROLLER
 * Handles Staff Reports, Calendar Analytics, and Configuration
 * -------------------------------------------------------------------
 */

/**
 * Scans multiple calendars for events within a specified date range.
 * Automatically categorizes events based on 'Event_Types' settings.
 */
function api_inspectCalendars(startStr, endStr, calIds) {
    try {
      if (!calIds || !Array.isArray(calIds) || calIds.length === 0) {
        return { success: false, message: "No calendars were selected to scan." };
      }
      if (!startStr || !endStr) {
        return { success: false, message: "A start and end date are required." };
      }
  
      const start = new Date(startStr);
      const end = new Date(endStr);
      end.setHours(23, 59, 59, 999);
  
      // 1. Fetch Event Types for Categorization
      const eventTypesRes = api_getEventTypes();
      const eventTypes = eventTypesRes.success ? eventTypesRes.data : [];
  
      let allEvents = [];
  
      calIds.forEach(id => {
        try {
          const cal = CalendarApp.getCalendarById(id);
          if (cal) {
            const calName = cal.getName();
            const events = cal.getEvents(start, end);
  
            const mapped = events.map(e => {
              const title = e.getTitle();
              const durationMinutes = (e.getEndTime() - e.getStartTime()) / (1000 * 60);
              
              // 2. Categorize Event
              let category = "Uncategorized";
              for (const type of eventTypes) {
                  if (!type.keywords) continue;
                  const keywords = type.keywords.split(',').map(k => k.trim().toLowerCase());
                  const titleLower = title.toLowerCase();
                  if (keywords.some(k => k && titleLower.includes(k))) {
                      category = type.name;
                      break; 
                  }
              }
  
              return {
                  summary: title,
                  start: e.getStartTime().toISOString(),
                  end: e.getEndTime().toISOString(),
                  calendar: calName,
                  durationMinutes: durationMinutes,
                  category: category
              };
            });
  
            allEvents = allEvents.concat(mapped);
          }
        } catch (err) {
          console.error(`Error processing calendar ID ${id}: ${err.message}`);
        }
      });
  
      allEvents.sort((a, b) => new Date(a.start) - new Date(b.start));
  
      return { success: true, data: allEvents };
  
    } catch (e) {
      console.error(`Error in api_inspectCalendars: ${e.message}`);
      return { success: false, message: "An unexpected server error occurred: " + e.message };
    }
  }
  
  /**
   * EXPORT FUNCTION
   * Creates a Google Sheet from the inspected events.
   */
  function api_exportReport(events) {
    try {
      if (!events || events.length === 0) return { success: false, message: "No data to export." };
      
      // Create new Spreadsheet
      const filename = "Calendar Report " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      const ss = SpreadsheetApp.create(filename);
      const sheet = ss.getActiveSheet();
      
      // Headers
      const headers = ["Date", "Start Time", "End Time", "Duration (Hrs)", "Category", "Event Title", "Source Calendar"];
      sheet.appendRow(headers);
      
      // Format Data Rows
      const rows = events.map(e => {
          const start = new Date(e.start);
          const end = new Date(e.end);
          const dateStr = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd");
          const startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), "HH:mm");
          const endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), "HH:mm");
          const durationHrs = (e.durationMinutes / 60).toFixed(2);
  
          return [
              dateStr,
              startStr,
              endStr,
              durationHrs,
              e.category || "Uncategorized",
              e.summary,
              e.calendar
          ];
      });
      
      // Write Data in bulk
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      
      // Styling
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold").setBackground("#003057").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, headers.length);
      
      return { success: true, url: ss.getUrl() };
  
    } catch (e) {
      return { success: false, message: e.message };
    }
  }
  
  /**
   * Fetches defined Event Types from the "Event_Types" tab.
   */
  function api_getEventTypes() {
    try {
      const sheet = getSheet('Event_Types');
      if (!sheet) return { success: true, data: [] }; 
      
      const data = sheet.getDataRange().getValues();
      const types = [];
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && String(data[i][0]).trim() !== "") {
          types.push({
            name: String(data[i][0]).trim(),
            keywords: String(data[i][1] || "")
          });
        }
      }
      return { success: true, data: types };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }
  
  /**
   * Saves the Event Types list back to the "Event_Types" tab.
   */
  function api_saveEventTypes(newTypes) {
      try {
          const sheet = getSheet('Event_Types');
          if (!sheet) throw new Error("Sheet 'Event_Types' not found.");
  
          const rows = newTypes.map(t => [t.name, t.keywords]);
  
          const lastRow = sheet.getLastRow();
          if (lastRow > 1) {
              sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
          }
  
          if (rows.length > 0) {
              sheet.getRange(2, 1, rows.length, 2).setValues(rows);
          }
  
          return { success: true, message: "Settings saved successfully." };
      } catch (e) {
          return { success: false, message: e.message };
      }
  }
  
  /**
   * Generates internal staff assignment reports.
   */
  function api_generateStaffReport(startDateString, endDateString) {
      try {
          const reportSheet = getSheet('Report_Data');
          if (!reportSheet) throw new Error("Sheet 'Report_Data' not found.");
  
          const startDate = new Date(startDateString);
          const endDate = new Date(endDateString);
          if (isNaN(startDate) || isNaN(endDate) || startDate > endDate) throw new Error("Invalid date range.");
  
          const staffAssignmentsSheet = getSheet('Staff_Assignments');
          const staffListSheet = getSheet('Staff_List');
          const shiftDataSheet = getSheet('TechHub_Shifts');
          const courseDataSheet = getSheet('Course_Schedule');
  
          if (!staffAssignmentsSheet || !staffListSheet || !shiftDataSheet || !courseDataSheet) {
              throw new Error("One or more required source sheets are missing.");
          }
  
          const staffAssignments = staffAssignmentsSheet.getDataRange().getValues().slice(1);
          const staffList = staffListSheet.getDataRange().getValues();
          const shiftData = shiftDataSheet.getDataRange().getValues();
          const courseData = courseDataSheet.getDataRange().getValues();
          
          const staffMap = createDataMap(staffList, 0);
          const shiftMap = createDataMap(shiftData, 0);
          const courseMap = createDataMap(courseData, 0);
  
          const reportRows = [];
          for (const assignmentRow of staffAssignments) {
              const assignment = { id: assignmentRow[0], staffId: assignmentRow[1], assignmentType: assignmentRow[2], referenceId: assignmentRow[3] };
              const staffDetails = staffMap[assignment.staffId];
              if (!staffDetails) continue;
  
              let context = { description: 'N/A', durationMinutes: 0, dayOfWeek: null, startTime: null, endTime: null };
              
              if (assignment.assignmentType === 'Tech Hub') {
                  const shift = shiftMap[assignment.referenceId];
                  if (shift) { 
                      const startTime = new Date(shift[2]);
                      const endTime = new Date(shift[3]);
                      context = { 
                          description: shift[1], 
                          dayOfWeek: shift[4], 
                          startTime: startTime.toLocaleTimeString(), 
                          endTime: endTime.toLocaleTimeString(), 
                          durationMinutes: (endTime - startTime) / 60000 
                      }; 
                  }
              } else if (assignment.assignmentType === 'Course') {
                  const course = courseMap[assignment.referenceId];
                  if (course) { 
                      const startTime = new Date(course[3]);
                      const endTime = new Date(course[4]);
                      context = { 
                          description: course[1], 
                          dayOfWeek: course[2] || 'Monday', 
                          startTime: startTime.toLocaleTimeString(), 
                          endTime: endTime.toLocaleTimeString(), 
                          durationMinutes: 60 
                      }; 
                  }
              }
  
              if (!context.dayOfWeek) continue;
  
              const WEEKDAYS = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
              const targetDayIndex = WEEKDAYS.indexOf(context.dayOfWeek.toUpperCase());
              if (targetDayIndex === -1) continue;
              
              let currentDate = new Date(startDate);
              while (currentDate <= endDate) {
                  if (currentDate.getDay() === targetDayIndex) {
                      const eventDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                      reportRows.push([
                          eventDate, 
                          staffDetails[0], 
                          staffDetails[2], 
                          staffDetails[4], 
                          assignment.assignmentType, 
                          assignment.referenceId, 
                          context.description, 
                          context.startTime, 
                          context.endTime, 
                          context.durationMinutes / 60, 
                          'PLANNED', 
                          0
                      ]);
                  }
                  currentDate.setDate(currentDate.getDate() + 1);
              }
          }
  
          if (reportSheet.getLastRow() > 1) reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, 12).clearContent();
          if (reportRows.length > 0) reportSheet.getRange(2, 1, reportRows.length, 12).setValues(reportRows);
          
          return { success: true, message: `Generated ${reportRows.length} records.` };
  
      } catch (e) { 
          return { success: false, message: e.message }; 
      }
  }