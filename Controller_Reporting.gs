/**
 * -------------------------------------------------------------------
 * REPORTING CONTROLLER
 * -------------------------------------------------------------------
 */

function api_generateStaffReport(startDateString, endDateString) {
    try {
        const reportSheet = getSheet('Report_Data');
        const startDate = new Date(startDateString);
        const endDate = new Date(endDateString);
        if (isNaN(startDate) || isNaN(endDate) || startDate > endDate) throw new Error("Invalid date range.");

        const staffAssignments = getSheet('Staff_Assignments').getDataRange().getValues().slice(1);
        const staffList = getSheet('Staff_List').getDataRange().getValues();
        const shiftData = getSheet('TechHub_Shifts').getDataRange().getValues();
        const courseData = getSheet('Course_Schedule').getDataRange().getValues();
        
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

            const WEEKDAYS = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];
            const targetDayIndex = WEEKDAYS.indexOf(context.dayOfWeek.toUpperCase());
            if (targetDayIndex === -1) continue;
            
            let currentDate = new Date(startDate);
            while (currentDate <= endDate) {
                if (currentDate.getDay() === targetDayIndex) {
                    const eventDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                    reportRows.push([eventDate, staffDetails[0], staffDetails[2], staffDetails[4], assignment.assignmentType, assignment.referenceId, context.description, context.startTime, context.endTime, context.durationMinutes / 60, 'PLANNED', 0]);
                }
                currentDate.setDate(currentDate.getDate() + 1);
            }
        }

        if (reportSheet.getLastRow() > 1) reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, 12).clearContent();
        if (reportRows.length > 0) reportSheet.getRange(2, 1, reportRows.length, 12).setValues(reportRows);
        return { success: true, message: `Generated ${reportRows.length} records.` };
    } catch (e) { return { success: false, message: e.message }; }
}
