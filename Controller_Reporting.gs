/**
 * -------------------------------------------------------------------
 * REPORTING CONTROLLER
 * -------------------------------------------------------------------
 */

function api_generateStaffReport(startDateString, endDateString) {
    try {
        const ss = getMasterDataHub();
        const reportSheet = getOrCreateSheet(ss, CONFIG.TABS.REPORT_DATA, CONFIG.HEADERS.REPORTING);
        const startDate = new Date(startDateString);
        const endDate = new Date(endDateString);
        if (isNaN(startDate) || isNaN(endDate) || startDate > endDate) throw new Error("Invalid date range.");

        const staffAssignments = getRequiredSheetData(ss, CONFIG.TABS.STAFF_ASSIGNMENTS).slice(1);
        const staffList = getRequiredSheetData(ss, CONFIG.TABS.STAFF_LIST);
        const shiftData = getRequiredSheetData(ss, CONFIG.TABS.TECH_HUB_SHIFTS);
        const courseData = getRequiredSheetData(ss, CONFIG.TABS.COURSE_SCHEDULE);
        
        const staffMap = createDataMap(staffList, 'StaffID');
        const shiftMap = createDataMap(shiftData, 'ShiftID');
        const courseMap = createDataMap(courseData, 'CourseID');

        const reportRows = [];
        for (const assignmentRow of staffAssignments) {
            const assignment = { id: assignmentRow[0], staffId: assignmentRow[1], assignmentType: assignmentRow[2], referenceId: assignmentRow[3] };
            const staffDetails = staffMap[assignment.staffId];
            if (!staffDetails) continue;

            let context = { description: 'N/A', durationMinutes: 0, dayOfWeek: null, startTime: null, endTime: null };
            if (assignment.assignmentType === 'Tech Hub') {
                const shift = shiftMap[assignment.referenceId];
                if (shift) { context = { description: shift.Description, dayOfWeek: shift.DayOfWeek, startTime: shift.StartTime, endTime: shift.EndTime, durationMinutes: parseTime(shift.EndTime) - parseTime(shift.StartTime) }; }
            } else if (assignment.assignmentType === 'Course') {
                const course = courseMap[assignment.referenceId];
                if (course) { context = { description: course.CourseName, dayOfWeek: course.DayOfWeek || 'Monday', startTime: course.StartTime || '10:00', endTime: course.EndTime || '11:00', durationMinutes: 60 }; }
            }

            const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
            const targetDayIndex = DAY_NAMES.indexOf(context.dayOfWeek);
            if (targetDayIndex === -1) continue;
            
            let currentDate = new Date(startDate);
            while (currentDate <= endDate) {
                if (currentDate.getDay() === targetDayIndex) {
                    const eventDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                    reportRows.push([eventDate, staffDetails.StaffID, staffDetails.FullName, staffDetails.Roles.split(',')[0], assignment.assignmentType, assignment.referenceId, context.description, context.startTime, context.endTime, context.durationMinutes / 60, 'Planned', 0]);
                }
                currentDate.setDate(currentDate.getDate() + 1);
            }
        }

        if (reportSheet.getLastRow() > 1) reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, CONFIG.HEADERS.REPORTING.length).clearContent();
        if (reportRows.length > 0) reportSheet.getRange(2, 1, reportRows.length, CONFIG.HEADERS.REPORTING.length).setValues(reportRows);
        return { success: true, message: `Generated ${reportRows.length} records.` };
    } catch (e) { return { success: false, message: e.message }; }
}