/**
 * -------------------------------------------------------------------
 * MY SCHEDULE CONTROLLER
 * Fetches and aggregates the current user's assignments from 
 * MST (Courses) and Tech Hub (Shifts).
 * -------------------------------------------------------------------
 */

function api_getMyScheduleData() {
    try {
        const email = Session.getActiveUser().getEmail().toLowerCase();
        
        // 1. Get user's Staff ID from Staff_List
        const staffSheet = getSheet('Staff_List');
        if (!staffSheet) throw new Error("Staff_List sheet not found.");
        
        const staffData = staffSheet.getDataRange().getValues();
        const staffHeaders = getColumnMap(staffData[0]);
        
        let staffId = email; // Fallback to email if ID isn't set
        const userRow = staffData.find(row => String(row[staffHeaders.staffid]).toLowerCase() === email);
        if (userRow && userRow[staffHeaders.staffid]) {
            staffId = String(userRow[staffHeaders.staffid]).trim();
        }

        // 2. Fetch Assignments
        const assignSheet = getSheet('Staff_Assignments');
        if (!assignSheet) throw new Error("Staff_Assignments sheet not found.");
        
        const assignData = assignSheet.getDataRange().getValues();
        const assignHeaders = getColumnMap(assignData[0]);
        
        // Filter assignments strictly for the active user
        const myAssignments = assignData.filter(row => {
            return String(row[assignHeaders.staffid]).toLowerCase().trim() === staffId.toLowerCase();
        });

        // Separate IDs by tool type
        const courseIds = myAssignments
            .filter(row => row[assignHeaders.assignmenttype] === 'Course')
            .map(row => String(row[assignHeaders.referenceid]).trim());
            
        const shiftIds = myAssignments
            .filter(row => row[assignHeaders.assignmenttype] === 'Tech Hub')
            .map(row => String(row[assignHeaders.referenceid]).trim());

        const allItems =[];
        const tz = getMasterDataHub().getSpreadsheetTimeZone();

        // 3. Fetch MST Courses (NOW USING THE MERGED OVERRIDE LAYER)
        if (courseIds.length > 0) {
            let mergedCourses =[];
            try {
                // This function lives in Controller_MST.js and handles all overrides
                if (typeof getMergedMstData === 'function') {
                    mergedCourses = getMergedMstData();
                }
            } catch(e) {
                console.warn("Could not fetch merged MST data: " + e.message);
            }

            mergedCourses.forEach(courseObj => {
                const id = String(courseObj.eventid).trim();
                
                if (courseIds.includes(id)) {
                    // Extract dates for UI display
                    let dateStr = "";
                    const startD = courseObj.startdate;
                    const endD = courseObj.enddate;
                    
                    if (startD && endD) {
                        const sd = new Date(startD);
                        const ed = new Date(endD);
                        if (!isNaN(sd.getTime()) && !isNaN(ed.getTime())) {
                            dateStr = `${sd.getMonth()+1}/${sd.getDate()} - ${ed.getMonth()+1}/${ed.getDate()}`;
                        }
                    }

                    // Extract and format times
                    let timeDisplay = String(courseObj.runtime || 'TBD');
                    if (courseObj.starttime && courseObj.endtime) {
                        const formatTime = (t) => {
                            if (!t) return '';
                            if (t instanceof Date) return Utilities.formatDate(t, tz, 'h:mm a');
                            if (typeof t === 'string' && t.includes('T')) {
                                const d = new Date(t);
                                if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'h:mm a');
                            }
                            return String(t);
                        };
                        timeDisplay = `${formatTime(courseObj.starttime)} - ${formatTime(courseObj.endtime)}`;
                    }

                    allItems.push({
                        id: id,
                        type: 'MST',
                        title: `${courseObj.course || 'Untitled'} - ${courseObj.faculty || 'Unassigned'}`,
                        rawDay: String(courseObj.day || ''),
                        timeString: timeDisplay,
                        location: String(courseObj.bxlocation || ''),
                        zoomLink: String(courseObj.zoomlink || ''), // Now correctly pulls from overrides
                        dateStr: dateStr
                    });
                }
            });
        }

        // 4. Fetch Tech Hub Shifts
        if (shiftIds.length > 0) {
            // FIXED: Using lowercase 's' to match your configuration key exactly
            const shiftSheet = getSheet('TechHub_shifts');
            if (shiftSheet) {
                const shiftData = shiftSheet.getDataRange().getValues();
                const sHeaders = getColumnMap(shiftData[0]);
                
                for (let i = 1; i < shiftData.length; i++) {
                    const row = shiftData[i];
                    const id = String(row[sHeaders.shiftid]).trim();
                    
                    if (shiftIds.includes(id)) {
                        let startDisplay = row[sHeaders.starttime];
                        let endDisplay = row[sHeaders.endtime];
                        
                        // Safely parse Dates to Time Strings
                        if (startDisplay instanceof Date) startDisplay = Utilities.formatDate(startDisplay, tz, "h:mm a");
                        if (endDisplay instanceof Date) endDisplay = Utilities.formatDate(endDisplay, tz, "h:mm a");
                        
                        const isZoom = String(row[sHeaders.zoom]).toLowerCase() === 'true';
                        
                        // --- Inject standard Tech Hub Zoom link if shift is marked as Zoom ---
                        const thZoomLink = isZoom ? "https://maine.zoom.us/j/2076213123" : "";
                        
                        allItems.push({
                            id: id,
                            type: 'TechHub',
                            title: String(row[sHeaders.description]),
                            rawDay: String(row[sHeaders.day] || row[sHeaders.dayofweek] || ''),
                            timeString: `${startDisplay} - ${endDisplay}`,
                            location: isZoom ? 'Zoom' : 'Tech Hub Desk',
                            zoomLink: thZoomLink,
                            dateStr: '' // Shifts don't use this specific date range display
                        });
                    }
                }
            }
        }

        // 5. Build Grouped Schedule (Monday - Friday)
        const schedule = {
            Monday: [],
            Tuesday: [],
            Wednesday:[],
            Thursday: [],
            Friday:[]
        };

        // Helper to accurately map various text formats (e.g. "M/W", "T/Th", "Tuesday") to standard days
        const mapDaysToStandard = (dayString) => {
            const mapped =[];
            if (!dayString) return mapped;
            
            // Replace slashes/commas with spaces so word boundaries work perfectly
            const str = dayString.toLowerCase().replace(/[^a-z]/g, ' '); 
            
            if (str.includes('monday') || /\bm\b/.test(str)) mapped.push('Monday');
            if (str.includes('tuesday') || /\btue\b|\btues\b|\btu\b/.test(str) || (/\bt\b/.test(str) && !/\bth\b/.test(str))) mapped.push('Tuesday');
            if (str.includes('wednesday') || /\bw\b|\bwed\b/.test(str)) mapped.push('Wednesday');
            if (str.includes('thursday') || /\bth\b|\bthu\b|\bthur\b|\bthurs\b|\br\b/.test(str)) mapped.push('Thursday');
            if (str.includes('friday') || /\bf\b|\bfri\b/.test(str)) mapped.push('Friday');
            
            return mapped;
        };

        // Populate the schedule object
        allItems.forEach(item => {
            const days = mapDaysToStandard(item.rawDay);
            // If an item occurs on multiple days (e.g. M/W), push it to both arrays
            days.forEach(day => {
                if (schedule[day]) {
                    schedule[day].push(item);
                }
            });
        });

        // 6. Return Payload
        return { 
            success: true, 
            data: schedule 
        };
        
    } catch (e) {
        console.error("api_getMyScheduleData Error: " + e.stack);
        return { success: false, message: `Failed to load schedule. Error: ${e.message}` };
    }
}