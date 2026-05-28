/**
 * -------------------------------------------------------------------
 * FILE: Controller_ProctorSchedule.js
 * DESCRIPTION: Aggregates Nursing and MLT data into a master schedule,
 * providing a 2-way sync with a "Shadow Database" spreadsheet.
 * -------------------------------------------------------------------
 */

const PS_SETTINGS_KEY = 'proctor_schedule_settings';

const PS_HEADERS = [
    "System_ID", // Hidden/Internal ID for merging
    "INTERNAL: SORT BY DATE (First Day Available)",
    "In-class Exam or Handout Date",
    "Day of Week",
    "Exam Window Dates",
    "Course",
    "Exam Name",
    "Campus",
    "Instructor (Last Name)",
    "Document Type - Instructions",
    "Instruction Mode",
    "Time and Duration",
    "Exam Instructions",
    "UMAAL Notes",
    "added to UMAAL Calendar",
    "Proofed Exam Instructions",
    "Non Accomodation Room at UMA"
];

function _getPSSettings() {
    let settings = { folderId: '', sheetId: '' };
    try {
        const raw = PropertiesService.getScriptProperties().getProperty(PS_SETTINGS_KEY);
        if (raw) settings = JSON.parse(raw);
    } catch(e) {}
    
    // Auto-extract IDs if they pasted full URLs
    const extractId = (input) => {
        if (!input) return '';
        const match = input.match(/[-\w]{25,}/);
        return match ? match[0] : input;
    };
    
    settings.folderId = extractId(settings.folderId);
    settings.sheetId = extractId(settings.sheetId);
    return settings;
}

function api_savePSSettings(settings) {
    try {
        PropertiesService.getScriptProperties().setProperty(PS_SETTINGS_KEY, JSON.stringify(settings));
        return { success: true, message: "Settings saved successfully." };
    } catch(e) {
        return { success: false, message: e.message };
    }
}

function api_createMasterSheet(folderId) {
    try {
        if (!folderId) throw new Error("Please provide a Folder ID first.");
        folderId = folderId.match(/[-\w]{25,}/) ? folderId.match(/[-\w]{25,}/)[0] : folderId;
        const folder = DriveApp.getFolderById(folderId);
        
        const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        const ss = SpreadsheetApp.create("Master Exam Schedule - " + dateStr);
        const file = DriveApp.getFileById(ss.getId());
        
        folder.addFile(file);
        DriveApp.getRootFolder().removeFile(file); // Move to target folder
        
        const sheet = ss.getActiveSheet();
        sheet.setName("Master Schedule");
        
        // Write Headers
        const headerRange = sheet.getRange(1, 1, 1, PS_HEADERS.length);
        headerRange.setValues([PS_HEADERS]);
        headerRange.setFontWeight("bold").setBackground("#003057").setFontColor("white");
        sheet.setFrozenRows(1);
        sheet.hideColumns(1); // Hide the System_ID column
        
        // Save to settings automatically
        const settings = _getPSSettings();
        settings.folderId = folderId;
        settings.sheetId = ss.getId();
        api_savePSSettings(settings);
        
        return { success: true, message: "Sheet created successfully!", sheetId: ss.getId() };
    } catch(e) {
        return { success: false, message: e.message };
    }
}

function api_getProctorScheduleData() {
    try {
        const settings = _getPSSettings();
        const liveMap = {};
        
        // 1. Fetch Live Data
        const nursingRes = api_getNursingData();
        const mltRes = api_getMLTData();
        
        const processLiveSheet = (program, sheetObj) => {
            if (!sheetObj || !sheetObj.exams) return;
            
            let courseCode = "Unknown";
            let courseTitle = "Unknown Course";
            let facultyName = sheetObj.faculty || sheetObj._meta?.facultyName || "Faculty Unassigned";
            
            if (sheetObj.course) {
                courseCode = sheetObj.course.code || "Unknown";
                courseTitle = sheetObj.course.name || sheetObj.fullTitle || "Unknown Course";
            }
            
            sheetObj.exams.forEach(exam => {
                const systemId = `${courseCode}|${exam.name}`.trim();
                
                // Formatting Helpers
                let lastName = facultyName;
                if (facultyName !== "Faculty Unassigned") {
                    const parts = facultyName.trim().split(' ');
                    lastName = parts[parts.length - 1];
                }
                
                let sortDate = "";
                let displayDate = "";
                let dayOfWeek = "";
                if (exam.date && exam.date !== "TBD" && exam.date !== "-") {
                    try {
                        const d = new Date(exam.date);
                        if (!isNaN(d.getTime())) {
                            sortDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
                            displayDate = Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
                            dayOfWeek = Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE");
                        }
                    } catch(e) {}
                }
                
                let td = exam.siteTime || "TBD";
                if (exam.duration && exam.duration !== "-") td += ` (${exam.duration})`;
                
                let inst = [];
                if (exam.password && exam.password !== "-") inst.push(`PW: ${exam.password}`);
                if (exam.itemsAllowed) inst.push(`Items: ${exam.itemsAllowed}`);
                if (exam.generalNotes) inst.push(`Notes: ${exam.generalNotes}`);
                const examInst = inst.join(' | ');

                liveMap[systemId] = {
                    systemId: systemId,
                    sortDate: sortDate,
                    date: exam.date || "",
                    displayDate: displayDate,
                    docUrl: exam.docUrl || null,
                    dayOfWeek: dayOfWeek,
                    windowDates: exam.date || "", // Default
                    course: `${courseCode} ${courseTitle}`,
                    examName: exam.name,
                    campus: "",
                    instructor: lastName,
                    docType: "",
                    instructionMode: "",
                    timeDuration: td,
                    instructions: examInst,
                    umaalNotes: "",
                    calendar: false,
                    proofed: false,
                    room: exam.room || "",
                    isOrphan: false,
                    program: program
                };
            });
        };

        if (nursingRes.success && nursingRes.data?.sheets) nursingRes.data.sheets.forEach(s => processLiveSheet("Nursing", s));
        if (mltRes.success && mltRes.data?.sheets) mltRes.data.sheets.forEach(s => processLiveSheet("MLT", s));

        // 2. Fetch Shadow Database (if configured)
        const sheetMap = {};
        if (settings.sheetId) {
            try {
                const ss = SpreadsheetApp.openById(settings.sheetId);
                const sheet = ss.getSheets()[0];
                const data = sheet.getDataRange().getValues();
                
                if (data.length > 0) {
                    const synonymConfig = {
                        systemId: ['system_id'],
                        sortDate: ['sort by date'],
                        date: ['handout date', 'in-class exam'],
                        dayOfWeek: ['day of week'],
                        windowDates: ['window dates'],
                        course: ['course'],
                        examName: ['exam name'],
                        campus: ['campus'],
                        instructor: ['instructor'],
                        docType: ['document type'],
                        instructionMode: ['instruction mode'],
                        timeDuration: ['time and duration'],
                        instructions: ['exam instructions'],
                        umaalNotes: ['umaal notes'],
                        calendar: ['calendar'],
                        proofed: ['proofed'],
                        room: ['room']
                    };
                    
                    const headerIdx = SheetReader.findHeaderRowHeuristic(data, ['course', 'date', 'campus'], 10);
                    if (headerIdx > -1) {
                        const colMap = SheetReader.mapColumnsBySynonyms(data[headerIdx], synonymConfig);
                        
                        for (let i = headerIdx + 1; i < data.length; i++) {
                            const row = data[i];
                            const sId = colMap.systemId > -1 ? String(row[colMap.systemId]).trim() : "";
                            if (!sId) continue;
                            
                            const parseBool = (val) => String(val).toLowerCase() === 'true' || val === true;
                            
                            sheetMap[sId] = {
                                systemId: sId,
                                windowDates: colMap.windowDates > -1 ? String(row[colMap.windowDates]) : "",
                                campus: colMap.campus > -1 ? String(row[colMap.campus]) : "",
                                docType: colMap.docType > -1 ? String(row[colMap.docType]) : "",
                                instructionMode: colMap.instructionMode > -1 ? String(row[colMap.instructionMode]) : "",
                                umaalNotes: colMap.umaalNotes > -1 ? String(row[colMap.umaalNotes]) : "",
                                calendar: colMap.calendar > -1 ? parseBool(row[colMap.calendar]) : false,
                                proofed: colMap.proofed > -1 ? parseBool(row[colMap.proofed]) : false,
                                // Keep old values for orphans
                                sortDate: colMap.sortDate > -1 ? String(row[colMap.sortDate]) : "",
                                date: colMap.date > -1 ? String(row[colMap.date]) : "",
                                dayOfWeek: colMap.dayOfWeek > -1 ? String(row[colMap.dayOfWeek]) : "",
                                course: colMap.course > -1 ? String(row[colMap.course]) : "",
                                examName: colMap.examName > -1 ? String(row[colMap.examName]) : "",
                                instructor: colMap.instructor > -1 ? String(row[colMap.instructor]) : "",
                                timeDuration: colMap.timeDuration > -1 ? String(row[colMap.timeDuration]) : "",
                                instructions: colMap.instructions > -1 ? String(row[colMap.instructions]) : "",
                                room: colMap.room > -1 ? String(row[colMap.room]) : ""
                            };
                        }
                    }
                }
            } catch(e) {
                console.warn("Failed to read Master Sheet: " + e.message);
            }
        }

        // 3. Merge Data
        const finalArray = [];
        
        // Pass 1: Live Data (Overwrite with user inputs if they exist)
        for (const [sId, liveObj] of Object.entries(liveMap)) {
            if (sheetMap[sId]) {
                const sheetObj = sheetMap[sId];
                // Apply user overrides
                liveObj.windowDates = sheetObj.windowDates || liveObj.windowDates;
                liveObj.campus = sheetObj.campus;
                liveObj.docType = sheetObj.docType;
                liveObj.instructionMode = sheetObj.instructionMode;
                liveObj.umaalNotes = sheetObj.umaalNotes;
                liveObj.calendar = sheetObj.calendar;
                liveObj.proofed = sheetObj.proofed;
            }
            finalArray.push(liveObj);
        }
        
        // Pass 2: Orphans (In sheet, but missing from Live)
        for (const [sId, sheetObj] of Object.entries(sheetMap)) {
            if (!liveMap[sId]) {
                sheetObj.isOrphan = true;
                finalArray.push(sheetObj);
            }
        }
        
        // 4. Sort by Date
        finalArray.sort((a, b) => {
            if (!a.sortDate) return 1;
            if (!b.sortDate) return -1;
            return a.sortDate.localeCompare(b.sortDate);
        });

        return { success: true, data: finalArray, settings: settings };

    } catch (e) {
        console.error("Master Schedule Error: " + e.stack);
        return { success: false, message: e.message };
    }
}

function api_exportProctorSchedule(dataArray) {
    try {
        const settings = _getPSSettings();
        if (!settings.sheetId) throw new Error("No Master Sheet configured. Please create one in settings.");
        
        const ss = SpreadsheetApp.openById(settings.sheetId);
        const sheet = ss.getSheets()[0];
        
        // Prepare Data Array
        const outputRows = [];
        dataArray.forEach(obj => {
            outputRows.push([
                obj.systemId || "",
                obj.sortDate || "",
                obj.date || "",
                obj.dayOfWeek || "",
                obj.windowDates || "",
                obj.course || "",
                obj.examName || "",
                obj.campus || "",
                obj.instructor || "",
                obj.docType || "",
                obj.instructionMode || "",
                obj.timeDuration || "",
                obj.instructions || "",
                obj.umaalNotes || "",
                obj.calendar ? true : false,
                obj.proofed ? true : false,
                obj.room || ""
            ]);
        });
        
        // Write to Sheet
        // We clear everything below row 1 to handle deletions cleanly
        if (sheet.getLastRow() > 1) {
            sheet.getRange(2, 1, sheet.getLastRow() - 1, PS_HEADERS.length).clearContent().setBackground(null);
        }
        
        if (outputRows.length > 0) {
            const range = sheet.getRange(2, 1, outputRows.length, PS_HEADERS.length);
            range.setValues(outputRows);
            
            // Format checkboxes
            sheet.getRange(2, 15, outputRows.length, 2).insertCheckboxes();
            
            // Highlight Orphans in Red
            const backgrounds = [];
            dataArray.forEach(obj => {
                const rowColors = new Array(PS_HEADERS.length).fill(obj.isOrphan ? '#ffcdd2' : null);
                backgrounds.push(rowColors);
            });
            range.setBackgrounds(backgrounds);
        }
        
        return { success: true, message: "Schedule synced to Google Sheets successfully!" };
        
    } catch(e) {
        return { success: false, message: e.message };
    }
}