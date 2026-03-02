/**
 * -------------------------------------------------------------------
 * AD-HOC DOCUMENT GENERATOR CONTROLLER
 * Decoupled logic to generate a single proctoring doc from a UI payload.
 * Reads directly from Google Drive for the View tab (No Database).
 * -------------------------------------------------------------------
 */

/**
 * Main API Endpoint called by the frontend.
 * @param {Object} payload The form data and roster array.
 */
function api_generateSingleDoc(payload) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); 

        // 1. Resolve Folder ID from Settings
        const allSettings = getSettings();
        let sdSettings = {};
        try { 
            sdSettings = JSON.parse(allSettings.single_doc_settings || '{}'); 
        } catch(e) {}

        const folderInput = sdSettings.targetFolderId;
        if (!folderInput) throw new Error("Output folder not configured. Please set it in the Settings tab.");

        const folderId = _singleDoc_extractFolderId(folderInput);
        if (!folderId) throw new Error("Could not extract a valid Folder ID from the Settings input.");

        let targetFolder;
        try {
            targetFolder = DriveApp.getFolderById(folderId);
        } catch (e) {
            throw new Error("Cannot access the Target Folder. Check the ID in Settings or your permissions.");
        }

        // 2. Create the Document
        const docTitle = payload.docTitle || "Untitled Proctoring Doc";
        const doc = DocumentApp.create(docTitle);
        const docId = doc.getId();

        // 3. Move Document to Target Folder
        const file = DriveApp.getFileById(docId);
        targetFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file); 

        // 4. Populate Document Content
        _singleDoc_populateDoc(doc, payload);
        const docUrl = file.getUrl();

        // 5. Log to global system audit (if available)
        if (typeof logSystemAction === 'function') {
            logSystemAction("Ad-Hoc Generator", "Created Doc", docTitle, docId, `Target Folder: ${folderId}`);
        }

        return { success: true, url: docUrl };

    } catch (e) {
        return { success: false, message: e.message };
    } finally {
        lock.releaseLock();
    }
}

/**
 * Fetches all documents directly from the target Google Drive folder.
 */
function api_getSingleDocs() {
    try {
        const allSettings = getSettings();
        let sdSettings = {};
        try { 
            sdSettings = JSON.parse(allSettings.single_doc_settings || '{}'); 
        } catch(e) {}

        const folderInput = sdSettings.targetFolderId;
        if (!folderInput) {
            return { success: true, data: [], message: "Folder not configured." };
        }

        const folderId = _singleDoc_extractFolderId(folderInput);
        if (!folderId) {
            return { success: false, message: "Invalid Folder ID in settings." };
        }

        let folder;
        try {
            folder = DriveApp.getFolderById(folderId);
        } catch (e) {
            return { success: false, message: "Cannot access the target folder. Please check permissions." };
        }

        // Fetch only Google Docs from the folder
        const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
        const results = [];

        while (files.hasNext()) {
            const file = files.next();
            results.push({
                DocTitle: file.getName(),
                DocUrl: file.getUrl(),
                // Format Date for safe transport to frontend
                DateCreated: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "yyyy-MM-dd h:mm a"),
                FolderId: folderId
            });
        }

        // Sort newest first
        results.sort((a, b) => new Date(b.DateCreated) - new Date(a.DateCreated));

        return { success: true, data: results };
        
    } catch (e) {
        return { success: false, message: e.message };
    }
}


// ===================================================================
// DOC POPULATION ENGINE
// ===================================================================

function _singleDoc_populateDoc(doc, payload) {
    const body = doc.getBody();
    
    // 1. Title (Course)
    const displayTitle = payload.courseTitle || "Exam Details";
    const titlePara = body.getChild(0).asParagraph();
    titlePara.setText(displayTitle);
    titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    // 2. Heading 1 (Faculty)
    const facultyPara = body.insertParagraph(1, payload.facultyName || "Faculty Unassigned");
    facultyPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // 3. Exam Details (Heading 2)
    const examDetailsMap = {};
    if (payload.date) examDetailsMap['Date'] = payload.date;
    if (payload.siteTime) examDetailsMap['Start Time (On Site)'] = payload.siteTime;
    if (payload.zoomTime) examDetailsMap['Start Time (Zoom)'] = payload.zoomTime;
    if (payload.duration) examDetailsMap['Duration'] = payload.duration;
    if (payload.password) examDetailsMap['Password'] = payload.password;
    
    _singleDoc_updateKeyValueSection(body, "Exam Details", examDetailsMap, DocumentApp.ParagraphHeading.HEADING2);
    
    // 4. Important Links
    _singleDoc_insertImportantLinks(body);
    
    // 5. General Accommodations / Notes (Green Box)
    if (payload.generalNotes) {
        _singleDoc_updateOrCreateHighlightedSection(body, "Exam Accommodations", payload.generalNotes);
    }
    
    // 6. Roster / Students
    _singleDoc_buildLocationRosters(body, payload.students);
    
    doc.saveAndClose();
}

// ===================================================================
// FORMATTING HELPERS
// ===================================================================

function _singleDoc_extractFolderId(input) {
    if (!input) return null;
    const trimmed = input.trim();
    if (!trimmed.includes('/') && !trimmed.includes('http')) return trimmed;
    const match = trimmed.match(/\/folders\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
}

function _singleDoc_updateKeyValueSection(body, sectionTitle, dataMap, headingLevel) {
    const insertAt = body.getNumChildren();
    body.insertParagraph(insertAt, sectionTitle).setHeading(headingLevel);
    
    let insertCount = 1;
    for (const [key, value] of Object.entries(dataMap)) {
        const text = `${key}: ${value}`;
        const li = body.insertListItem(insertAt + insertCount, text);
        li.setGlyphType(DocumentApp.GlyphType.BULLET);

        const keyCheck = key.toLowerCase();
        if (keyCheck.includes("date") || keyCheck.includes("time") || keyCheck.includes("password")) {
            li.editAsText().setBackgroundColor('#ffff00');
        }
        insertCount++;
    }
}

function _singleDoc_insertImportantLinks(body) {
    const insertAt = body.getNumChildren();
    body.insertParagraph(insertAt, "Important Links").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    const links = [
      { text: "Red Flag Reporting Form", url: "https://docs.google.com/forms/d/e/1FAIpQLSfORKCKol8SsRldNKfvsDy3ILNs9HcFv3gKb8TuxrNrlqxijw/viewform" }
    ];
    
    links.forEach((link, i) => {
      const li = body.insertListItem(insertAt + 1 + i, link.text);
      li.setLinkUrl(link.url);
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
    });
}

function _singleDoc_updateOrCreateHighlightedSection(body, sectionTitle, content) {
    const insertAt = body.getNumChildren();
    body.insertParagraph(insertAt, sectionTitle).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const p = body.insertParagraph(insertAt + 1, content);
    
    const style = {};
    style[DocumentApp.Attribute.BACKGROUND_COLOR] = '#e8f5e9'; // Light Green
    p.setAttributes(style);
    
    p.setIndentStart(20);
    p.setIndentEnd(20);
    p.setSpacingBefore(10);
    p.setSpacingAfter(10);
}

function _singleDoc_buildLocationRosters(body, studentsArray) {
    if (!studentsArray || studentsArray.length === 0) return;

    const grouped = {};
    studentsArray.forEach(student => {
        const loc = student.location || "Unassigned";
        if (!grouped[loc]) grouped[loc] = [];
        grouped[loc].push(student);
    });

    const sortedLocations = Object.keys(grouped).sort((a, b) => a.localeCompare(b));

    const rostersIndex = body.getNumChildren();
    body.insertParagraph(rostersIndex, "Location Rosters").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    let currentIndex = rostersIndex + 1;

    sortedLocations.forEach(loc => {
        const displayLocation = loc.replace(/\w\S*/g, (w) => (w.replace(/^\w/, (c) => c.toUpperCase())));
        
        body.insertParagraph(currentIndex, displayLocation).setHeading(DocumentApp.ParagraphHeading.HEADING3);
        currentIndex++;

        const studentsInLoc = grouped[loc];
        studentsInLoc.forEach(student => {
            const studentName = student.name;
            const note = student.note;

            const li = body.insertListItem(currentIndex, studentName);
            if (note) {
                const t = li.appendText(` [${note}]`);
                t.setBold(true).setForegroundColor('#000000').setBackgroundColor('#fff59d'); 
            }
            currentIndex++;
        });
    });
}