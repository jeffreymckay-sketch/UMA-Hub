/**
 * DIAGNOSTIC TOOL: Check why Nursing docs aren't linking.
 * Run "debug_NursingDocMatcher_Fixed" to see the log report.
 */
function debug_NursingDocMatcher_Fixed() {
  console.log("--- STARTING DIAGNOSTIC (SELF-CONTAINED) ---");

  // --- LOCAL HELPER: Get Settings ---
  // This replaces the missing _getNursingSettings function
  const getConf = () => {
    try {
      // Access the global properties service directly
      const props = PropertiesService.getScriptProperties().getProperties();
      let raw = props['nursing_settings']; // Key from your Config
      let settings = {};

      if (raw) {
        try { settings = JSON.parse(raw); } catch (e) { settings = raw; }
      }
      
      // Helper to strip URL to ID
      const extract = (s) => {
        if (!s) return null;
        const m = s.match(/[-\w]{25,}/);
        return m ? m[0] : s;
      };

      return {
        sheetId: extract(settings.nursingSheetId),
        folderId: extract(settings.nursingFolderId)
      };
    } catch (e) {
      console.error("Error reading settings: " + e.message);
      return null;
    }
  };

  const config = getConf();
  
  if (!config || !config.folderId || !config.sheetId) {
    console.error("❌ CRITICAL ERROR: Settings are missing.");
    console.error("   > Please go to the 'Nursing > Settings' tab in the app and click 'Save Settings'.");
    console.error("   > Ensure both Sheet ID and Folder ID are filled in.");
    return;
  }

  console.log(`✅ Settings Found.`);
  console.log(`   > Folder ID: ${config.folderId}`);
  console.log(`   > Sheet ID:  ${config.sheetId}`);

  // --- STEP 1: SCAN DRIVE FOLDER ---
  try {
    const folder = DriveApp.getFolderById(config.folderId);
    console.log(`\n📂 SCANNING DRIVE FOLDER: "${folder.getName()}"`);
    
    const driveFiles = [];
    const files = folder.getFiles();
    let limit = 0;
    
    while (files.hasNext() && limit < 50) {
      const f = files.next();
      driveFiles.push(f.getName());
      limit++;
    }
    console.log(`   > Found ${driveFiles.length} files (showing first 5):`);
    driveFiles.slice(0, 5).forEach(n => console.log(`     - "${n}"`));

    // --- STEP 2: SCAN SPREADSHEET ---
    console.log(`\n📊 SCANNING SPREADSHEET...`);
    const ss = SpreadsheetApp.openById(config.sheetId);
    const sheets = ss.getSheets();
    
    let mismatchCount = 0;
    let matchCount = 0;

    // We only check the first valid sheet to keep the log readable
    for (const sheet of sheets) {
      if (sheet.getName().startsWith("_")) continue; // Skip hidden/DB sheets
      
      console.log(`   > Checking Tab: "${sheet.getName()}"`);
      
      // Basic parser to find exam names in the sheet (Col A usually)
      // This is a simplified version of your main parser for debugging
      const data = sheet.getDataRange().getValues();
      let examCount = 0;

      for (let i = 0; i < data.length; i++) {
        const rowVal = String(data[i][0]).trim(); // Assuming Exam Name is in Column A
        
        // Look for rows that look like exams (not headers, not "Augusta", not dates)
        if (rowVal && !rowVal.includes("Augusta") && !rowVal.includes("Date") && rowVal.length > 3) {
          
          // GENERATE EXPECTED NAME
          const expectedName = `${sheet.getName()} - ${rowVal}`;
          
          // CHECK FOR MATCH
          if (driveFiles.includes(expectedName)) {
            console.log(`     ✅ MATCH: "${expectedName}"`);
            matchCount++;
          } else {
            // Check for near-misses
            const fuzzy = driveFiles.find(f => f.includes(rowVal));
            
            console.log(`     ❌ MISSING: "${expectedName}"`);
            if (fuzzy) {
              console.log(`        👉 Did you mean: "${fuzzy}"?`);
              console.log(`           (Check for extra spaces, dashes vs hyphens, or tab name changes)`);
            }
            mismatchCount++;
          }
          examCount++;
          if (examCount >= 3) break; // Check max 3 exams per tab
        }
      }
      // Stop after one valid sheet
      if (examCount > 0) break;
    }

    console.log(`\n--- DIAGNOSTIC COMPLETE ---`);
    console.log(`Matches: ${matchCount}`);
    console.log(`Mismatches: ${mismatchCount}`);
    if (mismatchCount > 0) console.log("ACTION: Look at the 'MISSING' lines above. Compare the text exactly with the 'Did you mean' line.");

  } catch (e) {
    console.error("RUNTIME ERROR: " + e.message);
    console.error(e.stack);
  }
}