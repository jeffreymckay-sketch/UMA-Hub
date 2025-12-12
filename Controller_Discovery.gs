/**
 * -------------------------------------------------------------------
 * EVENT DISCOVERY CONTROLLER (The Scrubber)
 * -------------------------------------------------------------------
 */

/**
 * Scans a calendar for frequent keywords in event titles.
 * @param {string} calendarId - The ID of the calendar to scan.
 * @param {number} daysBack - How many days into the past to look.
 */
function api_discovery_scanCalendar(calendarId, daysBack) {
  try {
    if (!calendarId) throw new Error("No calendar selected.");
    
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - daysBack);
    
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error("Calendar not found or permission denied.");
    
    const events = cal.getEvents(start, end);
    if (events.length === 0) return { success: true, data: [], message: "No events found in range." };

    // 1. Tokenize Titles
    const stopWords = new Set(['and', 'the', 'for', 'with', 'at', 'by', 'to', 'in', 'on', 'of', 'a', 'an', 'is', '-', '|', ':', 'meeting', 'call']);
    const frequencyMap = {};
    
    events.forEach(evt => {
      const title = evt.getTitle().toLowerCase().replace(/[^a-z0-9\s]/g, ''); // Remove special chars
      const words = title.split(/\s+/);
      
      // Count individual words
      words.forEach(w => {
        if (w.length > 2 && !stopWords.has(w)) {
          frequencyMap[w] = (frequencyMap[w] || 0) + 1;
        }
      });
      
      // Optional: Count 2-word phrases for better context (e.g. "tech hub")
      for(let i=0; i<words.length-1; i++) {
        const phrase = words[i] + " " + words[i+1];
        if (!stopWords.has(words[i]) && !stopWords.has(words[i+1])) {
           frequencyMap[phrase] = (frequencyMap[phrase] || 0) + 1;
        }
      }
    });

    // 2. Convert to Array & Sort
    const sorted = Object.keys(frequencyMap)
      .map(key => ({ keyword: key, count: frequencyMap[key] }))
      .filter(item => item.count > 1) // Filter out one-offs
      .sort((a, b) => b.count - a.count) // Highest frequency first
      .slice(0, 50); // Top 50 patterns

    return { success: true, data: sorted, count: events.length };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Saves the classified rules to the Event_Types sheet.
 * @param {Array} newRules - [{ keyword: "ooo", category: "Out of Office" }, ...]
 */
function api_discovery_saveRules(newRules) {
  try {
    const ss = getMasterDataHub();
    let sheet = ss.getSheetByName(CONFIG.TABS.EVENT_TYPES);
    
    // Create sheet if missing
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.TABS.EVENT_TYPES);
      sheet.appendRow(["Category Name", "Keywords (Comma Separated)"]);
      sheet.getRange(1,1,1,2).setFontWeight("bold");
    }

    const data = sheet.getDataRange().getValues();
    // Map existing categories: { "Nursing": rowIndex }
    const categoryRowMap = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) categoryRowMap[data[i][0].toLowerCase()] = i + 1; // 1-based row index
    }

    // Process new rules
    newRules.forEach(rule => {
      const catName = rule.category.trim();
      const keyword = rule.keyword.trim();
      const catKey = catName.toLowerCase();

      if (categoryRowMap[catKey]) {
        // Update existing category
        const row = categoryRowMap[catKey];
        const currentKeywords = sheet.getRange(row, 2).getValue().toString();
        // Avoid duplicates
        if (!currentKeywords.toLowerCase().includes(keyword.toLowerCase())) {
          const updated = currentKeywords ? currentKeywords + ", " + keyword : keyword;
          sheet.getRange(row, 2).setValue(updated);
        }
      } else {
        // Create new category
        sheet.appendRow([catName, keyword]);
        // Update map so subsequent rules for this new category find it
        categoryRowMap[catKey] = sheet.getLastRow();
      }
    });

    return { success: true, message: "Rules saved successfully." };

  } catch (e) {
    return { success: false, message: e.message };
  }
}