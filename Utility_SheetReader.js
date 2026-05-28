/**
 * @file Utility_SheetReader.js
 * @description A standalone utility for parsing highly unstructured, messy Google Sheets
 * using 'Heuristic' pattern-matching instead of strict row numbers.
 */

const SheetReader = {

  /**
   * Helper to aggressively sanitize strings for comparison.
   * @param {string} str The string to sanitize.
   * @returns {string} Lowercased, trimmed string with excess whitespace removed.
   */
  sanitize: function(str) {
    if (str === null || str === undefined) return '';
    return String(str).toLowerCase().trim().replace(/\s+/g, ' ');
  },

  /**
   * Scans rows to find the most likely header row by scoring them based on keyword presence.
   * @param {Array<Array<any>>} data 2D array of sheet data.
   * @param {Array<string>} keywords Array of critical words expected in the header.
   * @param {number} maxScanDepth Maximum number of rows to scan (defaults to 20).
   * @returns {number} The index of the highest-scoring row, or -1 if no row scores > 0.
   */
  findHeaderRowHeuristic: function(data, keywords, maxScanDepth = 20) {
    if (!data || data.length === 0) return -1;
    
    let bestRowIndex = -1;
    let maxScore = 0;
    const sanitizedKeywords = keywords.map(k => this.sanitize(k));

    for (let i = 0; i < Math.min(data.length, maxScanDepth); i++) {
      if (!data[i]) continue;
      
      let rowScore = 0;
      const rowStr = data[i].map(cell => this.sanitize(cell)).join(' ');

      sanitizedKeywords.forEach(keyword => {
        if (rowStr.includes(keyword)) {
          rowScore++;
        }
      });

      // Tie-breaker goes to the earlier row
      if (rowScore > maxScore) {
        maxScore = rowScore;
        bestRowIndex = i;
      }
    }

    return maxScore > 0 ? bestRowIndex : -1;
  },

  /**
   * Maps column indices based on arrays of synonyms.
   * @param {Array<any>} headerRow The row array containing the headers.
   * @param {Object} synonymConfig An object where keys are your internal column names, 
   *                               and values are arrays of acceptable string synonyms.
   *                               Example: { startTime: ['start time', 'time', 'onsite'] }
   * @returns {Object} An object mapping internal names to column indices (or -1 if not found).
   */
  mapColumnsBySynonyms: function(headerRow, synonymConfig) {
    const colMap = {};
    const sanitizedHeaders = headerRow.map(h => this.sanitize(h));

    for (const [key, synonyms] of Object.entries(synonymConfig)) {
      colMap[key] = -1; // Default to not found
      const sanitizedSynonyms = synonyms.map(s => this.sanitize(s));

      for (let i = 0; i < sanitizedHeaders.length; i++) {
        const headerStr = sanitizedHeaders[i];
        if (!headerStr) continue;

        // Check if any synonym matches or is included in the header string
        const match = sanitizedSynonyms.some(syn => headerStr === syn || headerStr.includes(syn));
        if (match) {
          colMap[key] = i;
          break; // Stop looking for this key once found
        }
      }
    }
    return colMap;
  },

  /**
   * Scans rows to find a specific "anchor" row (like the start of a roster).
   * @param {Array<Array<any>>} data 2D array of sheet data.
   * @param {Array<string>} anchorWords Array of words to look for in the first column.
   * @param {number} startRow The index to start scanning from.
   * @returns {number} The index of the row containing an anchor word, or -1.
   */
  findAnchorRow: function(data, anchorWords, startRow = 0) {
    if (!data || data.length === 0) return -1;
    
    const sanitizedAnchors = anchorWords.map(w => this.sanitize(w));

    for (let i = startRow; i < data.length; i++) {
      if (!data[i] || data[i].length === 0) continue;
      
      const firstCell = this.sanitize(data[i][0]);
      if (!firstCell) continue;

      if (sanitizedAnchors.some(anchor => firstCell.includes(anchor))) {
        return i;
      }
    }
    return -1;
  },

  /**
   * Dynamically reads a list/roster starting from a specific row until a stopping condition is met.
   * @param {Array<Array<any>>} data 2D array of sheet data.
   * @param {number} startRow The row index to start reading data from.
   * @param {number} nameColIndex The column index where the primary data (e.g., student name) lives.
   * @param {Array<string>} stopWords Array of words that, if found in the name column, stop reading.
   * @param {number} maxBlankRows The number of consecutive blank rows before stopping (defaults to 3).
   * @returns {Array<Object>} An array of objects representing the rows read.
   */
  readDynamicRoster: function(data, startRow, nameColIndex, stopWords, maxBlankRows = 3) {
    const roster = [];
    let blankCount = 0;
    const sanitizedStopWords = stopWords.map(w => this.sanitize(w));

    for (let i = startRow; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length <= nameColIndex) {
        blankCount++;
        if (blankCount >= maxBlankRows) break;
        continue;
      }

      const primaryCell = this.sanitize(row[nameColIndex]);

      // Check for blank row
      if (primaryCell === '') {
        blankCount++;
        if (blankCount >= maxBlankRows) break;
        continue;
      }

      // Reset blank count since we found data
      blankCount = 0;

      // Check for stop words
      if (sanitizedStopWords.some(stopWord => primaryCell === stopWord || primaryCell.includes(stopWord))) {
        break; // Stop reading entirely
      }

      // Format the row, replacing empty cells with 'TBD'
      const formattedRow = row.map(cell => {
         const strCell = String(cell).trim();
         return strCell === '' ? 'TBD' : cell; // Keep original type if not empty string
      });

      roster.push({
        rowIndex: i,
        name: String(row[nameColIndex]).trim(), // Keep original capitalization for the name
        data: formattedRow
      });
    }

    return roster;
  }
};
