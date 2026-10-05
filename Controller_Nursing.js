/**
 * @file Controller_Nursing.gs
 * @description
 * Nursing controller, resilient spreadsheet parser, document generation,
 * accommodations persistence, and calendar synchronization.
 *
 * IMPORTANT DESIGN RULES
 *
 * 1. Nursing course tabs are identified by:
 *      NUR101
 *      NUR 101
 *
 * 2. The tab name is authoritative for the normalized course code.
 *
 * 3. The spreadsheet itself is searched across ALL columns for:
 *      - course title
 *      - exam header
 *      - roster/location header
 *
 * 4. No fixed Nursing row/column offsets are used.
 *
 * 5. Header matching is forgiving and synonym-based.
 *
 * 6. Roster rows are classified cell-by-cell.
 *
 * 7. Student font colors are preserved.
 *
 * 8. Existing Drive folder/file naming behavior is preserved by deriving
 *    legacy metadata from the original course title cell.
 */


/* ========================================================================
 * GENERAL NURSING HELPERS
 * ====================================================================== */


/**
 * Reads and normalizes Nursing settings.
 */
function _getNursingSettings() {
  const allProps = getSettings();

  let raw = allProps.nursing_settings;
  let settings = {};

  if (raw && typeof raw === "string") {
    try {
      settings = JSON.parse(raw);
    } catch (e) {
      console.error(
        "Nursing settings JSON parse error:",
        e
      );
      settings = {};
    }
  } else if (
    raw &&
    typeof raw === "object"
  ) {
    settings = raw;
  }

  const extractId = input => {
    if (!input) {
      return null;
    }

    const value =
      String(input).trim();

    const match =
      value.match(/[-\w]{25,}/);

    return match
      ? match[0]
      : value;
  };

  return {
    sheetId:
      extractId(
        settings.nursingSheetId
      ),

    folderId:
      extractId(
        settings.nursingFolderId
      ),

    calendarId:
      settings.nursingCalendarId ||
      "",

    customNotes:
      settings.customNotes ||
      "",

    /*
     * These remain configurable, but are now preferred synonyms rather
     * than rigid exact header requirements.
     */
    examHeader:
      settings.examHeader ||
      "exam",

    dateHeader:
      settings.dateHeader ||
      "date",

    passwordHeader:
      settings.passwordHeader ||
      "password",

    /*
     * Optional location hint. The parser does not require it.
     */
    rosterKeyword:
      settings.rosterKeyword ||
      "augusta",

    stopKeywords:
      settings.stopKeywords ||
      "total,totals,average",

    rosterStopKeywords:
      settings.rosterStopKeywords ||
      ""
  };
}


/**
 * Normalizes strings for filename/folder comparison.
 *
 * Preserved from original implementation because historical Drive
 * objects may contain punctuation/spacing variations.
 */
function _normalizeStr(str) {
  if (!str) {
    return "";
  }

  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}


/**
 * Simple title casing used by Nursing documents.
 */
function _toTitleCase(str) {
  if (!str) {
    return "";
  }

  return String(str).replace(
    /\w\S*/g,
    function(txt) {
      return (
        txt.charAt(0).toUpperCase() +
        txt.substr(1).toLowerCase()
      );
    }
  );
}


/**
 * Parses the legacy Nursing course-title convention:
 *
 *   NUR305 Health Assessment : Professor Lakey
 *
 * into the components historically used by Drive/document generation.
 *
 * IMPORTANT:
 * folderKey remains the complete course portion, NOT merely NUR305.
 */
function _parseCellA1(cellValue) {
  if (!cellValue) {
    return {
      folderKey: "Unknown",
      courseTitle: "Unknown",
      facultyName: "Faculty Unassigned",
      fullTitlePrefix:
        "Unknown : Faculty Unassigned"
    };
  }

  const str =
    String(cellValue).trim();

  let separator = ":";

  if (!str.includes(":")) {
    separator = "-";
  }

  const parts =
    str.split(separator);

  const courseInfo =
    parts[0].trim();

  let facultyInfo = "";

  if (parts.length > 1) {
    facultyInfo =
      parts
        .slice(1)
        .join(separator)
        .trim();
  }

  if (!facultyInfo) {
    facultyInfo =
      "Faculty Unassigned";
  }

  const fullTitlePrefix =
    `${courseInfo} : ${facultyInfo}`;

  return {
    folderKey:
      courseInfo,

    courseTitle:
      courseInfo,

    facultyName:
      facultyInfo,

    fullTitlePrefix:
      fullTitlePrefix
  };
}


/**
 * Finds a child folder by normalized name.
 *
 * This preserves the original Drive lookup behavior.
 */
function _findSubFolder(
  parentFolder,
  targetName
) {
  const targetNorm =
    _normalizeStr(targetName);

  const folders =
    parentFolder.getFolders();

  while (folders.hasNext()) {
    const folder =
      folders.next();

    if (
      _normalizeStr(
        folder.getName()
      ) === targetNorm
    ) {
      return folder;
    }
  }

  return null;
}


/**
 * Existing date display helper retained.
 */
function formatDateToPlainLanguage(
  input
) {
  if (!input) {
    return "";
  }

  let dateObj;

  if (input instanceof Date) {
    dateObj = input;
  } else {
    const str =
      String(input)
        .replace(
          /(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|st|nd|rd|th),?/gi,
          ""
        )
        .trim();

    dateObj = new Date(str);
  }

  if (
    isNaN(
      dateObj.getTime()
    )
  ) {
    return String(input);
  }

  return Utilities.formatDate(
    dateObj,
    Session.getScriptTimeZone(),
    "MMMM d, yyyy"
  );
}


/**
 * Returns true for valid Nursing class tab names.
 */
function _isNursingCourseTab(
  sheetName
) {
  return /^NUR\s?\d{3}$/i.test(
    String(sheetName || "").trim()
  );
}


/**
 * Normalizes:
 *
 *   NUR305
 *   NUR 305
 *
 * to:
 *
 *   NUR305
 */
function _normalizeNursingCourseCode(
  value
) {
  const match =
    String(value || "")
      .trim()
      .match(
        /^NUR\s*(\d{3})$/i
      );

  return match
    ? "NUR" + match[1]
    : "";
}


/* ========================================================================
 * DOCUMENT NAMING COMPATIBILITY
 * ====================================================================== */


/**
 * Returns the legacy document metadata used by the original system.
 *
 * Priority:
 *
 * 1. Payload _meta
 * 2. Original title cell retained as rawTitle/rawA1
 * 3. Course object
 *
 * This is critical for historical Drive compatibility.
 */
function _getNursingDocumentMeta(
  sheetData
) {
  if (
    sheetData &&
    sheetData._meta &&
    sheetData._meta.folderKey &&
    sheetData._meta.fullTitlePrefix
  ) {
    return sheetData._meta;
  }

  /*
   * The resilient parser retains the original course-title cell.
   */
  const rawTitle =
    sheetData &&
    (
      sheetData.rawTitle ||
      sheetData.rawA1
    );

  if (rawTitle) {
    const meta =
      _parseCellA1(
        rawTitle
      );

    /*
     * Avoid returning a meaningless Unknown structure if we have actual
     * course data available.
     */
    if (
      meta.folderKey &&
      meta.folderKey !== "Unknown"
    ) {
      return meta;
    }
  }

  const course =
    sheetData &&
    sheetData.course
      ? sheetData.course
      : {};

  const code =
    course.code ||
    "Unknown";

  const name =
    course.name ||
    code;

  const faculty =
    course.faculty ||
    "Faculty Unassigned";

  /*
   * This fallback is only used when the original title cell is unavailable.
   */
  const coursePortion =
    `${code} ${name}`.trim();

  return {
    folderKey:
      coursePortion,

    courseTitle:
      coursePortion,

    facultyName:
      faculty,

    fullTitlePrefix:
      `${coursePortion} : ${faculty}`
  };
}


/* ========================================================================
 * MAIN NURSING DATA API
 * ====================================================================== */


/**
 * Main SPA endpoint.
 *
 * Returns:
 *
 * {
 *   success: true,
 *   data: {
 *     settings: {...},
 *     sheets: [
 *       {
 *         sheetName: "NUR305",
 *         rawA1: "...",
 *         rawTitle: "...",
 *         folderKey: "...",
 *         _meta: {...},
 *         course: {...},
 *         exams: [...]
 *       }
 *     ],
 *     diagnostics: [...]
 *   }
 * }
 */
function api_getNursingData() {
  try {
    const config =
      _getNursingSettings();

    if (
      !config.sheetId ||
      !config.folderId
    ) {
      return {
        success: false,
        message:
          "Nursing settings missing. Please save Sheet/Folder IDs in Settings."
      };
    }

    const ss =
      SpreadsheetApp.openById(
        config.sheetId
      );

    const sheets =
      ss.getSheets();

    const rootFolder =
      DriveApp.getFolderById(
        config.folderId
      );

    /*
     * IMPORTANT:
     * Use the real accommodations database.
     *
     * The previous parser's temporary empty object would discard
     * saved notes/tags.
     */
    const dbMap =
      getAccommodationsDBMap();

    const courseData = [];
    const diagnostics = [];

    sheets.forEach(sheet => {
      const sheetName =
        sheet.getName();

      /*
       * Only Nursing class tabs are eligible.
       * Instructions/template/etc. tabs are silently ignored.
       */
      if (
        !_isNursingCourseTab(
          sheetName
        )
      ) {
        return;
      }

      try {
        const parsed =
          parseNursingSheet_V2(
            sheet,
            dbMap,
            config
          );

        if (!parsed) {
          diagnostics.push({
            sheetName:
              sheetName,

            success:
              false,

            message:
              "No recognizable Nursing exam table was found."
          });

          return;
        }

        /*
         * Reconstruct the exact legacy metadata used by the old
         * Drive/document code.
         */
        const meta =
          _getNursingDocumentMeta(
            parsed
          );

        parsed._meta =
          meta;

        /*
         * Find the historical course folder.
         *
         * This intentionally uses meta.folderKey rather than
         * course.code.
         */
        const targetFolder =
          _findSubFolder(
            rootFolder,
            meta.folderKey
          );

        /*
         * Load existing files once per course folder.
         */
        const existingFiles = [];

        if (targetFolder) {
          const files =
            targetFolder.getFiles();

          while (files.hasNext()) {
            const f =
              files.next();

            existingFiles.push({
              name:
                f.getName(),

              nameNorm:
                _normalizeStr(
                  f.getName()
                ),

              url:
                f.getUrl()
            });
          }
        }

        /*
         * Attach existing document URLs.
         */
        parsed.exams.forEach(
          exam => {
            const examNameNorm =
              _normalizeStr(
                exam.name
              );

            const expectedTitle =
              `${meta.fullTitlePrefix} - ${exam.name}`;

            const expectedTitleNorm =
              _normalizeStr(
                expectedTitle
              );

            /*
             * 1. Exact normalized canonical filename.
             */
            let match =
              existingFiles.find(
                f =>
                  f.nameNorm ===
                  expectedTitleNorm
              );

            /*
             * 2. Legacy/fuzzy exam-name match.
             */
            if (!match) {
              match =
                existingFiles.find(
                  f =>
                    f.nameNorm.includes(
                      examNameNorm
                    )
                );
            }

            exam.docUrl =
              match
                ? match.url
                : null;
          }
        );

        /*
         * Preserve the old behavior of only displaying courses with
         * recognizable exams.
         */
        if (
          parsed.exams.length > 0
        ) {
          courseData.push({
            sheetName:
              sheetName,

            course:
              parsed.course,

            exams:
              parsed.exams,

            /*
             * Required by document creation/update.
             */
            _meta:
              meta,

            /*
             * Preserve original title information.
             */
            rawA1:
              parsed.rawA1,

            rawTitle:
              parsed.rawTitle,

            folderKey:
              meta.folderKey
          });
        }

      } catch (sheetError) {
        diagnostics.push({
          sheetName:
            sheetName,

          success:
            false,

          message:
            sheetError &&
            sheetError.message
              ? sheetError.message
              : String(sheetError)
        });

        console.error(
          "Nursing parser error on '" +
          sheetName +
          "': " +
          (
            sheetError &&
            sheetError.stack
              ? sheetError.stack
              : sheetError
          )
        );
      }
    });

    return {
      success: true,

      data: {
        sheets:
          courseData,

        settings:
          config,

        diagnostics:
          diagnostics
      }
    };

  } catch (e) {
    console.error(
      "Critical Error in api_getNursingData: " +
      (
        e &&
        e.stack
          ? e.stack
          : e
      )
    );

    return {
      success: false,
      message:
        e &&
        e.message
          ? e.message
          : String(e)
    };
  }
}


/* ========================================================================
 * COURSE DISCOVERY
 * ====================================================================== */


/**
 * Finds the course-title cell somewhere in the first several rows and
 * across every column.
 */
function _findAndParseCourseInfo(
  data,
  sheetName
) {
  const courseCode =
    _normalizeNursingCourseCode(
      sheetName
    );

  if (!courseCode) {
    return null;
  }

  const digits =
    courseCode.substring(3);

  const courseRegex =
    new RegExp(
      "\\bNUR\\s*" +
      digits +
      "\\b",
      "i"
    );

  const searchLimit =
    Math.min(
      data.length,
      15
    );

  let bestCandidate =
    null;

  for (
    let r = 0;
    r < searchLimit;
    r++
  ) {
    const row =
      data[r] || [];

    for (
      let c = 0;
      c < row.length;
      c++
    ) {
      const rawValue =
        row[c];

      if (
        rawValue === null ||
        rawValue === undefined ||
        rawValue === ""
      ) {
        continue;
      }

      const text =
        String(rawValue).trim();

      if (
        !text ||
        !courseRegex.test(text)
      ) {
        continue;
      }

      let score = 10;

      /*
       * Prefer cells that begin with NUR###.
       */
      if (
        new RegExp(
          "^\\s*NUR\\s*" +
          digits +
          "\\b",
          "i"
        ).test(text)
      ) {
        score += 5;
      }

      const withoutCode =
        text
          .replace(
            courseRegex,
            ""
          )
          .trim();

      if (
        withoutCode.length >= 3
      ) {
        score += 5;
      }

      if (
        text.length <= 120
      ) {
        score += 2;
      }

      if (
        !bestCandidate ||
        score >
          bestCandidate.score
      ) {
        bestCandidate = {
          row:
            r,

          col:
            c,

          rawValue:
            text,

          score:
            score
        };
      }
    }
  }

  /*
   * No title cell is acceptable.
   * The tab name remains authoritative.
   */
  if (!bestCandidate) {
    return {
      rawA1: "",
      rawTitle: "",
      folderKey:
        courseCode,
      courseCode:
        courseCode,
      courseTitle:
        courseCode,
      faculty:
        "",
      foundTitle:
        false,
      position: {
        row: 0,
        col: -1
      }
    };
  }

  const rawTitle =
    bestCandidate.rawValue;

  let remainder =
    rawTitle
      .replace(
        courseRegex,
        ""
      )
      .trim()
      .replace(
        /^[\s:;,\-–—]+/,
        ""
      )
      .trim();

  let courseTitle =
    remainder;

  let faculty = "";

  const colonIndex =
    remainder.indexOf(":");

  if (
    colonIndex > 0
  ) {
    courseTitle =
      remainder
        .substring(
          0,
          colonIndex
        )
        .trim();

    faculty =
      remainder
        .substring(
          colonIndex + 1
        )
        .trim();

  } else {
    /*
     * Tolerate:
     * Health Assessment - Professor Lakey
     */
    const facultySplit =
      remainder.match(
        /^(.*?)\s+[-–—]\s+((?:professor|prof\.?|instructor|faculty)\b.*)$/i
      );

    if (facultySplit) {
      courseTitle =
        facultySplit[1].trim();

      faculty =
        facultySplit[2].trim();
    }
  }

  if (!courseTitle) {
    courseTitle =
      courseCode;
  }

  return {
    rawA1:
      rawTitle,

    rawTitle:
      rawTitle,

    /*
     * IMPORTANT:
     * This legacy field is used only as a fallback.
     * The API reconstructs the exact legacy _meta with _parseCellA1.
     */
    folderKey:
      courseCode,

    courseCode:
      courseCode,

    courseTitle:
      courseTitle,

    faculty:
      faculty,

    foundTitle:
      true,

    position: {
      row:
        bestCandidate.row,

      col:
        bestCandidate.col
    }
  };
}


/* ========================================================================
 * MAIN RESILIENT PARSER
 * ====================================================================== */


/**
 * Parses a single Nursing worksheet.
 */
function parseNursingSheet_V2(
  sheet,
  dbMap,
  config
) {
  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    !lastRow ||
    !lastColumn
  ) {
    return null;
  }

  const range =
    sheet.getRange(
      1,
      1,
      lastRow,
      lastColumn
    );

  const data =
    range.getValues();

  const fontColors =
    range.getFontColors();

  const fontLines =
    range.getFontLines();

  let actualLastRow =
    data.length;

  while (
    actualLastRow > 0 &&
    !_nursingRowHasContent(
      data[
        actualLastRow - 1
      ]
    )
  ) {
    actualLastRow--;
  }

  if (
    actualLastRow === 0
  ) {
    return null;
  }

  const finalData =
    data.slice(
      0,
      actualLastRow
    );

  const finalFontColors =
    fontColors.slice(
      0,
      actualLastRow
    );

  const finalFontLines =
    fontLines.slice(
      0,
      actualLastRow
    );

  const sheetName =
    sheet.getName();

  const courseCode =
    _normalizeNursingCourseCode(
      sheetName
    );

  if (!courseCode) {
    return null;
  }

  const courseInfo =
    _findAndParseCourseInfo(
      finalData,
      sheetName
    );

  if (!courseInfo) {
    return null;
  }

  /*
   * Find exam table.
   */
  const examHeader =
    _findNursingExamHeader(
      finalData,
      config
    );

  if (!examHeader) {
    console.warn(
      "No Nursing exam header found on sheet: " +
      sheetName
    );

    return null;
  }

  const headerRowIndex =
    examHeader.rowIndex;

  const colMap =
    examHeader.colMap;

  /*
   * Find roster table below exam table.
   */
  const rosterHeaderInfo =
    _findNursingRosterHeader(
      finalData,
      headerRowIndex + 1,
      config
    );

  const rosterHeaderIdx =
    rosterHeaderInfo
      ? rosterHeaderInfo.rowIndex
      : -1;

  /*
   * Prevent exam parser from reading into roster data.
   */
  const examScanEnd =
    rosterHeaderIdx !== -1
      ? rosterHeaderIdx
      : finalData.length;

  let sheetRoster = {};

  if (rosterHeaderInfo) {
    sheetRoster =
      _parseNursingRoster(
        finalData,
        finalFontColors,
        rosterHeaderInfo,
        config
      );
  }

  const rawExamRows =
    _readNursingExamRows(
      finalData,
      headerRowIndex + 1,
      examScanEnd,
      colMap,
      config
    );

  const exams = [];

  rawExamRows.forEach(
    parsedRow => {
      const rIndex =
        parsedRow.rowIndex;

      const rowData =
        parsedRow.data;

      const examName =
        parsedRow.name;

      const rawDate =
        colMap.date > -1
          ? rowData[
              colMap.date
            ]
          : "";

      const isDone =
        _isNursingExamDone(
          rawDate,
          finalFontLines,
          rIndex,
          colMap
        );

      const dateVal =
        _formatNursingDate(
          rawDate
        );

      const password =
        colMap.password > -1
          ? _nursingDisplayValue(
              rowData[
                colMap.password
              ]
            )
          : "-";

      const startTime =
        colMap.startTime > -1
          ? _formatNursingStartTime(
              rowData[
                colMap.startTime
              ]
            )
          : "-";

      const duration =
        colMap.duration > -1
          ? _formatNursingDuration(
              rowData[
                colMap.duration
              ]
            )
          : "-";

      const dbKey =
        courseCode +
        "|" +
        examName;

      const dbEntry =
        dbMap &&
        dbMap[dbKey]
          ? dbMap[dbKey]
          : {};

      exams.push({
        name:
          examName,

        date:
          dateVal,

        password:
          password,

        startTime:
          startTime,

        duration:
          duration,

        generalNotes:
          dbEntry.generalNotes ||
          "",

        studentTags:
          dbEntry.studentTags ||
          {},

        docUrl:
          dbEntry.docUrl ||
          "",

        rosters:
          sheetRoster,

        isDone:
          isDone
      });
    }
  );

  return {
    sheetName:
      sheetName,

    rawA1:
      courseInfo.rawA1 ||
      "",

    rawTitle:
      courseInfo.rawTitle ||
      "",

    /*
     * Retained for compatibility.
     * The true legacy folder key is reconstructed by _parseCellA1()
     * before document operations.
     */
    folderKey:
      courseInfo.folderKey ||
      courseCode,

    course: {
      code:
        courseCode,

      name:
        courseInfo.courseTitle ||
        courseCode,

      faculty:
        courseInfo.faculty ||
        ""
    },

    exams:
      exams
  };
}


/* ========================================================================
 * EXAM HEADER DISCOVERY
 * ====================================================================== */


/**
 * Returns acceptable Nursing exam-header synonyms.
 */
function _getNursingExamSynonyms(
  config
) {
  return {
    exam:
      _uniqueNursingSynonyms([
        config.examHeader,
        "exam",
        "test",
        "assessment",
        "exam name",
        "test name",
        "assessment name",
        "exam/test",
        "test/exam"
      ]),

    date:
      _uniqueNursingSynonyms([
        config.dateHeader,
        "date",
        "exam date",
        "test date",
        "assessment date"
      ]),

    password:
      _uniqueNursingSynonyms([
        config.passwordHeader,
        "password",
        "passcode",
        "exam password",
        "test password"
      ]),

    startTime:
      _uniqueNursingSynonyms([
        "start time",
        "exam start time",
        "test start time",
        "exam time",
        "test time",
        "time"
      ]),

    duration:
      _uniqueNursingSynonyms([
        "duration",
        "duration mins",
        "duration/mins",
        "duration minutes",
        "minutes",
        "mins",
        "length",
        "test length",
        "exam length"
      ])
  };
}


/**
 * Finds the most likely Nursing exam-header row.
 */
function _findNursingExamHeader(
  data,
  config
) {
  if (
    !data ||
    !data.length
  ) {
    return null;
  }

  const synonyms =
    _getNursingExamSynonyms(
      config
    );

  const scanLimit =
    Math.min(
      data.length,
      40
    );

  const preferredExamHeader =
    SheetReader.normalizeHeader(
      config.examHeader ||
      "exam"
    );

  const preferredDateHeader =
    SheetReader.normalizeHeader(
      config.dateHeader ||
      "date"
    );

  let best = null;

  for (
    let r = 0;
    r < scanLimit;
    r++
  ) {
    const row =
      data[r] || [];

    const mapped =
      SheetReader.mapColumnsBySynonymsScored(
        row,
        synonyms
      );

    const colMap =
      mapped.columns;

    const scores =
      mapped.scores;

    if (
      colMap.exam === -1 ||
      colMap.date === -1
    ) {
      continue;
    }

    if (
      colMap.exam ===
      colMap.date
    ) {
      continue;
    }

    const examHeaderText =
      SheetReader.normalizeHeader(
        row[colMap.exam]
      );

    const dateHeaderText =
      SheetReader.normalizeHeader(
        row[colMap.date]
      );

    /*
     * Prevent generic "test" from winning against a more appropriate
     * non-exam header such as "Test Time".
     */
    if (
      examHeaderText !==
        preferredExamHeader &&
      /\b(date|time|duration|minute|minutes|mins|password|passcode|room)\b/i
        .test(
          examHeaderText
        )
    ) {
      continue;
    }

    if (
      dateHeaderText !==
        preferredDateHeader &&
      !/\bdate\b/i.test(
        dateHeaderText
      )
    ) {
      continue;
    }

    const matchedKeys =
      Object.keys(colMap)
        .filter(
          key =>
            colMap[key] !== -1
        );

    let totalScore = 0;

    matchedKeys.forEach(
      key => {
        totalScore +=
          scores[key] || 0;
      }
    );

    totalScore += 100;

    totalScore +=
      matchedKeys.length *
      15;

    if (
      !best ||
      totalScore >
        best.score
    ) {
      best = {
        rowIndex:
          r,

        colMap:
          colMap,

        scores:
          scores,

        score:
          totalScore
      };
    }
  }

  return best;
}


/**
 * Reads exam rows after the header.
 */
function _readNursingExamRows(
  data,
  startRow,
  endRow,
  colMap,
  config
) {
  const exams = [];

  const stopWords =
    _splitNursingKeywords(
      config.stopKeywords
    );

  let consecutiveEmptyRows =
    0;

  for (
    let r = startRow;
    r < endRow;
    r++
  ) {
    const row =
      data[r] || [];

    const rawExamName =
      colMap.exam > -1
        ? row[colMap.exam]
        : "";

    const examName =
      String(
        rawExamName ===
          null ||
        rawExamName ===
          undefined
          ? ""
          : rawExamName
      ).trim();

    if (!examName) {
      consecutiveEmptyRows++;

      if (
        consecutiveEmptyRows >= 5
      ) {
        break;
      }

      continue;
    }

    consecutiveEmptyRows = 0;

    const normalizedExamName =
      SheetReader.sanitize(
        examName
      );

    if (
      stopWords.some(
        word =>
          normalizedExamName ===
            word ||
          normalizedExamName.includes(
            word
          )
      )
    ) {
      break;
    }

    if (
      _looksLikeRepeatedNursingExamHeader(
        row,
        colMap,
        config
      )
    ) {
      continue;
    }

    const formattedRow =
      row.map(
        cell =>
          cell === null ||
          cell === undefined ||
          String(cell).trim() === ""
            ? ""
            : cell
      );

    exams.push({
      rowIndex:
        r,

      name:
        examName,

      data:
        formattedRow
    });
  }

  return exams;
}


/**
 * Identifies a repeated header row inside the exam table.
 */
function _looksLikeRepeatedNursingExamHeader(
  row,
  colMap,
  config
) {
  if (
    !row ||
    !row.length
  ) {
    return false;
  }

  const examCell =
    colMap.exam > -1
      ? SheetReader.normalizeHeader(
          row[colMap.exam]
        )
      : "";

  const dateCell =
    colMap.date > -1
      ? SheetReader.normalizeHeader(
          row[colMap.date]
        )
      : "";

  const synonyms =
    _getNursingExamSynonyms(
      config
    );

  const examSynonyms =
    synonyms.exam.map(
      value =>
        SheetReader.normalizeHeader(
          value
        )
    );

  const dateSynonyms =
    synonyms.date.map(
      value =>
        SheetReader.normalizeHeader(
          value
        )
    );

  return (
    examSynonyms.includes(
      examCell
    ) &&
    dateSynonyms.includes(
      dateCell
    )
  );
}


/* ========================================================================
 * ROSTER HEADER DISCOVERY
 * ====================================================================== */


/**
 * Finds the most likely location-header row.
 */
function _findNursingRosterHeader(
  data,
  startRow,
  config
) {
  if (
    !data ||
    !data.length ||
    startRow >= data.length
  ) {
    return null;
  }

  let best = null;

  for (
    let r = startRow;
    r < data.length;
    r++
  ) {
    const row =
      data[r] || [];

    const rowResult =
      _scoreNursingRosterHeaderRow(
        row,
        config
      );

    if (
      rowResult
        .candidateColumns
        .length === 0
    ) {
      continue;
    }

    let score =
      rowResult.score;

    let supportingJunkCount =
      0;

    const lookAheadEnd =
      Math.min(
        data.length,
        r + 5
      );

    for (
      let rr = r + 1;
      rr < lookAheadEnd;
      rr++
    ) {
      rowResult
        .candidateColumns
        .forEach(
          locationColumn => {
            const value =
              data[rr] &&
              data[rr].length >
                locationColumn.colIndex
                ? data[rr][
                    locationColumn.colIndex
                  ]
                : "";

            if (
              _nursingHasValue(
                value
              ) &&
              _isNursingRosterJunkValue(
                value
              )
            ) {
              supportingJunkCount++;
            }
          }
        );
    }

    score +=
      Math.min(
        supportingJunkCount,
        6
      );

    const structurallyPlausible =
      rowResult.strongCount > 0 ||
      rowResult
        .candidateColumns
        .length >= 3;

    if (!structurallyPlausible) {
      continue;
    }

    if (
      !best ||
      score > best.score
    ) {
      best = {
        rowIndex:
          r,

        score:
          score,

        locationColumns:
          rowResult.candidateColumns
      };
    }
  }

  return best;
}


/**
 * Scores potential roster-header rows.
 */
function _scoreNursingRosterHeaderRow(
  row,
  config
) {
  let score = 0;
  let strongCount = 0;

  const candidateColumns = [];

  for (
    let c = 0;
    c < row.length;
    c++
  ) {
    const cellResult =
      _scoreNursingLocationHeaderCell(
        row[c],
        config
      );

    if (
      cellResult.score <= 0
    ) {
      continue;
    }

    score +=
      cellResult.score;

    if (
      cellResult.strong
    ) {
      strongCount++;
    }

    candidateColumns.push({
      colIndex:
        c,

      name:
        String(row[c]).trim(),

      score:
        cellResult.score,

      strong:
        cellResult.strong
    });
  }

  if (
    candidateColumns.length >= 2
  ) {
    score +=
      candidateColumns.length *
      2;
  }

  return {
    score:
      score,

    strongCount:
      strongCount,

    candidateColumns:
      candidateColumns
  };
}


/**
 * Scores one potential Nursing location heading.
 */
function _scoreNursingLocationHeaderCell(
  value,
  config
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return {
      score: 0,
      strong: false
    };
  }

  if (
    value instanceof Date ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return {
      score: 0,
      strong: false
    };
  }

  const raw =
    String(value).trim();

  const normalized =
    SheetReader.sanitize(
      raw
    );

  if (!normalized) {
    return {
      score: 0,
      strong: false
    };
  }

  if (
    [
      "location",
      "locations",
      "site",
      "sites",
      "student",
      "students",
      "name",
      "names"
    ].includes(
      normalized
    )
  ) {
    return {
      score: 0,
      strong: false
    };
  }

  if (
    /^(exam|test|assessment|date|password|passcode|start time|duration|duration mins|room)$/i
      .test(
        normalized
      )
  ) {
    return {
      score: 0,
      strong: false
    };
  }

  if (
    _isNursingRosterJunkValue(
      raw,
      true
    )
  ) {
    return {
      score: 0,
      strong: false
    };
  }

  const knownLocations =
    _getKnownNursingLocations();

  const rosterHint =
    SheetReader.sanitize(
      config.rosterKeyword ||
      ""
    );

  let score = 0;
  let strong = false;

  if (
    rosterHint &&
    (
      normalized ===
        rosterHint ||
      normalized.includes(
        rosterHint
      )
    )
  ) {
    score += 12;
    strong = true;
  }

  if (
    knownLocations.some(
      location =>
        normalized ===
          location ||
        normalized.includes(
          location
        )
    )
  ) {
    score += 10;
    strong = true;
  }

  if (
    /\btesting\s+center\b/i.test(
      raw
    )
  ) {
    score += 7;
    strong = true;
  }

  if (
    /^[A-Z][A-Z0-9]{2,9}$/.test(
      raw
    )
  ) {
    score += 5;
    strong = true;
  }

  const words =
    raw.split(/\s+/)
      .filter(Boolean);

  if (
    raw.length <= 45 &&
    words.length <= 5 &&
    /^[A-Za-z0-9&.'’()\/\-\s]+$/
      .test(raw)
  ) {
    score += 2;
  }

  return {
    score:
      score,

    strong:
      strong
  };
}


/* ========================================================================
 * ROSTER PARSING
 * ====================================================================== */


/**
 * Parses location rosters without any fixed offset.
 */
function _parseNursingRoster(
  data,
  fontColors,
  rosterHeaderInfo,
  config
) {
  const roster = {};

  const locationColumns =
    rosterHeaderInfo
      .locationColumns ||
    [];

  locationColumns.forEach(
    location => {
      if (
        location.name &&
        !roster[
          location.name
        ]
      ) {
        roster[
          location.name
        ] = [];
      }
    }
  );

  const rosterStopWords =
    _splitNursingKeywords(
      config.rosterStopKeywords
    );

  let rosterStarted =
    false;

  let consecutiveInactiveRows =
    0;

  const MAX_INACTIVE_ROWS =
    4;

  for (
    let r =
      rosterHeaderInfo.rowIndex + 1;
    r < data.length;
    r++
  ) {
    const row =
      data[r] || [];

    if (
      rosterStopWords.length > 0 &&
      _nursingRosterRowHasStopMarker(
        row,
        locationColumns,
        rosterStopWords
      )
    ) {
      break;
    }

    let rowStudentCount =
      0;

    locationColumns.forEach(
      location => {
        const c =
          location.colIndex;

        const rawValue =
          row.length > c
            ? row[c]
            : "";

        if (
          !_looksLikeNursingStudentName(
            rawValue
          )
        ) {
          return;
        }

        const studentName =
          String(
            rawValue
          ).trim();

        const color =
          fontColors[r] &&
          fontColors[r].length > c
            ? fontColors[r][c]
            : "#000000";

        if (
          !roster[
            location.name
          ]
        ) {
          roster[
            location.name
          ] = [];
        }

        roster[
          location.name
        ].push({
          name:
            studentName,

          color:
            color ||
            "#000000"
        });

        rowStudentCount++;
      }
    );

    if (
      rowStudentCount > 0
    ) {
      rosterStarted = true;
      consecutiveInactiveRows = 0;
      continue;
    }

    /*
     * Do not terminate merely because the rows before the first student
     * contain addresses, notes, or test-time information.
     */
    if (!rosterStarted) {
      continue;
    }

    consecutiveInactiveRows++;

    if (
      consecutiveInactiveRows >=
      MAX_INACTIVE_ROWS
    ) {
      break;
    }
  }

  return roster;
}


/**
 * Tests configured roster stop words across every location column.
 */
function _nursingRosterRowHasStopMarker(
  row,
  locationColumns,
  stopWords
) {
  for (
    let i = 0;
    i < locationColumns.length;
    i++
  ) {
    const c =
      locationColumns[i]
        .colIndex;

    const value =
      row.length > c
        ? row[c]
        : "";

    const normalized =
      SheetReader.sanitize(
        value
      );

    if (!normalized) {
      continue;
    }

    if (
      stopWords.some(
        word =>
          normalized ===
            word ||
          normalized.includes(
            word
          )
      )
    ) {
      return true;
    }
  }

  return false;
}


/**
 * Attempts to recognize a student name.
 */
function _looksLikeNursingStudentName(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return false;
  }

  if (
    value instanceof Date ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return false;
  }

  const raw =
    String(value).trim();

  if (!raw) {
    return false;
  }

  if (
    _isNursingRosterJunkValue(
      raw
    )
  ) {
    return false;
  }

  if (
    !/[A-Za-z]/.test(raw) ||
    /\d/.test(raw)
  ) {
    return false;
  }

  if (
    /[@;]/.test(raw) ||
    /https?:\/\//i.test(raw) ||
    /\bwww\./i.test(raw)
  ) {
    return false;
  }

  const words =
    raw.split(/\s+/)
      .filter(Boolean);

  if (
    words.length === 0 ||
    words.length > 6
  ) {
    return false;
  }

  if (
    raw.length > 70
  ) {
    return false;
  }

  const normalized =
    SheetReader.sanitize(
      raw
    );

  const facilityWords = [
    "testing center",
    "student center",
    "campus",
    "building",
    "library",
    "college",
    "university",
    "room ",
    "zoom",
    "remote",
    "online",
    "proctor",
    "proctoring"
  ];

  if (
    facilityWords.some(
      word =>
        normalized.includes(
          word
        )
    )
  ) {
    return false;
  }

  const instructionStarts = [
    "please ",
    "student ",
    "students ",
    "call ",
    "contact ",
    "note ",
    "notes ",
    "instruction ",
    "instructions ",
    "must ",
    "use ",
    "bring ",
    "email "
  ];

  if (
    instructionStarts.some(
      phrase =>
        normalized.startsWith(
          phrase
        )
    )
  ) {
    return false;
  }

  if (
    _getKnownNursingLocations()
      .some(
        location =>
          normalized ===
          location
      )
  ) {
    return false;
  }

  if (
    !/^[A-Za-zÀ-ÿ.'’,\-\s]+$/
      .test(raw)
  ) {
    return false;
  }

  return true;
}


/**
 * Identifies obvious roster junk.
 */
function _isNursingRosterJunkValue(
  value,
  allowLocationHeader
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return true;
  }

  if (
    value instanceof Date ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return true;
  }

  const raw =
    String(value).trim();

  const normalized =
    SheetReader.sanitize(
      raw
    );

  if (!normalized) {
    return true;
  }

  const exactJunk = [
    "confirmed",
    "confirmation",
    "tbd",
    "n/a",
    "na",
    "n.a.",
    "none",
    "not applicable",
    "yes",
    "no"
  ];

  if (
    exactJunk.includes(
      normalized
    )
  ) {
    return true;
  }

  /*
   * Time/test-time values.
   */
  if (
    /^\d{3,4}\s*(?:test|exam)?\s*time$/i
      .test(raw) ||
    /^\d{3,4}$/i.test(raw) ||
    /^\d{1,2}:\d{2}\s*(?:am|pm)?$/i
      .test(raw) ||
    /^\d{1,2}:\d{2}\s*(?:am|pm)?\s*(?:test|exam)?\s*time$/i
      .test(raw)
  ) {
    return true;
  }

  if (
    /\b(?:test|exam)\s*time\b/i
      .test(raw)
  ) {
    return true;
  }

  if (
    /\bconfirmed\b/i.test(
      raw
    )
  ) {
    return true;
  }

  /*
   * URLs/email.
   */
  if (
    /@/.test(raw) ||
    /https?:\/\//i.test(raw) ||
    /\bwww\./i.test(raw)
  ) {
    return true;
  }

  /*
   * Phone-like content.
   */
  if (
    /^\+?[\d\s().-]{7,}$/
      .test(raw)
  ) {
    return true;
  }

  /*
   * ZIP.
   */
  if (
    /\b\d{5}(?:-\d{4})?\b/
      .test(raw)
  ) {
    return true;
  }

  /*
   * Address.
   */
  if (
    /^\d+\s+.+\b(?:street|st|road|rd|avenue|ave|drive|dr|lane|ln|boulevard|blvd|route|rt|highway|hwy|way|court|ct)\b/i
      .test(raw)
  ) {
    return true;
  }

  /*
   * Long text.
   */
  const words =
    raw.split(/\s+/)
      .filter(Boolean);

  if (
    words.length > 8 ||
    raw.length > 90
  ) {
    return true;
  }

  if (
    !allowLocationHeader
  ) {
    if (
      /\b(?:instructions?|notes?|contact|please|must|proctoring instructions?)\b/i
        .test(raw)
    ) {
      return true;
    }
  }

  return false;
}


/**
 * Known Nursing locations are scoring signals, not hard requirements.
 */
function _getKnownNursingLocations() {
  return [
    "augusta",
    "umaal",
    "umf testing center",
    "bangor",
    "ellsworth",
    "lewiston",
    "rockland",
    "saco"
  ];
}


/* ========================================================================
 * EXAM STATUS / FORMATTING
 * ====================================================================== */


/**
 * Preserves the existing Nursing done logic:
 *
 * - strike-through exam/date
 * - OR exam date before today
 */
function _isNursingExamDone(
  rawDate,
  fontLines,
  rowIndex,
  colMap
) {
  let isDone = false;

  const rowFontLines =
    fontLines[rowIndex] ||
    [];

  if (
    (
      colMap.date > -1 &&
      rowFontLines[
        colMap.date
      ] === "line-through"
    ) ||
    (
      colMap.exam > -1 &&
      rowFontLines[
        colMap.exam
      ] === "line-through"
    )
  ) {
    isDone = true;
  }

  const dateObj =
    _parseNursingDate(
      rawDate
    );

  if (
    dateObj &&
    !isNaN(
      dateObj.getTime()
    )
  ) {
    const today =
      new Date();

    today.setHours(
      0,
      0,
      0,
      0
    );

    const comparisonDate =
      new Date(
        dateObj.getFullYear(),
        dateObj.getMonth(),
        dateObj.getDate()
      );

    if (
      comparisonDate <
      today
    ) {
      isDone = true;
    }
  }

  return isDone;
}


/**
 * Converts a Nursing date cell to a Date object for comparison.
 */
function _parseNursingDate(
  rawDate
) {
  if (!rawDate) {
    return null;
  }

  if (
    rawDate instanceof Date
  ) {
    return rawDate;
  }

  const text =
    String(rawDate).trim();

  if (
    !text ||
    /^(tbd|n\/a|na)$/i.test(
      text
    )
  ) {
    return null;
  }

  const cleaned =
    text.replace(
      /(\d)(st|nd|rd|th)\b/gi,
      "$1"
    );

  const parsed =
    new Date(cleaned);

  return isNaN(
    parsed.getTime()
  )
    ? null
    : parsed;
}


/**
 * Formats the exam date for the SPA.
 */
function _formatNursingDate(
  rawDate
) {
  if (
    rawDate === null ||
    rawDate === undefined ||
    rawDate === ""
  ) {
    return "-";
  }

  try {
    if (
      typeof formatDateToPlainLanguage ===
      "function"
    ) {
      const result =
        formatDateToPlainLanguage(
          rawDate
        );

      if (
        result !== null &&
        result !== undefined &&
        String(result).trim() !== ""
      ) {
        return result;
      }
    }
  } catch (e) {
    console.warn(
      "Could not format Nursing date:",
      e
    );
  }

  if (
    rawDate instanceof Date
  ) {
    return Utilities.formatDate(
      rawDate,
      Session.getScriptTimeZone(),
      "MMMM d, yyyy"
    );
  }

  return (
    String(rawDate).trim() ||
    "-"
  );
}


/**
 * Formats the single exam-table Start Time.
 *
 * Roster-level times are intentionally ignored.
 */
function _formatNursingStartTime(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "-";
  }

  try {
    if (
      typeof normalizeTime ===
      "function"
    ) {
      const result =
        normalizeTime(
          value
        );

      if (
        result !== null &&
        result !== undefined &&
        String(result).trim() !== ""
      ) {
        return result;
      }
    }
  } catch (e) {
    console.warn(
      "Could not normalize Nursing start time:",
      e
    );
  }

  return (
    String(value).trim() ||
    "-"
  );
}


/**
 * Duration is deliberately preserved as supplied by the spreadsheet.
 */
function _formatNursingDuration(
  value
) {
  return _nursingDisplayValue(
    value
  );
}


function _nursingDisplayValue(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "-";
  }

  const text =
    String(value).trim();

  return text || "-";
}


function _nursingRowHasContent(
  row
) {
  if (
    !row ||
    !row.length
  ) {
    return false;
  }

  return row.some(
    value =>
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ""
  );
}


function _nursingHasValue(
  value
) {
  return (
    value !== null &&
    value !== undefined &&
    String(value).trim() !== ""
  );
}


function _splitNursingKeywords(
  input
) {
  if (!input) {
    return [];
  }

  return String(input)
    .split(",")
    .map(
      value =>
        SheetReader.sanitize(
          value
        )
    )
    .filter(Boolean);
}


function _uniqueNursingSynonyms(
  values
) {
  const seen = {};

  return values.filter(
    value => {
      const normalized =
        SheetReader.normalizeHeader(
          value
        );

      if (
        !normalized ||
        seen[normalized]
      ) {
        return false;
      }

      seen[normalized] = true;

      return true;
    }
  );
}


/* ========================================================================
 * NURSING DOCUMENT CREATION
 * ====================================================================== */


/**
 * Creates missing Nursing proctoring documents.
 *
 * Existing folder/file naming strategy is intentionally preserved.
 */
function createNursingProctoringDocuments(
  payload
) {
  try {
    const config =
      _getNursingSettings();

    const rootFolder =
      DriveApp.getFolderById(
        config.folderId
      );

    let createdCount = 0;

    const createdUrls = {};

    payload.sheets.forEach(
      sheetData => {
        const meta =
          _getNursingDocumentMeta(
            sheetData
          );

        let targetFolder =
          _findSubFolder(
            rootFolder,
            meta.folderKey
          );

        if (!targetFolder) {
          targetFolder =
            rootFolder.createFolder(
              meta.folderKey
            );
        }

        sheetData.exams.forEach(
          exam => {
            /*
             * Canonical historical filename:
             *
             * Course : Faculty - Exam
             */
            const docTitle =
              `${meta.fullTitlePrefix} - ${exam.name}`;

            let targetFile =
              null;

            /*
             * 1. Exact filename.
             */
            const filesByName =
              targetFolder.getFilesByName(
                docTitle
              );

            if (
              filesByName.hasNext()
            ) {
              targetFile =
                filesByName.next();
            }

            /*
             * 2. Fuzzy legacy search.
             */
            if (!targetFile) {
              const suffix =
                ` - ${exam.name}`;

              const files =
                targetFolder.getFiles();

              while (
                files.hasNext()
              ) {
                const f =
                  files.next();

                if (
                  f.getName()
                    .includes(
                      suffix
                    ) ||
                  f.getName()
                    .trim() ===
                    exam.name
                ) {
                  targetFile = f;

                  /*
                   * Preserve the existing automatic filename correction.
                   */
                  if (
                    targetFile.getName() !==
                    docTitle
                  ) {
                    targetFile.setName(
                      docTitle
                    );
                  }

                  break;
                }
              }
            }

            /*
             * Existing document found.
             */
            if (targetFile) {
              createdUrls[
                exam.name
              ] =
                targetFile.getUrl();

              return;
            }

            /*
             * Create new document.
             */
            const doc =
              DocumentApp.create(
                docTitle
              );

            _populateNursingDoc(
              doc,
              meta.courseTitle,
              meta.facultyName,
              exam,
              config.customNotes
            );

            const file =
              DriveApp.getFileById(
                doc.getId()
              );

            targetFolder.addFile(
              file
            );

            DriveApp
              .getRootFolder()
              .removeFile(file);

            createdUrls[
              exam.name
            ] =
              file.getUrl();

            if (
              typeof logSystemAction ===
              "function"
            ) {
              logSystemAction(
                "Nursing",
                "Created Doc",
                docTitle,
                doc.getId(),
                `Date: ${exam.date}`
              );
            }

            createdCount++;
          }
        );
      }
    );

    return {
      success:
        true,

      message:
        `Created ${createdCount} documents.`,

      createdUrls:
        createdUrls
    };

  } catch (e) {
    return {
      success:
        false,

      message:
        e &&
        e.message
          ? e.message
          : String(e)
    };
  }
}


/* ========================================================================
 * NURSING DOCUMENT UPDATES
 * ====================================================================== */


/**
 * Updates existing Nursing documents.
 *
 * Existing folder/file lookup behavior is preserved.
 */
function updateAllNursingDocuments(
  payload
) {
  try {
    const config =
      _getNursingSettings();

    const rootFolder =
      DriveApp.getFolderById(
        config.folderId
      );

    let updatedCount = 0;

    payload.sheets.forEach(
      sheetData => {
        const meta =
          _getNursingDocumentMeta(
            sheetData
          );

        const targetFolder =
          _findSubFolder(
            rootFolder,
            meta.folderKey
          );

        /*
         * No folder means there is nothing to update.
         */
        if (!targetFolder) {
          return;
        }

        sheetData.exams.forEach(
          exam => {
            const docTitle =
              `${meta.fullTitlePrefix} - ${exam.name}`;

            let targetFile =
              null;

            /*
             * First preference: URL supplied by API.
             */
            if (
              exam.docUrl
            ) {
              try {
                targetFile =
                  DriveApp.getFileByUrl(
                    exam.docUrl
                  );
              } catch (e) {
                targetFile =
                  null;
              }
            }

            /*
             * 1. Exact filename if URL unavailable.
             */
            if (!targetFile) {
              const filesByName =
                targetFolder.getFilesByName(
                  docTitle
                );

              if (
                filesByName.hasNext()
              ) {
                targetFile =
                  filesByName.next();
              }
            }

            /*
             * 2. Fuzzy legacy search.
             */
            if (!targetFile) {
              const suffix =
                ` - ${exam.name}`;

              const files =
                targetFolder.getFiles();

              while (
                files.hasNext()
              ) {
                const f =
                  files.next();

                if (
                  f.getName()
                    .includes(
                      suffix
                    ) ||
                  f.getName()
                    .trim() ===
                    exam.name
                ) {
                  targetFile = f;

                  if (
                    targetFile.getName() !==
                    docTitle
                  ) {
                    targetFile.setName(
                      docTitle
                    );
                  }

                  break;
                }
              }
            }

            if (!targetFile) {
              return;
            }

            const doc =
              DocumentApp.openById(
                targetFile.getId()
              );

            _populateNursingDoc(
              doc,
              meta.courseTitle,
              meta.facultyName,
              exam,
              config.customNotes
            );

            if (
              typeof logSystemAction ===
              "function"
            ) {
              logSystemAction(
                "Nursing",
                "Updated Doc",
                targetFile.getName(),
                targetFile.getId(),
                `Date: ${exam.date}`
              );
            }

            updatedCount++;
          }
        );
      }
    );

    return {
      success:
        true,

      message:
        `Updated ${updatedCount} documents.`
    };

  } catch (e) {
    return {
      success:
        false,

      message:
        e &&
        e.message
          ? e.message
          : String(e)
    };
  }
}


/* ========================================================================
 * NURSING DOCUMENT POPULATION
 * ====================================================================== */


/**
 * Populates/updates a Nursing proctoring document.
 *
 * The document now uses the same single Start Time model as the SPA.
 */
function _populateNursingDoc(
  doc,
  courseTitle,
  facultyName,
  exam,
  customNotes
) {
  const body =
    doc.getBody();

  /*
   * 1. Document title
   */
  const displayTitle =
    `${courseTitle} - ${exam.name}`;

  const displayFaculty =
    facultyName ||
    "Faculty Unassigned";

  const existingTitle =
    body.getChild(0);

  if (
    existingTitle.getType() ===
    DocumentApp.ElementType.PARAGRAPH
  ) {
    const titlePara =
      existingTitle.asParagraph();

    if (
      titlePara.getText() !==
      displayTitle
    ) {
      titlePara.setText(
        displayTitle
      );
    }

    titlePara.setHeading(
      DocumentApp.ParagraphHeading.TITLE
    );
  }

  /*
   * 2. Faculty heading
   */
  let facultyPara =
    null;

  if (
    body.getNumChildren() > 1
  ) {
    const child =
      body.getChild(1);

    if (
      child.getType() ===
      DocumentApp.ElementType.PARAGRAPH
    ) {
      facultyPara =
        child.asParagraph();
    }
  }

  if (!facultyPara) {
    facultyPara =
      body.insertParagraph(
        1,
        displayFaculty
      );
  } else {
    if (
      facultyPara.getText() !==
      displayFaculty
    ) {
      facultyPara.setText(
        displayFaculty
      );
    }
  }

  facultyPara.setHeading(
    DocumentApp.ParagraphHeading.HEADING1
  );

  /*
   * 3. Exam Details
   *
   * IMPORTANT:
   * There is now one Start Time rather than separate Site/Zoom times.
   */
  const examDetailsMap = {
    "Date":
      exam.date || "N/A",

    "Start Time":
      exam.startTime || "N/A",

    "Duration":
      exam.duration || "-",

    "Password":
      exam.password || "-"
  };

  _updateKeyValueSection(
    body,
    "Exam Details",
    examDetailsMap,
    DocumentApp.ParagraphHeading.HEADING2
  );

  /*
   * 4. General Instructions
   */
  if (customNotes) {
    _updateOrCreateTextSection(
      body,
      "General Instructions",
      customNotes,
      true
    );
  }

  /*
   * 5. Important Links
   */
  _updateImportantLinks(
    body
  );

  /*
   * 6. Accommodations
   */
  if (exam.generalNotes) {
    _updateOrCreateHighlightedSection(
      body,
      "Exam Accommodations",
      exam.generalNotes
    );
  } else {
    _removeSectionIfExists(
      body,
      "Exam Accommodations"
    );
  }

  /*
   * 7. Location rosters
   */
  _updateLocationRosters(
    body,
    exam
  );

  doc.saveAndClose();
}


/* ========================================================================
 * DOCUMENT SECTION HELPERS
 * ====================================================================== */


/**
 * Finds a heading by normalized text.
 */
function _findHeadingIndex(
  body,
  headingText
) {
  const numChildren =
    body.getNumChildren();

  const targetNorm =
    headingText
      .toLowerCase()
      .replace(
        /[^a-z0-9]/g,
        ""
      );

  for (
    let i = 0;
    i < numChildren;
    i++
  ) {
    const child =
      body.getChild(i);

    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {
      continue;
    }

    const para =
      child.asParagraph();

    const text =
      para.getText();

    if (
      text &&
      text
        .toLowerCase()
        .replace(
          /[^a-z0-9]/g,
          ""
        )
        .includes(
          targetNorm
        )
    ) {
      const h =
        para.getHeading();

      if (
        h !==
          DocumentApp.ParagraphHeading.NORMAL &&
        h !==
          DocumentApp.ParagraphHeading.TITLE
      ) {
        return i;
      }
    }
  }

  return -1;
}


/**
 * Maintains the standard Nursing links.
 */
function _updateImportantLinks(
  body
) {
  const sectionTitle =
    "Important Links";

  const links = [
    {
      text:
        "Red Flag Reporting Form",

      url:
        "https://docs.google.com/forms/d/e/1FAIpQLSfORKCKol8SsRldNKfvsDy3ILNs9HcFv3gKb8TuxrNrlqxijw/viewform"
    },

    {
      text:
        "Nursing Protocol",

      url:
        "https://docs.google.com/document/d/1TgKtmoDFqXLK0lBFPNirOAz_TW4S3E_BFhS934VcjOo/edit"
    }
  ];

  let sectionIndex =
    _findHeadingIndex(
      body,
      sectionTitle
    );

  if (
    sectionIndex === -1
  ) {
    let insertAt =
      _findHeadingIndex(
        body,
        "Exam Accommodations"
      );

    if (
      insertAt === -1
    ) {
      insertAt =
        _findHeadingIndex(
          body,
          "Location Rosters"
        );
    }

    if (
      insertAt === -1
    ) {
      insertAt =
        body.getNumChildren();
    }

    body
      .insertParagraph(
        insertAt,
        sectionTitle
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    sectionIndex =
      insertAt;

  } else {
    body
      .getChild(
        sectionIndex
      )
      .asParagraph()
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    /*
     * Remove old normal content under this heading.
     */
    let scanIndex =
      sectionIndex + 1;

    while (
      scanIndex <
      body.getNumChildren()
    ) {
      const child =
        body.getChild(
          scanIndex
        );

      if (
        child.getType() ===
        DocumentApp.ElementType.PARAGRAPH
      ) {
        const h =
          child
            .asParagraph()
            .getHeading();

        if (
          h !==
          DocumentApp.ParagraphHeading.NORMAL
        ) {
          break;
        }
      }

      body.removeChild(
        child
      );
    }
  }

  /*
   * Write fresh links.
   */
  links.forEach(
    (link, i) => {
      const li =
        body.insertListItem(
          sectionIndex +
            1 +
            i,
          link.text
        );

      li.setLinkUrl(
        link.url
      );

      li.setGlyphType(
        DocumentApp.GlyphType.BULLET
      );
    }
  );
}


/**
 * Updates a key/value section.
 */
function _updateKeyValueSection(
  body,
  sectionTitle,
  dataMap,
  headingLevel
) {
  let sectionIndex =
    _findHeadingIndex(
      body,
      sectionTitle
    );

  if (
    sectionIndex === -1
  ) {
    const rostersIndex =
      _findHeadingIndex(
        body,
        "Location Rosters"
      );

    const insertAt =
      rostersIndex > -1
        ? rostersIndex
        : body.getNumChildren();

    body
      .insertParagraph(
        insertAt,
        sectionTitle
      )
      .setHeading(
        headingLevel
      );

    sectionIndex =
      insertAt;

  } else {
    body
      .getChild(
        sectionIndex
      )
      .asParagraph()
      .setHeading(
        headingLevel
      );

    let scanIndex =
      sectionIndex + 1;

    while (
      scanIndex <
      body.getNumChildren()
    ) {
      const child =
        body.getChild(
          scanIndex
        );

      if (
        child.getType() ===
        DocumentApp.ElementType.PARAGRAPH
      ) {
        const h =
          child
            .asParagraph()
            .getHeading();

        if (
          h !==
          DocumentApp.ParagraphHeading.NORMAL
        ) {
          break;
        }
      }

      body.removeChild(
        child
      );
    }
  }

  let insertCount = 1;

  for (
    const [key, value]
    of Object.entries(
      dataMap
    )
  ) {
    const text =
      `${key}: ${value}`;

    const li =
      body.insertListItem(
        sectionIndex +
          insertCount,
        text
      );

    li.setGlyphType(
      DocumentApp.GlyphType.BULLET
    );

    const keyCheck =
      key.toLowerCase();

    if (
      keyCheck.includes("date") ||
      keyCheck.includes("time") ||
      keyCheck.includes("password")
    ) {
      li.setBackgroundColor(
        "#ffff00"
      );
    }

    insertCount++;
  }
}


/**
 * Updates or creates a plain text section.
 */
function _updateOrCreateTextSection(
  body,
  sectionTitle,
  content,
  isItalic
) {
  let sectionIndex =
    _findHeadingIndex(
      body,
      sectionTitle
    );

  if (
    sectionIndex === -1
  ) {
    const rostersIndex =
      _findHeadingIndex(
        body,
        "Location Rosters"
      );

    const insertAt =
      rostersIndex > -1
        ? rostersIndex
        : body.getNumChildren();

    body
      .insertParagraph(
        insertAt,
        sectionTitle
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    const para =
      body.insertParagraph(
        insertAt + 1,
        content
      );

    if (isItalic) {
      para.setItalic(true);
    }

  } else {
    body
      .getChild(
        sectionIndex
      )
      .asParagraph()
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    let scanIndex =
      sectionIndex + 1;

    if (
      scanIndex <
      body.getNumChildren()
    ) {
      const nextChild =
        body.getChild(
          scanIndex
        );

      if (
        nextChild.getType() ===
        DocumentApp.ElementType.PARAGRAPH &&
        nextChild
          .asParagraph()
          .getHeading() ===
          DocumentApp.ParagraphHeading.NORMAL
      ) {
        nextChild
          .asParagraph()
          .setText(
            content
          );

        if (isItalic) {
          nextChild
            .asParagraph()
            .setItalic(true);
        }

        scanIndex++;

      } else {
        const para =
          body.insertParagraph(
            scanIndex,
            content
          );

        if (isItalic) {
          para.setItalic(true);
        }

        scanIndex++;
      }

    } else {
      const para =
        body.insertParagraph(
          scanIndex,
          content
        );

      if (isItalic) {
        para.setItalic(true);
      }

      scanIndex++;
    }

    while (
      scanIndex <
      body.getNumChildren()
    ) {
      const child =
        body.getChild(
          scanIndex
        );

      if (
        child.getType() ===
        DocumentApp.ElementType.PARAGRAPH
      ) {
        if (
          child
            .asParagraph()
            .getHeading() !==
          DocumentApp.ParagraphHeading.NORMAL
        ) {
          break;
        }
      }

      body.removeChild(
        child
      );
    }
  }
}


/**
 * Updates or creates highlighted text section.
 */
function _updateOrCreateHighlightedSection(
  body,
  sectionTitle,
  content
) {
  let sectionIndex =
    _findHeadingIndex(
      body,
      sectionTitle
    );

  if (
    sectionIndex === -1
  ) {
    const rostersIndex =
      _findHeadingIndex(
        body,
        "Location Rosters"
      );

    const insertAt =
      rostersIndex > -1
        ? rostersIndex
        : body.getNumChildren();

    body
      .insertParagraph(
        insertAt,
        sectionTitle
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    const p =
      body.insertParagraph(
        insertAt + 1,
        content
      );

    p.setBackgroundColor(
      "#e8f5e9"
    );

    p
      .setPaddingTop(5)
      .setPaddingBottom(5)
      .setPaddingLeft(10)
      .setPaddingRight(10);

  } else {
    body
      .getChild(
        sectionIndex
      )
      .asParagraph()
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    const nextChild =
      body.getChild(
        sectionIndex + 1
      );

    if (
      nextChild &&
      nextChild.getType() ===
      DocumentApp.ElementType.PARAGRAPH
    ) {
      const para =
        nextChild.asParagraph();

      if (
        para.getText() !==
        content
      ) {
        para.setText(
          content
        );

        para.setBackgroundColor(
          "#e8f5e9"
        );

        para
          .setPaddingTop(5)
          .setPaddingBottom(5)
          .setPaddingLeft(10)
          .setPaddingRight(10);
      }
    }
  }
}


/**
 * Removes a Nursing document section.
 */
function _removeSectionIfExists(
  body,
  sectionTitle
) {
  const sectionIndex =
    _findHeadingIndex(
      body,
      sectionTitle
    );

  if (
    sectionIndex > -1
  ) {
    body.removeChild(
      body.getChild(
        sectionIndex
      )
    );

    if (
      sectionIndex <
      body.getNumChildren()
    ) {
      const nextChild =
        body.getChild(
          sectionIndex
        );

      if (
        nextChild.getType() ===
        DocumentApp.ElementType.PARAGRAPH
      ) {
        if (
          nextChild
            .asParagraph()
            .getHeading() ===
          DocumentApp.ParagraphHeading.NORMAL
        ) {
          body.removeChild(
            nextChild
          );
        }
      }
    }
  }
}


/* ========================================================================
 * LOCATION ROSTERS IN DOCUMENTS
 * ====================================================================== */


/**
 * Updates the document's location rosters.
 *
 * Student objects retain:
 *
 *   {
 *     name: "...",
 *     color: "#..."
 *   }
 */
function _updateLocationRosters(
  body,
  exam
) {
  if (
    !exam.rosters ||
    Object.keys(
      exam.rosters
    ).length === 0
  ) {
    _removeSectionIfExists(
      body,
      "Location Rosters"
    );

    return;
  }

  let rostersIndex =
    _findHeadingIndex(
      body,
      "Location Rosters"
    );

  if (
    rostersIndex === -1
  ) {
    rostersIndex =
      body.getNumChildren();

    body
      .insertParagraph(
        rostersIndex,
        "Location Rosters"
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

  } else {
    body
      .getChild(
        rostersIndex
      )
      .asParagraph()
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );
  }

  /*
   * Preserve existing sorting behavior:
   * UMAAL first, then alphabetical.
   */
  const sortedLocations =
    Object.keys(
      exam.rosters
    ).sort(
      (a, b) => {
        if (
          a === "UMAAL"
        ) {
          return -1;
        }

        if (
          b === "UMAAL"
        ) {
          return 1;
        }

        return a.localeCompare(
          b
        );
      }
    );

  let currentIndex =
    rostersIndex + 1;

  sortedLocations.forEach(
    location => {
      const students =
        exam.rosters[
          location
        ] || [];

      /*
       * Preserve existing accommodation exclusion behavior.
       */
      const activeStudents =
        students.filter(
          s => {
            const tagData =
              (
                exam.studentTags &&
                exam.studentTags[
                  s.name
                ]
              )
                ? exam.studentTags[
                    s.name
                  ]
                : null;

            return !(
              tagData &&
              typeof tagData ===
                "object" &&
              tagData.excluded
            );
          }
        );

      let locationIndex =
        -1;

      const locNorm =
        location
          .toLowerCase()
          .replace(
            /[^a-z0-9]/g,
            ""
          );

      for (
        let i = currentIndex;
        i < body.getNumChildren();
        i++
      ) {
        const child =
          body.getChild(i);

        if (
          child.getType() ===
          DocumentApp.ElementType.PARAGRAPH
        ) {
          const para =
            child.asParagraph();

          const h =
            para.getHeading();

          const t =
            para.getText()
              .toLowerCase()
              .replace(
                /[^a-z0-9]/g,
                ""
              );

          if (
            t.includes(
              locNorm
            ) &&
            (
              h ===
                DocumentApp.ParagraphHeading.HEADING3 ||
              h ===
                DocumentApp.ParagraphHeading.HEADING2
            )
          ) {
            locationIndex =
              i;

            break;
          }

          if (
            h ===
              DocumentApp.ParagraphHeading.HEADING2 ||
            h ===
              DocumentApp.ParagraphHeading.HEADING1
          ) {
            break;
          }
        }
      }

      const displayLocation =
        _toTitleCase(
          location
        );

      if (
        locationIndex === -1
      ) {
        body
          .insertParagraph(
            currentIndex,
            displayLocation
          )
          .setHeading(
            DocumentApp.ParagraphHeading.HEADING3
          );

        locationIndex =
          currentIndex;

        currentIndex++;

      } else {
        const para =
          body
            .getChild(
              locationIndex
            )
            .asParagraph();

        para.setHeading(
          DocumentApp.ParagraphHeading.HEADING3
        );

        if (
          para.getText() !==
          displayLocation
        ) {
          para.setText(
            displayLocation
          );
        }

        currentIndex =
          locationIndex + 1;
      }

      const existingLines = [];

      let scanIndex =
        currentIndex;

      while (
        scanIndex <
        body.getNumChildren()
      ) {
        const child =
          body.getChild(
            scanIndex
          );

        if (
          child.getType() ===
          DocumentApp.ElementType.PARAGRAPH
        ) {
          const para =
            child.asParagraph();

          if (
            para.getHeading() !==
            DocumentApp.ParagraphHeading.NORMAL
          ) {
            break;
          }
        }

        if (
          child.getType() ===
          DocumentApp.ElementType.LIST_ITEM
        ) {
          existingLines.push({
            index:
              scanIndex,

            listItem:
              child.asListItem(),

            text:
              child
                .asListItem()
                .getText()
          });
        }

        scanIndex++;
      }

      const processedDocIndices =
        new Set();

      let insertPosition =
        currentIndex;

      activeStudents.forEach(
        studentObj => {
          const studentName =
            studentObj.name;

          let note = "";
          let isHighlighted =
            false;

          let isLocked =
            false;

          let color =
            studentObj.color;

          if (
            exam.studentTags &&
            exam.studentTags[
              studentName
            ]
          ) {
            const tag =
              exam.studentTags[
                studentName
              ];

            if (
              typeof tag ===
              "object"
            ) {
              note =
                tag.note ||
                "";

              isHighlighted =
                tag.highlighted ||
                false;

              isLocked =
                tag.locked ||
                false;
            }
          }

          if (
            !color ||
            color ===
              "#000000"
          ) {
            color = null;
          }

          const foundLine =
            existingLines.find(
              line =>
                !processedDocIndices.has(
                  line.index
                ) &&
                line.text.includes(
                  studentName
                )
            );

          if (foundLine) {
            processedDocIndices.add(
              foundLine.index
            );

            const li =
              foundLine.listItem;

            if (!isLocked) {
              const expectedText =
                note
                  ? `${studentName} [${note}]`
                  : studentName;

              if (
                li.getText() !==
                expectedText
              ) {
                li.clear();

                li.setText(
                  studentName
                );

                if (note) {
                  const t =
                    li.appendText(
                      ` [${note}]`
                    );

                  t.setBold(
                    true
                  );

                  t.setForegroundColor(
                    "#000000"
                  );

                  if (
                    isHighlighted
                  ) {
                    t.setBackgroundColor(
                      "#ffff00"
                    );
                  } else {
                    t.setBackgroundColor(
                      "#fff59d"
                    );
                  }
                }
              }

              if (color) {
                li.setForegroundColor(
                  color
                );
              } else {
                li.setForegroundColor(
                  "#000000"
                );
              }
            }

            insertPosition =
              Math.max(
                insertPosition,
                foundLine.index + 1
              );

          } else {
            const li =
              body.insertListItem(
                insertPosition,
                studentName
              );

            if (note) {
              const t =
                li.appendText(
                  ` [${note}]`
                );

              t.setBold(
                true
              );

              t.setForegroundColor(
                "#000000"
              );

              if (
                isHighlighted
              ) {
                t.setBackgroundColor(
                  "#ffff00"
                );
              } else {
                t.setBackgroundColor(
                  "#fff59d"
                );
              }
            }

            if (color) {
              li.setForegroundColor(
                color
              );
            }

            insertPosition++;
          }
        }
      );

      /*
       * Remove students no longer present in the current roster.
       */
      for (
        let i =
          existingLines.length - 1;
        i >= 0;
        i--
      ) {
        const line =
          existingLines[i];

        if (
          !processedDocIndices.has(
            line.index
          )
        ) {
          const elementIndex =
            body.getChildIndex(
              line.listItem
            );

          if (
            elementIndex ===
            body.getNumChildren() - 1
          ) {
            body.appendParagraph(
              " "
            );
          }

          body.removeChild(
            line.listItem
          );

          if (
            line.index <
            insertPosition
          ) {
            insertPosition--;
          }
        }
      }

      currentIndex =
        insertPosition;
    }
  );
}


/* ========================================================================
 * NURSING CALENDAR
 * ====================================================================== */


/**
 * Syncs Nursing exams to the configured calendar.
 *
 * The original event window behavior is preserved.
 *
 * Event descriptions now use the new Start Time field.
 */
function api_syncNursingCalendar(
  payload
) {
  try {
    const config =
      _getNursingSettings();

    if (
      !config.calendarId
    ) {
      return {
        success:
          false,

        message:
          "No Calendar ID."
      };
    }

    const cal =
      CalendarApp.getCalendarById(
        config.calendarId
      );

    if (!cal) {
      return {
        success:
          false,

        message:
          "Calendar not found."
      };
    }

    let count = 0;

    const sheetsToProcess =
      Array.isArray(
        payload.sheets
      )
        ? payload.sheets
        : [payload.sheets];

    sheetsToProcess.forEach(
      sheetData => {
        sheetData.exams.forEach(
          exam => {
            if (!exam.date) {
              return;
            }

            const dateObj =
              new Date(
                String(
                  exam.date
                ).replace(
                  /(st|nd|rd|th)/gi,
                  ""
                )
              );

            if (
              isNaN(
                dateObj.getTime()
              )
            ) {
              return;
            }

            /*
             * Preserve original broad calendar window.
             */
            const start =
              new Date(
                dateObj.getFullYear(),
                dateObj.getMonth(),
                dateObj.getDate(),
                8,
                0,
                0
              );

            const end =
              new Date(
                start
              );

            end.setHours(
              17,
              0,
              0
            );

            const meta =
              _getNursingDocumentMeta(
                sheetData
              );

            const titlePrefix =
              meta.fullTitlePrefix ||
              sheetData.sheetName;

            const title =
              `Proctor: ${titlePrefix} - ${exam.name}`;

            const events =
              cal.getEvents(
                start,
                end
              );

            const exists =
              events.some(
                e =>
                  e.getTitle() ===
                  title
              );

            if (!exists) {
              cal.createEvent(
                title,
                start,
                end,
                {
                  description:
                    `Password: ${exam.password || "-"}\n` +
                    `Start Time: ${exam.startTime || "-"}\n` +
                    `Duration: ${exam.duration || "-"}\n` +
                    `Notes: ${exam.generalNotes || ""}`
                }
              );

              count++;
            }
          }
        );
      }
    );

    if (
      typeof logSystemAction ===
      "function"
    ) {
      logSystemAction(
        "Nursing",
        "Calendar Sync",
        "Batch",
        config.calendarId,
        `Synced ${count}`
      );
    }

    return {
      success:
        true,

      message:
        `Synced ${count} events.`
    };

  } catch (e) {
    return {
      success:
        false,

      message:
        e &&
        e.message
          ? e.message
          : String(e)
    };
  }
}


/* ========================================================================
 * ACCOMMODATIONS DATABASE
 * ====================================================================== */


/**
 * Loads the existing Nursing accommodations database.
 *
 * Unique_ID format:
 *
 *   NUR305|Exam 1
 */
function getAccommodationsDBMap() {
  const map = {};

  try {
    const ss =
      getMasterDataHub();

    const sheet =
      ss.getSheetByName(
        "_DB_ACCOMMODATIONS"
      );

    if (!sheet) {
      return map;
    }

    const data =
      sheet.getDataRange()
        .getValues();

    for (
      let i = 1;
      i < data.length;
      i++
    ) {
      const row =
        data[i];

      const id =
        String(
          row[0] || ""
        ).trim();

      if (!id) {
        continue;
      }

      let tags = {};

      try {
        tags =
          JSON.parse(
            row[4] || "{}"
          );
      } catch (e) {
        tags = {};
      }

      map[id] = {
        generalNotes:
          row[3] || "",

        studentTags:
          tags
      };
    }

  } catch (e) {
    console.warn(
      "Nursing accommodations DB error: " +
      e.message
    );
  }

  return map;
}


/**
 * Saves Nursing accommodation information.
 */
function api_saveNursingAccommodations(
  payload
) {
  try {
    const ss =
      getMasterDataHub();

    let sheet =
      ss.getSheetByName(
        "_DB_ACCOMMODATIONS"
      );

    if (!sheet) {
      sheet =
        ss.insertSheet(
          "_DB_ACCOMMODATIONS"
        );

      sheet.appendRow([
        "Unique_ID",
        "Course_Code",
        "Exam_Name",
        "General_Notes",
        "Student_Data"
      ]);
    }

    const uniqueId =
      `${payload.courseCode}|${payload.examName}`;

    const studentJson =
      JSON.stringify(
        payload.studentTags ||
        {}
      );

    const data =
      sheet.getDataRange()
        .getValues();

    let rowIndex =
      -1;

    for (
      let i = 1;
      i < data.length;
      i++
    ) {
      if (
        String(
          data[i][0]
        ) === uniqueId
      ) {
        rowIndex =
          i + 1;

        break;
      }
    }

    if (
      rowIndex > -1
    ) {
      sheet
        .getRange(
          rowIndex,
          4,
          1,
          2
        )
        .setValues([
          [
            payload.generalNotes ||
              "",

            studentJson
          ]
        ]);

    } else {
      sheet.appendRow([
        uniqueId,

        payload.courseCode,

        payload.examName,

        payload.generalNotes ||
          "",

        studentJson
      ]);
    }

    return {
      success:
        true,

      message:
        "Saved to Database!"
    };

  } catch (e) {
    return {
      success:
        false,

      message:
        e &&
        e.message
          ? e.message
          : String(e)
    };
  }
}