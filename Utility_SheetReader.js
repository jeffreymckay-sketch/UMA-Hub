/**
 * @file Utility_SheetReader.js
 * @description
 * Standalone utilities for parsing unstructured Google Sheets
 * using heuristic matching instead of fixed row/column positions.
 */

const SheetReader = {

  /**
   * Aggressively sanitizes strings for loose comparisons.
   *
   * Example:
   *   "  Start   Time " -> "start time"
   */
  sanitize: function(str) {
    if (
      str === null ||
      str === undefined
    ) {
      return "";
    }

    return String(str)
      .toLowerCase()
      .trim()
      .replace(/\s+/g, " ");
  },


  /**
   * Normalizes header text more aggressively than sanitize().
   *
   * Punctuation and separators become spaces so values such as:
   *
   *   Duration/Mins
   *   Duration - Mins
   *   duration_mins
   *
   * can all compare as:
   *
   *   duration mins
   */
  normalizeHeader: function(str) {
    if (
      str === null ||
      str === undefined
    ) {
      return "";
    }

    return String(str)
      .toLowerCase()
      .trim()
      .replace(/[_\/\\|:;,\-–—]+/g, " ")
      .replace(/[^a-z0-9.'&() ]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  },


  /**
   * Scans rows to find the most likely header row by scoring keyword
   * presence.
   *
   * Legacy utility retained for compatibility with other parsers.
   *
   * @param {Array<Array<any>>} data
   * @param {Array<string>} keywords
   * @param {number} maxScanDepth
   * @returns {number}
   */
  findHeaderRowHeuristic: function(
    data,
    keywords,
    maxScanDepth = 20
  ) {
    if (
      !data ||
      data.length === 0
    ) {
      return -1;
    }

    let bestRowIndex = -1;
    let maxScore = 0;

    const sanitizedKeywords =
      keywords
        .map(
          keyword =>
            this.sanitize(keyword)
        )
        .filter(Boolean);

    for (
      let i = 0;
      i <
      Math.min(
        data.length,
        maxScanDepth
      );
      i++
    ) {
      if (!data[i]) {
        continue;
      }

      let rowScore = 0;

      const rowStr =
        data[i]
          .map(
            cell =>
              this.sanitize(cell)
          )
          .join(" ");

      sanitizedKeywords.forEach(
        keyword => {
          if (
            rowStr.includes(
              keyword
            )
          ) {
            rowScore++;
          }
        }
      );

      /*
       * Tie-breaker goes to the earlier row.
       */
      if (
        rowScore > maxScore
      ) {
        maxScore = rowScore;
        bestRowIndex = i;
      }
    }

    return maxScore > 0
      ? bestRowIndex
      : -1;
  },


  /**
   * Legacy synonym mapper retained for compatibility.
   *
   * This uses the original first-match behavior.
   */
  mapColumnsBySynonyms: function(
    headerRow,
    synonymConfig
  ) {
    const colMap = {};

    const sanitizedHeaders =
      headerRow.map(
        header =>
          this.sanitize(header)
      );

    for (
      const [key, synonyms]
      of Object.entries(
        synonymConfig
      )
    ) {
      colMap[key] = -1;

      const sanitizedSynonyms =
        synonyms
          .map(
            synonym =>
              this.sanitize(
                synonym
              )
          )
          .filter(Boolean);

      for (
        let i = 0;
        i <
        sanitizedHeaders.length;
        i++
      ) {
        const headerStr =
          sanitizedHeaders[i];

        if (!headerStr) {
          continue;
        }

        const match =
          sanitizedSynonyms.some(
            synonym =>
              headerStr ===
                synonym ||
              headerStr.includes(
                synonym
              )
          );

        if (match) {
          colMap[key] = i;
          break;
        }
      }
    }

    return colMap;
  },


  /**
   * Scores one header value against a group of synonyms.
   *
   * Higher score = stronger semantic match.
   *
   * Exact matches heavily outrank loose substring matches.
   *
   * This prevents generic synonyms such as "exam" from beating
   * a stronger exact header such as "exam date" when mapping fields.
   */
  scoreHeaderAgainstSynonyms:
    function(
      headerValue,
      synonyms
    ) {
      const header =
        this.normalizeHeader(
          headerValue
        );

      if (!header) {
        return 0;
      }

      let bestScore = 0;

      synonyms.forEach(
        synonymValue => {
          const synonym =
            this.normalizeHeader(
              synonymValue
            );

          if (!synonym) {
            return;
          }

          let score = 0;

          /*
           * Perfect normalized match.
           */
          if (
            header === synonym
          ) {
            score =
              120 +
              Math.min(
                synonym.length,
                20
              );
          }

          /*
           * Header starts or ends with the complete synonym.
           *
           * Examples:
           *   "exam date" matches "exam"
           *   "scheduled start time" matches "start time"
           */
          else if (
            header.startsWith(
              synonym + " "
            ) ||
            header.endsWith(
              " " + synonym
            )
          ) {
            score =
              90 +
              Math.min(
                synonym.length,
                15
              );
          }

          /*
           * Whole phrase/token match inside a longer header.
           */
          else if (
            (
              " " +
              header +
              " "
            ).includes(
              " " +
              synonym +
              " "
            )
          ) {
            score =
              80 +
              Math.min(
                synonym.length,
                15
              );
          }

          /*
           * Loose substring match is allowed, but intentionally weak.
           */
          else if (
            synonym.length >= 4 &&
            header.includes(
              synonym
            )
          ) {
            score =
              55 +
              Math.min(
                synonym.length,
                10
              );
          }

          if (
            score > bestScore
          ) {
            bestScore = score;
          }
        }
      );

      return bestScore;
    },


  /**
   * Maps columns using scored synonym matching.
   *
   * Return format:
   *
   * {
   *   columns: {
   *     exam: 2,
   *     date: 3
   *   },
   *   scores: {
   *     exam: 124,
   *     date: 129
   *   }
   * }
   *
   * Unlike mapColumnsBySynonyms(), this searches every possible header
   * and returns the strongest match rather than the first substring hit.
   */
  mapColumnsBySynonymsScored:
    function(
      headerRow,
      synonymConfig
    ) {
      const columns = {};
      const scores = {};

      const row =
        headerRow || [];

      for (
        const [key, synonyms]
        of Object.entries(
          synonymConfig
        )
      ) {
        let bestColumn = -1;
        let bestScore = 0;

        for (
          let c = 0;
          c < row.length;
          c++
        ) {
          const score =
            this.scoreHeaderAgainstSynonyms(
              row[c],
              synonyms
            );

          if (
            score > bestScore
          ) {
            bestScore = score;
            bestColumn = c;
          }
        }

        columns[key] =
          bestColumn;

        scores[key] =
          bestScore;
      }

      return {
        columns: columns,
        scores: scores
      };
    },


  /**
   * Generic scored-header-row finder.
   *
   * Not required by the Nursing parser directly, but useful for future
   * resilient parsers.
   *
   * options:
   * {
   *   startRow: 0,
   *   maxRows: 40,
   *   requiredKeys: ["exam", "date"],
   *   requireDistinctRequiredColumns: true
   * }
   */
  findBestHeaderRowBySynonyms:
    function(
      data,
      synonymConfig,
      options
    ) {
      if (
        !data ||
        !data.length
      ) {
        return null;
      }

      options =
        options || {};

      const startRow =
        Math.max(
          0,
          Number(
            options.startRow || 0
          )
        );

      const maxRows =
        options.maxRows
          ? Math.min(
              data.length,
              startRow +
                Number(
                  options.maxRows
                )
            )
          : data.length;

      const requiredKeys =
        options.requiredKeys ||
        [];

      const requireDistinct =
        options
          .requireDistinctRequiredColumns !==
        false;

      let best = null;

      for (
        let r = startRow;
        r < maxRows;
        r++
      ) {
        const mapped =
          this.mapColumnsBySynonymsScored(
            data[r] || [],
            synonymConfig
          );

        let valid = true;

        requiredKeys.forEach(
          key => {
            if (
              mapped.columns[key] ===
              -1
            ) {
              valid = false;
            }
          }
        );

        if (!valid) {
          continue;
        }

        if (
          requireDistinct &&
          requiredKeys.length > 1
        ) {
          const requiredColumns =
            requiredKeys.map(
              key =>
                mapped.columns[key]
            );

          const uniqueColumns =
            Array.from(
              new Set(
                requiredColumns
              )
            );

          if (
            uniqueColumns.length !==
            requiredColumns.length
          ) {
            continue;
          }
        }

        let totalScore = 0;
        let matchedKeys = 0;

        Object.keys(
          mapped.columns
        ).forEach(key => {
          if (
            mapped.columns[key] !==
            -1
          ) {
            matchedKeys++;

            totalScore +=
              mapped.scores[key] ||
              0;
          }
        });

        /*
         * Reward rows that match several semantic fields.
         */
        totalScore +=
          matchedKeys * 10;

        if (
          !best ||
          totalScore >
            best.score
        ) {
          best = {
            rowIndex: r,
            columns:
              mapped.columns,
            scores:
              mapped.scores,
            matchedKeys:
              matchedKeys,
            score:
              totalScore
          };
        }
      }

      return best;
    },


  /**
   * Scans rows to find a specific anchor word in the first column.
   *
   * Legacy function retained for compatibility with other code.
   * The Nursing parser intentionally does NOT use this function.
   */
  findAnchorRow: function(
    data,
    anchorWords,
    startRow = 0
  ) {
    if (
      !data ||
      data.length === 0
    ) {
      return -1;
    }

    const sanitizedAnchors =
      anchorWords
        .map(
          word =>
            this.sanitize(word)
        )
        .filter(Boolean);

    for (
      let i = startRow;
      i < data.length;
      i++
    ) {
      if (
        !data[i] ||
        data[i].length === 0
      ) {
        continue;
      }

      const firstCell =
        this.sanitize(
          data[i][0]
        );

      if (!firstCell) {
        continue;
      }

      if (
        sanitizedAnchors.some(
          anchor =>
            firstCell.includes(
              anchor
            )
        )
      ) {
        return i;
      }
    }

    return -1;
  },


  /**
   * Dynamically reads a list based on one primary column.
   *
   * Legacy function retained for compatibility.
   *
   * The Nursing student roster parser intentionally no longer uses this
   * because Nursing rosters require cell-by-cell classification.
   */
  readDynamicRoster: function(
    data,
    startRow,
    nameColIndex,
    stopWords,
    maxBlankRows = 3
  ) {
    const roster = [];

    let blankCount = 0;

    const sanitizedStopWords =
      stopWords
        .map(
          word =>
            this.sanitize(word)
        )
        .filter(Boolean);

    for (
      let i = startRow;
      i < data.length;
      i++
    ) {
      const row =
        data[i];

      if (
        !row ||
        row.length <=
          nameColIndex
      ) {
        blankCount++;

        if (
          blankCount >=
          maxBlankRows
        ) {
          break;
        }

        continue;
      }

      const primaryCell =
        this.sanitize(
          row[nameColIndex]
        );

      if (
        primaryCell === ""
      ) {
        blankCount++;

        if (
          blankCount >=
          maxBlankRows
        ) {
          break;
        }

        continue;
      }

      blankCount = 0;

      if (
        sanitizedStopWords.some(
          stopWord =>
            primaryCell ===
              stopWord ||
            primaryCell.includes(
              stopWord
            )
        )
      ) {
        break;
      }

      const formattedRow =
        row.map(cell => {
          const strCell =
            String(
              cell === null ||
              cell === undefined
                ? ""
                : cell
            ).trim();

          return strCell === ""
            ? "TBD"
            : cell;
        });

      roster.push({
        rowIndex: i,
        name:
          String(
            row[nameColIndex]
          ).trim(),
        data:
          formattedRow
      });
    }

    return roster;
  }
};