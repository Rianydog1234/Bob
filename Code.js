/**********************
 * CONFIG
 **********************/

const SPREADSHEET_ID = "1dolNsNrpgD0AoWTCBz-cijXvZp5uorFoRYPccKRWpoY";

const SHEET_GIDS = [
  0,          // Sheet 1: Last First
  1493058526, // Sheet 2: First Last
  1220633850  // Sheet 3: First Last
];

const PROJECT_COLUMNS = [1, 3, 5, 7, 9, 11]; // B D F H J L

// Default fallback target (kept for legacy fallback).
const TARGET_SPREADSHEET_ID = "1dolNsNrpgD0AoWTCBz-cijXvZp5uorFoRYPccKRWpoY";
const SOURCE_INFO_SHEET_GID = 1446473767;
// NOTE: We no longer hard-code TARGET_SHEET_GID constant; instead we read it from config if present.

const GRADE_ORDER = {
  "2A": 1,
  "2B": 2,
  "3A": 3,
  "3B": 4
};

/**********************
 * CONFIG helpers (store the active target sheet id)
 * We persist chosen target sheet id in a sheet named "ScriptConfig" in the SOURCE spreadsheet (A1).
 **********************/
function getConfigSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("ScriptConfig");
  if (!sheet) {
    sheet = ss.insertSheet("ScriptConfig");
    sheet.getRange("A1").setValue("TARGET_SHEET_ID");
    sheet.getRange("B1").setValue(""); // store gid as number
  }
  return sheet;
}

function setTargetSheetId(gid) {
  const sheet = getConfigSheet();
  sheet.getRange("B1").setValue(gid);
}

function getTargetSheetId() {
  const sheet = getConfigSheet();
  const gid = sheet.getRange("B1").getValue();
  if (gid) return Number(gid);
  // fallback: try to find a sheet named "Target" in the target spreadsheet
  // or fall back to the first sheet in TARGET_SPREADSHEET_ID
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const s = ss.getSheets()[0];
  return s ? s.getSheetId() : null;
}

function getTargetSheet() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const gid = getTargetSheetId();
  if (!gid) throw new Error("Target sheet ID not configured");
  const sheet = ss.getSheets().find(s => s.getSheetId() === gid);
  if (!sheet) throw new Error("Target sheet not found (gid=" + gid + ")");
  return sheet;
}

/**********************
 * Get project column mapping (reads from Sheet1 row 2)
 **********************/
function getProjectColumnMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SHEET_GIDS[0]);
  const data = sheet.getDataRange().getValues();
  const projectRow = data[1]; // row 2

  const map = {};
  PROJECT_COLUMNS.forEach(col => {
    const name = projectRow[col];
    if (name) {
      map[name.trim()] = col;
    }
  });

  return map;
}

/**********************
 * WEB ENTRY
 **********************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

/**********************
 * HELPERS
 **********************/
function normalizeName(name) {
  if (!name) return "";
  return name
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function swapNameIfNeeded(name, sheetIndex) {
  if (!name) return "";
  if (sheetIndex === 1 || sheetIndex === 2) {
    const p = name.trim().split(/\s+/);
    if (p.length < 2) return name;
    return `${p[1]} ${p[0]}`;
  }
  return name;
}

/************
 * CONDITIONAL FORMATTING
 * Applies a single set of grade rules for the entire grade column in the sheet
 */
function setGradeConditionalFormatting(sheet, col) {
  const rules = sheet.getConditionalFormatRules();

  // Remove old grade rules that target this column (col is zero-based index from map)
  const filteredRules = rules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() === col + 2); // grade column = nameCol + 1 -> +2 here
  });

  const startRow = 3;
  const maxRows = 100; // safely covers future entries
  const gradeRange = sheet.getRange(startRow, col + 2, maxRows, 1);

  const newRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('2A')
      .setBackground('#b6d7a8')
      .setRanges([gradeRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('2B')
      .setBackground('#a4c2f4')
      .setRanges([gradeRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('3A')
      .setBackground('#ffe599')
      .setRanges([gradeRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('3B')
      .setBackground('#ea9999')
      .setRanges([gradeRange])
      .build()
  ];

  sheet.setConditionalFormatRules([...filteredRules, ...newRules]);
}

/**********************
 * LOAD STUDENT + PROJECT DATA
 **********************/
function getStudentData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const students = {};
  const projects = new Set();

  SHEET_GIDS.forEach((gid, sheetIndex) => {
    const sheet = ss.getSheets().find(s => s.getSheetId() === gid);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const projectRow = data[1];

    PROJECT_COLUMNS.forEach(col => {
      const projectName = projectRow[col];
      if (!projectName) return;

      projects.add(projectName);

      for (let r = 2; r < data.length; r++) {
        let fullName = data[r][col];
        if (!fullName) continue;

        fullName = swapNameIfNeeded(fullName, sheetIndex);
        const key = normalizeName(fullName);

        if (!students[key]) {
          students[key] = {
            fullName,
            projects: {}
          };
        }

        students[key].projects[projectName] = true;
      }
    });
  });

  return {
    students: Object.values(students),
    projects: Array.from(projects)
  };
}

/**********************
 * GET STUDENT INFO (minimal)
 **********************/
function getStudentInfo(fullName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SOURCE_INFO_SHEET_GID);
  if (!sheet) throw new Error("Info sheet not found");

  const data = sheet.getDataRange().getValues();
  const key = normalizeName(fullName);

  for (let r = 1; r < data.length; r++) {
    const last = data[r][0];
    const first = data[r][1];
    if (!first || !last) continue;

    if (normalizeName(`${last} ${first}`) === key) {
      return { grade: "", gender: "" };
    }
  }

  throw new Error("Student info not found: " + fullName);
}

/**********************
 * GET STUDENT GRADE + GENDER
 * Reads info from the info sheet (columns: Last | First | Grade | Gender)
 **********************/
function getStudentGradeGender(studentName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SOURCE_INFO_SHEET_GID);
  if (!sheet) {
    throw new Error(
      `‼️ Info sheet not found (expected GID=${SOURCE_INFO_SHEET_GID}).\n` +
      `Check that the info sheet exists and its GID hasn’t changed!`
    );
  }

  const data = sheet.getDataRange().getValues();
  const keyNormalized = normalizeName(studentName);

  const matches = [];
  for (let r = 1; r < data.length; r++) {
    const last = data[r][0];
    const first = data[r][1];
    if (!last || !first) continue;
    const combinedLF = normalizeName(`${last} ${first}`);
    const combinedFL = normalizeName(`${first} ${last}`);
    if (combinedLF === keyNormalized || combinedFL === keyNormalized) {
      return {
        grade: data[r][2],
        gender: data[r][3]
      };
    }
    matches.push(combinedLF);
  }

  // Failed lookup — throw with all possible info
  throw new Error(
    `🛑 Student info lookup failed!\n` +
    `Input name: "${studentName}"\n` +
    `Normalized lookup key: "${keyNormalized}"\n` +
    `Info sheet rows scanned: ${data.length - 1}\n` +
    `Sample matches available (first 10): [${matches.slice(0,10).join(", ")}]\n` +
    `Ensure the name matches exactly (accents, spaces, order).\n` +
    `Check full list of names in info sheet for typos or missing entries.`
  );
}



/**********************
 * WRITE STUDENT TO PROJECT
 * Writes name + grade, colors by gender, sorts by grade
 **********************/
function writeStudentToProject(fullName, projectName, grade, gender) {
  const projectMap = getProjectColumnMap();
  const col = projectMap[projectName.trim()];

  if (col === undefined) throw new Error("Invalid project (mapping failed)");

  const sheet = getTargetSheet();
  if (!sheet) throw new Error("Target sheet not found");

  const startRow = 3;
  const maxRows = 16;

  // Check if student was already on this project
  const range = sheet.getRange(startRow, col + 1, maxRows, 2);
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeName(values[i][0]) === normalizeName(fullName)) {
      throw new Error(`${fullName} has already been on ${projectName}`);
    }
  }

  // Find first empty row
  let insertRow = startRow;
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      insertRow = startRow + i;
      break;
    }
  }

  // Clear any old formatting first only on the insertion row
  sheet.getRange(insertRow, col + 1, 1, 2).clearFormat();

  // Write name + grade
  sheet.getRange(insertRow, col + 1).setValue(fullName); // name column
  sheet.getRange(insertRow, col + 2).setValue(grade);    // grade column

  // Set gender color manually (kept from original)
  const genderColor =
    (gender || "").toString().toLowerCase() === "f"
      ? "#ff000033"
      : (gender || "").toString().toLowerCase() === "m"
      ? "#0000ff33"
      : "#ffffff00"; // transparent if missing
  sheet.getRange(insertRow, col + 1).setBackground(genderColor);

  // Apply conditional formatting for grade colors (single-shot)
  setGradeConditionalFormatting(sheet, col);

  // Remove manual backgrounds in full column range before sorting (prevents ghost colors)
  sheet.getRange(startRow, col + 1, maxRows, 2).setBackground(null);

  // Re-write the student (since we cleared background)
  sheet.getRange(insertRow, col + 1).setValue(fullName);
  sheet.getRange(insertRow, col + 2).setValue(grade);
  sheet.getRange(insertRow, col + 1).setBackground(genderColor);

  // Sort by grade using safe ascending order (grade column)
  sheet.getRange(startRow, col + 1, maxRows, 2).sort([{ column: col + 2, ascending: true }]);

  // After sorting, reapply formatting rules (conditional formatting already set),
  // gender colors may be inconsistent after sort so we reapply gender for all rows:
  applyGenderColors(sheet, col, startRow, maxRows);

  return "OK";
}

/**********************
 * applyGenderColors (repaints gender backgrounds after sorts)
 **********************/
function applyGenderColors(sheet, col, startRow, maxRows) {
  const names = sheet.getRange(startRow, col + 1, maxRows, 1).getValues();

  for (let i = 0; i < names.length; i++) {
    const fullName = names[i][0];
    if (!fullName) {
      // clear any leftover background
      sheet.getRange(startRow + i, col + 1).setBackground(null);
      continue;
    }
    let gender = "";
    try {
      gender = (getStudentGradeGender(fullName).gender || "").toString().toLowerCase();
    } catch (e) {
      gender = "";
    }

    const color =
      gender === "f" ? "#ff000033" :
      gender === "m"   ? "#0000ff33" :
      null;

    if (color) {
      sheet.getRange(startRow + i, col + 1).setBackground(color);
    } else {
      sheet.getRange(startRow + i, col + 1).setBackground(null);
    }
  }
}

/**********************
 * SUBMIT STUDENT TO PROJECT (wrapper)
 * Combines grade/gender fetch + writing
 **********************/
function submitStudentToProject(fullName, projectName) {
  const { grade, gender } = getStudentGradeGender(fullName);
  return writeStudentToProject(fullName, projectName, grade, gender);
}
/**********************
 * Read All groups
 **********************/
function readAllGroups() {
  const sheet = getTargetSheet(); // uses your TARGET_SPREADSHEET_ID + GID from ScriptConfig
  const startRow = 3;
  const maxRows = 16;
  const groups = {};
  const map = getProjectColumnMap();

  Object.entries(map).forEach(([project, col]) => {
    const values = sheet.getRange(startRow, col + 1, maxRows, 2).getValues();

    groups[project] = values
      .map((r, index) => {
        const rawNameCell = (r[0] || "").toString().trim();
        const rawGradeCell = (r[1] || "").toString().trim();
        const sheetRow = startRow + index;

        // Skip empty cells
        if (!rawNameCell) return null;

        // Count words to ensure at least 2-word names
        const wordCount = rawNameCell.split(/\s+/).filter(Boolean).length;
        if (wordCount < 2) {
          Logger.log(
            "SKIP single-word name -> Project: '%s', Row: %s, NameCell: '%s', GradeCell: '%s'",
            project,
            sheetRow,
            rawNameCell,
            rawGradeCell
          );
          return null;
        }

        // Multi-word: lookup
        try {
          const info = getStudentGradeGender(rawNameCell); // existing function
          return {
            name: rawNameCell,
            gender: (info.gender || "").toString().toLowerCase(),
            grade: info.grade || ""
          };
        } catch (e) {
          const detailed = [
            `READ GROUPS LOOKUP FAILURE!`,
            `Project: "${project}"`,
            `Sheet row: ${sheetRow}`,
            `Raw name cell: "${rawNameCell}"`,
            `Raw grade cell: "${rawGradeCell}"`,
            `Underlying error: ${e.message || e.toString()}`
          ].join("\n");
          Logger.log(detailed);
          throw new Error(detailed);
        }
      })
      .filter(Boolean); // remove skipped single-word or empty entries
  });

  return groups;
}


/**********************
 * Write Balanced groups
 **********************/
function writeBalancedGroups(groups) {
  const sheet = getTargetSheet();
  const map = getProjectColumnMap();
  const startRow = 3;
  const maxRows = 16;

  Object.entries(groups).forEach(([project, students]) => {
    const col = map[project];

    // clear old data
    sheet.getRange(startRow, col + 1, maxRows, 2).clearContent();

    // write new
    students.slice(0, maxRows).forEach((s, i) => {
      sheet.getRange(startRow + i, col + 1).setValue(s.name);
      sheet.getRange(startRow + i, col + 2).setValue(s.grade);
    });

    // reapply formatting
    setGradeConditionalFormatting(sheet, col);
  });

  // reapply gender colors for all
  Object.entries(map).forEach(([project, col]) => {
    applyGenderColors(sheet, col, startRow, maxRows);
  });
}

/**********************
 * CONSTANTS
 **********************/
const MAX_STUDENTS = 16;
const MIN_STUDENTS = 15;
const MAX_GENDER_DIFF = 2;
const MAX_GRADE_DIFF = 3;
const MAX_SHUFFLE_PASSES = 20;

/**********************
 * ANALYZE GROUP
 **********************/
function analyzeGroup(group) {
  let girls = 0, boys = 0, g2 = 0, g3 = 0;

  group.forEach(s => {
    if ((s.gender || "").toLowerCase() === "f" || (s.gender || "").toLowerCase() === "female") girls++;
    else boys++;

    if ((s.grade || "").toString().startsWith("2")) g2++;
    else if ((s.grade || "").toString().startsWith("3")) g3++;
  });

  return {
    total: group.length,
    girls,
    boys,
    g2,
    g3,
    genderDiff: Math.abs(girls - boys),
    gradeDiff: Math.abs(g3 - g2),
    valid: (
      group.length >= MIN_STUDENTS &&
      group.length <= MAX_STUDENTS &&
      Math.abs(girls - boys) <= MAX_GENDER_DIFF &&
      Math.abs(g3 - g2) <= MAX_GRADE_DIFF
    )
  };
}

/**********************
 * BUILD HISTORY MAP
 **********************/
function buildHistoryMapFromWasOnProject() {
  const sh = getWasOnProjectSheet();
  const rows = sh.getDataRange().getValues().slice(1);
  const history = {};

  rows.forEach(([name, project]) => {
    if (!name) return;
    const key = normalizeName(name);
    if (!history[key]) history[key] = {};
    history[key][project] = true;
  });

  return history;
}

/**********************
 * RANDOMIZED SWAP ENGINE
 **********************/
function shuffleProjectsRandomly(groups, history) {
  const projectNames = Object.keys(groups);

  for (let pass = 0; pass < MAX_SHUFFLE_PASSES; pass++) {
    let changed = false;

    projectNames.forEach(p1 => {
      projectNames.forEach(p2 => {
        if (p1 === p2) return;

        const g1 = groups[p1];
        const g2 = groups[p2];

        const a1 = analyzeGroup(g1);
        const a2 = analyzeGroup(g2);

        // Skip if both valid
        if (a1.valid && a2.valid) return;

        // Randomize candidates
        const c1 = [...g1].sort(() => Math.random() - 0.5);
        const c2 = [...g2].sort(() => Math.random() - 0.5);

        for (const s1 of c1) {
          for (const s2 of c2) {
            // HARD RULE: history violation check
            if (history[s1.name]?.[p2] || history[s2.name]?.[p1]) continue;

            // Swap candidates
            const newG1 = g1.filter(s => s !== s1).concat(s2);
            const newG2 = g2.filter(s => s !== s2).concat(s1);

            const na1 = analyzeGroup(newG1);
            const na2 = analyzeGroup(newG2);

            // Accept swap if improves or fixes any invalid rule
            const improves = 
              (!a1.valid || na1.valid) && (!a2.valid || na2.valid) ||
              na1.genderDiff < a1.genderDiff || na2.genderDiff < a2.genderDiff ||
              na1.gradeDiff < a1.gradeDiff || na2.gradeDiff < a2.gradeDiff ||
              (na1.total >= MIN_STUDENTS && na1.total <= MAX_STUDENTS) ||
              (na2.total >= MIN_STUDENTS && na2.total <= MAX_STUDENTS);

            if (improves) {
              groups[p1] = newG1;
              groups[p2] = newG2;
              changed = true;
              return;
            }
          }
        }
      });
    });

    if (!changed) break;
  }

  // FINAL CHECK: If any group is still invalid, do a random override to satisfy min/max students
  projectNames.forEach(p => {
    const g = groups[p];
    let ag = analyzeGroup(g);
    if (g.length < MIN_STUDENTS) {
      // Pull random students from other groups not in history
      for (const otherP of projectNames) {
        if (otherP === p) continue;
        const donorGroup = groups[otherP];
        for (const s of donorGroup) {
          if (!history[s.name]?.[p]) {
            g.push(s);
            donorGroup.splice(donorGroup.indexOf(s),1);
            ag = analyzeGroup(g);
            if (ag.total >= MIN_STUDENTS) break;
          }
        }
        if (ag.total >= MIN_STUDENTS) break;
      }
    }
    if (g.length > MAX_STUDENTS) {
      // Push random students to other groups
      for (const otherP of projectNames) {
        if (otherP === p) continue;
        const targetGroup = groups[otherP];
        for (const s of [...g]) {
          if (!history[s.name]?.[otherP] && targetGroup.length < MAX_STUDENTS) {
            targetGroup.push(s);
            g.splice(g.indexOf(s),1);
            ag = analyzeGroup(g);
            if (ag.total <= MAX_STUDENTS) break;
          }
        }
        if (ag.total <= MAX_STUDENTS) break;
      }
    }
  });

  return groups;
}

/**********************
 * SHUFFLE WITH RULES
 **********************/
function shuffleWithRules() {
  const groups = readAllGroups();
  const history = buildHistoryMapFromWasOnProject();

  const balanced = shuffleProjectsRandomly(groups, history);

  // HARD VALIDATION: no history repeats
  Object.entries(balanced).forEach(([project, students]) => {
    students.forEach(s => {
      if (history[s.name]?.[project]) {
        throw new Error(`History violation: ${s.name} already did ${project}`);
      }
    });
  });

  writeBalancedGroups(balanced);
  return true;
}

/**********************
 * ADMIN CLOSE SHEET
 **********************/
function adminCloseSheet() {
  try {
    shuffleWithRules();
    archiveClosedSheet();
    return { ok: true, message: "Shuffle & close completed successfully." };
  } catch (e) {
    return { ok: false, message: e.message || e.toString() };
  }
}


