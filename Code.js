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
function getStudentGradeGender(fullName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SOURCE_INFO_SHEET_GID);
  if (!sheet) throw new Error("Info sheet not found");

  const data = sheet.getDataRange().getValues();
  const key = normalizeName(fullName); // dropdown gives Last First

  for (let r = 1; r < data.length; r++) { // skip header row
    const last = data[r][0];
    const first = data[r][1];
    const grade = data[r][2];
    const gender = data[r][3];

    if (!first || !last) continue;

    const combined = normalizeName(`${last} ${first}`); // Last First
    if (combined === key) {
      return { grade: grade || "", gender: (gender || "").toString() };
    }
  }

  throw new Error("Student info not found: " + fullName);
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
    (gender || "").toString().toLowerCase() === "female"
      ? "#ff000033"
      : (gender || "").toString().toLowerCase() === "male"
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
      gender === "female" ? "#ff000033" :
      gender === "male"   ? "#0000ff33" :
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
 * READ ALL GROUPS
 **********************/
function readAllGroups() {
  const sheet = getTargetSheet();

  const startRow = 3;
  const maxRows = 16;
  const groups = {};
  const map = getProjectColumnMap();

  Object.entries(map).forEach(([name, col]) => {
    const values = sheet.getRange(startRow, col + 1, maxRows, 2).getValues();

    groups[name] = values
      .filter(r => r[0])
      .map(r => {
        const info = getStudentGradeGender(r[0]);
        return {
          name: r[0],
          gender: (info.gender || "").toString().toLowerCase(),
          grade: info.grade || ""
        };
      });
  });

  return groups;
}

/**********************
 * ANALYZE GROUP
 **********************/
function analyzeGroup(group) {
  let girls = 0, boys = 0, g2 = 0, g3 = 0;

  group.forEach(s => {
    if ((s.gender || "").toLowerCase() === "female") girls++;
    if ((s.gender || "").toLowerCase() === "male") boys++;
    if ((s.grade || "").toString().startsWith("2")) g2++;
    if ((s.grade || "").toString().startsWith("3")) g3++;
  });

  return {
    girls, boys, g2, g3,
    extraGirls: girls - boys,
    extraGrades: g3 - g2
  };
}

/**********************
 * WRITE GROUPS BACK
 **********************/
function writeBalancedGroups(groups) {
  const sheet = getTargetSheet();
  const map = getProjectColumnMap();

  const startRow = 3;
  const maxRows = 16;

  Object.entries(groups).forEach(([group, students]) => {
    const col = map[group];

    sheet.getRange(startRow, col + 1, maxRows, 2).clearContent();

    students.slice(0, 16).forEach((s, i) => {
      sheet.getRange(startRow + i, col + 1).setValue(s.name);
      sheet.getRange(startRow + i, col + 2).setValue(s.grade);
    });

    // Re-apply conditional formatting for the grade column
    setGradeConditionalFormatting(sheet, col);
  });

  // After writing everything, reapply gender colors to be deterministic
  Object.entries(map).forEach(([project, col]) => {
    applyGenderColors(sheet, col, startRow, maxRows);
  });
}


/**********************
 * PROJECT HISTORY MAP (safe)
 **********************/
function getWasOnProjectMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("WasOnProject");
  if (!sheet) {
    sheet = ss.insertSheet("WasOnProject");
    sheet.appendRow(["Name", "Project", "Date"]);
  }

  const data = sheet.getDataRange().getValues();
  const map = {};

  for (let r = 1; r < data.length; r++) {
    const rawName = data[r][0];
    if (!rawName) continue;
    const name = normalizeName(rawName);
    const project = (data[r][1] || "").toString();
    if (!map[name]) map[name] = [];
    if (project && !map[name].includes(project)) map[name].push(project);
  }
  return map;
}

/**********************
 * ARCHIVE CLOSED SHEET (append only non-duplicates)
 **********************/
function archiveClosedSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("WasOnProject");
  if (!sheet) {
    sheet = ss.insertSheet("WasOnProject");
    sheet.appendRow(["Name", "Project", "Date"]);
  }

  const history = getWasOnProjectMap(); // existing
  const groups = readAllGroups();

  Object.entries(groups).forEach(([project, students]) => {
    students.forEach(s => {
      const key = normalizeName(s.name);
      // only append if this project not already recorded for the student
      if (!(history[key] || []).includes(project)) {
        sheet.appendRow([s.name, project, new Date()]);
      }
    });
  });
}

/*******
 * NEW THINGS
 */
const MAX_STUDENTS = 16;
const MIN_STUDENTS = 15;
const MAX_GENDER_DIFF = 2;
const MAX_GRADE_DIFF = 3;
const MAX_SHUFFLE_PASSES = 20;
/****
 * Analyze group
 */
function analyzeGroup(group) {
  let girls = 0, boys = 0, g2 = 0, g3 = 0;

  group.forEach(s => {
    if (s.gender === "F") girls++;
    else boys++;

    if (s.grade.startsWith("2")) g2++;
    else if (s.grade.startsWith("3")) g3++;
  });

  return {
    total: group.length,
    girls,
    boys,
    g2,
    g3,
    genderDiff: Math.abs(girls - boys),
    gradeDiff: Math.abs(g2 - g3),
    valid:
      group.length >= MIN_STUDENTS &&
      group.length <= MAX_STUDENTS &&
      Math.abs(girls - boys) <= MAX_GENDER_DIFF &&
      Math.abs(g2 - g3) <= MAX_GRADE_DIFF
  };
}
/**
 * Project map
 */
function buildHistoryMapFromWasOnProject() {
  const sh = getWasOnProjectSheet();
  const rows = sh.getDataRange().getValues().slice(1);
  const history = {};

  rows.forEach(([name, project]) => {
    if (!history[name]) history[name] = {};
    history[name][project] = true;
  });

  return history;
}
/***\
 * Shuffle projects randomly
 */
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

        if (a1.valid && a2.valid) return;

        // Randomize candidate order
        const c1 = [...g1].sort(() => Math.random() - 0.5);
        const c2 = [...g2].sort(() => Math.random() - 0.5);

        for (const s1 of c1) {
          for (const s2 of c2) {
            // HARD RULE: no history violations
            if (
              history[s1.name]?.[p2] ||
              history[s2.name]?.[p1]
            ) continue;

            // Simulate swap
            const newG1 = g1.filter(s => s !== s1).concat(s2);
            const newG2 = g2.filter(s => s !== s2).concat(s1);

            const na1 = analyzeGroup(newG1);
            const na2 = analyzeGroup(newG2);

            // Accept if it improves or fixes something
            if (
              na1.genderDiff <= a1.genderDiff &&
              na2.genderDiff <= a2.genderDiff &&
              na1.gradeDiff <= a1.gradeDiff &&
              na2.gradeDiff <= a2.gradeDiff &&
              newG1.length <= MAX_STUDENTS &&
              newG2.length <= MAX_STUDENTS
            ) {
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

  return groups;
}

/****\
 * Shuffle with rules
 */
function shuffleWithRules() {
  const groups = readAllGroups(); // your existing function
  const history = buildHistoryMapFromWasOnProject();

  const balanced = shuffleProjectsRandomly(groups, history);

  // Final hard validation
  Object.entries(balanced).forEach(([project, students]) => {
    students.forEach(s => {
      if (history[s.name]?.[project]) {
        throw new Error(
          `History violation: ${s.name} already did ${project}`
        );
      }
    });
  });

  writeAllGroups(balanced); // your existing writer
  return true;
}
/****
 * Admin close sheet
 */
function adminCloseSheet() {
  try {
    shuffleWithRules();
    archiveClosedSheet(); // your existing function
    return { ok: true, message: "Shuffle & close completed successfully." };
  } catch (e) {
    return { ok: false, message: e.message || e.toString() };
  }
}



