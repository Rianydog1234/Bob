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

