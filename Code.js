/**********************
 * CONFIG
 **********************/
const SPREADSHEET_ID = "1dolNsNrpgD0AoWTCBz-cijXvZp5uorFoRYPccKRWpoY";
const SHEET_GIDS = [
  681181988,
  1493058526,
  1220633850
];
const PROJECT_COLUMNS = [1, 3, 5, 7, 9, 11]; // B D F H J L
const TARGET_SPREADSHEET_ID = "1dolNsNrpgD0AoWTCBz-cijXvZp5uorFoRYPccKRWpoY";
const SOURCE_INFO_SHEET_GID = 1446473767;
const GRADE_ORDER = { "2A": 1, "2B": 2, "3A": 3, "3B": 4 };

/**********************
 * CONFIG helpers
 **********************/
function getConfigSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("ScriptConfig");
  if (!sheet) {
    sheet = ss.insertSheet("ScriptConfig");
    sheet.getRange("A1").setValue("TARGET_SHEET_ID");
    sheet.getRange("B1").setValue("");
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
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const s = ss.getSheets()[0];
  return s ? s.getSheetId() : null;
}
function getTargetSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const gid = getTargetSheetId();
  if (!gid) throw new Error("Active target sheet GID not configured.");
  const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
  if (!sheet) throw new Error("Target sheet not found (gid=" + gid + ")");
  return sheet;
}




/**********************
 * PROJECT column mapping (read from sheet1 row 2)
 **********************/
function getProjectColumnMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Skip GID entries that are 0 (placeholders)
  const sheetGid = SHEET_GIDS.find(g => g && g !== 0);
  const sheet = ss.getSheets().find(s => s.getSheetId() === sheetGid);
  if (!sheet) throw new Error("Project definitions sheet not found (check SHEET_GIDS)");
  const data = sheet.getDataRange().getValues();
  const projectRow = data[1] || []; // row 2
  const map = {};
  PROJECT_COLUMNS.forEach(col => {
    const name = projectRow[col];
    if (name) map[name.toString().trim()] = col;
  });
  return map;
}

function getProjectColumnMapForSheet(sheet) {
  const lastCol = sheet.getLastColumn() || (Math.max(...PROJECT_COLUMNS) + 3);
  const row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0] || [];
  const map = {};
  PROJECT_COLUMNS.forEach(col => {
    const name = row2[col];
    if (name) map[String(name).trim()] = col;
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
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}


/**
 * swapNameIfNeeded(name, sheetIndex)
 * Some of your source sheets store "Last First" vs "First Last".
 * When sheetIndex indicates a "First Last" sheet but we want Last First
 * swap if needed. The original code used sheetIndex semantics — we preserve it.
 */
function swapNameIfNeeded(name, sheetIndex) {
  if (!name) return "";
  if (sheetIndex === 1 || sheetIndex === 2) {
    const p = name.trim().split(/\s+/);
    if (p.length < 2) return name;
    return `${p[1]} ${p[0]}`;
  }
  return name;
}

/**********************
 * CONDITIONAL FORMATTING (grade colors)
 **********************/
function setGradeConditionalFormatting(sheet, col) {
  const rules = sheet.getConditionalFormatRules();
  const filteredRules = rules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() === col + 2);
  });
  const startRow = 3;
  const maxRows = 100;
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
 * LOAD STUDENT + PROJECT DATA (reads multiple source sheets)
 **********************/
function getStudentData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const students = {};
  const projects = new Set();

  SHEET_GIDS.forEach((gid, sheetIndex) => {
    if (!gid || gid === 0) return; // skip placeholder
    const sheet = ss.getSheets().find(s => s.getSheetId() === gid);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const projectRow = data[1] || [];
    PROJECT_COLUMNS.forEach(col => {
      const projectName = projectRow[col];
      if (!projectName) return;
      projects.add(projectName);
      for (let r = 2; r < data.length; r++) {
        let fullName = data[r][col];
        if (!fullName) continue;
        fullName = swapNameIfNeeded(fullName, sheetIndex);
        const key = normalizeName(fullName);
        if (!students[key]) students[key] = { fullName, projects: {} };
        students[key].projects[projectName] = true;
      }
    });
  });

  return { students: Object.values(students), projects: Array.from(projects) };
}

/**********************
 * GET STUDENT INFO (light check)
 **********************/
function getStudentInfo(fullName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SOURCE_INFO_SHEET_GID);
  if (!sheet) throw new Error("Info sheet not found");
  const data = sheet.getDataRange().getValues();
  const key = normalizeName(fullName);
  for (let r = 1; r < data.length; r++) {
    const last = data[r][0], first = data[r][1];
    if (!first || !last) continue;
    if (normalizeName(`${last} ${first}`) === key) return { grade: "", gender: "" };
  }
  throw new Error("Student info not found: " + fullName);
}

/**********************
 * GET STUDENT GRADE + GENDER
 * (reads Info sheet columns: Last | First | Grade | Gender)
 **********************/
function getStudentGradeGender(studentName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === SOURCE_INFO_SHEET_GID);
  if (!sheet) {
    throw new Error(
      `‼️ Info sheet not found (expected GID=${SOURCE_INFO_SHEET_GID}).\nCheck that the info sheet exists and its GID hasn’t changed!`
    );
  }
  const data = sheet.getDataRange().getValues();
  const keyNormalized = normalizeName(studentName);
  const matches = [];
  for (let r = 1; r < data.length; r++) {
    const last = data[r][0], first = data[r][1];
    if (!last || !first) continue;
    const combinedLF = normalizeName(`${last} ${first}`);
    const combinedFL = normalizeName(`${first} ${last}`);
    if (combinedLF === keyNormalized || combinedFL === keyNormalized) {
      return { grade: data[r][2], gender: data[r][3] };
    }
    matches.push(combinedLF);
  }
  throw new Error(
    `🛑 Student info lookup failed!\n` +
    `Input name: "${studentName}"\n` +
    `Normalized lookup key: "${keyNormalized}"\n` +
    `Info sheet rows scanned: ${data.length - 1}\n` +
    `Sample matches available (first 10): [${matches.slice(0,10).join(", ")}]\n` +
    `Ensure the name matches exactly (accents, spaces, order).`
  );
}

/**********************
 * WRITE STUDENT TO PROJECT
 **********************/
function writeStudentToProject(fullName, projectName, grade, gender) {
  const target = getTargetSheet();
  const projMap = getProjectColumnMapForSheet(target);
  if (projMap[projectName] === undefined) {
    throw new Error("writeStudentToProject: invalid project '" + projectName + "' in target sheet");
  }

  // Global history check
  const history = buildHistoryMapFromSheets();
  const key = normalizeName(fullName);
  if (history[key]) {
    if (history[key][projectName]) {
      throw new Error(`${fullName} already did ${projectName} previously (history).`);
    }
    throw new Error(`${fullName} is already present in project history; cannot assign to another project.`);
  }

  // Ensure not already in this active target
  const startRow = 3;
  const maxRows = 100;
  const map = getProjectColumnMapForSheet(target);
  for (const [pName, col] of Object.entries(map)) {
    const vals = target.getRange(startRow, col + 1, maxRows, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      const nm = (vals[i][0] || "").toString().trim();
      if (!nm) continue;
      if (normalizeName(nm) === key) {
        throw new Error(`${fullName} is already listed in the active sheet under project "${pName}".`);
      }
    }
  }

  // Find first empty row in this project block and write
  const col = projMap[projectName];
  const block = target.getRange(startRow, col + 1, maxRows, 2).getValues();
  let insertIdx = -1;
  for (let i = 0; i < block.length; i++) {
    if (!block[i][0]) { insertIdx = i; break; }
  }
  if (insertIdx === -1) insertIdx = block.length;
  const writeRow = startRow + insertIdx;
  target.getRange(writeRow, col + 1).setValue(fullName);
  target.getRange(writeRow, col + 2).setValue(grade || "");
  const gLower = (gender || "").toString().toLowerCase();
  const color = gLower === "f" ? "#ff000033" : gLower === "m" ? "#0000ff33" : null;
  if (color) target.getRange(writeRow, col + 1).setBackground(color);

  return { ok: true, message: `Wrote ${fullName} to ${projectName} at row ${writeRow}` };
}


/**********************
 * applyGenderColors
 **********************/
function applyGenderColors(sheet, col, startRow, maxRows) {
  const names = sheet.getRange(startRow, col + 1, maxRows, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    const fullName = names[i][0];
    const cell = sheet.getRange(startRow + i, col + 1);
    if (!fullName) {
      cell.setBackground(null);
      continue;
    }
    let gender = "";
    try {
      gender = (getStudentGradeGender(fullName).gender || "").toString().toLowerCase();
    } catch (e) {
      gender = "";
    }
    const color = gender === "f" ? "#ff000033" : gender === "m" ? "#0000ff33" : null;
    if (color) cell.setBackground(color);
    else cell.setBackground(null);
  }
}

/**********************
 * submitStudentToProject (wrapper)
 **********************/
function submitStudentToProject(fullName, projectName) {
  try {
    if (!fullName || !projectName) throw new Error("Invalid name or project");
    if ((fullName || "").toString().trim().split(/\s+/).filter(Boolean).length < 2) {
      return { ok: false, message: "Please choose a full name (first + last) from the dropdown." };
    }
    const info = getStudentGradeGender(fullName); // throws if not found
    const res = writeStudentToProject(fullName, projectName, info.grade, info.gender);
    return { ok: true, message: res.message };
  } catch (e) {
    return { ok: false, message: e.message || e.toString() };
  }
}


/**********************
 * ReadAllGroups (reads current target sheet project blocks)
 **********************/
function readAllGroups() {
  const sheet = getTargetSheet();
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
        if (!rawNameCell) return null;

        const wordCount = rawNameCell.split(/\s+/).filter(Boolean).length;
        if (wordCount < 2) {
          Logger.log(
            "SKIP single-word name -> Project: '%s', Row: %s, NameCell: '%s', GradeCell: '%s'",
            project, sheetRow, rawNameCell, rawGradeCell
          );
          return null;
        }

        try {
          const info = getStudentGradeGender(rawNameCell);
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
      .filter(Boolean);
  });

  return groups;
}
/**********************
 * CLOSE SHEET LOGIC
 **********************/
 function adminCloseAndCreateNewTarget() {
  try {
    const res = closeCreateNewTarget();
    return { ok: true, message: res.message, log: res.commitLog };
  } catch (e) {
    return { ok: false, message: e.message || e.toString() };
  }
}

 /**********************
 * Helper: getProjectColumnMapForSheet
 * Reads row 2 of the passed sheet and maps projectName -> column index
 **********************/
function getProjectColumnMapForSheet(sheet) {
  const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const map = {};
  PROJECT_COLUMNS.forEach(col => {
    const name = row2[col];
    if (name) map[name.toString().trim()] = col;
  });
  return map;
}

/**********************
 * commitGroupsToSheet(destGid, groups)
 * Append current groups into the destination sheet (project blocks)
 * destGid should be one of the sheets included in SHEET_GIDS (so buildHistoryMapFromSheets sees it)
 **********************/
function commitGroupsToSheet(destGid, groups) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const destSheet = ss.getSheets().find(s => s.getSheetId() === Number(destGid));
  if (!destSheet) throw new Error("commitGroupsToSheet: destination sheet GID not found: " + destGid);

  const startRow = 3;
  const maxRows = 16;
  const projectMap = getProjectColumnMapForSheet(destSheet);
  const log = [];

  Object.entries(groups).forEach(([project, students]) => {
    const col = projectMap[project];
    if (col === undefined) {
      log.push(`WARN: project "${project}" not found in destination sheet (skipped)`);
      return;
    }

    // read the block to find first empty row
    const block = destSheet.getRange(startRow, col + 1, maxRows, 2).getValues();
    let insertIdx = -1;
    for (let i = 0; i < block.length; i++) {
      if (!block[i][0]) { insertIdx = i; break; }
    }
    if (insertIdx === -1) insertIdx = block.length; // will append after block

    // write students one-per-row starting at startRow+insertIdx
    for (let k = 0; k < students.length; k++) {
      const r = startRow + insertIdx + k;
      destSheet.getRange(r, col + 1).setValue(students[k].name || "");
      destSheet.getRange(r, col + 2).setValue(students[k].grade || "");
    }

    log.push(`Committed ${students.length} to "${project}" at col ${col} starting row ${startRow + insertIdx}`);
  });

  Logger.log("commitGroupsToSheet: \n" + log.join("\n"));
  return log;
}

function commitGroupsToSheet(destGid, groups) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const destSheet = ss.getSheets().find(s => s.getSheetId() === Number(destGid));
  if (!destSheet) throw new Error("commitGroupsToSheet: destination sheet not found (gid=" + destGid + ")");

  const startRow = 3;
  const maxRows = 100;
  const projectMap = getProjectColumnMapForSheet(destSheet);
  const log = [];

  Object.entries(groups).forEach(([project, students]) => {
    const col = projectMap[project];
    if (col === undefined) {
      log.push(`WARN: project "${project}" not found in dest sheet (skipped)`);
      return;
    }

    const block = destSheet.getRange(startRow, col + 1, maxRows, 2).getValues();
    let insertIdx = -1;
    for (let i = 0; i < block.length; i++) {
      if (!block[i][0]) { insertIdx = i; break; }
    }
    if (insertIdx === -1) insertIdx = block.length;

    for (let k = 0; k < students.length; k++) {
      const r = startRow + insertIdx + k;
      destSheet.getRange(r, col + 1).setValue(students[k].name || "");
      destSheet.getRange(r, col + 2).setValue(students[k].grade || "");
    }
    log.push(`Committed ${students.length} to "${project}" at col ${col} starting row ${startRow + insertIdx}`);
  });

  Logger.log("commitGroupsToSheet:\n" + log.join("\n"));
  return log;
}


/**********************
 * createNewTargetSheetFromTemplate(newName)
 * Create a fresh sheet, copy project titles row from the project-definition sheet,
 * and clear its project blocks (so admin gets a blank board).
 * Returns the new sheet object.
 **********************/
function createNewTargetSheetFromTemplate(newName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const newSheet = ss.insertSheet(newName);

  const defGid = SHEET_GIDS.find(g => g && g !== 0);
  if (!defGid) throw new Error("createNewTargetSheetFromTemplate: no def sheet in SHEET_GIDS");
  const defSheet = ss.getSheets().find(s => s.getSheetId() === Number(defGid));
  if (!defSheet) throw new Error("createNewTargetSheetFromTemplate: def sheet missing for gid " + defGid);

  const maxCol = Math.max(...PROJECT_COLUMNS) + 3;
  const header = defSheet.getRange(2, 1, 1, maxCol).getValues();
  newSheet.getRange(2, 1, 1, maxCol).setValues(header);

  const startRow = 3;
  const maxRows = 16;
  PROJECT_COLUMNS.forEach(col => {
    newSheet.getRange(startRow, col + 1, maxRows, 2).clearContent().clearFormat();
  });

  return newSheet;
}


/**********************
 * closeCreateNewTarget()
 *
 * - reads current assignments
 * - commits them into the first configured SHEET_GIDS sheet (so dynamic history picks them up)
 * - creates a new blank sheet and sets it as active target 
 * - clears the new sheet's project blocks so the admin has a fresh board
 */
function closeCreateNewTarget() {
  Logger.log("closeCreateNewTarget starting at " + new Date().toISOString());

  const groups = readAllGroups();
  const totalAssigned = Object.values(groups).reduce((acc, arr) => acc + (arr ? arr.length : 0), 0);
  if (totalAssigned === 0) throw new Error("closeCreateNewTarget: no assignments found in current target (aborting).");

  const destGid = SHEET_GIDS.find(g => g && g !== 0);
  if (!destGid) throw new Error("closeCreateNewTarget: no valid dest in SHEET_GIDS.");

  const commitLog = commitGroupsToSheet(destGid, groups);

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const newName = `Target_${ts}`;
  const newSheet = createNewTargetSheetFromTemplate(newName);

  setTargetSheetId(newSheet.getSheetId());
  Logger.log("closeCreateNewTarget: set new target gid=" + newSheet.getSheetId());

  Logger.log("closeCreateNewTarget completed. commitLog:\n" + commitLog.join("\n"));
  return { ok: true, message: "Closed and created new target " + newName, commitLog };
}


/**********************
 * writeBalancedGroups (writes groups back to target sheet)
 **********************/
function writeBalancedGroups(groups) {
  const sheet = getTargetSheet();
  const map = getProjectColumnMap();
  const startRow = 3;
  const maxRows = 16;

  // Helper: parse grade into sortable value
  function gradeSortKey(grade) {
    if (!grade) return Number.MAX_SAFE_INTEGER;
    grade = grade.toString().trim();
    // Prefer explicit mapping if present
    if (typeof GRADE_ORDER !== "undefined" && GRADE_ORDER[grade] !== undefined) {
      return GRADE_ORDER[grade] * 1000; // keep room for letter ordering
    }
    // Fallback: parse "number + letter" like 2A, 10B
    const m = grade.match(/^(\d+)\s*([A-Za-z])?$/);
    if (m) {
      const num = Number(m[1]);
      const letter = (m[2] || "A").toUpperCase();
      return num * 1000 + (letter.charCodeAt(0) - 64);
    }
    // Last fallback: sort lexicographically at the end
    return Number.MAX_SAFE_INTEGER - 1;
  }

  Object.entries(groups).forEach(([project, students]) => {
    const col = map[project];
    if (col === undefined) {
      Logger.log("writeBalancedGroups: unknown project mapping for '%s' - skipping", project);
      return;
    }

    // Clear old data
    sheet.getRange(startRow, col + 1, maxRows, 2).clearContent();

    // Write unsorted students into the block (so rows exist)
    students.slice(0, maxRows).forEach((s, i) => {
      sheet.getRange(startRow + i, col + 1).setValue(s.name);
      sheet.getRange(startRow + i, col + 2).setValue(s.grade);
    });

    // Read the full block (so we can sort the non-empty rows)
    const blockRange = sheet.getRange(startRow, col + 1, maxRows, 2);
    const blockValues = blockRange.getValues(); // array of [name, grade]

    // Separate non-empty rows (we consider a row non-empty if name cell has content)
    const populated = [];
    const emptyRowsCount = [];
    for (let i = 0; i < blockValues.length; i++) {
      const row = blockValues[i];
      const name = (row[0] || "").toString().trim();
      const grade = (row[1] || "").toString().trim();
      if (name) {
        populated.push({ name, grade, origIndex: i });
      } else {
        emptyRowsCount.push(i);
      }
    }

    // Sort populated rows by the grade order (then by name for stability)
    populated.sort((a, b) => {
      const ka = gradeSortKey(a.grade);
      const kb = gradeSortKey(b.grade);
      if (ka !== kb) return ka - kb;
      // fallback alphabetical by normalized name
      return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
    });

    // Rebuild blockValues: first fill with sorted populated rows, then blanks
    const newBlock = [];
    for (let i = 0; i < populated.length; i++) {
      newBlock.push([populated[i].name, populated[i].grade]);
    }
    // Fill remaining slots with empty rows up to maxRows
    while (newBlock.length < maxRows) newBlock.push(["", ""]);

    // Write sorted block back into sheet
    blockRange.setValues(newBlock);

    // Reapply conditional formatting for grade colors
    setGradeConditionalFormatting(sheet, col);
  });

  // Reapply gender colors after writing all blocks
  Object.entries(map).forEach(([project, col]) => {
    applyGenderColors(sheet, col, startRow, maxRows);
  });
}


/**********************
 * Constants for shuffle rules
 **********************/
const MAX_STUDENTS = 16;
const MIN_STUDENTS = 15;
const MAX_GENDER_DIFF = 2;
const MAX_GRADE_DIFF = 2;
const MAX_SHUFFLE_PASSES = 20;

/**********************
 * analyzeGroup
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
    girls, boys,
    g2, g3,
    genderDiff: Math.abs(girls - boys),
    gradeDiff: Math.abs(g3 - g2),
    valid: (group.length >= MIN_STUDENTS &&
            group.length <= MAX_STUDENTS &&
            Math.abs(girls - boys) <= MAX_GENDER_DIFF &&
            Math.abs(g3 - g2) <= MAX_GRADE_DIFF)
  };
}

/**********************
 * getWasOnProjectSheet (IMPLEMENTED)
 * Finds a sheet named 'WasOnProject' (or similar) in the source SPREADSHEET_ID
 **********************/
function getWasOnProjectSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Try common names
  const candidates = ['WasOnProject', 'Was On Project', 'WasOnProjects', 'WasOnProjectList', 'WasOn'];
  for (const name of candidates) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  // Fallback: if a sheet named 'History' or 'WasOn' exists
  const alt = ss.getSheetByName('History') || ss.getSheetByName('was_on_project');
  if (alt) return alt;
  // As a last resort, search for a sheet with at least two columns with header 'Name' or 'Project'
  const sheets = ss.getSheets();
  for (const s of sheets) {
    const hdr = s.getRange(1,1,1,5).getValues()[0].map(c => (c||"").toString().toLowerCase());
    if (hdr.includes('name') && hdr.includes('project')) return s;
  }
  throw new Error("WasOnProject sheet not found. Create a sheet named 'WasOnProject' with columns [Name, Project].");
}

/**********************
 * buildHistoryMapFromSheets
 **********************/
/**********************
 * buildHistoryMapFromSheets (REPLACEMENT)
 * Build history keys for BOTH "Last First" and "First Last" forms so
 * name-ordering differences in sheets won't break history checks.
 **********************/
function buildHistoryMapFromSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const history = {};

  SHEET_GIDS.forEach((gid, sheetIndex) => {
    if (!gid || gid === 0) return; // skip placeholder
    const sheet = ss.getSheets().find(s => s.getSheetId() === gid);
    if (!sheet) return;

    PROJECT_COLUMNS.forEach(col => {
      const projectName = sheet.getRange(2, col + 1).getValue(); // row 2 = project names
      if (!projectName) return;

      const values = sheet.getRange(3, col + 1, 16, 1).getValues(); // rows 3-18
      values.forEach(r => {
        const name = (r[0] || "").toString().trim();
        if (!name) return;

        // Primary normalized key (as-is order)
        const key1 = normalizeName(name);

        // Also create a swapped-first-last variant when possible
        let key2 = null;
        const parts = name.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
          // swap first two tokens only (covers Last First vs First Last)
          key2 = normalizeName(parts.slice(1).join(" ") + " " + parts[0]);
        }

        // Record both variants in history so lookups are order-agnostic
        if (!history[key1]) history[key1] = {};
        history[key1][projectName] = true;

        if (key2) {
          if (!history[key2]) history[key2] = {};
          history[key2][projectName] = true;
        }
      });
    });
  });

  return history;
}



/**********************
 * shuffleProjectsRandomly (random swap engine)
 **********************/
/**********************
 * shuffleProjectsRandomly (REPLACEMENT)
 * More conservative: only perform swaps that:
 *  - make both groups valid, OR
 *  - reduce a combined penalty by a configurable margin.
 * Also prevents swapping the same student multiple times per pass.
 **********************/
function shuffleProjectsRandomly(groups, history) {
  const projectNames = Object.keys(groups);

  // penalty weights - tune if desired
  const WEIGHT_GENDER = 2;
  const WEIGHT_GRADE = 3;
  const WEIGHT_SIZE_PENALTY = 5; // penalize being outside min/max
  const MIN_IMPROVEMENT = 2; // require at least this much reduction in combined penalty

  function penaltyForGroup(g) {
    const a = analyzeGroup(g);
    // size penalty: if inside bounds => 0, else proportional to distance
    const sizePenalty = a.total < MIN_STUDENTS ? (MIN_STUDENTS - a.total) : (a.total > MAX_STUDENTS ? (a.total - MAX_STUDENTS) : 0);
    return a.genderDiff * WEIGHT_GENDER + a.gradeDiff * WEIGHT_GRADE + sizePenalty * WEIGHT_SIZE_PENALTY;
  }

  for (let pass = 0; pass < MAX_SHUFFLE_PASSES; pass++) {
    let changed = false;
    const swappedThisPass = new Set(); // normalized student keys swapped already this pass

    for (let i = 0; i < projectNames.length; i++) {
      for (let j = i + 1; j < projectNames.length; j++) { // only pairs once (p1,p2)
        const p1 = projectNames[i];
        const p2 = projectNames[j];
        const g1 = groups[p1];
        const g2 = groups[p2];
        const a1 = analyzeGroup(g1);
        const a2 = analyzeGroup(g2);

        // If both already valid, skip pair
        if (a1.valid && a2.valid) continue;

        // Shuffle shallow copies for candidate selection
        const c1 = [...g1].sort(() => Math.random() - 0.5);
        const c2 = [...g2].sort(() => Math.random() - 0.5);

        const basePenalty = penaltyForGroup(g1) + penaltyForGroup(g2);

        let madeSwap = false;

        // Try limited number of candidate pairs (avoid O(n^2) explosion)
        const MAX_CANDIDATE_TRIES = Math.min(8, c1.length * c2.length);

        let tries = 0;
        outer:
        for (const s1 of c1) {
          for (const s2 of c2) {
            if (tries++ > MAX_CANDIDATE_TRIES) break outer;

            const s1Key = normalizeName(s1.name);
            const s2Key = normalizeName(s2.name);

            // don't re-swap the same student multiple times this pass
            if (swappedThisPass.has(s1Key) || swappedThisPass.has(s2Key)) continue;

            // history check: don't put someone on a project they already did
            if ((history[s1Key] && history[s1Key][p2]) || (history[s2Key] && history[s2Key][p1])) continue;

            // prepare candidate new groups
            const newG1 = g1.filter(s => s !== s1).concat(s2);
            const newG2 = g2.filter(s => s !== s2).concat(s1);

            const na1 = analyzeGroup(newG1);
            const na2 = analyzeGroup(newG2);

            const newPenalty = penaltyForGroup(newG1) + penaltyForGroup(newG2);

            // Accept the swap if:
            //  - it makes BOTH groups valid, OR
            //  - it reduces combined penalty by MIN_IMPROVEMENT (conservative)
            const bothValid = na1.valid && na2.valid;
            const penaltyDrop = basePenalty - newPenalty;

            if (bothValid || penaltyDrop >= MIN_IMPROVEMENT) {
              // apply swap
              groups[p1] = newG1;
              groups[p2] = newG2;
              swappedThisPass.add(s1Key);
              swappedThisPass.add(s2Key);
              changed = true;
              madeSwap = true;
              break outer;
            }
            // otherwise, skip this candidate pair
          }
        }

        // continue to next pair
        if (madeSwap) continue;
      }
    }

    if (!changed) break;
  }

  // Final balancing for min/max sizes (use same conservative history checks)
  const projectNamesFinal = Object.keys(groups);
  projectNamesFinal.forEach(p => {
    const g = groups[p];
    let ag = analyzeGroup(g);

    if (g.length < MIN_STUDENTS) {
      for (const otherP of projectNamesFinal) {
        if (otherP === p) continue;
        const donorGroup = groups[otherP];
        for (const s of [...donorGroup]) {
          const sKey = normalizeName(s.name);
          if (!history[sKey] || !history[sKey][p]) {
            g.push(s);
            donorGroup.splice(donorGroup.indexOf(s), 1);
            ag = analyzeGroup(g);
            if (ag.total >= MIN_STUDENTS) break;
          }
        }
        if (ag.total >= MIN_STUDENTS) break;
      }
    }

    if (g.length > MAX_STUDENTS) {
      for (const otherP of projectNamesFinal) {
        if (otherP === p) continue;
        const targetGroup = groups[otherP];
        for (const s of [...g]) {
          const sKey = normalizeName(s.name);
          if ((!history[sKey] || !history[sKey][otherP]) && targetGroup.length < MAX_STUDENTS) {
            targetGroup.push(s);
            g.splice(g.indexOf(s), 1);
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
 * shuffleWithRules
 **********************/
function shuffleWithRules() {
  const groups = readAllGroups();
  const history = buildHistoryMapFromSheets();
  const balanced = shuffleProjectsRandomly(groups, history);

  // HARD VALIDATION: ensure no history repeats (use normalized keys)
  Object.entries(balanced).forEach(([project, students]) => {
    students.forEach(s => {
      const key = normalizeName(s.name);
      if (history[key] && history[key][project]) {
        throw new Error(`History violation: ${s.name} already did ${project}`);
      }
    });
  });

  writeBalancedGroups(balanced);
  return true;
}

/**********************
 * archiveClosedSheet (safe implementation)
 * Copies the current target sheet into a new sheet named 'Archive_<timestamp>'
 **********************/
function archiveClosedSheet() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const target = getTargetSheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const archiveName = `Archive_${timestamp}`;
    const copy = target.copyTo(ss).setName(archiveName);
    // Optionally remove extra rows/cols — keep as snapshot
    Logger.log("Archived target sheet as: " + archiveName);
    return archiveName;
  } catch (e) {
    Logger.log("archiveClosedSheet failed: " + (e.message || e.toString()));
    // Do not block closing if archiving fails; just warn
    return null;
  }
}

/**********************
 * ADMIN / ENTRY wrappers
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
