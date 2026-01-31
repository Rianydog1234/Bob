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
  // Replace this with the specific sheet GID you want
  const TARGET_SHEET_GID = 1759753668;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Find the sheet with the specific GID
  const sheet = ss.getSheets().find(s => s.getSheetId() === TARGET_SHEET_GID);

  if (!sheet) {
    throw new Error(
      `Target sheet with GID ${TARGET_SHEET_GID} not found in spreadsheet ${SPREADSHEET_ID}`
    );
  }

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
  const projectMap = getProjectColumnMap();
  const col = projectMap[projectName.trim()];
  if (col === undefined) throw new Error("Invalid project (mapping failed): " + projectName);
  const sheet = getTargetSheet();
  if (!sheet) throw new Error("Target sheet not found");
  const startRow = 3;
  const maxRows = 16;

  const range = sheet.getRange(startRow, col + 1, maxRows, 2);
  const values = range.getValues();

  // Prevent duplicate presence
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
    if (i === values.length - 1) insertRow = startRow + values.length;
  }

  // Clear format on insertion row
  sheet.getRange(insertRow, col + 1, 1, 2).clearFormat();

  // Write name + grade
  sheet.getRange(insertRow, col + 1).setValue(fullName);
  sheet.getRange(insertRow, col + 2).setValue(grade);

  // Gender color
  const genderColor = (gender || "").toString().toLowerCase() === "f" ? "#ff000033"
    : (gender || "").toString().toLowerCase() === "m" ? "#0000ff33" : "#ffffff00";
  sheet.getRange(insertRow, col + 1).setBackground(genderColor);

  // Apply conditional formatting for grade colors
  setGradeConditionalFormatting(sheet, col);

  // Ensure entire column has no manual backgrounds (pre-sort)
  sheet.getRange(startRow, col + 1, maxRows, 2).setBackground(null);
  // Re-write the student entries
  sheet.getRange(insertRow, col + 1).setValue(fullName);
  sheet.getRange(insertRow, col + 2).setValue(grade);
  sheet.getRange(insertRow, col + 1).setBackground(genderColor);

  // Sort by grade (grade is at col+2)
  sheet.getRange(startRow, col + 1, maxRows, 2).sort([{ column: col + 2, ascending: true }]);

  // Reapply gender colors after sort
  applyGenderColors(sheet, col, startRow, maxRows);
  return "OK";
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
  const { grade, gender } = getStudentGradeGender(fullName);
  return writeStudentToProject(fullName, projectName, grade, gender);
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
        const name = r[0];
        if (!name) return;

        const key = normalizeName(name);
        if (!history[key]) history[key] = {};
        history[key][projectName] = true; // record student has been on this project
      });
    });
  });

  return history;
}


/**********************
 * shuffleProjectsRandomly (random swap engine)
 **********************/
function shuffleProjectsRandomly(groups, history) {
  const projectNames = Object.keys(groups);
  for (let pass = 0; pass < MAX_SHUFFLE_PASSES; pass++) {
    let changed = false;
    for (const p1 of projectNames) {
      for (const p2 of projectNames) {
        if (p1 === p2) continue;
        const g1 = groups[p1];
        const g2 = groups[p2];
        const a1 = analyzeGroup(g1);
        const a2 = analyzeGroup(g2);
        if (a1.valid && a2.valid) continue;

        const c1 = [...g1].sort(() => Math.random() - 0.5);
        const c2 = [...g2].sort(() => Math.random() - 0.5);

        let swapped = false;
        for (const s1 of c1) {
          for (const s2 of c2) {
            // Use normalized keys when checking history
            const s1Key = normalizeName(s1.name);
            const s2Key = normalizeName(s2.name);
            if ((history[s1Key] && history[s1Key][p2]) || (history[s2Key] && history[s2Key][p1])) continue;

            const newG1 = g1.filter(s => s !== s1).concat(s2);
            const newG2 = g2.filter(s => s !== s2).concat(s1);
            const na1 = analyzeGroup(newG1);
            const na2 = analyzeGroup(newG2);

            const improves =
              ((!a1.valid || na1.valid) && (!a2.valid || na2.valid)) ||
              (na1.genderDiff < a1.genderDiff) ||
              (na2.genderDiff < a2.genderDiff) ||
              (na1.gradeDiff < a1.gradeDiff) ||
              (na2.gradeDiff < a2.gradeDiff) ||
              (na1.total >= MIN_STUDENTS && na1.total <= MAX_STUDENTS) ||
              (na2.total >= MIN_STUDENTS && na2.total <= MAX_STUDENTS);

            if (improves) {
              groups[p1] = newG1;
              groups[p2] = newG2;
              changed = true;
              swapped = true;
              break;
            }
          }
          if (swapped) break;
        }
      }
    }
    if (!changed) break;
  }

  // Final fixes to meet min/max by moving students if needed
  projectNames.forEach(p => {
    const g = groups[p];
    let ag = analyzeGroup(g);

    if (g.length < MIN_STUDENTS) {
      for (const otherP of projectNames) {
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
      for (const otherP of projectNames) {
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
