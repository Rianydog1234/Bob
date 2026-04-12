/***************************************************************
 * Code.gs - Fix: restored swap-name logic + full backend + debug + doGet
 *
 * Paste this entire file into Apps Script (replace existing).
 * Set SPREADSHEET_ID = "" to use the active spreadsheet you're editing in the browser.
 **************************************************************/

/* ============= CONFIG ============= */
const SPREADSHEET_ID = ""; // "" => use active spreadsheet; otherwise set explicit id
const PROJECT_COLUMNS = [1, 3, 5, 7, 9, 11]; // zero-based offsets where project names appear in row 2
const SOURCE_INFO_SHEET_GID = 1446473767;
const MAX_STUDENTS = 16;
const GRADE_ORDER = {
  "2A":1,
  "2B":2,
  "3A":3,
  "3B":4
};
/* GET SHEET GIDS*/

function getSheetGids() {
  const cfg = getConfigSheet();
  const vals = cfg.getRange("B1:B").getValues().map(r => r[0]);
  const gids = [];
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (v === undefined || v === null || v === "") break;
    const n = Number(v);
    if (!isNaN(n) && n !== 0) gids.push(n);
  }
  return gids;
}

function getCommitDestinationGid() {
  const gids = getSheetGids();
  const active = getTargetSheetId();
  for (let i = 0; i < gids.length; i++) {
    if (Number(gids[i]) !== Number(active)) return Number(gids[i]);
  }
  return null;
}

function appendActiveToFirstEmptyB(gid) {
  if (!gid) return;
  const cfg = getConfigSheet();
  const vals = cfg.getRange("B1:B").getValues().map(r => r[0]);
  let emptyIdx = -1;
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (v === undefined || v === null || v === "") { emptyIdx = i; break; }
  }
  if (emptyIdx === -1) {
    // append after last read value
    const appendRow = vals.length + 1;
    cfg.getRange(appendRow, 2).setValue(gid);
  } else {
    cfg.getRange(emptyIdx + 1, 2).setValue(gid);
  }
}

/* ADMIN SIGN IN  */
const ADMIN_EMAILS = [
  "biskumar1@duhovkagymnazium.cz",
  "martina.coufalova@duhovkagymnazium.cz",
  "Zdena.Tejkalova@duhovkagymnazium.cz"
];

function isAdmin() {
  const email = Session.getActiveUser().getEmail();
  return ADMIN_EMAILS.includes(email);
}

function checkAdminAccess() {
  return isAdmin();
}
/*SWAP NAMES --*/
function swapNameIfNeeded(name, sheetIdentifier) {
  if (!name) return "";
  const raw = String(name).trim();

  // handle "Last, First" formats
  if (raw.indexOf(",") !== -1) {
    const parts = raw.split(",").map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2) return (parts[1] + " " + parts[0]).trim();
    return raw;
  }

  const PROBLEMATIC_GIDS = [1493058526, 1220633850];

  // Determine whether to swap
  let shouldSwap = false;
  if (sheetIdentifier != null) {
    const n = Number(sheetIdentifier);
    if (!isNaN(n) && (PROBLEMATIC_GIDS.indexOf(n) !== -1 || n === 1 || n === 2)) {
      shouldSwap = true;
    }
  }

  if (shouldSwap) {
    const parts = raw.split(/\s+/).filter(Boolean);
    if (parts.length < 2) return raw;
    return (parts.slice(1).join(" ") + " " + parts[0]).trim();
  }

  return raw;
}
/* ============= UTILITIES ============= */
function getSpreadsheet() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && String(SPREADSHEET_ID).trim()) {
    try { return SpreadsheetApp.openById(String(SPREADSHEET_ID).trim()); }
    catch (e) { Logger.log("getSpreadsheet: openById failed, falling back to active: " + e); }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function normalizeName(name) {
  if (!name) return "";
  return name.toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/* ============= ScriptConfig helpers ============= */
function getConfigSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("ScriptConfig");
  if (!sheet) {
    sheet = ss.insertSheet("ScriptConfig");
    sheet.getRange(1,1,1,2).setValues([["TARGET_SHEET_ID",""]]);
  }
  return sheet;
}
function setTargetSheetId(gid) {
  const cfg = getConfigSheet();
  cfg.getRange(1,2).setValue(gid);
}
function getTargetSheetId() {
  const cfg = getConfigSheet();
  const v = cfg.getRange(1,2).getValue();
  if (v) return Number(v);
  // fallback to first non-config sheet
  const ss = getSpreadsheet();
  const s = ss.getSheets().find(sh => sh.getName() !== "ScriptConfig" && sh.getName() !== "WasOnProject");
  return s ? s.getSheetId() : null;
}
function getTargetSheet() {
  const ss = getSpreadsheet();
  const gid = getTargetSheetId();
  if (!gid) throw new Error("Active target sheet GID not configured in ScriptConfig! (B1)");
  const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
  if (!sheet) throw new Error("Target sheet not found for gid=" + gid);
  return sheet;
}

/* ============= Project Column Mapping ============= */
function getProjectColumnMap() {
  const ss = getSpreadsheet();
  const defGid = getSheetGids().find(g => g && g !== 0);
  if (!defGid) throw new Error("No definition GID configured in getSheetGids()");
  const defSheet = ss.getSheets().find(s => s.getSheetId() === Number(defGid));
  if (!defSheet) throw new Error("Definition sheet not found (check getSheetGids())");
  const row = defSheet.getRange(2, 1, 1, Math.max(...PROJECT_COLUMNS) + 3).getValues()[0] || [];
  const map = {};
  PROJECT_COLUMNS.forEach(col => {
    const name = row[col];
    if (name) map[String(name).trim()] = col;
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

/* ============= Student grade/gender lookup ============= */
function getStudentGradeGender(studentName) {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  let infoSheet = sheets.find(s => s.getSheetId && s.getSheetId() === Number(SOURCE_INFO_SHEET_GID));
  if (!infoSheet) throw new Error(`Info sheet not found (expected gid=${SOURCE_INFO_SHEET_GID})`);
  const data = infoSheet.getDataRange().getValues();
  const key = normalizeName(studentName);
  const sample = [];
  for (let r = 1; r < data.length; r++) {
    const last = data[r][0], first = data[r][1];
    if (!first && !last) continue;
    const combinedLF = normalizeName(`${last} ${first}`);
    const combinedFL = normalizeName(`${first} ${last}`);
    if (combinedLF === key || combinedFL === key) {
      return { grade: data[r][2], gender: data[r][3] };
    }
    if (combinedLF) sample.push(combinedLF);
  }
  throw new Error(`Student grade/gender lookup failed for "${studentName}". Sample keys: ${sample.slice(0,10).join(", ")}`);
}

/* ============= History builders (multi-sheet) ============= */
function buildHistoryMapFromSheets() {
  const ss = getSpreadsheet();
  const history = {};

  // 1) build the list of gids to check:
  // start with configured list (filter out falsy / 0 entries)
  const configured = Array.isArray(getSheetGids()) ? getSheetGids().filter(g => g && Number(g) !== 0).map(Number) : [];

  // add target sheet gid (if configured in ScriptConfig and not already in the list)
  let combined = configured.slice(); // copy
  try {
    const targetGid = getTargetSheetId();
    if (targetGid && !combined.includes(Number(targetGid))) {
      combined.push(Number(targetGid));
    }
  } catch (e) {
    // ignore errors resolving target gid (we'll proceed with configured list)
    Logger.log("buildHistoryMapFromSheets: getTargetSheetId() failed: " + String(e));
  }

  // --- OPTIONAL: if you prefer to scan *all* sheets in the spreadsheet automatically,
  // uncomment the block below. This will replace the combined list with all sheet GIDs
  // except ScriptConfig and WasOnProject. (Be careful: this may include sheets you don't want scanned.)
  /*
  try {
    const allGids = ss.getSheets()
      .filter(sh => {
        const n = (sh.getName()||"").toString().toLowerCase();
        return n !== "scriptconfig" && !/wasonproject|was on project|history|archive/i.test(n);
      })
      .map(sh => sh.getSheetId());
    combined = Array.from(new Set(allGids)); // unique
  } catch (e) {
    Logger.log("buildHistoryMapFromSheets: fallback reading all sheets failed: " + String(e));
  }
  */

  // 2) iterate the combined gid list and gather names
  combined.forEach((gid, sheetIndex) => {
    if (!gid) return;
    const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
    if (!sheet) {
      Logger.log("buildHistoryMapFromSheets: skipping missing sheet gid=" + gid);
      return;
    }

    // For each project column, read header and MAX_STUDENTS student rows (same as original)
    PROJECT_COLUMNS.forEach(col => {
      // Project column indexes in your code are zero-based offsets; when reading ranges we use col+1
      let projectName;
      try {
        projectName = sheet.getRange(2, col + 1).getValue();
      } catch (e) {
        projectName = null;
      }
      if (!projectName) return;
      let values = [];
      try {
        values = sheet.getRange(3, col + 1, MAX_STUDENTS, 1).getValues();
      } catch (e) {
        values = [];
      }

      values.forEach(r => {
        let name = (r[0] || "").toString().trim();
        if (!name) return;
        // apply sheet-specific swap heuristic
        name = swapNameIfNeeded(name, gid);
        const key1 = normalizeName(name);
        history[key1] = history[key1] || {};
        history[key1][projectName] = true;

        // also store reversed variant for safety (same behaviour as original)
        const parts = name.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
          const swapped = normalizeName(parts.slice(1).join(" ") + " " + parts[0]);
          history[swapped] = history[swapped] || {};
          history[swapped][projectName] = true;
        }
      });
    });
  });

  return history;
}


/* ============= WasOnProject helper ============= */
function getWasOnProjectSheet() {
  const ss = getSpreadsheet();
  const candidates = ['WasOnProject', 'Was On Project', 'WasOnProjects', 'WasOnProjectList', 'WasOn'];
  for (const name of candidates) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  const alt = ss.getSheetByName('History') || ss.getSheetByName('was_on_project');
  if (alt) return alt;
  // not found
  throw new Error("WasOnProject sheet not found. Create a sheet named 'WasOnProject' or update code.");
}

/* ============= Write student to project (fixed + swap-aware) ============= */
function writeStudentToProject(fullName, projectName, grade, gender) {
  const target = getTargetSheet();
  const projMap = getProjectColumnMapForSheet(target);
  if (projMap[projectName] === undefined) throw new Error("writeStudentToProject: project not found on active target: " + projectName);

  // history-per-project check
  const history = buildHistoryMapFromSheets();
  const key = normalizeName(fullName);
  if (history[key] && history[key][projectName]) throw new Error(`${fullName} already did ${projectName} previously (history).`);

  // ensure not already in active target (scan each project column)
  const startRow = 3; const maxRows = MAX_STUDENTS;
  for (const [pName, col] of Object.entries(projMap)) {
    const vals = target.getRange(startRow, col + 1, maxRows, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      const nmRaw = (vals[i][0] || "").toString().trim();
      if (!nmRaw) continue;
      // apply swap heuristic for reading from this sheet (but here sheetIndex unknown; use swapNameIfNeeded with undefined to not swap)
      const nm = swapNameIfNeeded(nmRaw, undefined);
      if (normalizeName(nm) === key) throw new Error(`${fullName} is already listed in the active sheet under project "${pName}".`);
    }
  }

  // capacity check
  const col = projMap[projectName];
  const block = target.getRange(startRow, col + 1, MAX_STUDENTS, 2).getValues();
  const countAssigned = block.reduce((acc, r) => acc + ((r[0] || "").toString().trim() ? 1 : 0), 0);
  if (countAssigned >= MAX_STUDENTS) throw new Error(`${projectName} is full (${MAX_STUDENTS} students).`);

  // find first empty slot
  let insertIdx = -1;
  for (let i = 0; i < block.length; i++) { if (!block[i][0] || !block[i][0].toString().trim()) { insertIdx = i; break; } }
  if (insertIdx === -1) insertIdx = block.length;
  const writeRow = startRow + insertIdx;

  target.getRange(writeRow, col + 1).setValue(fullName);
  target.getRange(writeRow, col + 2).setValue(grade || "");

  // gender coloring (same as before)
  const gLower = (gender || "").toString().toLowerCase();
  if (gLower) {
    if (gLower.indexOf("f") === 0) target.getRange(writeRow, col + 1).setBackground("#ffebee");
    else if (gLower.indexOf("m") === 0) target.getRange(writeRow, col + 1).setBackground("#e8f0fe");
  }

  // grade coloring (NEW)
  const gradeUpper = (grade || "").toString().toUpperCase();
  const gradeCell = target.getRange(writeRow, col + 2);

  if (gradeUpper === "2A") gradeCell.setBackground("#b6d7a8");
  else if (gradeUpper === "2B") gradeCell.setBackground("#a4c2f4");
  else if (gradeUpper === "3A") gradeCell.setBackground("#ffe599");
  else if (gradeUpper === "3B") gradeCell.setBackground("#ea9999");
  try { sortProjectBlock(target, col); } catch (e) { Logger.log("sortProjectBlock failed: " + e.toString()); }
  return { ok: true, message: `Wrote ${fullName} to ${projectName} at row ${writeRow}` };
}

/* ============= API for UI ============= */
function getStudentData() {
  // returns { students: [...], projects: [...] } as front-end expects
  const ss = getSpreadsheet();
  const students = {};
  const projects = new Set();
  getSheetGids().forEach((gid, sheetIndex) => {
    if (!gid || gid === 0) return;
    const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
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
        // apply possible sheet-specific swap heuristic
        fullName = swapNameIfNeeded(fullName, gid);
        const key = normalizeName(fullName);
        if (!students[key]) students[key] = { fullName, projects: {} };
        students[key].projects[projectName] = true;
      }
    });
  });
  return { students: Object.values(students), projects: Array.from(projects) };
}

/* front-end calls this */
function submitStudentToProject(fullName, projectName) {
  try {
    if (!fullName || !projectName) throw new Error("Invalid name or project.");
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

/* ============= Read / Commit Groups ============= */
function readAllGroups() {
  const sheet = getTargetSheet();
  const startRow = 3; const maxRows = MAX_STUDENTS;
  const groups = {};
  const map = getProjectColumnMapForSheet(sheet);
  Object.entries(map).forEach(([project, col]) => {
    const values = sheet.getRange(startRow, col + 1, maxRows, 2).getValues();
    groups[project] = values
      .map((r, index) => {
        const rawNameCell = (r[0] || "").toString().trim();
        const rawGradeCell = (r[1] || "").toString().trim();
        if (!rawNameCell) return null;
        const parts = rawNameCell.split(/\s+/).filter(Boolean);
        if (parts.length < 2) {
          Logger.log("READ skip single-word name -> Project: %s, row %s, cell:'%s'", project, startRow + index, rawNameCell);
          return null;
        }
        try {
          const info = getStudentGradeGender(rawNameCell);
          return { name: rawNameCell, gender: (info.gender || "").toString().toLowerCase(), grade: info.grade || rawGradeCell || "" };
        } catch (e) {
          Logger.log("READ group lookup failed for %s: %s", rawNameCell, e.toString());
          return { name: rawNameCell, gender: "", grade: rawGradeCell || "" };
        }
      })
      .filter(Boolean);
  });
  return groups;
}

function commitGroupsToSheet(destGid, groups) {
  const ss = getSpreadsheet();
  const destSheet = ss.getSheets().find(s => s.getSheetId() === Number(destGid));
  if (!destSheet) throw new Error("commitGroupsToSheet: destination sheet not found (gid=" + destGid + ")");
  const startRow = 3; const maxRows = 100;
  const projectMap = getProjectColumnMapForSheet(destSheet);
  const log = [];
  Object.entries(groups).forEach(([project, students]) => {
    const col = projectMap[project];
    if (col === undefined) { log.push(`WARN: project "${project}" not found in dest sheet (skipped)`); return; }
    const block = destSheet.getRange(startRow, col + 1, maxRows, 2).getValues();
    let insertIdx = -1;
    for (let i = 0; i < block.length; i++) { if (!block[i][0]) { insertIdx = i; break; } }
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

/* ============= Create new target ============= */
function createNewTargetSheetFromTemplate(newName) {
  const ss = getSpreadsheet();
  const newSheet = ss.insertSheet(newName);
  const defGid = getSheetGids().find(g => g && g !== 0);
  if (!defGid) throw new Error("createNewTargetSheetFromTemplate: no def sheet in getSheetGids()");
  const defSheet = ss.getSheets().find(s => s.getSheetId() === Number(defGid));
  if (!defSheet) throw new Error("createNewTargetSheetFromTemplate: def sheet missing for gid " + defGid);
  const maxCol = Math.max(...PROJECT_COLUMNS) + 3;
  const header = defSheet.getRange(2, 1, 1, maxCol).getValues();
  newSheet.getRange(2, 1, 1, maxCol).setValues(header);
  const startRow = 3; const maxRows = MAX_STUDENTS;
  PROJECT_COLUMNS.forEach(col => {
    newSheet.getRange(startRow, col + 1, maxRows, 2).clearContent().clearFormat();
  });
  PROJECT_COLUMNS.forEach(col => {
  try { setGradeConditionalFormatting(newSheet, col); } catch (e) { Logger.log("createNewTarget CF: " + e.toString()); }
});
  return newSheet;
}

function closeCreateNewTarget() {
  // capture old active BEFORE we create the new sheet
  const oldActive = getTargetSheetId();

  const groups = readAllGroups();
  const totalAssigned = Object.values(groups).reduce((acc, arr) => acc + (arr ? arr.length : 0), 0);
  if (totalAssigned === 0) throw new Error("closeCreateNewTarget: no assignments found in current target (aborting).");

  const destGid = getCommitDestinationGid();
  if (!destGid) throw new Error("closeCreateNewTarget: no valid dest in getSheetGids().");

  // store commit log
  let commitLog = [];
  try {
    commitLog = commitGroupsToSheet(destGid, groups) || [];
  } catch (e) {
    Logger.log("closeCreateNewTarget: commitGroupsToSheet failed: " + e.toString());
    // depending on your desired behavior you can re-throw or continue — here we continue
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const newName = `Target_${ts}`;
  const newSheet = createNewTargetSheetFromTemplate(newName);
  setTargetSheetId(newSheet.getSheetId());

  // Move old active into ScriptConfig B column so the rotation/history is preserved
  try {
    appendActiveToFirstEmptyB(oldActive);
  } catch (e) {
    Logger.log("appendActiveToFirstEmptyB failed: " + e.toString());
  }

  return { ok: true, message: "Closed and created new target " + newName, commitLog };
}

/* ============= Balanced write & formatting ============= */
/**
 * Replace existing setGradeConditionalFormatting in THEBLOB.js with this.
 * - sheet: Google Sheet object
 * - col: zero-based index from PROJECT_COLUMNS (same mapping used elsewhere)
 *
 * This version uses the same whenTextEqualTo rules as Code.js (works nicely for exact text).
 */
// REPLACE: existing setGradeConditionalFormatting with this
function setGradeConditionalFormatting(sheet, col) {
  // defensive: col is the zero-based PROJECT_COLUMNS index
  if (col === undefined || col === null) return;

  const rules = sheet.getConditionalFormatRules() || [];
  const gradeCol = col + 2; // 1-based grade column index

  // remove any existing rules that target/overlap this grade column
  const filteredRules = rules.filter(rule => {
    const ranges = rule.getRanges() || [];
    return !ranges.some(r => {
      const start = r.getColumn();
      const end = r.getLastColumn();
      return gradeCol >= start && gradeCol <= end;
    });
  });

  const startRow = 3;
  const maxRows = 100; // you can tighten to MAX_STUDENTS if you prefer
  const gradeRange = sheet.getRange(startRow, gradeCol, maxRows, 1);

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

  sheet.setConditionalFormatRules(filteredRules.concat(newRules));
}

function gradeKey(g){
  if(!g) return 999;
  g = g.toString().trim().toUpperCase();
  return GRADE_ORDER[g] || 999;
}
/**
 * Sort the project block (name + grade) in-place for the provided sheet/col.
 * - sheet: Sheet object
 * - col: zero-based index from PROJECT_COLUMNS (same mapping used elsewhere)
 *
 * Uses gradeKey() already defined in THEBLOB.js for ordering.
 */

/**
 * Sort the project block (name + grade) in-place and preserve cell backgrounds.
 * - sheet: Sheet object
 * - col: zero-based index from PROJECT_COLUMNS (same mapping used elsewhere)
 *
 * Behavior:
 * - Reads the block values and backgrounds.
 * - Builds queues of backgrounds keyed by student name (handles duplicate names).
 * - Sorts rows by gradeKey() then name.
 * - Writes sorted values back and reassigns backgrounds so colours move with students.
 */
function sortProjectBlock(sheet, col) {
  if (col === undefined || col === null) return;

  var startRow = 3;
  var maxRows = (typeof MAX_STUDENTS !== "undefined" && MAX_STUDENTS) ? MAX_STUDENTS : 100;

  var range = sheet.getRange(startRow, col + 1, maxRows, 2); // name, grade
  var values = range.getValues();           // [[name,grade],...]
  var bgs = range.getBackgrounds();        // [[nameBg,gradeBg],...]

  // Build list of populated rows
  var populated = [];
  for (var i = 0; i < values.length; i++) {
    var name = (values[i][0] || "").toString().trim();
    var grade = (values[i][1] || "").toString().trim();
    if (name) populated.push({ name: name, grade: grade });
  }

  // If nothing to sort, still ensure backgrounds for blanks are cleared (optional)
  if (populated.length === 0) {
    // clear any stray backgrounds in this block (optional)
    try {
      sheet.getRange(startRow, col + 1, maxRows, 2).setBackgrounds(
        (function() { var arr = []; for (var k=0;k<maxRows;k++) arr.push(["",""]); return arr; })()
      );
    } catch (e) { /* ignore */ }
    return;
  }

  // Build background queues keyed by name from the original block (preserve occurrence order)
  var nameBgQueue = {};   // name -> [bg1, bg2, ...]
  var gradeBgQueue = {};  // name -> [bg1, bg2, ...]
  for (var j = 0; j < values.length; j++) {
    var origName = (values[j][0] || "").toString().trim();
    if (!origName) continue;
    if (!nameBgQueue[origName]) nameBgQueue[origName] = [];
    if (!gradeBgQueue[origName]) gradeBgQueue[origName] = [];
    // push the background values so duplicates get preserved in order
    nameBgQueue[origName].push(bgs[j][0] || "");
    gradeBgQueue[origName].push(bgs[j][1] || "");
  }

  // Sort populated rows by gradeKey then name
  populated.sort(function(a, b) {
    var ka = gradeKey(a.grade);
    var kb = gradeKey(b.grade);
    if (ka !== kb) return ka - kb;
    return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
  });

  // Build output values and the matching background arrays (2D arrays)
  var outValues = [];
  var outNameBgs = [];
  var outGradeBgs = [];

  for (var m = 0; m < maxRows; m++) {
    if (m < populated.length) {
      var p = populated[m];
      outValues.push([p.name, p.grade]);

      // pop one background for this name (or empty string if none)
      var nBg = "";
      var gBg = "";
      if (nameBgQueue[p.name] && nameBgQueue[p.name].length) nBg = nameBgQueue[p.name].shift();
      if (gradeBgQueue[p.name] && gradeBgQueue[p.name].length) gBg = gradeBgQueue[p.name].shift();

      outNameBgs.push([nBg]);
      outGradeBgs.push([gBg]);
    } else {
      outValues.push(["", ""]);
      outNameBgs.push([""]);
      outGradeBgs.push([""]);
    }
  }

  // Write sorted values back
  range.setValues(outValues);

  // Apply backgrounds for the two columns separately
  try {
    var nameRange = sheet.getRange(startRow, col + 1, maxRows, 1);
    var gradeRange = sheet.getRange(startRow, col + 2, maxRows, 1);
    nameRange.setBackgrounds(outNameBgs);
    gradeRange.setBackgrounds(outGradeBgs);
  } catch (e) {
    Logger.log("sortProjectBlock: failed to set backgrounds: " + e.toString());
  }

  // (Optional) Reapply CF for grade column if you use CF rules elsewhere
  try { setGradeConditionalFormatting(sheet, col); } catch (e) { /* ignore */ }

  return;
}

function submitStudent(name, grade, project){
  const sheet = getTargetSheet();
  const map = getProjectColumnMapForSheet(sheet);
  const col = map[project];
  if (col === undefined || col === null) return { ok:false, message: "project not found" };

  const startRow = 3;
  const maxRows = MAX_STUDENTS || 100;

  // read existing block (name, grade)
  const blockRange = sheet.getRange(startRow, col + 1, maxRows, 2);
  const block = blockRange.getValues();

  // collect non-empty rows
  const filled = block
    .filter(r => (r[0] || "").toString().trim())
    .map(r => ({ name: r[0].toString().trim(), grade: (r[1] || "").toString().trim() }));

  // add new entry and sort by grade order
  filled.push({ name: name, grade: (grade || "").toString() });
  filled.sort((a, b) => gradeKey(a.grade) - gradeKey(b.grade));

  // prepare output rows with blanks after filled entries
  const out = [];
  for (let i = 0; i < maxRows; i++) {
    if (i < filled.length) out.push([filled[i].name, filled[i].grade]);
    else out.push(["", ""]);
  }

  // write back block
  blockRange.setValues(out);

  // reapply conditional formatting for this project column
  try { setGradeConditionalFormatting(sheet, col); } catch (e) { Logger.log("submitStudent CF: " + e.toString()); }

  return { ok: true, message: `Inserted ${name} into "${project}" and sorted (${filled.length} rows now)` };
}

function sortProject(sheet,nameCol,gradeCol){

  const range = sheet.getRange(START_ROW,nameCol,MAX_ROWS,2);
  const values = range.getValues();

  const filled = values
    .filter(r=>r[0])
    .map(r=>({name:r[0],grade:r[1]}));

  filled.sort((a,b)=>gradeKey(a.grade)-gradeKey(b.grade));

  const out = filled.map(s=>[s.name,s.grade]);

  while(out.length<MAX_ROWS) out.push(["",""]);

  range.setValues(out);
}

function applyGradeFormatting(sheet,col){

  const range = sheet.getRange(START_ROW,col,MAX_ROWS,1);

  const rules = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("2A")
      .setBackground("#b6d7a8")
      .setRanges([range])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("2B")
      .setBackground("#a4c2f4")
      .setRanges([range])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("3A")
      .setBackground("#ffe599")
      .setRanges([range])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("3B")
      .setBackground("#ea9999")
      .setRanges([range])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

/* ============= Shuffle engine ============= */
const MIN_STUDENTS = 15;
const MAX_GENDER_DIFF = 2;
const MAX_GRADE_DIFF = 2;
const MAX_SHUFFLE_PASSES = 20;

function analyzeGroup(group) {
  let girls = 0, boys = 0, g2 = 0, g3 = 0;
  group.forEach(s => { const g = ((s.gender || "") + "").toLowerCase(); if (g.indexOf("f") === 0) girls++; else if (g.indexOf("m") === 0) boys++; if ((s.grade || "").toString().startsWith("2")) g2++; else if ((s.grade || "").toString().startsWith("3")) g3++; });
  return { total: group.length, girls, boys, g2, g3, genderDiff: Math.abs(girls-boys), gradeDiff: Math.abs(g3-g2), valid: (group.length >= MIN_STUDENTS && group.length <= MAX_STUDENTS && Math.abs(girls-boys) <= MAX_GENDER_DIFF && Math.abs(g3-g2) <= MAX_GRADE_DIFF) };
}

function shuffleProjectsRandomly(groupsObj) {
  const projects = Object.keys(groupsObj);
  const groups = {}; projects.forEach(p => groups[p] = (groupsObj[p] || []).slice());
  function groupPenalty(arr) { const an = analyzeGroup(arr); let penalty = 0; if (an.total < MIN_STUDENTS) penalty += (MIN_STUDENTS - an.total) * 1000; if (an.total > MAX_STUDENTS) penalty += (an.total - MAX_STUDENTS) * 1000; penalty += an.genderDiff * 50; penalty += an.gradeDiff * 50; return penalty; }
  const initialPenalty = projects.reduce((acc,p)=>acc+groupPenalty(groups[p]),0);
  let best = JSON.parse(JSON.stringify(groups)); let bestPenalty = initialPenalty;
  for (let pass=0; pass<MAX_SHUFFLE_PASSES; pass++) {
    let improved = false; const attemptedPairs = {};
    for (let i=0;i<projects.length;i++){
      for (let j=i+1;j<projects.length;j++){
        const pA = projects[i], pB = projects[j]; const arrA = groups[pA], arrB = groups[pB];
        if (!arrA.length || !arrB.length) continue;
        const idxA = Math.floor(Math.random() * arrA.length), idxB = Math.floor(Math.random() * arrB.length);
        const a = arrA[idxA], b = arrB[idxB]; if (!a || !b) continue;
        const keyPair = `${pA}|${idxA}:${pB}|${idxB}`; if (attemptedPairs[keyPair]) continue; attemptedPairs[keyPair] = true;
        const hist = buildHistoryMapFromSheets();
        if ((hist[normalizeName(a.name)] && hist[normalizeName(a.name)][pB]) || (hist[normalizeName(b.name)] && hist[normalizeName(b.name)][pA])) continue;
        const newA = arrA.slice(); newA[idxA] = b; const newB = arrB.slice(); newB[idxB] = a;
        const newPenalty = groupPenalty(newA) + groupPenalty(newB) + projects.reduce((acc,p)=> { if (p!==pA && p!==pB) return acc + groupPenalty(groups[p]); return acc;}, 0);
        if (newPenalty < bestPenalty) { groups[pA][idxA] = b; groups[pB][idxB] = a; best = JSON.parse(JSON.stringify(groups)); bestPenalty = newPenalty; improved = true; }
      }
    }
    if (!improved) break;
  }
  const histFinal = buildHistoryMapFromSheets();
  for (const p of Object.keys(best)) for (const s of best[p]) if (histFinal[normalizeName(s.name)] && histFinal[normalizeName(s.name)][p]) throw new Error(`shuffleProjectsRandomly: would assign ${s.name} to project ${p} but history forbids it.`);
  return best;
}
/*============== Write Balanced Groups ==============*/
function writeBalancedGroups(groups) {
  if (!groups || typeof groups !== 'object') {
    throw new Error('writeBalancedGroups: invalid groups argument');
  }

  const sheet = getTargetSheet();
  const projMap = getProjectColumnMapForSheet(sheet);
  const startRow = 3;
  const maxRows = (typeof MAX_STUDENTS !== 'undefined' && MAX_STUDENTS) ? MAX_STUDENTS : 16;

  Object.entries(groups).forEach(([project, students]) => {
    if (!project) return;
    const col = projMap[project];
    if (col === undefined || col === null) {
      Logger.log('writeBalancedGroups: project "%s" not found on active sheet - skipping', project);
      return;
    }

    // Normalize students array
    const arr = Array.isArray(students) ? students.slice() : [];
    const normalized = arr.map(s => ({
      name: (s && s.name) ? String(s.name).trim() : '',
      grade: (s && s.grade) ? String(s.grade).trim() : '',
      gender: (s && s.gender) ? String(s.gender).toLowerCase().trim() : ''
    })).filter(s => s.name); // drop empties

    // Sort by gradeKey() (existing helper) then name
    normalized.sort((a, b) => {
      try {
        const ka = typeof gradeKey === 'function' ? gradeKey(a.grade) : 0;
        const kb = typeof gradeKey === 'function' ? gradeKey(b.grade) : 0;
        if (ka !== kb) return ka - kb;
      } catch (e) {
        // fallback: ignore gradeKey errors
      }
      return (a.name || '').localeCompare(b.name || '', undefined, { sensitivity: 'base' });
    });

    // Build output values and name/grade background arrays
    const outValues = [];
    const nameBgs = [];
    const gradeBgs = [];
    for (let i = 0; i < maxRows; i++) {
      if (i < normalized.length) {
        const s = normalized[i];
        outValues.push([s.name, s.grade]);
        // apply simple gender color for name cell
        if (s.gender && s.gender[0] === 'f') nameBgs.push(['#ffebee']);
        else if (s.gender && s.gender[0] === 'm') nameBgs.push(['#e8f0fe']);
        else nameBgs.push(['']);
        gradeBgs.push(['']);
      } else {
        outValues.push(['', '']);
        nameBgs.push(['']);
        gradeBgs.push(['']);
      }
    }

    try {
      const blockRange = sheet.getRange(startRow, col + 1, maxRows, 2);
      blockRange.setValues(outValues);

      // Set background colors for name and grade columns (if available)
      try {
        const nameRange = sheet.getRange(startRow, col + 1, maxRows, 1);
        const gradeRange = sheet.getRange(startRow, col + 2, maxRows, 1);
        nameRange.setBackgrounds(nameBgs);
        gradeRange.setBackgrounds(gradeBgs);
      } catch (e) {
        Logger.log('writeBalancedGroups: failed to set backgrounds for project "%s": %s', project, e.toString());
      }

      // Reapply conditional formatting for grade column
      try {
        setGradeConditionalFormatting(sheet, col);
      } catch (e) {
        Logger.log('writeBalancedGroups: setGradeConditionalFormatting failed for col %s: %s', col, e.toString());
      }
    } catch (e) {
      Logger.log('writeBalancedGroups: failed to write project "%s": %s', project, e.toString());
    }
  });

  return { ok: true, message: 'writeBalancedGroups: wrote ' + Object.keys(groups).length + ' projects' };
}
/* ============= Admin wrappers ============= */
function adminCloseSheet() {
  try {
    const groups = readAllGroups();
    let shuffled;
    try { shuffled = shuffleProjectsRandomly(groups); } catch(e) { Logger.log("shuffle failed: " + e.toString()); shuffled = groups; }
    writeBalancedGroups(shuffled);
    const destGid = getSheetGids().find(g => g && g !== 0);
    if (destGid) { const commitLog = commitGroupsToSheet(destGid, shuffled); Logger.log("adminCloseSheet commit log:\n" + commitLog.join("\n")); }
    try { getConfigSheet().getRange(1,3).setValue("CLOSED"); } catch(e){}
    return { ok:true, message: "Closed and shuffled active target." };
  } catch (e) {
    Logger.log("adminCloseSheet error: " + e.toString());
    return { ok:false, message: e.message || e.toString() };
  }
}
function adminCloseAndCreateNewTarget() { try { const res = closeCreateNewTarget(); return { ok:true, message: res.message, log: res.commitLog }; } catch (e) { return { ok:false, message: e.message || e.toString() }; } }
function adminSetExistingTarget(gid) { try { const ss = getSpreadsheet(); const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid)); if (!sheet) throw new Error("Sheet with gid " + gid + " not found."); setTargetSheetId(Number(gid)); try { getConfigSheet().getRange(1,3).clearContent(); } catch(e){} return { ok:true, message: `Set active target to gid=${gid} (${sheet.getName()})` }; } catch (e) { return { ok:false, message: e.message || String(e) }; } }

/*================ Signing in missing students*/
/**
 * getAllRegisteredStudents()
 * 
 * 
 */
/* ---------- Auto-assign helpers (drop-in) ---------- */

/**
 * getAllRegisteredStudents()
 * - Reads the SOURCE_INFO_SHEET (canonical students list)
 * - Returns array of { fullName, normalized }
 */
function getAllRegisteredStudents() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const infoSheet = sheets.find(s => s.getSheetId && s.getSheetId() === Number(SOURCE_INFO_SHEET_GID));
  if (!infoSheet) throw new Error("getAllRegisteredStudents: info sheet not found (gid=" + SOURCE_INFO_SHEET_GID + ")");
  const data = infoSheet.getDataRange().getValues();
  const students = [];
  for (let r = 1; r < data.length; r++) {
    const last = (data[r][0] || "").toString().trim();
    const first = (data[r][1] || "").toString().trim();
    if (!last && !first) continue;
    const full = (first && last) ? (first + " " + last) : (first || last);
    const norm = normalizeName(full);
    students.push({ fullName: full, normalized: norm });
  }
  return students;
}

/**
 * getSignedInStudentsOnTarget()
 * - Returns a Set of normalized names currently on the active target sheet.
 */
function getSignedInStudentsOnTarget() {
  const sheet = getTargetSheet();
  const projMap = getProjectColumnMapForSheet(sheet);
  const startRow = 3;
  const maxRows = (typeof MAX_STUDENTS !== 'undefined' && MAX_STUDENTS) ? MAX_STUDENTS : 16;
  const signed = new Set();
  Object.values(projMap).forEach(col => {
    const block = sheet.getRange(startRow, col + 1, maxRows, 1).getValues();
    for (let i = 0; i < block.length; i++) {
      const raw = (block[i][0] || "").toString().trim();
      if (!raw) continue;
      signed.add(normalizeName(raw));
    }
  });
  return signed;
}

/**
 * findAvailableProjectsForTarget()
 * - Returns array: { projectName, col, assignedCount, capacityLeft }
 */
function findAvailableProjectsForTarget() {
  const sheet = getTargetSheet();
  const projMap = getProjectColumnMapForSheet(sheet);
  const startRow = 3;
  const maxRows = (typeof MAX_STUDENTS !== 'undefined' && MAX_STUDENTS) ? MAX_STUDENTS : 16;
  const list = [];
  Object.entries(projMap).forEach(([projectName, col]) => {
    const block = sheet.getRange(startRow, col + 1, maxRows, 1).getValues();
    let assigned = 0;
    for (let i = 0; i < block.length; i++) {
      if ((block[i][0] || "").toString().trim()) assigned++;
    }
    list.push({ projectName, col: Number(col), assignedCount: assigned, capacityLeft: Math.max(0, maxRows - assigned) });
  });
  return list;
}

/**
 * safeSignInStudent(fullName, projectName)
 * - Prefers existing writeStudentToProject() when available; otherwise writes directly.
 * - Returns { ok: true/false, message: ..., row, col }
 */
function safeSignInStudent(fullName, projectName) {
  try {
    if (!fullName || !projectName) return { ok: false, message: "invalid args" };

    // Prefer canonical API if present
    if (typeof writeStudentToProject === "function") {
      try {
        const info = (typeof getStudentGradeGender === "function") ? (getStudentGradeGender(fullName) || {}) : {};
        const res = writeStudentToProject(fullName, projectName, info.grade || "", info.gender || "");
        // normalize result
        if (res && typeof res === "object") return res;
        return { ok: !!res, message: String(res || "") };
      } catch (e) {
        Logger.log("safeSignInStudent: writeStudentToProject threw, falling back: " + e.toString());
        // fall through to fallback implementation
      }
    }

    const sheet = getTargetSheet();
    const projMap = getProjectColumnMapForSheet(sheet);
    const col = projMap[projectName];
    if (col === undefined || col === null) return { ok: false, message: "project not found on active sheet: " + projectName };

    const startRow = 3;
    const maxRows = (typeof MAX_STUDENTS !== 'undefined' && MAX_STUDENTS) ? MAX_STUDENTS : 16;
    const blockRange = sheet.getRange(startRow, col + 1, maxRows, 2);
    const vals = blockRange.getValues();

    let insertIdx = -1;
    for (let i = 0; i < vals.length; i++) {
      if (!vals[i][0] || !String(vals[i][0]).trim()) { insertIdx = i; break; }
    }
    if (insertIdx === -1) return { ok: false, message: "no free slot in " + projectName };

    const writeRow = startRow + insertIdx;
    vals[insertIdx][0] = fullName;
    try {
      const info = (typeof getStudentGradeGender === "function") ? (getStudentGradeGender(fullName) || {}) : {};
      vals[insertIdx][1] = info.grade || vals[insertIdx][1] || "";
      blockRange.setValues(vals);

      // set gender background on name cell (best-effort)
      if (info && info.gender) {
        const nameRange = sheet.getRange(writeRow, col + 1, 1, 1);
        const gLower = String(info.gender).toLowerCase();
        if (gLower.indexOf("f") === 0) nameRange.setBackground("#ffebee");
        else if (gLower.indexOf("m") === 0) nameRange.setBackground("#e8f0fe");
      }
      try { setGradeConditionalFormatting(sheet, col); } catch (e) {}
      return { ok: true, message: "signed " + fullName + " -> " + projectName + " (row " + writeRow + ")", row: writeRow, col: col };
    } catch (e) {
      return { ok: false, message: "write failed: " + e.toString() };
    }

  } catch (e) {
    return { ok: false, message: "safeSignInStudent failed: " + e.toString() };
  }
}

/**
 * safeRemoveStudentFromProject(fullName, projectName)
 * - Remove first matching row and compacts column
 */
function safeRemoveStudentFromProject(fullName, projectName) {
  try {
    if (!fullName || !projectName) return false;
    const sheet = getTargetSheet();
    const projMap = getProjectColumnMapForSheet(sheet);
    const col = projMap[projectName];
    if (col === undefined || col === null) return false;

    const startRow = 3;
    const maxRows = (typeof MAX_STUDENTS !== 'undefined' && MAX_STUDENTS) ? MAX_STUDENTS : 16;
    const blockRange = sheet.getRange(startRow, col + 1, maxRows, 2);
    const vals = blockRange.getValues();
    const bgs = blockRange.getBackgrounds();

    const targetNorm = normalizeName(fullName);
    let idx = -1;
    for (let i = 0; i < vals.length; i++) {
      const n = (vals[i][0] || "").toString().trim();
      if (n && normalizeName(n) === targetNorm) { idx = i; break; }
    }
    if (idx === -1) return false;

    // shift up
    for (let i = idx; i < vals.length - 1; i++) {
      vals[i][0] = vals[i + 1][0];
      vals[i][1] = vals[i + 1][1];
      bgs[i][0] = bgs[i + 1][0];
      bgs[i][1] = bgs[i + 1][1];
    }
    // clear last row
    vals[vals.length - 1][0] = "";
    vals[vals.length - 1][1] = "";
    bgs[bgs.length - 1][0] = "";
    bgs[bgs.length - 1][1] = "";

    // write back
    blockRange.setValues(vals);
    try {
      const nameRange = sheet.getRange(startRow, col + 1, maxRows, 1);
      const gradeRange = sheet.getRange(startRow, col + 2, maxRows, 1);
      const nameBgs = bgs.map(r => [r[0] || ""]);
      const gradeBgs = bgs.map(r => [r[1] || ""]);
      nameRange.setBackgrounds(nameBgs);
      gradeRange.setBackgrounds(gradeBgs);
    } catch (e) {
      Logger.log("safeRemoveStudentFromProject: background set failed: " + e.toString());
    }
    try { setGradeConditionalFormatting(sheet, col); } catch (e) {}
    return true;
  } catch (e) {
    Logger.log("safeRemoveStudentFromProject failed: " + e.toString());
    return false;
  }
}

/**
 * tryRebalanceForStudent(studentFullName, history)
 * - Single-step rebalance: move one occupant if it can be moved to free a slot.
 */
function tryRebalanceForStudent(studentFullName, history) {
  if (!studentFullName) return { ok: false, reason: "invalid student" };
  try {
    const sheet = getTargetSheet();
    const projMap = getProjectColumnMapForSheet(sheet);
    const startRow = 3;
    const maxRows = (typeof MAX_STUDENTS !== 'undefined' && MAX_STUDENTS) ? MAX_STUDENTS : 16;
    const studentNorm = normalizeName(studentFullName);
    history = history || {};
    const projects = Object.keys(projMap);

    for (const projectName of projects) {
      if ((history[studentNorm] || {})[projectName]) continue; // student already did this
      const col = projMap[projectName];
      const block = sheet.getRange(startRow, col + 1, maxRows, 1).getValues();
      // if free slot exists, no need to rebalance
      const freeIdx = block.findIndex(r => !r[0] || !String(r[0]).trim());
      if (freeIdx !== -1) return { ok: false, reason: "project has free slot" };

      for (let i = 0; i < block.length; i++) {
        const occupant = (block[i][0] || "").toString().trim();
        if (!occupant) continue;
        const occNorm = normalizeName(occupant);
        const occHist = history[occNorm] || {};

        for (const altProject of projects) {
          if (altProject === projectName) continue;
          if (occHist[altProject]) continue;
          const altCol = projMap[altProject];
          const altBlock = sheet.getRange(startRow, altCol + 1, maxRows, 1).getValues();
          const altFreeIdx = altBlock.findIndex(r => !r[0] || !String(r[0]).trim());
          if (altFreeIdx === -1) continue;

          // perform move
          const removed = safeRemoveStudentFromProject(occupant, projectName);
          if (!removed) continue;
          const signAlt = safeSignInStudent(occupant, altProject);
          if (!signAlt || !signAlt.ok) {
            // try revert
            safeRemoveStudentFromProject(occupant, altProject);
            safeSignInStudent(occupant, projectName);
            continue;
          }
          const signTarget = safeSignInStudent(studentFullName, projectName);
          if (!signTarget || !signTarget.ok) {
            // revert both
            safeRemoveStudentFromProject(occupant, altProject);
            safeSignInStudent(occupant, projectName);
            continue;
          }

          // update history
          if (!history[occNorm]) history[occNorm] = {};
          history[occNorm][altProject] = true;
          if (!history[studentNorm]) history[studentNorm] = {};
          history[studentNorm][projectName] = true;

          return {
            ok: true,
            moved: occupant,
            movedFrom: projectName,
            movedTo: altProject,
            placed: studentFullName,
            details: { movedRow: signAlt.row || null, placedRow: signTarget.row || null }
          };
        }
      }
    }
    return { ok: false, reason: "no single-step rebalance found" };
  } catch (e) {
    Logger.log("tryRebalanceForStudent failed: " + e.toString());
    return { ok: false, reason: e.toString() };
  }
}

/**
 * findAndSignUnsignedStudents(options)
 * - options.trySignIn (default true) actually writes, false does dry-run
 * - returns report { signed, moved, skipped, errors }
 */
function findAndSignUnsignedStudents(options) {
  options = options || {};
  const trySignIn = (typeof options.trySignIn === "boolean") ? options.trySignIn : true;

  const report = { signed: [], moved: [], skipped: [], errors: [] };

  // 1) registered students
  let allStudents = [];
  try { allStudents = getAllRegisteredStudents(); } catch (e) { throw new Error("findAndSignUnsignedStudents: failed to read registered students: " + e.toString()); }

  // 2) who is already signed in
  let signedSet = new Set();
  try { signedSet = getSignedInStudentsOnTarget(); } catch (e) { throw new Error("findAndSignUnsignedStudents: failed to read target sheet: " + e.toString()); }

  // 3) history map
  let history = {};
  try { history = buildHistoryMapFromSheets(); } catch (e) { Logger.log("findAndSignUnsignedStudents: buildHistoryMapFromSheets failed: " + e.toString()); history = {}; }

  // 4) project availability snapshot
  let projects = findAvailableProjectsForTarget();
  function refreshProjects() { projects = findAvailableProjectsForTarget(); }

  // iterate students
  for (let i = 0; i < allStudents.length; i++) {
    const s = allStudents[i];
    if (!s || !s.normalized) continue;
    if (signedSet.has(s.normalized)) continue;

    let placed = false;
    // try direct placement
    for (let p = 0; p < projects.length; p++) {
      const proj = projects[p];
      if (!proj || proj.capacityLeft <= 0) continue;
      const histForStudent = history[s.normalized] || {};
      if (histForStudent[proj.projectName]) continue;
      if (trySignIn) {
        const res = safeSignInStudent(s.fullName, proj.projectName);
        if (res && res.ok) {
          report.signed.push({ student: s.fullName, project: proj.projectName, message: res.message });
          signedSet.add(s.normalized);
          refreshProjects();
          placed = true;
          break;
        } else {
          report.errors.push({ student: s.fullName, project: proj.projectName, reason: res ? res.message : "unknown" });
          continue;
        }
      } else {
        report.signed.push({ student: s.fullName, project: proj.projectName, message: "dry-run would assign" });
        signedSet.add(s.normalized);
        refreshProjects();
        placed = true;
        break;
      }
    }
    if (placed) continue;

    // try single-step rebalance
    if (trySignIn) {
      const reb = tryRebalanceForStudent(s.fullName, history);
      if (reb && reb.ok) {
        report.moved.push({
          studentPlaced: reb.placed,
          movedStudent: reb.moved,
          movedFrom: reb.movedFrom,
          movedTo: reb.movedTo,
          details: reb.details
        });
        signedSet.add(normalizeName(reb.placed));
        if (!history[normalizeName(reb.moved)]) history[normalizeName(reb.moved)] = {};
        history[normalizeName(reb.moved)][reb.movedTo] = true;
        if (!history[normalizeName(reb.placed)]) history[normalizeName(reb.placed)] = {};
        history[normalizeName(reb.placed)][reb.movedFrom] = true;
        refreshProjects();
        continue;
      }
    }

    report.skipped.push({ student: s.fullName, reason: "no available project (or only ones student already did)" });
  }

  return report;
}

/* ---------- end auto-assign helpers ---------- */

/* ============= Web entry ============= */
function doGet(e) {
  // Expects a file named "index.html" in the project (same as your front-end)
  return HtmlService.createHtmlOutputFromFile('index').setTitle('PřV projekt Sign-In');
}

/**
 *   DEBUG FUNCTIONS 
 *  Functions to debug the script -- Will be deleted 
 * 
*/
function _runAutoSignAll() {
  const out = findAndSignUnsignedStudents({ trySignIn: true });
  Logger.log("LIVE signed=%d, moved=%d, skipped=%d, errors=%d", out.signed.length, out.moved.length, out.skipped.length, out.errors.length);
  return out;
}
/**
 * debugListMissingStudentsAccurate
 * - Uses your existing helpers exactly:
 *   getAllRegisteredStudents(), getSignedInStudentsOnTarget(), buildHistoryMapFromSheets(), normalizeName()
 * - Does NOT modify any sheet. Logs summary and returns an object:
 *   { totalRegistered, presentOnTargetCount, presentInHistoryCount, missingCount, missing (array), details (array) }
 *
 * Missing = registered student that is NOT on the active target sheet and NOT present in the history map.
 */
function debugListMissingStudentsAccurate() {
  try {
    // 1) get canonical registered students
    if (typeof getAllRegisteredStudents !== "function") {
      Logger.log("debugListMissingStudentsAccurate: getAllRegisteredStudents() not available.");
      return null;
    }
    const all = getAllRegisteredStudents() || [];
    // defensive: ensure each has normalized (function does this, but double-check)
    const students = all.map(s => {
      const full = (s && s.fullName) ? String(s.fullName).trim() : "";
      const norm = (s && s.normalized) ? String(s.normalized).trim() : (typeof normalizeName === "function" ? normalizeName(full) : full);
      return { fullName: full, normalized: norm };
    });

    // 2) get who is currently signed in on the active target
    let signedSet = new Set();
    if (typeof getSignedInStudentsOnTarget === "function") {
      try {
        signedSet = getSignedInStudentsOnTarget() || new Set();
      } catch (e) {
        Logger.log("debugListMissingStudentsAccurate: getSignedInStudentsOnTarget() threw: " + String(e));
        signedSet = new Set();
      }
    } else {
      Logger.log("debugListMissingStudentsAccurate: getSignedInStudentsOnTarget() not found. Assuming none signed on target.");
    }

    // 3) build history map (students that already did projects)
    let history = {};
    if (typeof buildHistoryMapFromSheets === "function") {
      try {
        history = buildHistoryMapFromSheets() || {};
      } catch (e) {
        Logger.log("debugListMissingStudentsAccurate: buildHistoryMapFromSheets() threw: " + String(e));
        history = {};
      }
    } else {
      Logger.log("debugListMissingStudentsAccurate: buildHistoryMapFromSheets() not found. Assuming empty history.");
    }

    // 4) classify each registered student
    const details = students.map(s => {
      const norm = (s.normalized || "").toString().trim();
      const onTarget = signedSet.has(norm);
      const histEntry = history[norm] || {};
      // history may have meta keys like "__..." — filter those out when listing projects
      const historyProjects = Object.keys(histEntry).filter(k => k && !k.startsWith("__"));
      const inHistory = historyProjects.length > 0;
      const presentAnywhere = onTarget || inHistory;
      return {
        fullName: s.fullName,
        normalized: norm,
        onTarget: !!onTarget,
        inHistory: inHistory,
        historyProjects: historyProjects,
        presentAnywhere: !!presentAnywhere
      };
    });

    const missing = details.filter(d => !d.presentAnywhere).map(d => d.fullName);

    // 5) logging summary + helpful debug snippets
    Logger.log("debugListMissingStudentsAccurate: total registered = " + students.length);
    Logger.log("debugListMissingStudentsAccurate: present on target = " + details.filter(d => d.onTarget).length);
    Logger.log("debugListMissingStudentsAccurate: present in history = " + details.filter(d => d.inHistory).length);
    Logger.log("debugListMissingStudentsAccurate: missing (not on target & not in history) = " + missing.length);

    if (missing.length) {
      // log full list (comma separated) and one-per-line for easier reading
      Logger.log("Missing (comma-separated): " + missing.join(", "));
      Logger.log("Missing (lines):\n" + missing.join("\n"));
    } else {
      Logger.log("No missing students found — every registered student is on target or in history.");
    }

    // Also include small samples of history keys for debugging (if history non-empty)
    try {
      const historyKeys = Object.keys(history || {});
      if (historyKeys.length) {
        Logger.log("Sample history keys (first 10 normalized names): " + historyKeys.slice(0, 10).join(", "));
      } else {
        Logger.log("History map appears empty (no prior project records found).");
      }
    } catch (e) {
      /* ignore */
    }

    // return a programmatic result for inspection in debugger
    return {
      totalRegistered: students.length,
      presentOnTargetCount: details.filter(d => d.onTarget).length,
      presentInHistoryCount: details.filter(d => d.inHistory).length,
      missingCount: missing.length,
      missing: missing,
      details: details
    };
  } catch (err) {
    Logger.log("debugListMissingStudentsAccurate: ERROR: " + (err && err.message ? err.message : String(err)));
    return { error: (err && err.message) ? err.message : String(err) };
  }
}
