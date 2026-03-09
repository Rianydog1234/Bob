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

/** Append old active gid into first empty B cell (B1..Bn). */
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

/**
 * swapNameIfNeeded(name, sheetIndex)
 * - If name contains comma "Last, First" => returns "First Last"
 * - If sheetIndex is set to legacy indices where order was reversed, swap.
 * - Otherwise returns original trimmed name.
 * This preserves compatibility with original repo where some source sheets had Last First order.
 */
function swapNameIfNeeded(name, sheetIndex) {
  if (!name) return "";
  const raw = String(name).trim();
  // 1) handle "Last, First" formats
  if (raw.indexOf(",") !== -1) {
    const parts = raw.split(",").map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2) return (parts[1] + " " + parts[0]).trim();
    return raw;
  }
  // 2) legacy sheet index-based swap (keeps original behavior)
  if (typeof sheetIndex !== 'undefined' && (sheetIndex === 1 || sheetIndex === 2)) {
    const parts = raw.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return (parts.slice(1).join(" ") + " " + parts[0]).trim();
  }
  return raw;
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
        name = swapNameIfNeeded(name, sheetIndex);
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
        fullName = swapNameIfNeeded(fullName, sheetIndex);
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
  const groups = readAllGroups();
  const totalAssigned = Object.values(groups).reduce((acc, arr) => acc + (arr ? arr.length : 0), 0);
  if (totalAssigned === 0) throw new Error("closeCreateNewTarget: no assignments found in current target (aborting).");
  const destGid = getCommitDestinationGid();
  if (!destGid) throw new Error("closeCreateNewTarget: no valid dest in getSheetGids().");
  commitGroupsToSheet(destGid, groups);
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const newName = `Target_${ts}`;
  const newSheet = createNewTargetSheetFromTemplate(newName);
  setTargetSheetId(newSheet.getSheetId());
  // Move old active into ScriptConfig B column first empty cell so the history/rotation is preserved
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

/* ============= Debug functions ============= */
/**
 * runCloseSheetDebug
 * Read-only diagnostic for the "close sheet -> save to history" flow.
 * Run from the Apps Script editor (select runCloseSheetDebug and Run).
 *
 * This collects:
 *  - resolved target sheet
 *  - project headers (row 2) and assignments (rows 3..lastRow)
 *  - candidate history/archive sheets and sample content
 *  - normalization-matching between current assignments and history rows
 *  - presence/absence of functions involved in the close/save flow
 *
 * DOES NOT MODIFY ANY SHEET.
 */
function runCloseSheetDebug() {
  var out = {
    generatedAt: new Date().toISOString(),
    warnings: [],
    info: {}
  };

  function safeRun(name, fn) {
    try { return { ok: true, value: fn() }; } catch (e) { return { ok: false, error: String(e && e.stack ? e.stack : e) }; }
  }

  // small normalizer used for name matching (matches NAME.js behaviour)
  function normalizeNameForDebug(s) {
    if (s === null || typeof s === "undefined") return "";
    var t = "" + s;
    if (t.normalize) t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    t = t.replace(/\s+/g, " ").trim().toLowerCase();
    return t;
  }

  // 1) Spreadsheet resolution (respect SPREADSHEET_ID if set, fallback to active)
  out.info.spreadsheet = safeRun("resolveSpreadsheet", function () {
    var ss;
    if (typeof SPREADSHEET_ID !== "undefined" && SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    return { id: ss.getId(), name: ss.getName(), url: ss.getUrl() };
  });

  if (!out.info.spreadsheet.ok) {
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  }

  var ss = (typeof SPREADSHEET_ID !== "undefined" && SPREADSHEET_ID)
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // 2) list sheets
  out.info.sheets = ss.getSheets().map(function (sh) {
    return { name: sh.getName(), gid: sh.getSheetId(), lastRow: sh.getLastRow(), lastCol: sh.getLastColumn() };
  });

  // 3) read ScriptConfig A1/B1 if present
  out.info.scriptConfig = safeRun("scriptConfig", function () {
    var cfg = {};
    var sh = ss.getSheetByName("ScriptConfig") || ss.getSheets().find(function(s){ return s.getName().toLowerCase() === "scriptconfig"; });
    if (!sh) return { exists: false };
    cfg.exists = true;
    cfg.A1 = sh.getRange(1,1).getDisplayValue();
    cfg.B1 = sh.getRange(1,2).getDisplayValue();
    return cfg;
  });

  // 4) computed target sheet (try getTargetSheet() if it exists)
  out.info.targetSheet = safeRun("targetSheetResolve", function () {
    if (typeof getTargetSheet === "function") {
      var t = getTargetSheet();
      if (!t) return { note: "getTargetSheet() returned null/undefined" };
      return { name: t.getName(), gid: t.getSheetId(), lastRow: t.getLastRow(), lastCol: t.getLastColumn() };
    }
    // fallback: use ScriptConfig B1 numeric gid if present
    var gidRaw = (out.info.scriptConfig.ok && out.info.scriptConfig.value && out.info.scriptConfig.value.B1) ? out.info.scriptConfig.value.B1 : null;
    if (gidRaw) {
      var gid = parseInt(gidRaw, 10);
      var found = ss.getSheets().find(function(sh){ return sh.getSheetId() === gid; });
      if (found) return { name: found.getName(), gid: found.getSheetId(), lastRow: found.getLastRow(), lastCol: found.getLastColumn() };
      return { error: "ScriptConfig B1 provided gid but sheet not found: " + gidRaw };
    }
    // fallback: first non-ScriptConfig sheet
    var fallback = ss.getSheets().find(function(sh){ return sh.getName().toLowerCase() !== "scriptconfig"; });
    if (fallback) return { name: fallback.getName(), gid: fallback.getSheetId(), lastRow: fallback.getLastRow(), lastCol: fallback.getLastColumn(), note: "fallback to first non-ScriptConfig sheet" };
    return { error: "No usable target sheet found" };
  });

  // 5) read project headers (row 2) and assignments (rows 3..lastRow)
  out.info.targetData = safeRun("targetData", function () {
    var tsObj = out.info.targetSheet.ok ? out.info.targetSheet.value : null;
    if (!tsObj || !tsObj.gid) throw new Error("No target sheet resolved");
    var targetSheet = ss.getSheets().find(function(sh){ return sh.getSheetId() === tsObj.gid; });
    if (!targetSheet) throw new Error("Target sheet object not found by gid");

    var lastCol = Math.max(1, targetSheet.getLastColumn());
    var lastRow = Math.max(3, targetSheet.getLastRow());
    var headers = targetSheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
    // read rows 3..lastRow, all columns
    var rows = [];
    if (lastRow >= 3) {
      var numRows = lastRow - 2;
      var rvals = targetSheet.getRange(3, 1, numRows, lastCol).getDisplayValues();
      rvals.forEach(function (r, i) {
        rows.push(r);
      });
    }
    // build project -> students map (collect non-empty cells per header)
    var projects = {};
    for (var c = 0; c < headers.length; c++) {
      var rawHeader = headers[c] || "";
      var headerTrim = (""+rawHeader).toString().trim();
      if (!headerTrim) continue;
      projects[headerTrim] = [];
    }
    // iterate rows, collect names per column where header exists
    for (var rr = 0; rr < rows.length; rr++) {
      for (var c2 = 0; c2 < headers.length; c2++) {
        var hdr = (headers[c2] || "").toString().trim();
        if (!hdr) continue;
        var cell = rows[rr][c2];
        if (cell !== null && (""+cell).toString().trim() !== "") {
          projects[hdr].push({ raw: cell, normalized: normalizeNameForDebug(cell), rowIndex: rr + 3, colIndex: c2 + 1 });
        }
      }
    }
    return { headers: headers, lastRow: lastRow, lastCol: lastCol, projects: projects, sampleRows: rows.slice(0,8) };
  });

  // 6) find candidate history/archive sheets (common names + contains)
  out.info.historyCandidates = safeRun("historyCandidates", function () {
    var candidates = ["WasOnProject", "History", "Archive", "WasOnProject Archive", "WasOnProject_old", "WasOnProject_Archive", "WasOnProject (Archive)"];
    // also include any sheet that contains history/was/archive
    var dynamic = ss.getSheets().map(function(sh){ return sh.getName(); }).filter(function(n){ return /history|was|archive/i.test(n); });
    var unique = {};
    (candidates.concat(dynamic)).forEach(function(n){ unique[n] = true; });
    var names = Object.keys(unique);
    // return those that exist and a sample from each
    var found = [];
    names.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) {
        // try case-insensitive
        var maybe = ss.getSheets().find(function(s){ return s.getName().toLowerCase() === (""+name).toLowerCase(); });
        if (maybe) sh = maybe;
      }
      if (!sh) return;
      var lastRow = Math.max(1, sh.getLastRow());
      var lastCol = Math.max(1, sh.getLastColumn());
      var sampleRangeRows = Math.min(500, lastRow); // limit
      var data = [];
      if (sampleRangeRows > 0) {
        try { data = sh.getRange(1,1, sampleRangeRows, lastCol).getDisplayValues(); } catch(e) { data = [["read error: "+String(e)]]; }
      }
      // normalize first two columns (typical structure)
      var normalizedIndex = {};
      for (var r = 0; r < Math.min(500, data.length); r++) {
        var nameCell = data[r] && data[r][0] ? data[r][0] : "";
        var projCell = data[r] && data[r][1] ? data[r][1] : "";
        var nn = normalizeNameForDebug(nameCell);
        if (nn) {
          if (!normalizedIndex[nn]) normalizedIndex[nn] = [];
          normalizedIndex[nn].push({ row: r+1, rawName: nameCell, rawProject: projCell });
        }
      }
      found.push({ sheetName: sh.getName(), gid: sh.getSheetId(), lastRow: lastRow, lastCol: lastCol, sampleRowsCount: Math.min(50, data.length), normalizedIndexSampleCount: Object.keys(normalizedIndex).length, normalizedIndexSample: Object.keys(normalizedIndex).slice(0,10), rawSampleFirstRows: data.slice(0,8) });
    });
    return found;
  });

  // 7) For each assigned student in targetData, check presence in history candidates
  out.info.historyMatching = safeRun("historyMatching", function () {
    var td = out.info.targetData.ok ? out.info.targetData.value : null;
    if (!td) return { note: "no target data to test" };
    var matches = {};
    var histSheets = out.info.historyCandidates.ok ? out.info.historyCandidates.value : [];
    var histBySheet = {};
    // build normalized index for each hist sheet (read limited rows again)
    histSheets.forEach(function(hs) {
      var sh = ss.getSheetByName(hs.sheetName) || ss.getSheets().find(function(s){ return s.getSheetId() === hs.gid; });
      if (!sh) return;
      var lastRow = Math.max(1, sh.getLastRow());
      var lastCol = Math.max(1, sh.getLastColumn());
      var rows = [];
      try { rows = sh.getRange(1,1, Math.min(1000,lastRow), lastCol).getDisplayValues(); } catch(e) { rows = [["read error: " + String(e)]]; }
      var idx = {};
      for (var r = 0; r < rows.length; r++) {
        var nameCell = rows[r] && rows[r][0] ? rows[r][0] : "";
        var projCell = rows[r] && rows[r][1] ? rows[r][1] : "";
        var nn = normalizeNameForDebug(nameCell);
        if (!nn) continue;
        if (!idx[nn]) idx[nn] = [];
        idx[nn].push({ row: r+1, rawName: nameCell, rawProject: projCell });
      }
      histBySheet[hs.sheetName] = { gid: hs.gid, index: idx, rawSampleRows: rows.slice(0,8) };
    });

    // now test each assigned student
    var summary = { totalAssigned: 0, foundInAnyHistory: 0, foundSameProject: 0, notFound: 0, details: [] };
    var projects = td.projects || {};
    Object.keys(projects).forEach(function(projectName) {
      var list = projects[projectName];
      list.forEach(function(entry) {
        summary.totalAssigned++;
        var nn = entry.normalized;
        var foundAny = false;
        var foundSame = false;
        var foundLocations = [];
        Object.keys(histBySheet).forEach(function(sheetName) {
          var idx = histBySheet[sheetName].index;
          if (idx[nn]) {
            foundAny = true;
            idx[nn].forEach(function(hrow) {
              foundLocations.push({ sheet: sheetName, row: hrow.row, rawProject: hrow.rawProject });
              if ((hrow.rawProject || "").toString().trim() === projectName) foundSame = true;
            });
          }
        });
        if (foundAny) summary.foundInAnyHistory++;
        if (foundSame) summary.foundSameProject++;
        if (!foundAny) summary.notFound++;
        summary.details.push({ studentRaw: entry.raw, studentNormalized: nn, project: projectName, foundAny: foundAny, foundSame: foundSame, foundLocations: foundLocations.slice(0,5) });
      });
    });

    return { histBySheetSummaryCount: Object.keys(histBySheet).length, summary: summary };
  });

  // 8) check presence of relevant server functions (so we know what code paths exist)
  out.info.serverFunctions = {};
  var fnNames = ["adminCloseSheet","adminCloseAndCreateNewTarget","closeSheet","commitGroupsToSheet","writeStudentToProject","writeHistory","saveHistory","archiveCurrentSheet","moveToHistory"];
  fnNames.forEach(function(n) {
    try {
      out.info.serverFunctions[n] = (typeof this[n] === "function") || (typeof globalThis !== "undefined" && typeof globalThis[n] === "function") || (typeof this[n] === "function");
    } catch(e) {
      out.info.serverFunctions[n] = "check-failed:" + String(e);
    }
  });

  // 9) permissions / session info (best-effort)
  out.info.session = safeRun("session", function () {
    var s = {};
    try { s.activeUser = Session.getActiveUser() && Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : null; } catch (e) { s.activeUserError = String(e); }
    try { s.effectiveUser = Session.getEffectiveUser() && Session.getEffectiveUser().getEmail ? Session.getEffectiveUser().getEmail() : null; } catch (e) { s.effectiveUserError = String(e); }
    return s;
  });

  // 10) Hints & hypotheses (based on data gathered)
  out.hypotheses = [];
  try {
    // If there are zero history candidate sheets -> clue
    var hc = out.info.historyCandidates.ok ? (out.info.historyCandidates.value || []) : [];
    if (!hc.length) out.hypotheses.push("No history/archive sheet found. The close flow may expect a 'WasOnProject' or 'History' sheet to append into.");
    // If many assigned students not found in any history entries, maybe history was never written
    var histMatch = out.info.historyMatching.ok ? out.info.historyMatching.value.summary : null;
    if (histMatch && histMatch.notFound === histMatch.totalAssigned) {
      out.hypotheses.push("All current assigned students are not present in history sheets. That suggests saving to history wasn't executed or used a different history sheet/format.");
    } else if (histMatch && histMatch.notFound > 0) {
      out.hypotheses.push(histMatch.notFound + " of " + histMatch.totalAssigned + " assigned students are not in any detected history sheet.");
    }
    // If project headers are empty / mismatched
    var th = out.info.targetData.ok ? out.info.targetData.value : null;
    if (th && Object.keys(th.projects).length === 0) out.hypotheses.push("Target sheet row 2 has no project headers (or they are blank). Close flow may rely on headers to map columns into history.");
  } catch(e) {
    out.hypothesesError = String(e);
  }

  // log result
  try {
    Logger.log("=== runCloseSheetDebug result ===\n" + JSON.stringify(out, null, 2));
  } catch (e) {
    Logger.log("runCloseSheetDebug: failed to stringify output: " + String(e));
  }

  return out; // returned to caller (Run -> Logs will show it)
}



/* ============= Web entry ============= */
function doGet(e) {
  // Expects a file named "Index.html" in the project (same as your front-end)
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Project Sign-In');
}

/* EOF */
function FULL_DEBUG() {

  const sheet = getTargetSheet();
  Logger.log("=== SHEET INFO ===");
  Logger.log("Sheet name: " + sheet.getName());

  Logger.log("=== PROJECT_COLUMNS ===");
  Logger.log(JSON.stringify(PROJECT_COLUMNS));

  Logger.log("=== HEADERS ROW 2 ===");

  PROJECT_COLUMNS.forEach(col=>{
    const header = sheet.getRange(2,col+1).getValue();
    Logger.log("Column "+(col+1)+" header: '"+header+"'");
  });

  Logger.log("=== PROJECT COLUMN MAP ===");

  const map = getProjectColumnMapForSheet(sheet);
  Logger.log(JSON.stringify(map));

  Logger.log("=== TEST WRITE PERMISSION ===");

  try{

    const testCell = sheet.getRange(1,1);
    const original = testCell.getValue();

    testCell.setValue("DEBUG_TEST");
    Utilities.sleep(500);
    testCell.setValue(original);

    Logger.log("Write test: SUCCESS");

  }catch(e){

    Logger.log("Write test FAILED: "+e);

  }

  Logger.log("=== CONDITIONAL FORMAT RULES ===");

  const rules = sheet.getConditionalFormatRules();

  Logger.log("Total rules: "+rules.length);

  rules.forEach((r,i)=>{

    const ranges = r.getRanges().map(rr=>rr.getA1Notation());
    Logger.log("Rule "+i+" ranges: "+ranges.join(", "));

  });

  Logger.log("=== PROJECT BLOCK CHECK ===");

  const START_ROW = 3;
  const CHECK_ROWS = 5;

  Object.entries(map).forEach(([project,col])=>{

    const vals = sheet.getRange(START_ROW,col+2,CHECK_ROWS,1).getValues().flat();

    Logger.log(project+" grade column sample: "+JSON.stringify(vals));

  });

  Logger.log("=== DEBUG COMPLETE ===");

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

/**
 * swapNameIfNeeded(name, sheetIndex)
 * - If name contains comma "Last, First" => returns "First Last"
 * - If sheetIndex is set to legacy indices where order was reversed, swap.
 * - Otherwise returns original trimmed name.
 * This preserves compatibility with original repo where some source sheets had Last First order.
 */
function swapNameIfNeeded(name, sheetIndex) {
  if (!name) return "";
  const raw = String(name).trim();
  // 1) handle "Last, First" formats
  if (raw.indexOf(",") !== -1) {
    const parts = raw.split(",").map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2) return (parts[1] + " " + parts[0]).trim();
    return raw;
  }
  // 2) legacy sheet index-based swap (keeps original behavior)
  if (typeof sheetIndex !== 'undefined' && (sheetIndex === 1 || sheetIndex === 2)) {
    const parts = raw.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return (parts.slice(1).join(" ") + " " + parts[0]).trim();
  }
  return raw;
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
  const defGid = SHEET_GIDS.find(g => g && g !== 0);
  if (!defGid) throw new Error("No definition GID configured in SHEET_GIDS");
  const defSheet = ss.getSheets().find(s => s.getSheetId() === Number(defGid));
  if (!defSheet) throw new Error("Definition sheet not found (check SHEET_GIDS)");
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
/**
 * buildHistoryMapFromSheets()
 * - reads each configured sheet (SHEET_GIDS)
 * - uses swapNameIfNeeded(sheetIndex) heuristics
 * - stores both name variants (so "first last" and "last first" are covered)
 * returns { normalizedName: { projectName: true, ... }, ... }
 */
/**
 * buildHistoryMapFromSheets - improved
 * - Scans configured SHEET_GIDS plus the current target sheet (if not already included)
 * - Keeps the same swapNameIfNeeded(sheetIndex) heuristic, where sheetIndex is the index
 *   into the combined list (configured sheets first, then the target sheet appended if needed)
 * - Returns a map: { normalizedName: { projectA: true, projectB: true, ... }, ... }
 */
function buildHistoryMapFromSheets() {
  const ss = getSpreadsheet();
  const history = {};

  // 1) build the list of gids to check:
  // start with configured list (filter out falsy / 0 entries)
  const configured = Array.isArray(SHEET_GIDS) ? SHEET_GIDS.filter(g => g && Number(g) !== 0).map(Number) : [];

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
        name = swapNameIfNeeded(name, sheetIndex);
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
  SHEET_GIDS.forEach((gid, sheetIndex) => {
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
        fullName = swapNameIfNeeded(fullName, sheetIndex);
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
  const defGid = SHEET_GIDS.find(g => g && g !== 0);
  if (!defGid) throw new Error("createNewTargetSheetFromTemplate: no def sheet in SHEET_GIDS");
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
  const groups = readAllGroups();
  const totalAssigned = Object.values(groups).reduce((acc, arr) => acc + (arr ? arr.length : 0), 0);
  if (totalAssigned === 0) throw new Error("closeCreateNewTarget: no assignments found in current target (aborting).");
  const destGid = SHEET_GIDS.find(g => g && g != 0);
  if (!destGid) throw new Error("closeCreateNewTarget: no valid dest in SHEET_GIDS.");
  const commitLog = commitGroupsToSheet(destGid, groups);
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const newName = `Target_${ts}`;
  const newSheet = createNewTargetSheetFromTemplate(newName);
  setTargetSheetId(newSheet.getSheetId());
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

/* ============= Admin wrappers ============= */
function adminCloseSheet() {
  try {
    const groups = readAllGroups();
    let shuffled;
    try { shuffled = shuffleProjectsRandomly(groups); } catch(e) { Logger.log("shuffle failed: " + e.toString()); shuffled = groups; }
    writeBalancedGroups(shuffled);
    const destGid = SHEET_GIDS.find(g => g && g !== 0);
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

/* ============= Debug functions ============= */
/**
 * runCloseSheetDebug
 * Read-only diagnostic for the "close sheet -> save to history" flow.
 * Run from the Apps Script editor (select runCloseSheetDebug and Run).
 *
 * This collects:
 *  - resolved target sheet
 *  - project headers (row 2) and assignments (rows 3..lastRow)
 *  - candidate history/archive sheets and sample content
 *  - normalization-matching between current assignments and history rows
 *  - presence/absence of functions involved in the close/save flow
 *
 * DOES NOT MODIFY ANY SHEET.
 */
function runCloseSheetDebug() {
  var out = {
    generatedAt: new Date().toISOString(),
    warnings: [],
    info: {}
  };

  function safeRun(name, fn) {
    try { return { ok: true, value: fn() }; } catch (e) { return { ok: false, error: String(e && e.stack ? e.stack : e) }; }
  }

  // small normalizer used for name matching (matches NAME.js behaviour)
  function normalizeNameForDebug(s) {
    if (s === null || typeof s === "undefined") return "";
    var t = "" + s;
    if (t.normalize) t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    t = t.replace(/\s+/g, " ").trim().toLowerCase();
    return t;
  }

  // 1) Spreadsheet resolution (respect SPREADSHEET_ID if set, fallback to active)
  out.info.spreadsheet = safeRun("resolveSpreadsheet", function () {
    var ss;
    if (typeof SPREADSHEET_ID !== "undefined" && SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    return { id: ss.getId(), name: ss.getName(), url: ss.getUrl() };
  });

  if (!out.info.spreadsheet.ok) {
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  }

  var ss = (typeof SPREADSHEET_ID !== "undefined" && SPREADSHEET_ID)
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // 2) list sheets
  out.info.sheets = ss.getSheets().map(function (sh) {
    return { name: sh.getName(), gid: sh.getSheetId(), lastRow: sh.getLastRow(), lastCol: sh.getLastColumn() };
  });

  // 3) read ScriptConfig A1/B1 if present
  out.info.scriptConfig = safeRun("scriptConfig", function () {
    var cfg = {};
    var sh = ss.getSheetByName("ScriptConfig") || ss.getSheets().find(function(s){ return s.getName().toLowerCase() === "scriptconfig"; });
    if (!sh) return { exists: false };
    cfg.exists = true;
    cfg.A1 = sh.getRange(1,1).getDisplayValue();
    cfg.B1 = sh.getRange(1,2).getDisplayValue();
    return cfg;
  });

  // 4) computed target sheet (try getTargetSheet() if it exists)
  out.info.targetSheet = safeRun("targetSheetResolve", function () {
    if (typeof getTargetSheet === "function") {
      var t = getTargetSheet();
      if (!t) return { note: "getTargetSheet() returned null/undefined" };
      return { name: t.getName(), gid: t.getSheetId(), lastRow: t.getLastRow(), lastCol: t.getLastColumn() };
    }
    // fallback: use ScriptConfig B1 numeric gid if present
    var gidRaw = (out.info.scriptConfig.ok && out.info.scriptConfig.value && out.info.scriptConfig.value.B1) ? out.info.scriptConfig.value.B1 : null;
    if (gidRaw) {
      var gid = parseInt(gidRaw, 10);
      var found = ss.getSheets().find(function(sh){ return sh.getSheetId() === gid; });
      if (found) return { name: found.getName(), gid: found.getSheetId(), lastRow: found.getLastRow(), lastCol: found.getLastColumn() };
      return { error: "ScriptConfig B1 provided gid but sheet not found: " + gidRaw };
    }
    // fallback: first non-ScriptConfig sheet
    var fallback = ss.getSheets().find(function(sh){ return sh.getName().toLowerCase() !== "scriptconfig"; });
    if (fallback) return { name: fallback.getName(), gid: fallback.getSheetId(), lastRow: fallback.getLastRow(), lastCol: fallback.getLastColumn(), note: "fallback to first non-ScriptConfig sheet" };
    return { error: "No usable target sheet found" };
  });

  // 5) read project headers (row 2) and assignments (rows 3..lastRow)
  out.info.targetData = safeRun("targetData", function () {
    var tsObj = out.info.targetSheet.ok ? out.info.targetSheet.value : null;
    if (!tsObj || !tsObj.gid) throw new Error("No target sheet resolved");
    var targetSheet = ss.getSheets().find(function(sh){ return sh.getSheetId() === tsObj.gid; });
    if (!targetSheet) throw new Error("Target sheet object not found by gid");

    var lastCol = Math.max(1, targetSheet.getLastColumn());
    var lastRow = Math.max(3, targetSheet.getLastRow());
    var headers = targetSheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
    // read rows 3..lastRow, all columns
    var rows = [];
    if (lastRow >= 3) {
      var numRows = lastRow - 2;
      var rvals = targetSheet.getRange(3, 1, numRows, lastCol).getDisplayValues();
      rvals.forEach(function (r, i) {
        rows.push(r);
      });
    }
    // build project -> students map (collect non-empty cells per header)
    var projects = {};
    for (var c = 0; c < headers.length; c++) {
      var rawHeader = headers[c] || "";
      var headerTrim = (""+rawHeader).toString().trim();
      if (!headerTrim) continue;
      projects[headerTrim] = [];
    }
    // iterate rows, collect names per column where header exists
    for (var rr = 0; rr < rows.length; rr++) {
      for (var c2 = 0; c2 < headers.length; c2++) {
        var hdr = (headers[c2] || "").toString().trim();
        if (!hdr) continue;
        var cell = rows[rr][c2];
        if (cell !== null && (""+cell).toString().trim() !== "") {
          projects[hdr].push({ raw: cell, normalized: normalizeNameForDebug(cell), rowIndex: rr + 3, colIndex: c2 + 1 });
        }
      }
    }
    return { headers: headers, lastRow: lastRow, lastCol: lastCol, projects: projects, sampleRows: rows.slice(0,8) };
  });

  // 6) find candidate history/archive sheets (common names + contains)
  out.info.historyCandidates = safeRun("historyCandidates", function () {
    var candidates = ["WasOnProject", "History", "Archive", "WasOnProject Archive", "WasOnProject_old", "WasOnProject_Archive", "WasOnProject (Archive)"];
    // also include any sheet that contains history/was/archive
    var dynamic = ss.getSheets().map(function(sh){ return sh.getName(); }).filter(function(n){ return /history|was|archive/i.test(n); });
    var unique = {};
    (candidates.concat(dynamic)).forEach(function(n){ unique[n] = true; });
    var names = Object.keys(unique);
    // return those that exist and a sample from each
    var found = [];
    names.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) {
        // try case-insensitive
        var maybe = ss.getSheets().find(function(s){ return s.getName().toLowerCase() === (""+name).toLowerCase(); });
        if (maybe) sh = maybe;
      }
      if (!sh) return;
      var lastRow = Math.max(1, sh.getLastRow());
      var lastCol = Math.max(1, sh.getLastColumn());
      var sampleRangeRows = Math.min(500, lastRow); // limit
      var data = [];
      if (sampleRangeRows > 0) {
        try { data = sh.getRange(1,1, sampleRangeRows, lastCol).getDisplayValues(); } catch(e) { data = [["read error: "+String(e)]]; }
      }
      // normalize first two columns (typical structure)
      var normalizedIndex = {};
      for (var r = 0; r < Math.min(500, data.length); r++) {
        var nameCell = data[r] && data[r][0] ? data[r][0] : "";
        var projCell = data[r] && data[r][1] ? data[r][1] : "";
        var nn = normalizeNameForDebug(nameCell);
        if (nn) {
          if (!normalizedIndex[nn]) normalizedIndex[nn] = [];
          normalizedIndex[nn].push({ row: r+1, rawName: nameCell, rawProject: projCell });
        }
      }
      found.push({ sheetName: sh.getName(), gid: sh.getSheetId(), lastRow: lastRow, lastCol: lastCol, sampleRowsCount: Math.min(50, data.length), normalizedIndexSampleCount: Object.keys(normalizedIndex).length, normalizedIndexSample: Object.keys(normalizedIndex).slice(0,10), rawSampleFirstRows: data.slice(0,8) });
    });
    return found;
  });

  // 7) For each assigned student in targetData, check presence in history candidates
  out.info.historyMatching = safeRun("historyMatching", function () {
    var td = out.info.targetData.ok ? out.info.targetData.value : null;
    if (!td) return { note: "no target data to test" };
    var matches = {};
    var histSheets = out.info.historyCandidates.ok ? out.info.historyCandidates.value : [];
    var histBySheet = {};
    // build normalized index for each hist sheet (read limited rows again)
    histSheets.forEach(function(hs) {
      var sh = ss.getSheetByName(hs.sheetName) || ss.getSheets().find(function(s){ return s.getSheetId() === hs.gid; });
      if (!sh) return;
      var lastRow = Math.max(1, sh.getLastRow());
      var lastCol = Math.max(1, sh.getLastColumn());
      var rows = [];
      try { rows = sh.getRange(1,1, Math.min(1000,lastRow), lastCol).getDisplayValues(); } catch(e) { rows = [["read error: " + String(e)]]; }
      var idx = {};
      for (var r = 0; r < rows.length; r++) {
        var nameCell = rows[r] && rows[r][0] ? rows[r][0] : "";
        var projCell = rows[r] && rows[r][1] ? rows[r][1] : "";
        var nn = normalizeNameForDebug(nameCell);
        if (!nn) continue;
        if (!idx[nn]) idx[nn] = [];
        idx[nn].push({ row: r+1, rawName: nameCell, rawProject: projCell });
      }
      histBySheet[hs.sheetName] = { gid: hs.gid, index: idx, rawSampleRows: rows.slice(0,8) };
    });

    // now test each assigned student
    var summary = { totalAssigned: 0, foundInAnyHistory: 0, foundSameProject: 0, notFound: 0, details: [] };
    var projects = td.projects || {};
    Object.keys(projects).forEach(function(projectName) {
      var list = projects[projectName];
      list.forEach(function(entry) {
        summary.totalAssigned++;
        var nn = entry.normalized;
        var foundAny = false;
        var foundSame = false;
        var foundLocations = [];
        Object.keys(histBySheet).forEach(function(sheetName) {
          var idx = histBySheet[sheetName].index;
          if (idx[nn]) {
            foundAny = true;
            idx[nn].forEach(function(hrow) {
              foundLocations.push({ sheet: sheetName, row: hrow.row, rawProject: hrow.rawProject });
              if ((hrow.rawProject || "").toString().trim() === projectName) foundSame = true;
            });
          }
        });
        if (foundAny) summary.foundInAnyHistory++;
        if (foundSame) summary.foundSameProject++;
        if (!foundAny) summary.notFound++;
        summary.details.push({ studentRaw: entry.raw, studentNormalized: nn, project: projectName, foundAny: foundAny, foundSame: foundSame, foundLocations: foundLocations.slice(0,5) });
      });
    });

    return { histBySheetSummaryCount: Object.keys(histBySheet).length, summary: summary };
  });

  // 8) check presence of relevant server functions (so we know what code paths exist)
  out.info.serverFunctions = {};
  var fnNames = ["adminCloseSheet","adminCloseAndCreateNewTarget","closeSheet","commitGroupsToSheet","writeStudentToProject","writeHistory","saveHistory","archiveCurrentSheet","moveToHistory"];
  fnNames.forEach(function(n) {
    try {
      out.info.serverFunctions[n] = (typeof this[n] === "function") || (typeof globalThis !== "undefined" && typeof globalThis[n] === "function") || (typeof this[n] === "function");
    } catch(e) {
      out.info.serverFunctions[n] = "check-failed:" + String(e);
    }
  });

  // 9) permissions / session info (best-effort)
  out.info.session = safeRun("session", function () {
    var s = {};
    try { s.activeUser = Session.getActiveUser() && Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : null; } catch (e) { s.activeUserError = String(e); }
    try { s.effectiveUser = Session.getEffectiveUser() && Session.getEffectiveUser().getEmail ? Session.getEffectiveUser().getEmail() : null; } catch (e) { s.effectiveUserError = String(e); }
    return s;
  });

  // 10) Hints & hypotheses (based on data gathered)
  out.hypotheses = [];
  try {
    // If there are zero history candidate sheets -> clue
    var hc = out.info.historyCandidates.ok ? (out.info.historyCandidates.value || []) : [];
    if (!hc.length) out.hypotheses.push("No history/archive sheet found. The close flow may expect a 'WasOnProject' or 'History' sheet to append into.");
    // If many assigned students not found in any history entries, maybe history was never written
    var histMatch = out.info.historyMatching.ok ? out.info.historyMatching.value.summary : null;
    if (histMatch && histMatch.notFound === histMatch.totalAssigned) {
      out.hypotheses.push("All current assigned students are not present in history sheets. That suggests saving to history wasn't executed or used a different history sheet/format.");
    } else if (histMatch && histMatch.notFound > 0) {
      out.hypotheses.push(histMatch.notFound + " of " + histMatch.totalAssigned + " assigned students are not in any detected history sheet.");
    }
    // If project headers are empty / mismatched
    var th = out.info.targetData.ok ? out.info.targetData.value : null;
    if (th && Object.keys(th.projects).length === 0) out.hypotheses.push("Target sheet row 2 has no project headers (or they are blank). Close flow may rely on headers to map columns into history.");
  } catch(e) {
    out.hypothesesError = String(e);
  }

  // log result
  try {
    Logger.log("=== runCloseSheetDebug result ===\n" + JSON.stringify(out, null, 2));
  } catch (e) {
    Logger.log("runCloseSheetDebug: failed to stringify output: " + String(e));
  }

  return out; // returned to caller (Run -> Logs will show it)
}



/* ============= Web entry ============= */
function doGet(e) {
  // Expects a file named "Index.html" in the project (same as your front-end)
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Project Sign-In');
}

/* EOF */
function FULL_DEBUG() {

  const sheet = getTargetSheet();
  Logger.log("=== SHEET INFO ===");
  Logger.log("Sheet name: " + sheet.getName());

  Logger.log("=== PROJECT_COLUMNS ===");
  Logger.log(JSON.stringify(PROJECT_COLUMNS));

  Logger.log("=== HEADERS ROW 2 ===");

  PROJECT_COLUMNS.forEach(col=>{
    const header = sheet.getRange(2,col+1).getValue();
    Logger.log("Column "+(col+1)+" header: '"+header+"'");
  });

  Logger.log("=== PROJECT COLUMN MAP ===");

  const map = getProjectColumnMapForSheet(sheet);
  Logger.log(JSON.stringify(map));

  Logger.log("=== TEST WRITE PERMISSION ===");

  try{

    const testCell = sheet.getRange(1,1);
    const original = testCell.getValue();

    testCell.setValue("DEBUG_TEST");
    Utilities.sleep(500);
    testCell.setValue(original);

    Logger.log("Write test: SUCCESS");

  }catch(e){

    Logger.log("Write test FAILED: "+e);

  }

  Logger.log("=== CONDITIONAL FORMAT RULES ===");

  const rules = sheet.getConditionalFormatRules();

  Logger.log("Total rules: "+rules.length);

  rules.forEach((r,i)=>{

    const ranges = r.getRanges().map(rr=>rr.getA1Notation());
    Logger.log("Rule "+i+" ranges: "+ranges.join(", "));

  });

  Logger.log("=== PROJECT BLOCK CHECK ===");

  const START_ROW = 3;
  const CHECK_ROWS = 5;

  Object.entries(map).forEach(([project,col])=>{

    const vals = sheet.getRange(START_ROW,col+2,CHECK_ROWS,1).getValues().flat();

    Logger.log(project+" grade column sample: "+JSON.stringify(vals));

  });

  Logger.log("=== DEBUG COMPLETE ===");

}
