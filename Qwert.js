/***************************************************************
 * Code.gs - Fix: restored swap-name logic + full backend + debug + doGet
 *
 * Paste this entire file into Apps Script (replace existing).
 * Set SPREADSHEET_ID = "" to use the active spreadsheet you're editing in the browser.
 **************************************************************/

/* ============= CONFIG ============= */
const SPREADSHEET_ID = ""; // "" => use active spreadsheet; otherwise set explicit id
const SHEET_GIDS = [681181988, 1493058526, 1220633850];
const PROJECT_COLUMNS = [1, 3, 5, 7, 9, 11]; // zero-based offsets where project names appear in row 2
const SOURCE_INFO_SHEET_GID = 1446473767;
const MAX_STUDENTS = 16;
const GRADE_ORDER = { "2A": 1, "2B": 2, "3A": 3, "3B": 4 };

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
function buildHistoryMapFromSheets() {
  const ss = getSpreadsheet();
  const history = {};
  SHEET_GIDS.forEach((gid, sheetIndex) => {
    if (!gid || gid === 0) return;
    const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid));
    if (!sheet) return;
    PROJECT_COLUMNS.forEach(col => {
      const projectName = sheet.getRange(2, col + 1).getValue();
      if (!projectName) return;
      const values = sheet.getRange(3, col + 1, MAX_STUDENTS, 1).getValues(); // rows 3..(3+MAX_STUDENTS-1)
      values.forEach(r => {
        let name = (r[0] || "").toString().trim();
        if (!name) return;
        // apply sheet-specific swap heuristic
        name = swapNameIfNeeded(name, sheetIndex);
        const key1 = normalizeName(name);
        history[key1] = history[key1] || {};
        history[key1][projectName] = true;
        // also store reversed variant for safety
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
  const gLower = (gender || "").toString().toLowerCase();
  if (gLower) {
    if (gLower.indexOf("f") === 0) target.getRange(writeRow, col + 1).setBackground("#ffebee");
    else if (gLower.indexOf("m") === 0) target.getRange(writeRow, col + 1).setBackground("#e8f0fe");
  }
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
  return newSheet;
}

function closeCreateNewTarget() {
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
  return { ok: true, message: "Closed and created new target " + newName, commitLog };
}

/* ============= Balanced write & formatting ============= */
function setGradeConditionalFormatting(sheet, col) {
  const rules = sheet.getConditionalFormatRules();
  const filteredRules = rules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() === col + 2);
  });
  const startRow = 3; const maxRows = MAX_STUDENTS;
  const gradeRange = sheet.getRange(startRow, col + 2, maxRows, 1);
  const newRules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('2A').setBackground('#b6d7a8').setRanges([gradeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('2B').setBackground('#a4c2f4').setRanges([gradeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('3A').setBackground('#ffe599').setRanges([gradeRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('3B').setBackground('#ea9999').setRanges([gradeRange]).build()
  ];
  sheet.setConditionalFormatRules([...filteredRules, ...newRules]);
}
function applyGenderColors(sheet, col, startRow, maxRows) {
  const names = sheet.getRange(startRow, col + 1, maxRows, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    const fullName = names[i][0];
    const cell = sheet.getRange(startRow + i, col + 1);
    if (!fullName) { cell.setBackground(null); continue; }
    let gender = "";
    try { gender = (getStudentGradeGender(fullName).gender || "").toString().toLowerCase(); } catch (e) { gender = ""; }
    const color = (gender.indexOf("f") === 0) ? "#ff000033" : (gender.indexOf("m") === 0 ? "#0000ff33" : null);
    if (color) cell.setBackground(color); else cell.setBackground(null);
  }
}
function writeBalancedGroups(groups) {
  const sheet = getTargetSheet();
  const map = getProjectColumnMapForSheet(sheet);
  const startRow = 3; const maxRows = MAX_STUDENTS;
  function gradeSortKey(grade) {
    if (!grade) return Number.MAX_SAFE_INTEGER;
    grade = grade.toString().trim();
    if (GRADE_ORDER && GRADE_ORDER[grade] !== undefined) return GRADE_ORDER[grade] * 1000;
    const m = grade.match(/^(\d+)\s*([A-Za-z])?$/);
    if (m) { const num = Number(m[1]); const letter = (m[2] || "A").toUpperCase(); return num * 1000 + (letter.charCodeAt(0) - 64); }
    return Number.MAX_SAFE_INTEGER - 1;
  }
  Object.entries(groups).forEach(([project, students]) => {
    const col = map[project];
    if (col === undefined) { Logger.log("writeBalancedGroups: unknown project mapping for '%s' - skipping", project); return; }
    sheet.getRange(startRow, col + 1, maxRows, 2).clearContent();
    students.slice(0, maxRows).forEach((s, i) => { sheet.getRange(startRow + i, col + 1).setValue(s.name); sheet.getRange(startRow + i, col + 2).setValue(s.grade); });
    const blockRange = sheet.getRange(startRow, col + 1, maxRows, 2);
    const blockValues = blockRange.getValues();
    const populated = [];
    for (let i = 0; i < blockValues.length; i++) {
      const row = blockValues[i];
      const name = (row[0] || "").toString().trim();
      const grade = (row[1] || "").toString().trim();
      if (name) populated.push({ name, grade, origIndex: i });
    }
    populated.sort((a,b) => { const ka = gradeSortKey(a.grade); const kb = gradeSortKey(b.grade); if (ka !== kb) return ka - kb; return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }); });
    const newBlock = [];
    for (let i = 0; i < populated.length; i++) newBlock.push([populated[i].name, populated[i].grade]);
    while (newBlock.length < maxRows) newBlock.push(["",""]);
    blockRange.setValues(newBlock);
    setGradeConditionalFormatting(sheet, col);
  });
  const mapEntries = getProjectColumnMapForSheet(sheet);
  Object.entries(mapEntries).forEach(([project, col]) => { applyGenderColors(sheet, col, startRow, maxRows); });
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
function DEBUG_FULL_SYSTEM(studentName, projectName) {
  const report = {};
  try {
    const ss = getSpreadsheet();
    report.spreadsheetId = ss.getId();
    report.spreadsheetName = ss.getName();
    report.sheets = ss.getSheets().map(s => ({ name: s.getName(), gid: s.getSheetId(), lastRow: s.getLastRow(), lastCol: s.getLastColumn() }));
    const cfg = ss.getSheetByName("ScriptConfig");
    report.configExists = !!cfg;
    report.configB1 = cfg ? cfg.getRange(1,2).getValue() : null;
    report.targetGidUsed = getTargetSheetId();
    try {
      const target = getTargetSheet();
      report.activeTargetName = target.getName();
      const data = target.getDataRange().getValues();
      report.targetRowCount = data.length;
      report.targetColCount = data[0] ? data[0].length : 0;
      report.targetRow2 = data[1] || [];
      const projMap = getProjectColumnMapForSheet(target);
      report.projectMap = projMap;
      const startRow = 3;
      report.projectOccupancy = {};
      report.studentLocations = [];
      Object.entries(projMap).forEach(([proj, col]) => {
        const names = target.getRange(startRow, col + 1, MAX_STUDENTS, 1).getValues().map(r => (r[0]||"").toString().trim());
        const grades = target.getRange(startRow, col + 2, MAX_STUDENTS, 1).getValues().map(r => (r[0]||"").toString().trim());
        const nameCount = names.reduce((acc,n)=>acc + (n?1:0),0);
        report.projectOccupancy[proj] = { col: col, nameCount: nameCount, sampleNames: names.slice(0,10), sampleGrades: grades.slice(0,10) };
        if (studentName) {
          names.forEach((n, idx) => {
            if (normalizeName(n) === normalizeName(studentName)) report.studentLocations.push({ project: proj, row: startRow + idx, col: col, gradeCell: grades[idx] || null });
          });
        }
      });
    } catch (e) { report.activeTargetError = e.message || String(e); }
    const hist = buildHistoryMapFromSheets();
    report.historySampleKeys = Object.keys(hist).slice(0,20);
    if (studentName) report.historyForStudent = hist[normalizeName(studentName)] || {};
    report.status = "ok";
  } catch (err) { report.error = err.message || String(err); }
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
function DEBUG_FULL_SYSTEM_RUN() { const studentName = "Biskup Martin"; const projectName = "Bi- Nikola Krýdová"; return DEBUG_FULL_SYSTEM(studentName, projectName); }
function getActiveTargetInfo() { try { const ss = getSpreadsheet(); const cfg = getConfigSheet(); const gid = cfg.getRange(1,2).getValue(); const sheet = ss.getSheets().find(s => s.getSheetId() === Number(gid)); return { ok:true, usingSpreadsheetId: ss.getId(), configuredGid: gid, sheetName: sheet ? sheet.getName() : null }; } catch (e) { return { ok:false, message: e.message || String(e) }; } }

/* ============= Web entry ============= */
function doGet(e) {
  // Expects a file named "Index.html" in the project (same as your front-end)
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Project Sign-In');
}

/* EOF */
