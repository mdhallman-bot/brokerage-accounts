/** Unit 3 Day 5 — How Do I Actually Start Investing?
 * Script Properties required:
 *   SHEET_ID
 *
 * Student identifier: teacher-assigned pseudonymous class code only.
 * No names, emails, financial account numbers, or real trading information.
 *
 * Keep SHEET_ID in Apps Script Properties. Never commit its value here.
 */
const STATE_SHEET = 'U3D5_State';

function doGet(e) {
  if (e && e.parameter && e.parameter.action) return apiGet_(e);
  return HtmlService.createHtmlOutput(
    '<h2>Unit 3 Day 5 backend is running.</h2><p>The student tool is hosted separately on GitHub Pages.</p>'
  ).setTitle('Day 5 — How Do I Actually Start Investing?');
}

function apiGet_(e) {
  const callback = String(e.parameter.callback || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback.');
  }
  let p = {};
  try { p = JSON.parse(e.parameter.payload || '{}'); } catch (err) {}
  let result;
  try {
    switch (String(e.parameter.action || '')) {
      case 'save_state': result = saveState(p.studentCode, p.state); break;
      case 'load_state': result = loadState(p.studentCode); break;
      default: result = {error:'Unknown action.'};
    }
  } catch (err) {
    console.error(err);
    result = {error:String(err && err.message ? err.message : err)};
  }
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function saveState(studentCode, state) {
  const code = cleanCode_(studentCode);
  if (!code) throw new Error('Missing or invalid class code.');
  const sh = stateSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const json = JSON.stringify(state || {});
    const row = findLatestCodeRow_(sh, code);
    if (row) sh.getRange(row, 1, 1, 3).setValues([[code, json, now]]);
    else sh.appendRow([code, json, now]);
    SpreadsheetApp.flush();
    return {ok:true, savedAt:Utilities.formatDate(now, Session.getScriptTimeZone(), 'h:mm:ss a')};
  } finally {
    lock.releaseLock();
  }
}

function loadState(studentCode) {
  const code = cleanCode_(studentCode);
  if (!code) throw new Error('Missing or invalid class code.');
  const sh = stateSheet_();
  const row = findLatestCodeRow_(sh, code);
  if (!row) return {ok:true, found:false};
  const raw = sh.getRange(row, 2).getValue();
  if (!raw) return {ok:true, found:false};
  let state = {};
  try { state = JSON.parse(raw); }
  catch (err) { throw new Error('Saved work could not be read.'); }
  return {ok:true, found:true, state:state};
}

function stateSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('SHEET_ID is not configured in Script Properties.');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(STATE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATE_SHEET);
    sh.getRange(1,1,1,3).setValues([['StudentCode','StateJSON','UpdatedAt']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findLatestCodeRow_(sh, code) {
  const normalized = normalizeCode_(code);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const values = sh.getRange(2,1,last-1,1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (normalizeCode_(values[i][0]) === normalized) return i + 2;
  }
  return 0;
}

function normalizeCode_(value) {
  return String(value || '').trim().toUpperCase();
}

function cleanCode_(value) {
  return normalizeCode_(value).replace(/[^A-Z0-9_-]/g,'').slice(0,30);
}
