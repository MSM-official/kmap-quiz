/**
 * K-Map Quiz: everything in Google Apps Script.
 * Files in this project:  Code.gs (this file)  +  index (HTML file named "index").
 * Deploy → Manage deployments → Edit → Version: New version → Deploy.
 *   Students:  your /exec link
 *   Teacher (paste result codes):  your /exec link + ?teacher=yes
 *   Sheet link check:  your /exec link + ?status=yes
 * Marks go to the Sheet this script is attached to, or (standalone script)
 * to a sheet named "K-Map Quiz Marks" that it creates in your Drive.
 */
const SHEET_NAME = 'Marks';
const LEVELS = ['Beginner', 'Easy', 'Medium', 'Hard', 'Expert'];
const Q_PER_LEVEL = 3;
const QMAX = 10;
const NQ = LEVELS.length * Q_PER_LEVEL;
const TOTAL = NQ * QMAX;

const HEADER = ['Student ID', 'Name / Roll']
  .concat(LEVELS.map((l, i) => 'L' + (i + 1) + ' ' + l))
  .concat(['Total', 'Out of', '%', 'Grade', 'Status', 'Source', 'Started', 'Last update'])
  .concat(Array.from({ length: NQ }, (_, i) => 'Q' + (i + 1)));

function hash_(s) {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0;
  return h.toString(36);
}

function grade_(p) {
  return p >= 90 ? 'A+' : p >= 80 ? 'A' : p >= 70 ? 'B' : p >= 60 ? 'C' : p >= 50 ? 'D' : 'F';
}

// Works both when the script is opened from a Sheet (Extensions → Apps Script)
// and as a standalone project: then it creates a sheet "K-Map Quiz Marks" in your Drive.
function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('K-Map Quiz Marks');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#eef2ff');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(2);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.status) {
    try { sheet_(); return json_({ ok: true, msg: 'K-Map quiz is running.', sheet: ss_().getUrl() }); }
    catch (err) { return json_({ ok: false, error: String(err) }); }
  }
  const t = HtmlService.createTemplateFromFile('index');
  t.teacher = p.teacher === 'yes' ? 'yes' : 'no';
  return t.evaluate()
    .setTitle('K-Map Quiz')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called from the quiz page with google.script.run
function submitMarks(body) {
  return handle_(body);
}

function doPost(e) {
  try { return json_(handle_(JSON.parse(e.postData.contents))); }
  catch (err) { return json_({ ok: false, error: String(err) }); }
}

function handle_(body) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (typeof body.p !== 'string' || hash_('eed-kmap|' + body.p) !== body.s) return ({ ok: false, error: 'bad signature' });
    const d = JSON.parse(body.p);
    if (!Array.isArray(d.q) || d.q.length !== NQ) return ({ ok: false, error: 'bad data' });

    const q = d.q.map(x => (typeof x === 'number' ? Math.max(-1, Math.min(QMAX, Math.round(x))) : -1));
    const sid = String(d.sid || '').slice(0, 60);
    const name = String(d.name || '').replace(/^[=+\-@]/, "'$&").slice(0, 80);
    if (!sid) return ({ ok: false, error: 'no id' });

    const levelMarks = LEVELS.map((_, i) => q.slice(i * Q_PER_LEVEL, (i + 1) * Q_PER_LEVEL).reduce((s, x) => s + Math.max(0, x), 0));
    const total = levelMarks.reduce((a, b) => a + b, 0);
    const answered = q.filter(x => x >= 0).length;
    const finished = !!d.finished || answered >= NQ;
    const pct = Math.round((total / TOTAL) * 100);
    let status = 'Finished';
    if (!finished) {
      const lv = Math.floor(answered / Q_PER_LEVEL) + 1, qq = (answered % Q_PER_LEVEL) + 1;
      status = 'On L' + lv + ' Q' + qq;
    }

    const sh = sheet_();
    const last = sh.getLastRow();
    let row = -1;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === sid) { row = i + 2; break; }
    }
    let started = d.started || new Date();
    if (row > 0) {
      const prev = sh.getRange(row, 1, 1, HEADER.length).getValues()[0];
      if (prev[HEADER.indexOf('Status')] === 'Finished' && d.src !== 'code') return ({ ok: true, note: 'already finished' });
      started = prev[HEADER.indexOf('Started')] || started;
    }
    const values = [sid, name].concat(levelMarks).concat([
      total, TOTAL, pct, finished ? grade_(pct) : '', status, d.src === 'code' ? 'result code' : 'online',
      started ? new Date(started) : '', new Date()
    ]).concat(q.map(x => (x < 0 ? '' : x)));

    if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
    else sh.appendRow(values);
    return ({ ok: true });
  } catch (err) {
    return ({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
