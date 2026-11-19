/**
 * K-Map Quiz: marks collector for Google Sheets.
 * Paste this into Extensions → Apps Script of your Google Sheet,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Each student gets ONE row that updates after every question.
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

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

function doGet() {
  return json_({ ok: true, msg: 'K-Map marks endpoint is running.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const body = JSON.parse(e.postData.contents);
    if (typeof body.p !== 'string' || hash_('eed-kmap|' + body.p) !== body.s) return json_({ ok: false, error: 'bad signature' });
    const d = JSON.parse(body.p);
    if (!Array.isArray(d.q) || d.q.length !== NQ) return json_({ ok: false, error: 'bad data' });

    const q = d.q.map(x => (typeof x === 'number' ? Math.max(-1, Math.min(QMAX, Math.round(x))) : -1));
    const sid = String(d.sid || '').slice(0, 60);
    const name = String(d.name || '').replace(/^[=+\-@]/, "'$&").slice(0, 80);
    if (!sid) return json_({ ok: false, error: 'no id' });

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
      if (prev[HEADER.indexOf('Status')] === 'Finished' && d.src !== 'code') return json_({ ok: true, note: 'already finished' });
      started = prev[HEADER.indexOf('Started')] || started;
    }
    const values = [sid, name].concat(levelMarks).concat([
      total, TOTAL, pct, finished ? grade_(pct) : '', status, d.src === 'code' ? 'result code' : 'online',
      started ? new Date(started) : '', new Date()
    ]).concat(q.map(x => (x < 0 ? '' : x)));

    if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
    else sh.appendRow(values);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
