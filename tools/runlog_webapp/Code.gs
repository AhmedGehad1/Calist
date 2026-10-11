/**
 * Calist run log: the receiving end.
 *
 * A web app bound to the owner's Google Sheet. Official Calist builds POST one small JSON record
 * per event (a Build press, a crash, a PIN lockout, a launch). This script appends it to that
 * kind's tab, numbers each new PC, and answers with the record's own id. That answer is the
 * receipt: Calist deletes its local copy only when the id comes back. If an answer is lost, the
 * same record arrives again, and the Record id column turns it into a "dup" answer instead of a
 * second row.
 *
 * Nothing here is secret. The token lives in Script Properties, never in this file, which sits in
 * Calist's public repository so the script can always be redeployed.
 *
 * SET-UP (once)
 *  1. Open the sheet > Extensions > Apps Script. Replace Code.gs with this file. In Project
 *     Settings tick "Show appsscript.json manifest file in editor" and replace appsscript.json
 *     with the one beside this file.
 *  2. Project Settings > Script Properties > add TOKENS = a long random string
 *     (PowerShell: [guid]::NewGuid().ToString('N')). Several tokens may be listed,
 *     comma-separated, so a new release can carry a new one while older copies keep working.
 *  3. Choose `setup` in the function list and Run it once. Google asks you to authorise it.
 *  4. Deploy > New deployment > gear > Web app. Execute as: Me. Who has access: Anyone.
 *     Copy the Web app URL: https://script.google.com/macros/s/<DEPLOYMENT-ID>/exec
 *  5. Share the sheet with its readers as Viewer and set General access to Restricted.
 *     An editor could open this script and read the token.
 *
 * CHANGING THIS SCRIPT LATER
 *  Deploy > Manage deployments > pencil > Version: New version > Deploy. Editing the existing
 *  deployment keeps the URL every installed Calist was built with; a NEW deployment gets a new
 *  URL and silently strands them.
 *
 * KILL SWITCH AND LIMITS (Script Properties, read on every request)
 *  STOP=true          every Calist stops sending and queuing for STOP_HOURS (default 24), then
 *                     asks again
 *  MIN_INTERVAL_S=n   seconds between records from one PC (default 3)
 *  CAP_PER_DAY=n      rows accepted per day from all PCs together (default 2000); the rest are
 *                     told to try again later
 */

const SCHEMA = 1;
const MAX_BODY = 16 * 1024;
const RECORD_ID = /^[0-9a-f]{32}$/;
const INSTALL_ID = /^[0-9a-f]{16}$/;

// Each tab's columns, in order: [header, format]. T text, N number, D date-time. Rows are written
// by header name, so the columns can be reordered (or new ones added) in the sheet itself.
const TABS = {
  Builds: [
    ['Built at', 'D'], ['Received', 'D'], ['Outcome', 'T'], ['Devices', 'N'], ['Rows', 'N'],
    ['Took (s)', 'N'], ['Took', 'T'], ['Reading (s)', 'N'], ['Writing (s)', 'N'],
    ['Saved to', 'T'], ['User', 'T'], ['PC #', 'N'], ['Computer', 'T'], ['Errors', 'N'],
    ['Unknown code', 'N'], ['Unsupported', 'N'], ['Left out', 'N'], ['Copies', 'N'],
    ['Duplicates removed', 'N'], ['Unknown codes', 'T'], ['Version', 'T'], ['Shape', 'T'],
    ['CPUs', 'N'], ['RAM (GB)', 'N'], ['Error', 'T'], ['Dropped offline', 'T'],
    ['Record id', 'T'],
  ],
  PCs: [
    ['PC #', 'N'], ['Install id', 'T'], ['Computer', 'T'], ['Users seen', 'T'],
    ['First seen', 'D'], ['Last seen', 'D'], ['Last version', 'T'], ['Last shape', 'T'],
    ['Last send error', 'T'],
  ],
  Crashes: [
    ['Time', 'D'], ['Received', 'D'], ['PC #', 'N'], ['Computer', 'T'], ['User', 'T'],
    ['Version', 'T'], ['Where', 'T'], ['Type', 'T'], ['Message', 'T'], ['Frames', 'T'],
    ['Record id', 'T'],
  ],
  Lockouts: [
    ['Time', 'D'], ['Received', 'D'], ['PC #', 'N'], ['Computer', 'T'], ['User', 'T'],
    ['Failures', 'N'], ['Wait (s)', 'N'], ['Version', 'T'], ['Record id', 'T'],
  ],
  Launches: [
    ['Time', 'D'], ['Received', 'D'], ['PC #', 'N'], ['Computer', 'T'], ['User', 'T'],
    ['Version', 'T'], ['Shape', 'T'], ['Record id', 'T'],
  ],
};
const TAB_OF_KIND = { build: 'Builds', crash: 'Crashes', lockout: 'Lockouts', launch: 'Launches' };

/** Run once from the editor: creates the tabs, their headers and formats, and the time zone. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Africa/Cairo');
  Object.keys(TABS).forEach(name => tab_(ss, name));
}

/** A health check for a browser: a private window should show this, not a Google sign-in. */
function doGet() {
  return json_({ ok: true, service: 'calist-runlog', v: SCHEMA });
}

function doPost(e) {
  let id = null;
  try {
    const body = (e && e.postData && e.postData.contents) || '';
    if (body.length > MAX_BODY) return reject_(null, 'too large');
    let rec;
    try {
      rec = JSON.parse(body);
    } catch (err) {
      return reject_(null, 'not JSON');
    }
    if (!rec || typeof rec !== 'object') return reject_(null, 'not an object');
    id = typeof rec.id === 'string' && RECORD_ID.test(rec.id) ? rec.id : null;
    if (!id) return reject_(null, 'bad id');
    if (rec.v !== SCHEMA) return reject_(id, 'schema');
    const tabName = TAB_OF_KIND[rec.kind];
    if (!tabName) return reject_(id, 'kind');
    // The token before the install id, always: `python runlog.py --check`
    // sends a bad install id to learn whether the token is known (refused
    // for 'install' means it is), and that way never writes a row.
    if (!tokenOk_(rec.token)) return reject_(id, 'token');
    const pc = rec.pc || {};
    if (typeof pc.install !== 'string' || !INSTALL_ID.test(pc.install)) {
      return reject_(id, 'install');
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return retry_(id, 'busy');
    try {
      const ss = SpreadsheetApp.getActive();
      const sheet = tab_(ss, tabName);
      // A record the sheet already holds is answered as received: its first answer was lost.
      if (findRow_(sheet, 'Record id', id)) {
        return answer_(id, { dup: true, pc: pcNumber_(ss, pc.install) });
      }
      if (!underDailyCap_()) return retry_(id, 'daily cap');
      const number = upsertPc_(ss, rec);
      appendByHeader_(sheet, rowFor_(rec, number));
      SpreadsheetApp.flush();
      if (!findRow_(sheet, 'Record id', id)) return retry_(id, 'not written');
      countToday_();
      return answer_(id, { dup: false, pc: number });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return retry_(id, String((err && err.message) || err).slice(0, 200));
  }
}

// ── answers ─────────────────────────────────────────────────────────────────────────────────────
// Every answer carries the record's id and the current switches, so a client can tell exactly
// which record was confirmed, and a STOP reaches every PC whatever happened to its record.

function answer_(id, extra) {
  return json_(Object.assign({ ok: true, id: id, dup: false, pc: null, cfg: config_() }, extra));
}

function retry_(id, why) {
  return json_({ ok: false, retry: true, id: id, why: why, cfg: config_() });
}

function reject_(id, why) {
  return json_({ ok: false, reject: true, id: id, why: why, cfg: config_() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function config_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  return {
    stop: String(p.STOP || '').trim().toLowerCase() === 'true',
    stop_hours: number_(p.STOP_HOURS, 24),
    min_interval_s: number_(p.MIN_INTERVAL_S, 3),
  };
}

function tokenOk_(token) {
  if (typeof token !== 'string' || !token) return false;
  const tokens = String(PropertiesService.getScriptProperties().getProperty('TOKENS') || '')
    .split(',').map(s => s.trim()).filter(Boolean);
  return tokens.indexOf(token) !== -1;
}

// ── rows ────────────────────────────────────────────────────────────────────────────────────────

function rowFor_(rec, number) {
  const pc = rec.pc || {};
  const app = rec.app || {};
  const d = rec.data || {};
  const row = {
    'Received': new Date(),
    'PC #': number,
    'Computer': text_(pc.name, 64),
    'User': text_(pc.user, 64),
    'Version': text_(app.version, 20),
    'Shape': text_(app.shape, 10),
    'Record id': rec.id,
  };
  const when = when_(rec.at);
  if (rec.kind === 'build') {
    return Object.assign(row, {
      'Built at': when,
      'Outcome': text_(d.outcome, 20),
      'Devices': count_(d.devices),
      'Rows': count_(d.rows),
      'Took (s)': seconds_(d.took_s),
      'Took': took_(d.took_s),
      'Reading (s)': seconds_(d.read_s),
      'Writing (s)': seconds_(d.write_s),
      'Saved to': text_(d.saved_to, 400),
      'Errors': count_(d.errors),
      'Unknown code': count_(d.unknown),
      'Unsupported': count_(d.unsupported),
      'Left out': count_(d.left_out),
      'Copies': count_(d.copies),
      'Duplicates removed': count_(d.dupes_removed),
      'Unknown codes': text_(d.unknown_codes, 120),
      'CPUs': count_(app.cpus),
      'RAM (GB)': seconds_(app.ram_gb),
      'Error': text_(d.error, 400),
      'Dropped offline': dropped_(rec.health),
    });
  }
  if (rec.kind === 'crash') {
    const frames = Array.isArray(d.frames) ? d.frames.join('\n') : d.frames;
    return Object.assign(row, {
      'Time': when,
      'Where': text_(d.where, 20),
      'Type': text_(d.type, 80),
      'Message': text_(d.message, 400),
      'Frames': text_(frames, 2000),
    });
  }
  if (rec.kind === 'lockout') {
    return Object.assign(row, {
      'Time': when, 'Failures': count_(d.failures), 'Wait (s)': count_(d.wait_s),
    });
  }
  return Object.assign(row, { 'Time': when });                // launch
}

/** Each new install gets the next PC number; a known one has its last-seen details refreshed. */
function upsertPc_(ss, rec) {
  const sheet = tab_(ss, 'PCs');
  const pc = rec.pc || {};
  const app = rec.app || {};
  const name = text_(pc.name, 64);
  const user = text_(pc.user, 64);
  const lastError = text_((rec.health || {}).last_error, 120);
  const now = new Date();
  const found = findRow_(sheet, 'Install id', pc.install);
  if (!found) {
    const number = nextPcNumber_(sheet);
    appendByHeader_(sheet, {
      'PC #': number, 'Install id': pc.install, 'Computer': name, 'Users seen': user,
      'First seen': now, 'Last seen': now, 'Last version': text_(app.version, 20),
      'Last shape': text_(app.shape, 10), 'Last send error': lastError,
    });
    return number;
  }
  const headers = headers_(sheet);
  const range = sheet.getRange(found, 1, 1, headers.length);
  const values = range.getValues()[0];
  const at = h => headers.indexOf(h);
  const users = String(values[at('Users seen')] || '').split(', ').filter(Boolean);
  if (user && users.indexOf(user) === -1) users.push(user);
  const fresh = {
    'Computer': name, 'Users seen': users.join(', ').slice(0, 300), 'Last seen': now,
    'Last version': text_(app.version, 20), 'Last shape': text_(app.shape, 10),
    'Last send error': lastError,
  };
  range.setValues([values.map((v, i) => cell_(fresh.hasOwnProperty(headers[i])
    ? fresh[headers[i]] : v))]);
  return Number(values[at('PC #')]) || null;
}

function pcNumber_(ss, install) {
  const sheet = tab_(ss, 'PCs');
  const row = findRow_(sheet, 'Install id', install);
  if (!row) return null;
  const col = columnOf_(sheet, 'PC #');
  return col ? Number(sheet.getRange(row, col).getValue()) || null : null;
}

function nextPcNumber_(sheet) {
  const col = columnOf_(sheet, 'PC #');
  const last = sheet.getLastRow();
  if (!col || last < 2) return 1;
  const numbers = sheet.getRange(2, col, last - 1, 1).getValues().map(r => Number(r[0]) || 0);
  return Math.max(0, ...numbers) + 1;
}

// ── the daily cap (the only bound that holds if the URL and token are dug out of the exe) ──────

function underDailyCap_() {
  const cap = number_(PropertiesService.getScriptProperties().getProperty('CAP_PER_DAY'), 2000);
  return capState_().n < cap;
}

function countToday_() {
  const state = capState_();
  state.n += 1;
  PropertiesService.getDocumentProperties().setProperty('CAP_STATE', JSON.stringify(state));
}

function capState_() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let state;
  try {
    state = JSON.parse(PropertiesService.getDocumentProperties().getProperty('CAP_STATE') || '{}');
  } catch (err) {
    state = {};
  }
  return state.day === today ? state : { day: today, n: 0 };
}

// ── the sheet ───────────────────────────────────────────────────────────────────────────────────

function tab_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const columns = TABS[name];
    sheet.getRange(1, 1, 1, columns.length).setValues([columns.map(c => c[0])])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    columns.forEach(([, format], i) => {
      const range = sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1);
      if (format === 'T') range.setNumberFormat('@');
      else if (format === 'D') range.setNumberFormat('yyyy-mm-dd hh:mm');
    });
  }
  return sheet;
}

function headers_(sheet) {
  const width = sheet.getLastColumn();
  return width ? sheet.getRange(1, 1, 1, width).getValues()[0].map(String) : [];
}

function columnOf_(sheet, header) {
  return headers_(sheet).indexOf(header) + 1;
}

/** The row holding `value` in the `header` column, or 0. Ids are long random hex, so a plain
 *  substring match cannot hit the wrong row and tolerates a text marker on the cell. */
function findRow_(sheet, header, value) {
  const col = columnOf_(sheet, header);
  const last = sheet.getLastRow();
  if (!col || last < 2) return 0;
  const cell = sheet.getRange(2, col, last - 1, 1).createTextFinder(value).findNext();
  return cell ? cell.getRow() : 0;
}

function appendByHeader_(sheet, values) {
  sheet.appendRow(headers_(sheet).map(h => cell_(values.hasOwnProperty(h) ? values[h] : '')));
}

/** Numbers and dates go in as themselves. Every string gets Sheets' text marker, so nothing a
 *  client sends can become a formula (=, +, -, @) or be re-read as a number or a date. */
function cell_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date || typeof v === 'number') return v;
  const s = String(v);
  return s ? "'" + s : '';
}

// ── field coercion: anything not of the expected shape becomes blank ─────────────────────────────

function text_(v, max) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/[\u0000-\u0008\u000B-\u001F]/g, ' ').slice(0, max);
}

function count_(v) {
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 ? Math.round(n) : '';
}

function seconds_(v) {
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 ? Math.round(n * 10) / 10 : '';
}

function number_(v, fallback) {
  const n = Number(v);
  return v !== null && v !== undefined && v !== '' && Number.isFinite(n) ? n : fallback;
}

function when_(iso) {
  if (typeof iso !== 'string') return '';
  const d = new Date(iso);
  return isNaN(d.getTime()) ? text_(iso, 40) : d;
}

/** "45.2s", "1m 04s", "1h 02m": the readable twin of Took (s). */
function took_(v) {
  const s = Number(v);
  if (!Number.isFinite(s) || s < 0) return '';
  if (s < 60) return s.toFixed(1) + 's';
  const whole = Math.round(s);
  const pad = n => String(n).padStart(2, '0');
  if (whole < 3600) return Math.floor(whole / 60) + 'm ' + pad(whole % 60) + 's';
  return Math.floor(whole / 3600) + 'h ' + pad(Math.floor((whole % 3600) / 60)) + 'm';
}

/** "build 3, crash 1": records this PC had to drop while it could not reach the sheet. */
function dropped_(health) {
  const dropped = (health && health.dropped) || {};
  return Object.keys(dropped)
    .filter(k => Number(dropped[k]) > 0)
    .map(k => text_(k, 10) + ' ' + Math.round(Number(dropped[k])))
    .join(', ');
}
