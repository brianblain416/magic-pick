/**
 * Magic Pick — takes library (Google Apps Script web app)
 *
 * Stores recorded vocal and guitar-solo takes in a Drive folder called
 * "Magic Pick Takes" and keeps the list in this spreadsheet.
 *   Players sheet: handle, PIN hash, created
 *   Takes sheet:   one row per take
 *   Settings sheet: the admin key (the director's key that can hear private takes)
 *
 * Nothing is ever deleted: a removed take is marked removed and its file is
 * moved into a "_to_delete" folder for you to empty yourself.
 *
 * Setup: run setup() once from the editor, then Deploy > New deployment > Web app,
 * Execute as: Me, Who has access: Anyone.
 */

const FOLDER_NAME = 'Magic Pick Takes';
const TAKE_COLS = ['id', 'handle', 'act', 'song', 'kind', 'private', 'created', 'dur', 'offset', 'fileId', 'mime', 'removed'];
const MAX_BYTES = 30 * 1024 * 1024;

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_('Players', ['handle', 'pinHash', 'created']);
  sheet_('Takes', TAKE_COLS);
  const st = sheet_('Settings', ['key', 'value']);
  if (!setting_('ADMIN_KEY')) {
    const key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    st.appendRow(['ADMIN_KEY', key]);
    st.appendRow(['SALT', Utilities.getUuid()]);
  }
  folder_();
  Logger.log('Admin key: ' + setting_('ADMIN_KEY'));
  SpreadsheetApp.flush();
  return 'ok';
}

function doGet() { return out_({ ok: true, app: 'magic-pick-takes' }); }

function doPost(e) {
  try {
    const q = JSON.parse(e.postData.contents);
    const admin = !!q.admin && q.admin === setting_('ADMIN_KEY');
    switch (q.action) {
      case 'list': return out_(list_(q, admin));
      case 'login': return out_(login_(q));
      case 'upload': return out_(upload_(q));
      case 'get': return out_(get_(q, admin));
      case 'setPrivate': return out_(setPrivate_(q, admin));
      case 'remove': return out_(remove_(q, admin));
      default: return out_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// ---------- actions ----------
function list_(q, admin) {
  const who = verify_(q.handle, q.pin, false);
  const takes = rows_('Takes').filter(t => !t.removed && (!t.private || admin || (who && sameHandle_(t.handle, who))))
    .map(publicTake_);
  const players = rows_('Players').map(p => p.handle).sort();
  return { ok: true, takes, players, me: who || null, admin };
}

function login_(q) {
  const who = verify_(q.handle, q.pin, true);
  return { ok: true, handle: who };
}

function upload_(q) {
  const who = verify_(q.handle, q.pin, true);
  if (['vocal', 'solo'].indexOf(q.kind) < 0) throw new Error('bad_kind');
  if (!/^\d{2}$/.test(String(q.song)) || [1, 2].indexOf(Number(q.act)) < 0) throw new Error('bad_song');
  const bytes = Utilities.base64Decode(q.data);
  if (bytes.length > MAX_BYTES) throw new Error('too_big');
  const id = Utilities.getUuid();
  const created = new Date().toISOString();
  const name = ['A' + q.act, q.song, q.kind, who, created.slice(0, 16).replace(/[:T]/g, '')].join('_') + '.wav';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const file = actFolder_(q.act).createFile(Utilities.newBlob(bytes, q.mime || 'audio/wav', name));
    sheet_('Takes', TAKE_COLS).appendRow([id, who, Number(q.act), String(q.song), q.kind, !!q.private, created,
      Number(q.dur) || 0, Number(q.offset) || 0, file.getId(), q.mime || 'audio/wav', false]);
  } finally { lock.releaseLock(); }
  return { ok: true, id, handle: who };
}

function get_(q, admin) {
  const t = rows_('Takes').find(r => r.id === q.id && !r.removed);
  if (!t) throw new Error('not_found');
  if (t.private && !admin) {
    const who = verify_(q.handle, q.pin, false);
    if (!who || !sameHandle_(who, t.handle)) throw new Error('private');
  }
  const blob = DriveApp.getFileById(t.fileId).getBlob();
  return { ok: true, mime: t.mime, data: Utilities.base64Encode(blob.getBytes()) };
}

function setPrivate_(q, admin) {
  return editTake_(q, admin, (sh, rowIdx) => sh.getRange(rowIdx, TAKE_COLS.indexOf('private') + 1).setValue(!!q.private));
}

function remove_(q, admin) {
  return editTake_(q, admin, (sh, rowIdx, t) => {
    sh.getRange(rowIdx, TAKE_COLS.indexOf('removed') + 1).setValue(true);
    try { DriveApp.getFileById(t.fileId).moveTo(subfolder_(folder_(), '_to_delete')); } catch (e) {}
  });
}

function editTake_(q, admin, fn) {
  const sh = sheet_('Takes', TAKE_COLS), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const t = toObj_(data[r]);
    if (t.id !== q.id) continue;
    if (!admin) { const who = verify_(q.handle, q.pin, false); if (!who || !sameHandle_(who, t.handle)) throw new Error('not_yours'); }
    fn(sh, r + 1, t);
    return { ok: true };
  }
  throw new Error('not_found');
}

// ---------- players ----------
// Returns the canonical handle. create=true registers a new handle with this PIN.
function verify_(handle, pin, create) {
  handle = String(handle || '').trim();
  if (!handle) { if (create) throw new Error('need_handle'); return null; }
  if (handle.length < 2 || handle.length > 24 || /[<>]/.test(handle)) throw new Error('bad_handle');
  if (!/^\d{4,8}$/.test(String(pin || ''))) { if (create) throw new Error('bad_pin_format'); return null; }
  const hash = hash_(pin);
  const sh = sheet_('Players', ['handle', 'pinHash', 'created']), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (sameHandle_(data[r][0], handle)) {
      if (data[r][1] === hash) return data[r][0];
      if (create) throw new Error('bad_pin');
      return null;
    }
  }
  if (!create) return null;
  sh.appendRow([handle, hash, new Date().toISOString()]);
  return handle;
}

// ---------- helpers ----------
function sameHandle_(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }
function hash_(pin) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, setting_('SALT') + ':' + pin);
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function publicTake_(t) {
  return { id: t.id, handle: t.handle, act: Number(t.act), song: String(t.song).padStart(2, '0'), kind: t.kind,
    private: !!t.private, created: t.created, dur: Number(t.dur), offset: Number(t.offset) };
}
function toObj_(row) { const o = {}; TAKE_COLS.forEach((c, i) => o[c] = row[i]); o.song = String(o.song).padStart(2, '0'); return o; }
function rows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (name === 'Takes') return data.slice(1).filter(r => r[0]).map(toObj_);
  const head = data[0];
  return data.slice(1).filter(r => r[0]).map(r => { const o = {}; head.forEach((h, i) => o[h] = r[i]); return o; });
}
function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function setting_(k) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sh) return '';
  const row = sh.getDataRange().getValues().find(r => r[0] === k);
  return row ? String(row[1]) : '';
}
function folder_() { const it = DriveApp.getFoldersByName(FOLDER_NAME); return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME); }
function subfolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function actFolder_(a) { return subfolder_(folder_(), 'Act ' + Number(a)); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
