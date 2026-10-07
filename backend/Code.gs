/**
 * Magic Pick — takes library (Google Apps Script web app)
 *
 * Stores recorded vocal and guitar-solo takes in a Drive folder called
 * "Magic Pick Takes" and keeps the list in this spreadsheet.
 *   Players sheet:  name, email, created           (one row per singer)
 *   Devices sheet:  device, email, token, expires, verified, created
 *   Takes sheet:    one row per take
 *   Settings sheet: SALT, ADMIN_EMAILS (comma-separated; can hear every private take)
 *
 * Sign-in is by email: a singer types their email once per device and taps the
 * link we email them. After that the device is signed in. A private take can be
 * heard by the singer who made it (on any signed-in device) and by the admins.
 *
 * Nothing is ever deleted: a removed take is marked removed and its file is
 * moved into a "_to_delete" folder for you to empty yourself.
 *
 * Setup: run setup() once from the editor (allow the permissions it asks for),
 * then Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 */

const SITE_URL = 'https://brianblain416.github.io/magic-pick/';
const FOLDER_NAME = 'Magic Pick Takes';
const TAKE_COLS = ['id', 'handle', 'act', 'song', 'kind', 'private', 'created', 'dur', 'offset', 'fileId', 'mime', 'removed', 'owner'];
const DEV_COLS = ['device', 'email', 'token', 'expires', 'verified', 'created'];
const MAX_BYTES = 30 * 1024 * 1024;
const LINK_MINUTES = 60;

function setup() {
  sheet_('Players', ['name', 'email', 'created']);
  sheet_('Devices', DEV_COLS);
  sheet_('Takes', TAKE_COLS);
  const st = sheet_('Settings', ['key', 'value']);
  if (!setting_('SALT')) st.appendRow(['SALT', Utilities.getUuid()]);
  if (!setting_('ADMIN_EMAILS')) st.appendRow(['ADMIN_EMAILS', Session.getEffectiveUser().getEmail()]);
  folder_();
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Admin: ' + setting_('ADMIN_EMAILS'));
  SpreadsheetApp.flush();
  return 'ok';
}

function doGet() { return out_({ ok: true, app: 'magic-pick-takes' }); }

function doPost(e) {
  try {
    const q = JSON.parse(e.postData.contents);
    const me = device_(q.key);            // {email, name} if this device is signed in
    const admin = !!me && isAdmin_(me.email);
    switch (q.action) {
      case 'list': return out_(list_(me, admin));
      case 'signin': return out_(signin_(q));
      case 'verify': return out_(verify_(q));
      case 'signout': return out_(signout_(q));
      case 'upload': return out_(upload_(q, me));
      case 'get': return out_(get_(q, me, admin));
      case 'setPrivate': return out_(setPrivate_(q, me, admin));
      case 'remove': return out_(remove_(q, me, admin));
      default: return out_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// ---------- takes ----------
function list_(me, admin) {
  const takes = rows_('Takes', TAKE_COLS).filter(t => !t.removed && (!t.private || admin || (me && t.owner === me.email)))
    .map(t => Object.assign(publicTake_(t), { mine: !!me && t.owner === me.email }));
  const players = rows_('Players', ['name', 'email', 'created']).map(p => p.name).filter(String).sort();
  return { ok: true, takes, players, me: me ? { name: me.name, email: me.email } : null, admin };
}

function upload_(q, me) {
  if (!me) throw new Error('signed_out');
  if (['vocal', 'solo'].indexOf(q.kind) < 0) throw new Error('bad_kind');
  if (!/^\d{2}$/.test(String(q.song)) || [1, 2].indexOf(Number(q.act)) < 0) throw new Error('bad_song');
  const bytes = Utilities.base64Decode(q.data);
  if (bytes.length > MAX_BYTES) throw new Error('too_big');
  const id = Utilities.getUuid();
  const created = new Date().toISOString();
  const name = ['A' + q.act, q.song, q.kind, me.name, created.slice(0, 16).replace(/[:T]/g, '')].join('_') + '.wav';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const file = actFolder_(q.act).createFile(Utilities.newBlob(bytes, q.mime || 'audio/wav', name));
    sheet_('Takes', TAKE_COLS).appendRow([id, me.name, Number(q.act), String(q.song), q.kind, !!q.private, created,
      Number(q.dur) || 0, Number(q.offset) || 0, file.getId(), q.mime || 'audio/wav', false, me.email]);
  } finally { lock.releaseLock(); }
  return { ok: true, id, handle: me.name };
}

function get_(q, me, admin) {
  const t = rows_('Takes', TAKE_COLS).find(r => r.id === q.id && !r.removed);
  if (!t) throw new Error('not_found');
  if (t.private && !admin && !(me && t.owner === me.email)) throw new Error('private');
  const blob = DriveApp.getFileById(t.fileId).getBlob();
  return { ok: true, mime: t.mime, data: Utilities.base64Encode(blob.getBytes()) };
}

function setPrivate_(q, me, admin) {
  return editTake_(q, me, admin, (sh, rowIdx) => sh.getRange(rowIdx, TAKE_COLS.indexOf('private') + 1).setValue(!!q.private));
}

function remove_(q, me, admin) {
  return editTake_(q, me, admin, (sh, rowIdx, t) => {
    sh.getRange(rowIdx, TAKE_COLS.indexOf('removed') + 1).setValue(true);
    try { DriveApp.getFileById(t.fileId).moveTo(subfolder_(folder_(), '_to_delete')); } catch (e) {}
  });
}

function editTake_(q, me, admin, fn) {
  const sh = sheet_('Takes', TAKE_COLS), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const t = toObj_(data[r], TAKE_COLS);
    if (t.id !== q.id) continue;
    if (!admin && !(me && t.owner === me.email)) throw new Error('not_yours');
    fn(sh, r + 1, t);
    return { ok: true };
  }
  throw new Error('not_found');
}

// ---------- sign-in by email ----------
// signin: {email, name, key} — emails a one-time link. The device that opens the link becomes signed in.
function signin_(q) {
  const email = String(q.email || '').trim().toLowerCase();
  if (!/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(email) || email.length > 120) throw new Error('bad_email');
  const dev = keyHash_(q.key); if (!dev) throw new Error('no_key');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let name;
  try { name = player_(email, q.name); } finally { lock.releaseLock(); }
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('mail_quota');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expires = new Date(Date.now() + LINK_MINUTES * 60000).toISOString();
  sheet_('Devices', DEV_COLS).appendRow([dev, email, keyHash_(token), expires, false, new Date().toISOString()]);
  const link = SITE_URL + '?signin=' + token;
  MailApp.sendEmail({
    to: email,
    subject: 'Magic Pick rehearsal — sign-in link',
    body: 'Hi ' + name + ',\n\nOpen this link on the computer or phone you want to record on:\n\n' + link +
      '\n\nIt works once and expires in ' + LINK_MINUTES + ' minutes. If you didn\'t ask for this, ignore this email.',
    htmlBody: '<p>Hi ' + esc_(name) + ',</p><p>Open this link on the computer or phone you want to record on:</p>' +
      '<p><a href="' + link + '">Sign in to the Magic Pick rehearsal page</a></p>' +
      '<p style="color:#666">It works once and expires in ' + LINK_MINUTES + ' minutes. If you didn\'t ask for this, ignore this email.</p>',
    name: 'Magic Pick rehearsal'
  });
  return { ok: true, name };
}

// verify: {token, key} — signs in the device that opened the link
function verify_(q) {
  const th = keyHash_(q.token), dev = keyHash_(q.key);
  if (!th || !dev) throw new Error('bad_link');
  const sh = sheet_('Devices', DEV_COLS), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const d = toObj_(data[r], DEV_COLS);
    if (d.token !== th) continue;
    if (d.verified || new Date(d.expires) < new Date()) throw new Error('link_used_or_expired');
    sh.getRange(r + 1, 1, 1, DEV_COLS.length).setValues([[dev, d.email, '', d.expires, true, d.created]]);
    const p = rows_('Players', ['name', 'email', 'created']).find(x => x.email === d.email);
    return { ok: true, me: { email: d.email, name: p ? p.name : d.email } };
  }
  throw new Error('bad_link');
}

function signout_(q) {
  const dev = keyHash_(q.key); if (!dev) return { ok: true };
  const sh = sheet_('Devices', DEV_COLS), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) if (data[r][0] === dev && data[r][4] === true) sh.getRange(r + 1, 5).setValue(false);
  return { ok: true };
}

function device_(key) {
  const dev = keyHash_(key); if (!dev) return null;
  const d = rows_('Devices', DEV_COLS).find(x => x.device === dev && x.verified === true);
  if (!d) return null;
  const p = rows_('Players', ['name', 'email', 'created']).find(x => x.email === d.email);
  return { email: d.email, name: p ? p.name : d.email };
}

// Players: one name per email. A name already used by another email is refused.
function player_(email, name) {
  name = String(name || '').trim();
  const sh = sheet_('Players', ['name', 'email', 'created']), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) if (data[r][1] === email) return data[r][0];
  if (name.length < 2 || name.length > 24 || /[<>]/.test(name)) throw new Error('bad_name');
  for (let r = 1; r < data.length; r++) if (sameName_(data[r][0], name)) throw new Error('name_taken');
  sh.appendRow([name, email, new Date().toISOString()]);
  return name;
}

// ---------- helpers ----------
function isAdmin_(email) { return setting_('ADMIN_EMAILS').toLowerCase().split(/[\s,]+/).filter(String).indexOf(email) >= 0; }
function sameName_(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }
function keyHash_(key) {
  key = String(key || '');
  if (!/^[A-Za-z0-9]{16,128}$/.test(key)) return '';
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, setting_('SALT') + ':' + key);
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function publicTake_(t) {
  return { id: t.id, handle: t.handle, act: Number(t.act), song: String(t.song).padStart(2, '0'), kind: t.kind,
    private: !!t.private, created: t.created, dur: Number(t.dur), offset: Number(t.offset) };
}
function toObj_(row, cols) { const o = {}; cols.forEach((c, i) => o[c] = row[i]); if ('song' in o) o.song = String(o.song).padStart(2, '0'); return o; }
function rows_(name, cols) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '').map(r => toObj_(r, cols));
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
