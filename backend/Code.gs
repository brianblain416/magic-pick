/**
 * Magic Pick — takes library (Google Apps Script web app)
 *
 * Stores recorded vocal, guitar-solo and Syn2 (synth part) takes in a Drive folder called
 * "Magic Pick Takes" and keeps the list in this spreadsheet.
 *   Players sheet:  name, email, created           (one row per singer)
 *   Devices sheet:  device, email, token, expires, verified, created
 *   Takes sheet:    one row per take
 *   Playlists sheet: one row per playlist. Each person makes their own: per song, the vocal and guitar
 *                   takes it plays (more than one can be ticked), each with a name and Monitor on/off.
 *                   A hidden playlist is seen only by its owner. Removed playlists are marked removed.
 *   Settings sheet: SALT, ADMIN_EMAILS (comma-separated; can hear every private take)
 *   Subscribers sheet: Brian's Substack subscriber export (File > Import > Replace current sheet).
 *                   Only emails marked paid or comp there can sign in and hear the player.
 *   Allowed sheet:  extra emails let in by hand (column A, one per row), e.g. before the next export.
 *   Auditions sheet: one row per audition take (anyone can audition after confirming their email):
 *                   created, name, email, role, act, song, title, kind, takeId, link, note.
 *                   Audition takes are always private: only the auditioner and the admins hear them.
 *   Comments sheet: one row per comment on a lyric or dialogue line. A "replace" comment
 *                   is a suggested new wording (each person's suggestions form their own
 *                   version); a "note" just stays attached. Brian ticks suggestions
 *                   and presses Update, which marks them "applied" — the site then shows
 *                   the new wording. Nothing is deleted: removed comments are marked removed.
 *
 * Applied edits can be read by anyone at  <web app URL>?edits=1  (used to update Ableton).
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
const TAKE_COLS = ['id', 'handle', 'act', 'song', 'kind', 'private', 'created', 'dur', 'offset', 'fileId', 'mime', 'removed', 'owner', 'mode'];
const DEV_COLS = ['device', 'email', 'token', 'expires', 'verified', 'created'];
const MAX_BYTES = 30 * 1024 * 1024;
const LINK_MINUTES = 60;
const COM_COLS = ['id', 'act', 'song', 'part', 'idx', 'kind', 'text', 'dir', 'original', 'origDir', 'handle', 'owner', 'created', 'status', 'appliedAt', 'appliedBy'];
const PARTS = ['lyrics', 'before', 'during', 'after'];
const PL_COLS = ['id', 'owner', 'handle', 'name', 'hidden', 'data', 'created', 'updated', 'removed'];
const AUD_COLS = ['created', 'name', 'email', 'role', 'act', 'song', 'title', 'kind', 'takeId', 'link', 'note'];
const ROLES = ['Robbie', 'Melissa', 'Clive', 'Guitar'];

function setup() {
  sheet_('Players', ['name', 'email', 'created']);
  sheet_('Devices', DEV_COLS);
  sheet_('Takes', TAKE_COLS);
  sheet_('Auditions', AUD_COLS);
  comSheet_();
  plSheet_();
  const st = sheet_('Settings', ['key', 'value']);
  if (!setting_('SALT')) st.appendRow(['SALT', Utilities.getUuid()]);
  if (!setting_('ADMIN_EMAILS')) st.appendRow(['ADMIN_EMAILS', Session.getEffectiveUser().getEmail()]);
  folder_();
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Admin: ' + setting_('ADMIN_EMAILS'));
  SpreadsheetApp.flush();
  return 'ok';
}

function doGet(e) {
  if (e && e.parameter && e.parameter.edits) {
    return out_({ ok: true, edits: comments_().filter(c => c.status === 'applied' && c.kind === 'replace').map(publicComment_) });
  }
  return out_({ ok: true, app: 'magic-pick-takes' });
}

function doPost(e) {
  try {
    const q = JSON.parse(e.postData.contents);
    const me = device_(q.key);            // {email, name} if this device is signed in
    const admin = !!me && isAdmin_(me.email);
    // The player is for Substack paid (and comp) subscribers only. A signed-in device whose email
    // has dropped off the list sees the sign-in screen again.
    const allowed = !!me && (admin || allowed_(me.email));
    // Anyone who has confirmed their email can audition and comment, without being a subscriber ("guest").
    const GUEST_OK = ['audition', 'comment', 'uncomment', 'remove'];
    if (['list', 'signin', 'verify', 'signout'].indexOf(q.action) < 0 && !allowed && !(me && GUEST_OK.indexOf(q.action) >= 0))
      throw new Error(me ? 'not_subscriber' : 'signed_out');
    switch (q.action) {
      case 'list': return out_(allowed ? list_(me, admin) : guestList_(me));
      case 'audition': return out_(audition_(q, me));
      case 'signin': return out_(signin_(q));
      case 'verify': return out_(verify_(q));
      case 'signout': return out_(signout_(q));
      case 'upload': return out_(upload_(q, me));
      case 'get': return out_(get_(q, me, admin));
      case 'setPrivate': return out_(setPrivate_(q, me, admin));
      case 'remove': return out_(remove_(q, me, admin));
      case 'comment': return out_(comment_(q, me));
      case 'uncomment': return out_(uncomment_(q, me, admin));
      case 'applyEdits': return out_(applyEdits_(q, me, admin));
      case 'savePlaylist': return out_(savePlaylist_(q, me));
      case 'removePlaylist': return out_(removePlaylist_(q, me, admin));
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
  const comments = comments_().filter(c => c.status !== 'removed')
    .map(c => Object.assign(publicComment_(c), { mine: !!me && c.owner === me.email }));
  const playlists = rows_('Playlists', PL_COLS).filter(p => !truthy_(p.removed) && (!truthy_(p.hidden) || (me && p.owner === me.email)))
    .map(p => publicPlaylist_(p, me));
  return { ok: true, takes, players, comments, playlists, me: me ? { name: me.name, email: me.email } : null, admin };
}

// A signed-in non-subscriber (or nobody): the full player stays locked; they get the comments and their own audition takes.
function guestList_(me) {
  const takes = !me ? [] : rows_('Takes', TAKE_COLS).filter(t => !t.removed && t.owner === me.email)
    .map(t => Object.assign(publicTake_(t), { mine: true }));
  const auditions = !me ? [] : rows_('Auditions', AUD_COLS).filter(a => a.email === me.email)
    .map(a => ({ role: a.role, act: Number(a.act), song: String(a.song).padStart(2, '0'), kind: a.kind, takeId: a.takeId, created: a.created }));
  const comments = !me ? [] : comments_().filter(c => c.status !== 'removed')
    .map(c => Object.assign(publicComment_(c), { mine: c.owner === me.email }));
  return { ok: true, locked: true, guest: !!me, takes, auditions, players: [], comments, me: me ? { name: me.name, email: me.email } : null, admin: false };
}

// ---------- auditions ----------
// audition: an upload (always private) plus a row in the Auditions sheet for Brian, grouped by role.
// {act, song, kind: vocal|solo, mode, role, title, note, dur, offset, mime, data}
function audition_(q, me) {
  if (ROLES.indexOf(q.role) < 0) throw new Error('bad_role');
  if (['vocal', 'solo'].indexOf(q.kind) < 0) throw new Error('bad_kind');
  const up = upload_(Object.assign({}, q, { private: true }), me);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const row = [new Date().toISOString(), me.name, me.email, q.role, String(Number(q.act)), String(q.song),
      clean_(q.title, 80), q.kind === 'solo' ? 'guitar' : 'vocal', up.id, 'https://drive.google.com/file/d/' + up.fileId + '/view', clean_(q.note, 500)];
    // written as plain text so a name or note starting with = + - @ is never read as a formula
    const sh = sheet_('Auditions', AUD_COLS);
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setRichTextValues([row.map((v, i) => i === 9
      ? SpreadsheetApp.newRichTextValue().setText(String(v)).setLinkUrl(String(v)).build()
      : SpreadsheetApp.newRichTextValue().setText(String(v)).build())]);
  } finally { lock.releaseLock(); }
  return { ok: true, id: up.id, handle: me.name };
}

// ---------- playlists ----------
// savePlaylist: {id?, name, hidden, data: {songs: {"1-03": [{id, mon, name}], ...}}}. No id = a new playlist.
function savePlaylist_(q, me) {
  if (!me) throw new Error('signed_out');
  const name = clean_(q.name, 40); if (!name) throw new Error('bad_name');
  const data = JSON.stringify(cleanPlaylist_(q.data));
  if (data.length > 40000) throw new Error('too_big');
  const now = new Date().toISOString(), hidden = q.hidden ? 'TRUE' : 'FALSE';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = plSheet_(), rows = sh.getDataRange().getValues();
    const put = (r, row) => sh.getRange(r, 1, 1, row.length).setRichTextValues([row.map(v => SpreadsheetApp.newRichTextValue().setText(String(v)).build())]);
    if (q.id) {
      for (let r = 1; r < rows.length; r++) {
        const p = toObj_(rows[r], PL_COLS);
        if (p.id !== q.id || truthy_(p.removed)) continue;
        if (p.owner !== me.email) throw new Error('not_yours');
        put(r + 1, [p.id, p.owner, me.name, name, hidden, data, p.created, now, 'FALSE']);
        return { ok: true, id: p.id, updated: now };
      }
      throw new Error('not_found');
    }
    const id = Utilities.getUuid();
    put(sh.getLastRow() + 1, [id, me.email, me.name, name, hidden, data, now, now, 'FALSE']);
    return { ok: true, id, updated: now };
  } finally { lock.releaseLock(); }
}
function removePlaylist_(q, me, admin) {
  if (!me) throw new Error('signed_out');
  const sh = plSheet_(), rows = sh.getDataRange().getValues();
  for (let r = 1; r < rows.length; r++) {
    const p = toObj_(rows[r], PL_COLS);
    if (p.id !== q.id) continue;
    if (!admin && p.owner !== me.email) throw new Error('not_yours');
    sh.getRange(r + 1, PL_COLS.indexOf('removed') + 1).setValue('TRUE');
    return { ok: true };
  }
  throw new Error('not_found');
}
// Keeps only what the player uses: song keys like "1-03", up to 12 takes each with an id, Monitor and a name.
function cleanPlaylist_(d) {
  const songs = {}, src = d && typeof d === 'object' && d.songs && typeof d.songs === 'object' ? d.songs : {};
  Object.keys(src).slice(0, 80).forEach(k => {
    if (!/^[12]-\d{2}$/.test(k) || !Array.isArray(src[k])) return;
    const takes = src[k].filter(t => t && /^[\w-]{1,64}$/.test(String(t.id))).slice(0, 12)
      .map(t => ({ id: String(t.id), mon: t.mon !== false, name: clean_(t.name, 40) }));
    if (takes.length) songs[k] = takes;
  });
  return { songs };
}
function publicPlaylist_(p, me) {
  let data = { songs: {} }; try { data = JSON.parse(p.data); } catch (e) {}
  const iso = v => v instanceof Date ? v.toISOString() : String(v || '');
  return { id: p.id, handle: p.handle, name: String(p.name), hidden: truthy_(p.hidden), data, updated: iso(p.updated), mine: !!me && p.owner === me.email };
}
function plSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Playlists');
  if (!sh) {
    sh = ss.insertSheet('Playlists');
    sh.getRange(1, 1, 1, PL_COLS.length).setValues([PL_COLS]); sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), PL_COLS.length).setNumberFormat('@');
  }
  return sh;
}
function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }

// ---------- comments on lyric and dialogue lines ----------
// comment: {act, song, part, idx, kind: 'replace'|'note', text, dir, original, origDir}
function comment_(q, me) {
  if (!me) throw new Error('signed_out');
  const act = Number(q.act), song = String(q.song), part = String(q.part), idx = Number(q.idx);
  if ([1, 2].indexOf(act) < 0 || !/^\d{2}$/.test(song) || PARTS.indexOf(part) < 0 ||
      !(idx >= 0 && idx < 1000 && idx === Math.floor(idx))) throw new Error('bad_line');
  const kind = q.kind === 'note' ? 'note' : 'replace';
  const text = clean_(q.text, 500), dir = part === 'lyrics' ? '' : clean_(q.dir, 200);
  if (!text) throw new Error('empty');
  const row = [Utilities.getUuid(), act, song, part, idx, kind, text, dir, clean_(q.original, 500), clean_(q.origDir, 200),
    me.name, me.email, new Date().toISOString(), 'open', '', ''];
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = comSheet_();
    // written as plain text so lines starting with = + - @ are never read as formulas
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length)
      .setRichTextValues([row.map(v => SpreadsheetApp.newRichTextValue().setText(String(v)).build())]);
  } finally { lock.releaseLock(); }
  return { ok: true, id: row[0] };
}

function uncomment_(q, me, admin) {
  if (!me) throw new Error('signed_out');
  const sh = comSheet_(), data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const c = toObj_(data[r], COM_COLS);
    if (c.id !== q.id) continue;
    if (!admin && c.owner !== me.email) throw new Error('not_yours');
    if (c.status === 'applied') throw new Error('already_applied');
    sh.getRange(r + 1, COM_COLS.indexOf('status') + 1).setValue('removed');
    return { ok: true };
  }
  throw new Error('not_found');
}

// applyEdits: {ids: [...]} — Brian only. Makes the ticked suggestions the new wording.
function applyEdits_(q, me, admin) {
  if (!me || !admin) throw new Error('not_director');
  const ids = (Array.isArray(q.ids) ? q.ids : []).map(String).slice(0, 300);
  if (!ids.length) throw new Error('nothing_ticked');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = comSheet_(), data = sh.getDataRange().getValues(), rows = [], lines = {};
    for (let r = 1; r < data.length; r++) {
      const c = toObj_(data[r], COM_COLS);
      if (ids.indexOf(c.id) < 0) continue;
      if (c.status !== 'open' || c.kind !== 'replace') throw new Error('not_open');
      const k = [c.act, c.song, c.part, c.idx].join('|');
      if (lines[k]) throw new Error('two_on_one_line');
      lines[k] = true; rows.push(r + 1);
    }
    if (rows.length !== ids.length) throw new Error('not_found');
    const at = new Date().toISOString(), col = COM_COLS.indexOf('status') + 1;
    rows.forEach(r => sh.getRange(r, col, 1, 3).setRichTextValues([['applied', at, me.name]
      .map(v => SpreadsheetApp.newRichTextValue().setText(v).build())]));
    return { ok: true, applied: rows.length };
  } finally { lock.releaseLock(); }
}

function comments_() { return rows_('Comments', COM_COLS); }
function publicComment_(c) {
  const iso = v => v instanceof Date ? v.toISOString() : String(v || '');
  return { id: c.id, act: Number(c.act), song: String(c.song).padStart(2, '0'), part: String(c.part), idx: Number(c.idx),
    kind: c.kind, text: String(c.text), dir: String(c.dir || ''), original: String(c.original), origDir: String(c.origDir || ''),
    handle: c.handle, created: iso(c.created), status: c.status, appliedAt: iso(c.appliedAt) };
}
function comSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Comments');
  if (!sh) {
    sh = ss.insertSheet('Comments');
    sh.getRange(1, 1, 1, COM_COLS.length).setValues([COM_COLS]); sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), COM_COLS.length).setNumberFormat('@');
  }
  return sh;
}
function clean_(s, n) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, n); }

function upload_(q, me) {
  if (!me) throw new Error('signed_out');
  if (['vocal', 'solo', 'synth'].indexOf(q.kind) < 0) throw new Error('bad_kind');
  if (!/^\d{2}$/.test(String(q.song)) || [1, 2].indexOf(Number(q.act)) < 0) throw new Error('bad_song');
  const bytes = Utilities.base64Decode(q.data);
  if (bytes.length > MAX_BYTES) throw new Error('too_big');
  const id = Utilities.getUuid();
  const created = new Date().toISOString();
  const name = ['A' + q.act, q.song, q.kind, me.name, created.slice(0, 16).replace(/[:T]/g, '')].join('_') + '.wav';
  const lock = LockService.getScriptLock(); lock.waitLock(20000); let file;
  try {
    file = actFolder_(q.act).createFile(Utilities.newBlob(bytes, q.mime || 'audio/wav', name));
    sheet_('Takes', TAKE_COLS).appendRow([id, me.name, Number(q.act), String(q.song), q.kind, !!q.private, created,
      Number(q.dur) || 0, Number(q.offset) || 0, file.getId(), q.mime || 'audio/wav', false, me.email,
      ['full', 'loop'].indexOf(q.mode) >= 0 ? q.mode : '']);
  } finally { lock.releaseLock(); }
  return { ok: true, id, handle: me.name, fileId: file.getId() };
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
  const aud = !!q.audition;   // auditioners only need to confirm their email
  if (!aud && !isAdmin_(email) && !allowed_(email)) throw new Error('not_subscriber');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let name;
  try { name = player_(email, q.name); } finally { lock.releaseLock(); }
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('mail_quota');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expires = new Date(Date.now() + LINK_MINUTES * 60000).toISOString();
  sheet_('Devices', DEV_COLS).appendRow([dev, email, keyHash_(token), expires, false, new Date().toISOString()]);
  const link = SITE_URL + '?signin=' + token + (aud ? '&audition=1' : '');
  MailApp.sendEmail({
    to: email,
    subject: aud ? 'The Magic Pick — audition sign-in link' : 'Magic Pick Player — sign-in link',
    body: 'Hi ' + name + ',\n\nOpen this link on the computer or phone you want to record on:\n\n' + link +
      '\n\nIt works once and expires in ' + LINK_MINUTES + ' minutes. If you didn\'t ask for this, ignore this email.',
    htmlBody: '<p>Hi ' + esc_(name) + ',</p><p>Open this link on the computer or phone you want to record on:</p>' +
      '<p><a href="' + link + '">Sign in to the Magic Pick Player</a></p>' +
      '<p style="color:#666">It works once and expires in ' + LINK_MINUTES + ' minutes. If you didn\'t ask for this, ignore this email.</p>',
    name: 'Magic Pick Player'
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

// ---------- who may use the player ----------
// Subscribers: Substack's CSV export as imported. The email column is the one whose header says "email";
// a row is let in when a plan/type/status-like column says paid, comp, gift, founding, monthly or annual
// (or a "paid"/"comp" column is TRUE/yes). Free subscribers are not let in.
const PAID_RE = /\b(paid|comp|comped|complimentary|gift|gifted|founding|monthly|annual|yearly)\b/i;
function allowedEmails_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('allowed');
  if (hit != null) return hit ? hit.split(',') : [];
  const set = {}, ss = SpreadsheetApp.getActiveSpreadsheet();
  const sub = ss.getSheetByName('Subscribers');
  if (sub && sub.getLastRow() > 1) {
    const data = sub.getDataRange().getDisplayValues(), head = data[0].map(h => String(h).trim().toLowerCase());
    const ec = head.findIndex(h => /e-?mail/.test(h) && !/disabled|opens|sent/.test(h));
    // Substack's export has a "Type" column (Free, Comp, Yearly Subscriber, Monthly Subscriber, Founding Member, Author):
    // when it's there, only it counts, since "Subscription source (paid)" still says comp after a comp has lapsed to Free.
    const typeCol = head.indexOf('type');
    const planCols = typeCol >= 0 ? [typeCol] : head.map((h, i) => /plan|type|tier|status|subscription|paid|comp/.test(h) && i !== ec ? i : -1).filter(i => i >= 0);
    const flagCols = typeCol >= 0 ? [] : head.map((h, i) => /^(is_?)?(paid|comp|comped)$/.test(h.replace(/[\s?]/g, '')) ? i : -1).filter(i => i >= 0);
    if (ec >= 0) for (let r = 1; r < data.length; r++) {
      const em = String(data[r][ec]).trim().toLowerCase(); if (!em) continue;
      const paid = planCols.some(i => PAID_RE.test(data[r][i])) || flagCols.some(i => /^(true|yes|y|1)$/i.test(String(data[r][i]).trim()));
      if (paid) set[em] = true;
    }
  }
  const extra = ss.getSheetByName('Allowed');
  if (extra && extra.getLastRow() > 0) extra.getRange(1, 1, extra.getLastRow(), 1).getDisplayValues()
    .forEach(r => { const em = String(r[0]).trim().toLowerCase(); if (/@/.test(em)) set[em] = true; });
  const list = Object.keys(set);
  try { cache.put('allowed', list.join(','), 300); } catch (e) {}   // a re-import shows up within 5 minutes
  return list;
}
function allowed_(email) { return allowedEmails_().indexOf(String(email || '').trim().toLowerCase()) >= 0; }
// Run from the editor after importing a new export: logs how many emails are let in and which columns were used.
function checkSubscribers() {
  CacheService.getScriptCache().remove('allowed');
  const sub = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subscribers');
  Logger.log('Subscribers columns: ' + (sub ? sub.getRange(1, 1, 1, sub.getLastColumn()).getDisplayValues()[0].join(' | ') : '(no Subscribers sheet)'));
  const list = allowedEmails_();
  Logger.log(list.length + ' emails can sign in (plus admins): ' + list.slice(0, 20).join(', ') + (list.length > 20 ? ' …' : ''));
  return list.length;
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
    private: !!t.private, created: t.created, dur: Number(t.dur), offset: Number(t.offset), mode: String(t.mode || '') };
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
