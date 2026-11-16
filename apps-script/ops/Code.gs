/**
 * Herd Co. Operations — backend for herdcoindy.com/admin
 *
 * Lives in a PRIVATE Google Sheet in the Herd Co. Google account.
 * The public site never holds data; every call is checked against the password here.
 *
 * Setup (once):
 *   1. Create a Google Sheet "Herd Co. Operations" -> Extensions -> Apps Script -> paste this file.
 *   2. Run setup()  (creates tabs + a private Drive folder for contracts; approve permissions).
 *   3. Reload the Sheet -> menu "Herd Ops" -> "Set password…"
 *   4. Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone -> copy the /exec URL
 *   5. Paste that URL into assets/ops-config.js in the herdcoindy repo.
 *   Code changes later: Deploy -> Manage deployments -> edit -> Version: New version (keeps the URL).
 */

var TAB_PROJECTS = 'Projects';
var TAB_LOG = 'Log';
var TAB_FILES = 'Files';

var PROJECT_FIELDS = [
  'id', 'stage', 'name', 'org', 'contactName', 'email', 'phone', 'source', 'service',
  'value', 'deposit', 'paid', 'monthly',
  'nextStep', 'nextDate', 'firstContact', 'consultDate', 'contractDate', 'deliveryDue', 'deliveredDate', 'renewalDate',
  'liveUrl', 'adminUrl', 'repoUrl', 'sheetUrl', 'scriptUrl', 'domain', 'hosting', 'backend', 'auth', 'integrations', 'accounts',
  'architecture', 'scope', 'notes', 'tags', 'created', 'updated'
];
var LOG_FIELDS = ['id', 'projectId', 'ts', 'type', 'text'];
var FILE_FIELDS = ['id', 'projectId', 'ts', 'kind', 'name', 'date', 'url', 'driveId', 'mime', 'size'];
var NUMBER_FIELDS = { value: 1, deposit: 1, paid: 1, monthly: 1 };
var STAGES = ['lead', 'discussion', 'future', 'contract', 'live', 'closed'];

var MAX_FAILS = 10, LOCK_MINUTES = 15;
var SESSION_HOURS = 12, REMEMBER_DAYS = 30;
var MAX_UPLOAD_BYTES = 25 * 1024 * 1024;

/* ---------------- menu + setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Herd Ops')
    .addItem('Set password…', 'setPasswordPrompt')
    .addItem('Sign out all devices', 'signOutEverywhere')
    .addItem('Run setup (safe to re-run)', 'setup')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureTab_(ss, TAB_PROJECTS, PROJECT_FIELDS);
  ensureTab_(ss, TAB_LOG, LOG_FIELDS);
  ensureTab_(ss, TAB_FILES, FILE_FIELDS);
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  contractsFolder_();
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  return 'Setup done. Next: Herd Ops -> Set password…';
}

function ensureTab_(ss, name, fields) {
  var s = ss.getSheetByName(name) || ss.insertSheet(name);
  var lastCol = Math.max(s.getLastColumn(), 1);
  var head = s.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (!head[0]) head = [];
  fields.forEach(function (f) { if (head.indexOf(f) < 0) head.push(f); });
  s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  s.setFrozenRows(1);
  // Store as plain text so dates/URLs aren't reformatted; money columns stay numeric.
  if (s.getMaxRows() < 500) s.insertRowsAfter(s.getMaxRows(), 500 - s.getMaxRows());
  head.forEach(function (h, i) {
    s.getRange(2, i + 1, s.getMaxRows() - 1, 1).setNumberFormat(NUMBER_FIELDS[h] ? '0.00' : '@');
  });
  return s;
}

function contractsFolder_() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Herd Co. Operations — Contracts & Files (private)');
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function setPasswordPrompt() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Set admin password', 'At least 10 characters. A short phrase works well.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pw = r.getResponseText();
  if (!pw || pw.length < 10) { ui.alert('Password must be at least 10 characters. Nothing changed.'); return; }
  setPassword_(pw);
  ui.alert('Password set. All other devices were signed out.');
}

function signOutEverywhere() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('PW_VERSION', String(Number(p.getProperty('PW_VERSION') || 0) + 1));
}

function setPassword_(pw) {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  var salt = Utilities.getUuid();
  p.setProperty('PW_SALT', salt);
  p.setProperty('PW_HASH', hashPw_(pw, salt));
  p.setProperty('PW_VERSION', String(Number(p.getProperty('PW_VERSION') || 0) + 1));
  p.deleteProperty('FAILS'); p.deleteProperty('LOCK_UNTIL');
}

/* ---------------- crypto ---------------- */

function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function hashPw_(pw, salt) {
  var h = salt + '|' + pw;
  for (var i = 0; i < 2000; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + salt, Utilities.Charset.UTF_8));
  return h;
}
function sign_(payload) {
  var p = PropertiesService.getScriptProperties();
  return hex_(Utilities.computeHmacSha256Signature(payload, p.getProperty('SECRET')));
}
function safeEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0; for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function makeToken_(remember) {
  var p = PropertiesService.getScriptProperties();
  var exp = Date.now() + (remember ? REMEMBER_DAYS * 864e5 : SESSION_HOURS * 36e5);
  var payload = 'v1.' + exp + '.' + (p.getProperty('PW_VERSION') || '0');
  return { token: payload + '.' + sign_(payload), exp: exp };
}
function checkToken_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 4) return false;
  var payload = parts.slice(0, 3).join('.');
  var p = PropertiesService.getScriptProperties();
  if (!safeEq_(sign_(payload), parts[3])) return false;
  if (Number(parts[1]) < Date.now()) return false;
  return parts[2] === (p.getProperty('PW_VERSION') || '0');
}

/* ---------------- web app ---------------- */

function doGet() {
  return json_({ ok: true, app: 'herd-ops' });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  try {
    return json_(route_(req));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function route_(req) {
  var a = req.action;
  if (a === 'login') return login_(req);
  if (a === 'lead') return publicLead_(req);
  if (!checkToken_(req.token)) return { ok: false, auth: false, error: 'Signed out — please log in again.' };

  if (a === 'bootstrap') return { ok: true, projects: readTab_(TAB_PROJECTS), log: readTab_(TAB_LOG), files: readTab_(TAB_FILES), folderUrl: contractsFolder_().getUrl(), sheetUrl: SpreadsheetApp.getActive().getUrl() };
  return withLock_(function () {
    if (a === 'saveProject') return { ok: true, project: saveProject_(req.project) };
    if (a === 'deleteProject') return deleteProject_(req.id);
    if (a === 'addLog') return { ok: true, entry: addLog_(req.projectId, req.type || 'note', req.text) };
    if (a === 'deleteLog') { deleteRow_(TAB_LOG, req.id); return { ok: true }; }
    if (a === 'uploadFile') return { ok: true, file: uploadFile_(req) };
    if (a === 'addFileLink') return { ok: true, file: addFileLink_(req) };
    if (a === 'deleteFile') return deleteFile_(req.id);
    if (a === 'changePassword') return changePassword_(req);
    throw new Error('Unknown action');
  });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn) {
  var l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ---------------- auth ---------------- */

function login_(req) {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('PW_HASH')) return { ok: false, error: 'No password set yet. In the Sheet: Herd Ops -> Set password…' };
  var until = Number(p.getProperty('LOCK_UNTIL') || 0);
  if (until > Date.now()) return { ok: false, error: 'Too many wrong tries. Locked for ' + Math.ceil((until - Date.now()) / 60000) + ' more min.' };
  var ok = safeEq_(hashPw_(String(req.password || ''), p.getProperty('PW_SALT')), p.getProperty('PW_HASH'));
  if (!ok) {
    var fails = Number(p.getProperty('FAILS') || 0) + 1;
    if (fails >= MAX_FAILS) { p.setProperty('LOCK_UNTIL', String(Date.now() + LOCK_MINUTES * 60000)); fails = 0; }
    p.setProperty('FAILS', String(fails));
    Utilities.sleep(600);
    return { ok: false, error: 'Wrong password.' };
  }
  p.deleteProperty('FAILS');
  var t = makeToken_(!!req.remember);
  return { ok: true, token: t.token, exp: t.exp };
}

function changePassword_(req) {
  var p = PropertiesService.getScriptProperties();
  if (!safeEq_(hashPw_(String(req.oldPassword || ''), p.getProperty('PW_SALT')), p.getProperty('PW_HASH'))) return { ok: false, error: 'Current password is wrong.' };
  if (String(req.newPassword || '').length < 10) return { ok: false, error: 'New password must be at least 10 characters.' };
  setPassword_(String(req.newPassword));
  var t = makeToken_(!!req.remember);
  return { ok: true, token: t.token, exp: t.exp };
}

/* ---------------- public lead intake (consult form) ---------------- */

function publicLead_(req) {
  if (req._honey) return { ok: true };
  var c = CacheService.getScriptCache();
  var n = Number(c.get('leadCount') || 0);
  if (n >= 25) return { ok: false, error: 'Busy, try later' };
  c.put('leadCount', String(n + 1), 3600);
  var cut = function (v, max) { return String(v || '').slice(0, max); };
  var name = cut(req.name, 120), email = cut(req.email, 160);
  if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Name and email required' };
  return withLock_(function () {
    var proj = saveProject_({
      stage: 'lead', name: cut(req.group, 160) || name, contactName: name, email: email,
      service: cut(req.need, 60), source: 'Website form', firstContact: today_(),
      nextStep: 'Reply and book 30-min consult', nextDate: today_()
    });
    addLog_(proj.id, 'lead', 'Website consult request:\n' + cut(req.message, 4000));
    return { ok: true };
  });
}

/* ---------------- data ---------------- */

function sheet_(name) {
  var s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('Missing tab "' + name + '" — run setup()');
  return s;
}
function headers_(s) {
  return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String);
}
function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (typeof v === 'string' && /^'[=+\-@]/.test(v)) return v.slice(1);
  return v === null || v === undefined ? '' : v;
}
function readTab_(name) {
  var s = sheet_(name), last = s.getLastRow();
  if (last < 2) return [];
  var head = headers_(s);
  return s.getRange(2, 1, last - 1, head.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { var o = {}; head.forEach(function (h, i) { if (h) o[h] = cell_(r[i]); }); return o; });
}
function findRow_(s, id) {
  var last = s.getLastRow();
  if (last < 2) return -1;
  var ids = s.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function writeObj_(name, obj, allowed) {
  var s = sheet_(name), head = headers_(s);
  var row = obj.id ? findRow_(s, obj.id) : -1;
  var current = row > 0 ? s.getRange(row, 1, 1, head.length).getValues()[0] : head.map(function () { return ''; });
  var out = head.map(function (h, i) {
    if (allowed.indexOf(h) < 0 || !(h in obj)) return current[i];
    var v = obj[h];
    if (NUMBER_FIELDS[h]) return v === '' || v === null || isNaN(Number(v)) ? '' : Number(v);
    v = String(v === null || v === undefined ? '' : v);
    return /^[=+\-@]/.test(v) ? "'" + v : v; // never let a value become a formula
  });
  if (row > 0) s.getRange(row, 1, 1, head.length).setValues([out]);
  else s.appendRow(out);
  var res = {}; head.forEach(function (h, i) { if (h) res[h] = cell_(out[i]); });
  return res;
}
function deleteRow_(name, id) {
  var s = sheet_(name), r = findRow_(s, id);
  if (r > 0) s.deleteRow(r);
}
function newId_(prefix) { return prefix + Utilities.getUuid().slice(0, 8); }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function nowIso_() { return new Date().toISOString(); }

function saveProject_(p) {
  if (!p) throw new Error('Missing project');
  if (p.stage && STAGES.indexOf(p.stage) < 0) throw new Error('Bad stage');
  var s = sheet_(TAB_PROJECTS);
  var before = null;
  if (p.id) {
    var r = findRow_(s, p.id);
    if (r < 0) throw new Error('Project not found (deleted?)');
    var head = headers_(s), vals = s.getRange(r, 1, 1, head.length).getValues()[0];
    before = {}; head.forEach(function (h, i) { before[h] = cell_(vals[i]); });
  } else {
    p.id = newId_('p');
    p.created = nowIso_();
    p.stage = p.stage || 'lead';
  }
  p.updated = nowIso_();
  var saved = writeObj_(TAB_PROJECTS, p, PROJECT_FIELDS);
  if (before && p.stage && before.stage !== p.stage) addLog_(p.id, 'stage', before.stage + '→' + p.stage);
  if (!before) addLog_(p.id, 'stage', 'created→' + saved.stage);
  return saved;
}

function deleteProject_(id) {
  readTab_(TAB_FILES).forEach(function (f) {
    if (String(f.projectId) === String(id)) {
      if (f.driveId) { try { DriveApp.getFileById(f.driveId).setTrashed(true); } catch (e) {} }
      deleteRow_(TAB_FILES, f.id);
    }
  });
  readTab_(TAB_LOG).forEach(function (l) { if (String(l.projectId) === String(id)) deleteRow_(TAB_LOG, l.id); });
  deleteRow_(TAB_PROJECTS, id);
  return { ok: true };
}

function addLog_(projectId, type, text) {
  if (!projectId) throw new Error('Missing project');
  return writeObj_(TAB_LOG, { id: newId_('l'), projectId: projectId, ts: nowIso_(), type: type, text: String(text || '').slice(0, 20000) }, LOG_FIELDS);
}

function uploadFile_(req) {
  if (!req.projectId || !req.data) throw new Error('Missing file');
  var bytes = Utilities.base64Decode(req.data);
  if (bytes.length > MAX_UPLOAD_BYTES) throw new Error('File is over 25 MB');
  var name = String(req.name || 'file').slice(0, 200);
  var blob = Utilities.newBlob(bytes, req.mime || 'application/octet-stream', name);
  var file = contractsFolder_().createFile(blob); // private: only the owner can open it
  var rec = writeObj_(TAB_FILES, {
    id: newId_('f'), projectId: req.projectId, ts: nowIso_(), kind: req.kind || 'Other', name: name,
    date: req.date || '', url: file.getUrl(), driveId: file.getId(), mime: blob.getContentType(), size: String(bytes.length)
  }, FILE_FIELDS);
  addLog_(req.projectId, 'file', 'Added ' + (req.kind || 'file') + ': ' + name);
  return rec;
}

function addFileLink_(req) {
  if (!req.projectId || !/^https?:\/\//.test(String(req.url || ''))) throw new Error('Need a link that starts with https://');
  var rec = writeObj_(TAB_FILES, {
    id: newId_('f'), projectId: req.projectId, ts: nowIso_(), kind: req.kind || 'Other',
    name: String(req.name || req.url).slice(0, 200), date: req.date || '', url: req.url, driveId: '', mime: 'link', size: ''
  }, FILE_FIELDS);
  addLog_(req.projectId, 'file', 'Linked ' + (req.kind || 'file') + ': ' + rec.name);
  return rec;
}

function deleteFile_(id) {
  var f = readTab_(TAB_FILES).filter(function (x) { return String(x.id) === String(id); })[0];
  if (f && f.driveId) { try { DriveApp.getFileById(f.driveId).setTrashed(true); } catch (e) {} }
  deleteRow_(TAB_FILES, id);
  return { ok: true };
}
