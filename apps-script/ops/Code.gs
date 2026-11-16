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
  'architecture', 'scope', 'notes', 'tags', 'created', 'updated',
  'type', 'monitor', 'checklist', 'portfolioConsent', 'portfolioReady', 'closedReason', 'lastContact'
];
var LOG_FIELDS = ['id', 'projectId', 'ts', 'type', 'text', 'ref'];
var TAB_PAYMENTS = 'Payments';
var TAB_HEALTH = 'Health';
var PAYMENT_FIELDS = ['id', 'projectId', 'date', 'amount', 'method', 'invoice', 'note', 'ts'];
var HEALTH_FIELDS = ['projectId', 'name', 'url', 'kind', 'ok', 'code', 'detail', 'checkedAt'];
var SCHEMA_VERSION = '2';
var FILE_FIELDS = ['id', 'projectId', 'ts', 'kind', 'name', 'date', 'url', 'driveId', 'mime', 'size'];
var NUMBER_FIELDS = { value: 1, deposit: 1, paid: 1, monthly: 1, amount: 1 };
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
    .addSeparator()
    .addItem('Turn on automations (nightly checks, email sync, Monday digest)', 'installTriggers')
    .addItem('Run site checks now', 'healthCheck')
    .addItem('Sync client email now', 'syncEmail')
    .addItem('Send Monday digest now', 'weeklyDigest')
    .addItem('Create starter contract template', 'createContractTemplateMenu')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureTab_(ss, TAB_PROJECTS, PROJECT_FIELDS);
  ensureTab_(ss, TAB_LOG, LOG_FIELDS);
  ensureTab_(ss, TAB_FILES, FILE_FIELDS);
  ensureTab_(ss, TAB_PAYMENTS, PAYMENT_FIELDS);
  ensureTab_(ss, TAB_HEALTH, HEALTH_FIELDS);
  inferTypes_();
  var p = PropertiesService.getScriptProperties();
  p.setProperty('SCHEMA', SCHEMA_VERSION);
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

  ensureSchema_();
  if (a === 'bootstrap') return {
    ok: true, projects: readTab_(TAB_PROJECTS), log: readTab_(TAB_LOG), files: readTab_(TAB_FILES),
    payments: readTab_(TAB_PAYMENTS), health: readTab_(TAB_HEALTH), settings: getSettings_(),
    folderUrl: contractsFolder_().getUrl(), sheetUrl: SpreadsheetApp.getActive().getUrl()
  };
  if (a === 'runHealth') return { ok: true, health: healthCheck() };
  if (a === 'syncEmail') { var n = syncEmail(); return { ok: true, added: n, log: readTab_(TAB_LOG), projects: readTab_(TAB_PROJECTS) }; }
  if (a === 'sendDigest') { weeklyDigest(); return { ok: true }; }
  return withLock_(function () {
    if (a === 'saveProject') return { ok: true, project: saveProject_(req.project) };
    if (a === 'deleteProject') return deleteProject_(req.id);
    if (a === 'addLog') return { ok: true, entry: addLog_(req.projectId, req.type || 'note', req.text) };
    if (a === 'deleteLog') { deleteRow_(TAB_LOG, req.id); return { ok: true }; }
    if (a === 'uploadFile') return { ok: true, file: uploadFile_(req) };
    if (a === 'addFileLink') return { ok: true, file: addFileLink_(req) };
    if (a === 'deleteFile') return deleteFile_(req.id);
    if (a === 'changePassword') return changePassword_(req);
    if (a === 'addPayment') return { ok: true, payment: addPayment_(req.payment) };
    if (a === 'deletePayment') { deleteRow_(TAB_PAYMENTS, req.id); return { ok: true }; }
    if (a === 'generateContract') return { ok: true, file: generateContract_(req.projectId) };
    if (a === 'createContractTemplate') return { ok: true, settings: createContractTemplate_() };
    if (a === 'saveSettings') return { ok: true, settings: saveSettings_(req.settings || {}) };
    if (a === 'installTriggers') { installTriggers(); return { ok: true, settings: getSettings_() }; }
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

/* =====================================================================
   v2: schema migration, settings, payments, health checks, email sync,
   Monday digest, contract generation
   ===================================================================== */

var ADMIN_URL = 'https://herdcoindy.com/admin/';
var PROVIDER = 'Hoosier Prairie Holdings, LLC d/b/a Herd Co Indy';
var OWN_DOMAINS = ['herdcoindy.com', 'bluffroadbikeco.com'];

function ensureSchema_() {
  var p = PropertiesService.getScriptProperties();
  if (p.getProperty('SCHEMA') === SCHEMA_VERSION) return;
  withLock_(function () {
    if (p.getProperty('SCHEMA') === SCHEMA_VERSION) return;
    var ss = SpreadsheetApp.getActive();
    ensureTab_(ss, TAB_PROJECTS, PROJECT_FIELDS);
    ensureTab_(ss, TAB_LOG, LOG_FIELDS);
    ensureTab_(ss, TAB_PAYMENTS, PAYMENT_FIELDS);
    ensureTab_(ss, TAB_HEALTH, HEALTH_FIELDS);
    inferTypes_();
    p.setProperty('SCHEMA', SCHEMA_VERSION);
  });
}

/** Fill blank Project types once: "Personal" orgs -> personal, projects where Levi is the contact -> internal, else client. */
function inferTypes_() {
  readTab_(TAB_PROJECTS).forEach(function (pr) {
    if (pr.type) return;
    var t = /personal/i.test(pr.org) ? 'personal' : /^levi mcneely/i.test(String(pr.contactName)) ? 'internal' : 'client';
    writeObj_(TAB_PROJECTS, { id: pr.id, type: t }, PROJECT_FIELDS);
  });
}

/* ---------------- settings ---------------- */

function getSettings_() {
  var p = PropertiesService.getScriptProperties();
  var tpl = p.getProperty('CONTRACT_TEMPLATE_ID') || '';
  var triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  return {
    digestTo: p.getProperty('DIGEST_TO') || ownerEmail_(),
    staleDays: Number(p.getProperty('STALE_DAYS') || 14),
    contractTemplateId: tpl,
    contractTemplateUrl: tpl ? 'https://docs.google.com/document/d/' + tpl + '/edit' : '',
    automations: triggers.indexOf('nightly') >= 0 && triggers.indexOf('weeklyDigest') >= 0,
    lastHealth: p.getProperty('LAST_HEALTH') || '',
    lastEmailSync: p.getProperty('LAST_EMAIL_SYNC') || ''
  };
}
function saveSettings_(s) {
  var p = PropertiesService.getScriptProperties();
  if ('digestTo' in s) p.setProperty('DIGEST_TO', String(s.digestTo).trim());
  if ('staleDays' in s) p.setProperty('STALE_DAYS', String(Math.max(1, Number(s.staleDays) || 14)));
  if ('contractTemplateId' in s) {
    var id = String(s.contractTemplateId || '').trim();
    var m = id.match(/\/d\/([A-Za-z0-9_-]{20,})/); if (m) id = m[1];
    if (id) { try { DocumentApp.openById(id); } catch (e) { throw new Error('Could not open that Google Doc. Paste the Doc link from the Herd Co. account.'); } }
    p.setProperty('CONTRACT_TEMPLATE_ID', id);
  }
  return getSettings_();
}
function ownerEmail_() {
  try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

/* ---------------- triggers ---------------- */

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['nightly', 'weeklyDigest'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightly').timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger('weeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}
function nightly() {
  healthCheck();
  try { syncEmail(); } catch (e) { console.error('email sync failed: ' + e); }
}

/* ---------------- payments ---------------- */

function addPayment_(pay) {
  if (!pay || !pay.projectId) throw new Error('Missing project');
  var amt = Number(String(pay.amount).replace(/[$,]/g, ''));
  if (!(amt > 0)) throw new Error('Enter an amount');
  var rec = writeObj_(TAB_PAYMENTS, {
    id: newId_('$'), projectId: pay.projectId, date: pay.date || today_(), amount: amt,
    method: pay.method || '', invoice: pay.invoice || '', note: pay.note || '', ts: nowIso_()
  }, PAYMENT_FIELDS);
  addLog_(pay.projectId, 'payment', 'Payment received: $' + amt.toFixed(2) + (pay.method ? ' via ' + pay.method : '') + (pay.invoice ? ' (invoice ' + pay.invoice + ')' : ''));
  return rec;
}

/* ---------------- health checks ---------------- */

function healthTargets_() {
  var out = [];
  readTab_(TAB_PROJECTS).forEach(function (pr) {
    if (String(pr.monitor).toLowerCase() === 'no') return;
    if (pr.stage !== 'live' && pr.stage !== 'contract') return;
    var seen = {};
    [['site', pr.liveUrl], ['admin', pr.adminUrl], ['backend', pr.scriptUrl]].forEach(function (k) {
      var u = String(k[1] || '').trim();
      if (!/^https?:\/\//.test(u) || seen[u]) return;
      seen[u] = 1;
      out.push({ projectId: pr.id, name: pr.name, kind: k[0], url: u });
    });
  });
  return out;
}

/** Decide if a fetched response means "up". Apps Script backends get extra checks for Google's error pages. */
function judge_(t, code, body) {
  body = String(body || '').slice(0, 4000);
  if (code >= 400) return { ok: false, detail: 'HTTP ' + code };
  if (/script\.google\.com/.test(t.url)) {
    if (/unable to open the file/i.test(body)) return { ok: false, detail: 'Google says "unable to open the file" — deployment access is not "Anyone"' };
    // Google ran the script but it only handles POST — that still proves the deployment is public and alive.
    if (/Script function not found: doGet/i.test(body)) return { ok: true, detail: 'Reachable (POST-only backend)' };
    if (/Script function not found/i.test(body)) return { ok: false, detail: 'Deployed version is missing its entry point — redeploy as New version' };
    if (/accounts\.google\.com|ServiceLogin|Sign in - Google/i.test(body)) return { ok: false, detail: 'Asks visitors to sign in — set deployment access to "Anyone"' };
    if (/Authorization is required|needs your permission/i.test(body)) return { ok: false, detail: 'Script needs re-authorization — open Apps Script and run any function' };
    if (/^\s*<!DOCTYPE html|<html/i.test(body) && /error/i.test(body)) return { ok: false, detail: 'Script returned an error page' };
  }
  return { ok: true, detail: 'HTTP ' + code };
}

function healthCheck() {
  var targets = healthTargets_();
  var results = [];
  for (var i = 0; i < targets.length; i += 20) {
    var batch = targets.slice(i, i + 20);
    var resps = [];
    try {
      resps = UrlFetchApp.fetchAll(batch.map(function (t) { return { url: t.url, muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true }; }));
    } catch (e) {
      // One bad host can throw for the whole batch; retry one at a time.
      resps = batch.map(function (t) { try { return UrlFetchApp.fetch(t.url, { muteHttpExceptions: true, followRedirects: true }); } catch (err) { return { err: String(err.message || err) }; } });
    }
    batch.forEach(function (t, j) {
      var r = resps[j], v;
      if (!r || r.err) v = { ok: false, detail: (r && r.err) || 'No response', code: 0 };
      else { var code = r.getResponseCode(); v = judge_(t, code, r.getContentText()); v.code = code; }
      results.push({ projectId: t.projectId, name: t.name, url: t.url, kind: t.kind, ok: v.ok ? 'yes' : 'no', code: String(v.code), detail: v.detail, checkedAt: nowIso_() });
    });
  }
  var previous = {};
  readTab_(TAB_HEALTH).forEach(function (h) { previous[h.url] = h.ok; });
  withLock_(function () {
    var s = sheet_(TAB_HEALTH);
    if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    results.forEach(function (r) { s.appendRow(HEALTH_FIELDS.map(function (f) { return r[f]; })); });
  });
  PropertiesService.getScriptProperties().setProperty('LAST_HEALTH', nowIso_());

  var newlyDown = results.filter(function (r) { return r.ok === 'no' && previous[r.url] !== 'no'; });
  var recovered = results.filter(function (r) { return r.ok === 'yes' && previous[r.url] === 'no'; });
  if (newlyDown.length || recovered.length) {
    var to = getSettings_().digestTo;
    if (to) {
      var html = '<div style="font-family:Arial,sans-serif;font-size:14px">' +
        (newlyDown.length ? '<h3 style="color:#a33a2a">Down</h3><ul>' + newlyDown.map(function (r) { return '<li><b>' + h_(r.name) + '</b> (' + r.kind + ')<br><a href="' + h_(r.url) + '">' + h_(r.url) + '</a><br>' + h_(r.detail) + '</li>'; }).join('') + '</ul>' : '') +
        (recovered.length ? '<h3 style="color:#163823">Back up</h3><ul>' + recovered.map(function (r) { return '<li>' + h_(r.name) + ' (' + r.kind + ')</li>'; }).join('') + '</ul>' : '') +
        '<p><a href="' + ADMIN_URL + '">Open Herd Co. Operations</a></p></div>';
      MailApp.sendEmail({ to: to, subject: 'Herd Ops: ' + (newlyDown.length ? newlyDown.length + ' down' : '') + (newlyDown.length && recovered.length ? ', ' : '') + (recovered.length ? recovered.length + ' back up' : ''), htmlBody: html });
    }
  }
  return results;
}
function h_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

/* ---------------- client email sync ---------------- */

function isOwnAddress_(e) {
  e = String(e).toLowerCase();
  var me = ownerEmail_().toLowerCase();
  return e === me || OWN_DOMAINS.some(function (d) { return e.slice(-d.length - 1) === '@' + d; });
}

function syncEmail() {
  var p = PropertiesService.getScriptProperties();
  var last = p.getProperty('LAST_EMAIL_SYNC');
  var since = last ? new Date(new Date(last).getTime() - 2 * 864e5) : new Date(Date.now() - 60 * 864e5);
  var after = Utilities.formatDate(since, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  var seen = {};
  readTab_(TAB_LOG).forEach(function (l) { if (l.ref) seen[l.ref] = 1; });
  var added = 0, tz = Session.getScriptTimeZone();
  var work = [];

  readTab_(TAB_PROJECTS).forEach(function (pr) {
    if (pr.stage === 'closed' || pr.type === 'personal') return;
    var emails = String(pr.email || '').split(/[,;\s]+/).filter(function (e) { return /@/.test(e) && !isOwnAddress_(e); });
    if (!emails.length) return;
    var q = '(' + emails.map(function (e) { return 'from:' + e + ' OR to:' + e; }).join(' OR ') + ') after:' + after;
    var threads = GmailApp.search(q, 0, 25);
    var latest = pr.lastContact || '';
    threads.forEach(function (th) {
      th.getMessages().forEach(function (m) {
        var id = 'gmail:' + m.getId();
        var d = Utilities.formatDate(m.getDate(), tz, 'yyyy-MM-dd');
        var involved = (m.getFrom() + ' ' + m.getTo() + ' ' + m.getCc()).toLowerCase();
        if (!emails.some(function (e) { return involved.indexOf(e.toLowerCase()) >= 0; })) return;
        if (d > latest) latest = d;
        if (seen[id]) return;
        seen[id] = 1;
        var fromThem = emails.some(function (e) { return m.getFrom().toLowerCase().indexOf(e.toLowerCase()) >= 0; });
        var text = (fromThem ? '← ' : '→ ') + (m.getSubject() || '(no subject)') + '\n' + String(m.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 280);
        work.push({ log: { id: newId_('l'), projectId: pr.id, ts: m.getDate().toISOString(), type: 'email', text: text, ref: id } });
      });
    });
    if (latest && latest !== pr.lastContact) work.push({ project: { id: pr.id, lastContact: latest } });
  });

  withLock_(function () {
    work.forEach(function (w) {
      if (w.log) { writeObj_(TAB_LOG, w.log, LOG_FIELDS); added++; }
      if (w.project) writeObj_(TAB_PROJECTS, w.project, PROJECT_FIELDS);
    });
  });
  p.setProperty('LAST_EMAIL_SYNC', nowIso_());
  return added;
}

/* ---------------- Monday digest ---------------- */

function dayDiff_(s) {
  if (!s) return null;
  var a = new Date(String(s).slice(0, 10) + 'T12:00:00'), b = new Date(today_() + 'T12:00:00');
  return Math.round((a - b) / 864e5);
}
function money_(n) { return '$' + Number(n || 0).toLocaleString('en-US', { maximumFractionDigits: 2 }); }

function digestData_() {
  var P = readTab_(TAB_PROJECTS), pays = readTab_(TAB_PAYMENTS), health = readTab_(TAB_HEALTH);
  var stale = getSettings_().staleDays;
  var paid = {};
  pays.forEach(function (x) { paid[x.projectId] = (paid[x.projectId] || 0) + Number(x.amount || 0); });
  var paidFor = function (p) { return paid[p.id] != null ? paid[p.id] : Number(p.paid || 0); };
  var open = function (p) { return ['lead', 'discussion', 'contract', 'future'].indexOf(p.stage) >= 0; };
  var d = { followUps: [], deliveries: [], renewals: [], quiet: [], unpaid: [], down: [] };
  P.forEach(function (p) {
    var x;
    if (open(p) && p.nextDate && (x = dayDiff_(p.nextDate)) !== null && x <= 7) d.followUps.push({ p: p, d: x, t: p.nextStep || 'Next step' });
    if (p.stage === 'contract' && p.deliveryDue && !p.deliveredDate && (x = dayDiff_(p.deliveryDue)) !== null && x <= 14) d.deliveries.push({ p: p, d: x });
    if (p.renewalDate && p.stage === 'live' && (x = dayDiff_(p.renewalDate)) !== null && x <= 45) d.renewals.push({ p: p, d: x });
    if (open(p) && p.type !== 'personal' && p.type !== 'internal' && p.stage !== 'future') {
      var lc = p.lastContact || String(p.updated || '').slice(0, 10);
      if (lc && (x = dayDiff_(lc)) !== null && -x >= stale) d.quiet.push({ p: p, d: -x });
    }
    var bal = Number(p.value || 0) - paidFor(p);
    if ((p.stage === 'contract' || p.stage === 'live') && bal > 0) d.unpaid.push({ p: p, bal: bal });
  });
  health.forEach(function (h) { if (h.ok === 'no') d.down.push(h); });
  d.followUps.sort(function (a, b) { return a.d - b.d; });
  d.deliveries.sort(function (a, b) { return a.d - b.d; });
  return d;
}

function weeklyDigest() {
  var to = getSettings_().digestTo;
  if (!to) return;
  var d = digestData_();
  var rel = function (x) { return x < 0 ? Math.abs(x) + 'd late' : x === 0 ? 'today' : 'in ' + x + 'd'; };
  var sec = function (title, items, fn) { return items.length ? '<h3 style="margin:18px 0 6px;color:#163823">' + title + ' (' + items.length + ')</h3><ul style="margin:0;padding-left:18px">' + items.map(function (i) { return '<li style="margin:4px 0">' + fn(i) + '</li>'; }).join('') + '</ul>' : ''; };
  var body = sec('Sites down', d.down, function (h) { return '<b style="color:#a33a2a">' + h_(h.name) + '</b> (' + h_(h.kind) + ') — ' + h_(h.detail); }) +
    sec('Follow-ups this week', d.followUps, function (i) { return '<b>' + h_(i.p.name) + '</b> — ' + h_(i.t) + ' <i>(' + rel(i.d) + ')</i>'; }) +
    sec('Deliveries in the next 2 weeks', d.deliveries, function (i) { return '<b>' + h_(i.p.name) + '</b> — due ' + h_(i.p.deliveryDue) + ' <i>(' + rel(i.d) + ')</i>'; }) +
    sec('Renewals coming up', d.renewals, function (i) { return '<b>' + h_(i.p.name) + '</b> — ' + h_(i.p.renewalDate) + (i.p.domain ? ' · ' + h_(i.p.domain) : ''); }) +
    sec('Gone quiet', d.quiet, function (i) { return '<b>' + h_(i.p.name) + '</b> — no contact in ' + i.d + ' days'; }) +
    sec('Balances due', d.unpaid, function (i) { return '<b>' + h_(i.p.name) + '</b> — ' + money_(i.bal); });
  if (!body) body = '<p>Nothing due this week. Clear runway.</p>';
  MailApp.sendEmail({
    to: to, subject: 'Herd Co. week ahead — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d'),
    htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1b211d;max-width:620px">' + body +
      '<p style="margin-top:22px"><a href="' + ADMIN_URL + '" style="background:#163823;color:#fff;padding:9px 16px;border-radius:999px;text-decoration:none">Open Operations</a></p></div>'
  });
}

/* ---------------- contracts ---------------- */

var CONTRACT_KEYS = [
  ['provider', 'Hoosier Prairie Holdings, LLC d/b/a Herd Co Indy'], ['client', 'Client / project name'], ['org', 'Organization'],
  ['contactName', 'Client contact'], ['email', 'Client email'], ['phone', 'Client phone'], ['service', 'Service'],
  ['scope', 'What we\'re building'], ['value', 'Contract value ($)'], ['deposit', 'Deposit ($)'], ['balance', 'Value minus deposit ($)'],
  ['monthly', 'Recurring ($/mo)'], ['deliveryDue', 'Delivery date'], ['date', 'Today\'s date'], ['liveUrl', 'Live URL'], ['domain', 'Domain']
];

function longDate_(s) {
  if (!s) return '';
  var d = new Date(String(s).slice(0, 10) + 'T12:00:00');
  return isNaN(d) ? String(s) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM d, yyyy');
}

function contractValues_(pr) {
  var v = Number(pr.value || 0), dep = Number(pr.deposit || 0);
  return {
    provider: PROVIDER, client: pr.name || '', org: pr.org || pr.name || '', contactName: pr.contactName || '', email: pr.email || '',
    phone: pr.phone || '', service: pr.service || '', scope: pr.scope || '', value: money_(v), deposit: money_(dep),
    balance: money_(Math.max(0, v - dep)), monthly: pr.monthly ? money_(pr.monthly) : '$0', deliveryDue: longDate_(pr.deliveryDue),
    date: longDate_(today_()), liveUrl: pr.liveUrl || '', domain: pr.domain || ''
  };
}

function generateContract_(projectId) {
  var tpl = PropertiesService.getScriptProperties().getProperty('CONTRACT_TEMPLATE_ID');
  if (!tpl) throw new Error('No contract template set. Open ⋯ → Settings and add one (or create the starter template).');
  var pr = readTab_(TAB_PROJECTS).filter(function (x) { return String(x.id) === String(projectId); })[0];
  if (!pr) throw new Error('Project not found');
  var vals = contractValues_(pr);
  var name = 'Agreement — ' + (pr.name || 'Client') + ' — ' + today_();
  var copy = DriveApp.getFileById(tpl).makeCopy(name, contractsFolder_());
  var doc = DocumentApp.openById(copy.getId());
  var parts = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(function (x) { return x; });
  Object.keys(vals).forEach(function (k) {
    var val = String(vals[k]).replace(/\{\{/g, '{ {');
    var re = '\\{\\{\\s*' + k + '\\s*\\}\\}';
    parts.forEach(function (part) {
      var r, guard = 0;
      while ((r = part.findText(re)) && guard++ < 500) { // literal insert: no regex surprises with "$" in values
        var el = r.getElement().asText(), st = r.getStartOffset();
        el.deleteText(st, r.getEndOffsetInclusive());
        if (val) el.insertText(st, val);
      }
    });
  });
  doc.saveAndClose();
  var rec = writeObj_(TAB_FILES, {
    id: newId_('f'), projectId: pr.id, ts: nowIso_(), kind: 'Contract draft', name: name, date: today_(),
    url: 'https://docs.google.com/document/d/' + copy.getId() + '/edit', driveId: copy.getId(), mime: 'application/vnd.google-apps.document', size: ''
  }, FILE_FIELDS);
  addLog_(pr.id, 'file', 'Generated contract draft: ' + name);
  return rec;
}

function createContractTemplate_() {
  var doc = DocumentApp.create('Herd Co. — Client Agreement TEMPLATE');
  var b = doc.getBody();
  b.appendParagraph('HOW TO USE THIS TEMPLATE (delete this box when done)').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  b.appendParagraph('Paste your real client project agreement below this box. Wherever a client-specific detail goes, type the matching placeholder (with the double curly braces). "Generate contract" in /admin makes a copy of this Doc and fills every placeholder from the project.');
  var t = b.appendTable([['Placeholder', 'Filled with']].concat(CONTRACT_KEYS.map(function (k) { return ['{{' + k[0] + '}}', k[1]]; })));
  t.getRow(0).editAsText().setBold(true);
  b.appendHorizontalRule();
  b.appendParagraph('Client Project Agreement').setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph('This agreement is between {{provider}} ("Provider") and {{org}} ("Client"), dated {{date}}.');
  b.appendParagraph('Client contact: {{contactName}} · {{email}} · {{phone}}');
  b.appendParagraph('Project: {{client}} ({{service}})');
  b.appendParagraph('Scope: {{scope}}');
  b.appendParagraph('Price: {{value}} · Deposit: {{deposit}} · Due at delivery: {{balance}} · Recurring: {{monthly}}/mo');
  b.appendParagraph('Target delivery: {{deliveryDue}}');
  b.appendParagraph('[Paste the rest of your agreement terms here.]').editAsText().setItalic(true);
  doc.saveAndClose();
  var f = DriveApp.getFileById(doc.getId());
  f.moveTo(contractsFolder_());
  PropertiesService.getScriptProperties().setProperty('CONTRACT_TEMPLATE_ID', doc.getId());
  return getSettings_();
}
function createContractTemplateMenu() {
  var s = createContractTemplate_();
  SpreadsheetApp.getUi().alert('Template created and set:\n' + s.contractTemplateUrl);
}
