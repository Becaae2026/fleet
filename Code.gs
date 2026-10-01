/****************************************************************************************
 * AL ANSARI EXCHANGE — COMPANY-OWNED VEHICLE MANAGEMENT
 * Google Apps Script web app — two files: Code.gs (this file) + index.html (dashboard page).
 * Data: Google Sheets  |  Files: Google Drive  |  Access: username + password (hashed)
 *
 * QUICK START
 *  1. script.google.com → New project → paste this file into Code.gs.
 *     Then + (Add a file) → HTML → name it exactly: index → paste index.html into it.
 *  2. Project Settings (gear) → Time zone → "(GMT+03:00) Kuwait".
 *  3. Select function "setup" → Run → approve permissions.
 *     Copy the temporary admin password from the Execution log.
 *  4. Deploy → New deployment → type "Web app"
 *       Execute as: Me   |   Who has access: Anyone  (or: Anyone within your organisation)
 *  5. Open the web app URL → sign in as "admin" → set a new password →
 *     add users under Administration.
 *
 * ROLES
 *  Admin  – everything, plus users, settings, backups and audit log
 *  Editor – dashboard, view/download, add/edit/delete vehicles and documents
 *  Viewer – dashboard and view/download only (Civil ID numbers masked)
 *
 * RECOVERY: if the admin password is lost, run resetAdminPassword() from the editor.
 ****************************************************************************************/

const APP = {
  TITLE: 'Company Vehicle Management',
  ORG: 'Al Ansari Exchange',
  LOGO_URL: '',                 // Optional: public URL of the official logo artwork (white/negative version)
  ROOT_FOLDER_NAME: 'AAE Vehicle Management',
  SESSION_SECONDS: 6 * 60 * 60, // max 6 hours (CacheService limit), sliding
  IDLE_MINUTES: 30,             // auto sign-out after inactivity
  MAX_FILE_MB: 10,
  MAX_FAILED: 5,                // failed logins before lock
  LOCK_MINUTES: 15,
  BACKUP_KEEP: 30,              // number of daily spreadsheet backups kept
  HASH_ROUNDS: 500,
  ALLOWED_MIME: ['application/pdf', 'image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif']
};

const SH = { V: 'Vehicles', D: 'Documents', U: 'Users', A: 'AuditLog', S: 'Settings' };
const V_HEAD = ['VehicleID', 'Model', 'Year', 'RegNo', 'Department', 'Owner', 'Designation',
  'DaftarExpiry', 'InsuranceExpiry', 'LicenseExpiry', 'CivilIdNo', 'CivilIdExpiry', 'Remarks',
  'FolderId', 'CreatedAt', 'CreatedBy', 'UpdatedAt', 'UpdatedBy'];
const D_HEAD = ['DocID', 'VehicleID', 'DocType', 'FileName', 'MimeType', 'SizeKB', 'FileId', 'UploadedAt', 'UploadedBy'];
const U_HEAD = ['Username', 'FullName', 'Role', 'Active', 'PasswordHash', 'Salt', 'MustChange',
  'FailedAttempts', 'LockedUntil', 'LastLogin'];
const A_HEAD = ['Timestamp', 'User', 'Action', 'Details'];
const DOC_TYPES = ['Civil ID', 'Driving Licence', 'Daftar', 'Insurance', 'Other'];
const EXPIRY = [
  { key: 'DaftarExpiry', label: 'Daftar' },
  { key: 'InsuranceExpiry', label: 'Insurance' },
  { key: 'LicenseExpiry', label: 'Driving Licence' },
  { key: 'CivilIdExpiry', label: 'Civil ID' }
];
const ROLES = ['Admin', 'Editor', 'Viewer'];
const ALL = ROLES, EDIT = ['Admin', 'Editor'], ADMIN = ['Admin'];

/* ===================================== WEB APP ===================================== */

function doGet() {
  const html = HtmlService.createHtmlOutputFromFile('index').getContent()
    .replace('%%LOGO_JSON%%', JSON.stringify(APP.LOGO_URL || ''))
    .replace('%%MAX_MB%%', String(APP.MAX_FILE_MB))
    .replace('%%IDLE_MIN%%', String(APP.IDLE_MINUTES))
    .replace(/%%TITLE%%/g, APP.TITLE)
    .replace(/%%ORG%%/g, APP.ORG);
  return HtmlService.createHtmlOutput(html)
    .setTitle(APP.TITLE + ' | ' + APP.ORG)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ====================================== SETUP ====================================== */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SS_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('AAE Vehicle Management — Data');
    props.setProperty('SS_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());

  const root = folderProp_('ROOT_ID', () => DriveApp.createFolder(APP.ROOT_FOLDER_NAME));
  folderProp_('DOCS_ID', () => root.createFolder('Vehicle Documents'));
  folderProp_('BACKUP_ID', () => root.createFolder('Backups'));
  try { DriveApp.getFileById(ss.getId()).moveTo(root); } catch (e) { /* already there or not movable */ }

  const v = ensureSheet_(ss, SH.V, V_HEAD);
  ['DaftarExpiry', 'InsuranceExpiry', 'LicenseExpiry', 'CivilIdExpiry'].forEach(k =>
    v.getRange(2, V_HEAD.indexOf(k) + 1, v.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd'));
  ['RegNo', 'CivilIdNo'].forEach(k =>
    v.getRange(2, V_HEAD.indexOf(k) + 1, v.getMaxRows() - 1, 1).setNumberFormat('@'));
  ensureSheet_(ss, SH.D, D_HEAD);
  ensureSheet_(ss, SH.U, U_HEAD);
  ensureSheet_(ss, SH.A, A_HEAD);
  const s = ensureSheet_(ss, SH.S, ['Key', 'Value', 'Notes']);
  if (s.getLastRow() < 2) {
    s.getRange(2, 1, 2, 3).setValues([
      ['NEAR_EXPIRY_DAYS', 30, 'Documents expiring within this many days are flagged as near expiry'],
      ['ALERT_EMAILS', '', 'Comma-separated emails for the daily expiry digest (blank = off)']
    ]);
  }
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  let msg = 'Setup complete. Data sheet: ' + ss.getUrl();
  if (rows_(SH.U, U_HEAD).length === 0) {
    const temp = tempPw_(), salt = Utilities.getUuid();
    sheet_(SH.U).appendRow(['admin', 'Administrator', 'Admin', 'Yes', hash_(temp, salt), salt, 'Yes', 0, '', '']);
    msg += '\n\nADMIN LOGIN → username: admin   temporary password: ' + temp +
      '\n(You will be asked to set a new password at first sign-in.)';
  }

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyJob')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(6).create();

  audit_('system', 'SETUP', 'Setup run');
  Logger.log(msg);
  return msg;
}

function resetAdminPassword() {
  const sh = sheet_(SH.U);
  const u = rows_(SH.U, U_HEAD).find(x => String(x.Username).toLowerCase() === 'admin');
  const temp = tempPw_(), salt = Utilities.getUuid();
  if (u) {
    setCells_(sh, u._row, U_HEAD, { PasswordHash: hash_(temp, salt), Salt: salt, MustChange: 'Yes',
      FailedAttempts: 0, LockedUntil: '', Active: 'Yes', Role: 'Admin' });
  } else {
    sh.appendRow(['admin', 'Administrator', 'Admin', 'Yes', hash_(temp, salt), salt, 'Yes', 0, '', '']);
  }
  bumpUser_('admin');
  audit_('system', 'ADMIN_RESET', 'Admin password reset from script editor');
  Logger.log('username: admin   temporary password: ' + temp);
}

/* ====================================== AUTH ====================================== */

function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  if (!username || !password) throw new Error('Enter your username and password.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SH.U), now = new Date();
    const u = rows_(SH.U, U_HEAD).find(x => String(x.Username).toLowerCase() === username);
    if (!u) { hash_(password, 'x'); audit_(username, 'LOGIN_FAILED', 'Unknown user'); throw new Error('Invalid username or password.'); }
    if (u.LockedUntil instanceof Date && u.LockedUntil > now) {
      throw new Error('Account temporarily locked after repeated failed attempts. Try again after ' + fmtDT_(u.LockedUntil) + '.');
    }
    if (hash_(password, u.Salt) !== u.PasswordHash) {
      const f = Number(u.FailedAttempts || 0) + 1;
      if (f >= APP.MAX_FAILED) {
        setCells_(sh, u._row, U_HEAD, { FailedAttempts: 0, LockedUntil: new Date(now.getTime() + APP.LOCK_MINUTES * 60000) });
        audit_(u.Username, 'ACCOUNT_LOCKED', f + ' failed attempts');
      } else {
        setCells_(sh, u._row, U_HEAD, { FailedAttempts: f });
      }
      audit_(u.Username, 'LOGIN_FAILED', 'Wrong password');
      throw new Error('Invalid username or password.');
    }
    if (String(u.Active) !== 'Yes') throw new Error('This account is inactive. Contact the administrator.');
    setCells_(sh, u._row, U_HEAD, { FailedAttempts: 0, LockedUntil: '', LastLogin: now });
    audit_(u.Username, 'LOGIN', '');
    const mc = String(u.MustChange) === 'Yes';
    return { token: newSession_(u, mc), user: { username: u.Username, fullName: str_(u.FullName), role: u.Role, mustChange: mc } };
  } finally { lock.releaseLock(); }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s:' + token);
  return true;
}

function changePassword(token, oldPw, newPw) {
  const s = session_(token, ALL, true);
  const err = policy_(newPw);
  if (err) throw new Error(err);
  if (oldPw === newPw) throw new Error('The new password must be different from the current one.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SH.U);
    const u = rows_(SH.U, U_HEAD).find(x => x.Username === s.u);
    if (!u) throw new Error('SESSION_EXPIRED');
    if (hash_(String(oldPw || ''), u.Salt) !== u.PasswordHash) throw new Error('Current password is incorrect.');
    const salt = Utilities.getUuid();
    setCells_(sh, u._row, U_HEAD, { PasswordHash: hash_(newPw, salt), Salt: salt, MustChange: 'No' });
    bumpUser_(u.Username); // signs out every other session of this user
    CacheService.getScriptCache().remove('s:' + token);
    audit_(u.Username, 'PASSWORD_CHANGED', '');
    return { token: newSession_(u, false) };
  } finally { lock.releaseLock(); }
}

/* ====================================== DATA ====================================== */

function getData(token) {
  const s = session_(token, ALL);
  const st = settings_();
  const docs = rows_(SH.D, D_HEAD);
  const cnt = {};
  docs.forEach(d => { cnt[d.VehicleID] = (cnt[d.VehicleID] || 0) + 1; });
  const vehicles = rows_(SH.V, V_HEAD).map(v => ({
    id: str_(v.VehicleID), model: str_(v.Model), year: str_(v.Year), regNo: str_(v.RegNo),
    department: str_(v.Department), owner: str_(v.Owner), designation: str_(v.Designation),
    daftar: fmtD_(v.DaftarExpiry), insurance: fmtD_(v.InsuranceExpiry), license: fmtD_(v.LicenseExpiry),
    civilId: s.r === 'Viewer' ? mask_(v.CivilIdNo) : str_(v.CivilIdNo), civilExpiry: fmtD_(v.CivilIdExpiry),
    remarks: str_(v.Remarks), files: cnt[v.VehicleID] || 0,
    updatedAt: fmtDT_(v.UpdatedAt || v.CreatedAt), updatedBy: str_(v.UpdatedBy || v.CreatedBy)
  }));
  const documents = docs.map(d => ({
    id: str_(d.DocID), vehicleId: str_(d.VehicleID), type: str_(d.DocType), name: str_(d.FileName),
    mime: str_(d.MimeType), sizeKB: Number(d.SizeKB) || 0, uploadedAt: fmtDT_(d.UploadedAt), uploadedBy: str_(d.UploadedBy)
  }));
  const out = { user: { username: s.u, fullName: s.n, role: s.r }, today: fmtD_(new Date()),
    nearDays: st.NEAR_EXPIRY_DAYS, vehicles: vehicles, documents: documents, docTypes: DOC_TYPES };
  if (s.r === 'Admin') {
    out.sheetUrl = ss_().getUrl();
    out.folderUrl = 'https://drive.google.com/drive/folders/' + prop_('ROOT_ID');
    out.alertEmails = st.ALERT_EMAILS;
  }
  return out;
}

function saveVehicle(token, v) {
  const s = session_(token, EDIT);
  v = v || {};
  const rec = {
    Model: clean_(v.model, 80), RegNo: clean_(v.regNo, 30).toUpperCase(), Department: clean_(v.department, 80),
    Owner: clean_(v.owner, 100), Designation: clean_(v.designation, 80),
    DaftarExpiry: parseD_(v.daftar, 'Daftar expiry'), InsuranceExpiry: parseD_(v.insurance, 'Insurance expiry'),
    LicenseExpiry: parseD_(v.license, 'Driving licence expiry'), CivilIdNo: clean_(v.civilId, 20),
    CivilIdExpiry: parseD_(v.civilExpiry, 'Civil ID expiry'), Remarks: clean_(v.remarks, 1000)
  };
  if (!rec.Model || !rec.RegNo || !rec.Owner) throw new Error('Vehicle model, registration number and owner are required.');
  const y = String(v.year || '').trim();
  if (y && (!/^\d{4}$/.test(y) || +y < 1950 || +y > new Date().getFullYear() + 1)) throw new Error('Enter a valid year of manufacture.');
  rec.Year = y ? +y : '';
  if (rec.CivilIdNo && !/^\d{12}$/.test(rec.CivilIdNo)) throw new Error('Civil ID number must be 12 digits.');

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.V), all = rows_(SH.V, V_HEAD), now = new Date();
    const norm = x => String(x).replace(/[\s\-\/']/g, '').toUpperCase();
    const dup = all.find(x => norm(x.RegNo) === norm(rec.RegNo) && String(x.VehicleID) !== String(v.id || ''));
    if (dup) throw new Error('Registration number ' + rec.RegNo + ' already exists (' + dup.Owner + ').');

    if (!v.id) {
      rec.VehicleID = nextId_(all.map(x => x.VehicleID), 'VH');
      rec.FolderId = docsFolder_().createFolder(folderName_(rec)).getId();
      rec.CreatedAt = now; rec.CreatedBy = s.u; rec.UpdatedAt = now; rec.UpdatedBy = s.u;
      sh.appendRow(V_HEAD.map(h => rec[h] === undefined ? '' : rec[h]));
      audit_(s.u, 'VEHICLE_ADDED', rec.VehicleID + ' ' + rec.RegNo);
      return { id: rec.VehicleID, created: true };
    }
    const cur = all.find(x => String(x.VehicleID) === String(v.id));
    if (!cur) throw new Error('Vehicle not found. It may have been deleted.');
    const merged = {};
    V_HEAD.forEach(h => { merged[h] = rec[h] !== undefined ? rec[h] : cur[h]; });
    merged.UpdatedAt = now; merged.UpdatedBy = s.u;
    sh.getRange(cur._row, 1, 1, V_HEAD.length).setValues([V_HEAD.map(h => merged[h])]);
    if (cur.RegNo !== rec.RegNo || cur.Owner !== rec.Owner) {
      try { DriveApp.getFolderById(cur.FolderId).setName(folderName_(rec)); } catch (e) { /* folder missing */ }
    }
    audit_(s.u, 'VEHICLE_UPDATED', cur.VehicleID + ' ' + rec.RegNo);
    return { id: cur.VehicleID, created: false };
  } finally { lock.releaseLock(); }
}

function deleteVehicle(token, id) {
  const s = session_(token, EDIT);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = rows_(SH.V, V_HEAD).find(x => String(x.VehicleID) === String(id));
    if (!v) throw new Error('Vehicle not found.');
    try { if (v.FolderId) DriveApp.getFolderById(v.FolderId).setTrashed(true); } catch (e) { /* ignore */ }
    const dsh = sheet_(SH.D);
    rows_(SH.D, D_HEAD).filter(d => String(d.VehicleID) === String(id)).map(d => d._row)
      .sort((a, b) => b - a).forEach(r => dsh.deleteRow(r));
    sheet_(SH.V).deleteRow(v._row);
    audit_(s.u, 'VEHICLE_DELETED', v.VehicleID + ' ' + v.RegNo + ' (files moved to Drive trash)');
    return true;
  } finally { lock.releaseLock(); }
}

/* ==================================== DOCUMENTS ==================================== */

function uploadDocument(token, vehicleId, docType, fileName, mime, b64) {
  const s = session_(token, EDIT);
  if (DOC_TYPES.indexOf(docType) < 0) throw new Error('Choose a valid document type.');
  mime = String(mime || '').toLowerCase();
  if (APP.ALLOWED_MIME.indexOf(mime) < 0) throw new Error('Only PDF, JPG, PNG, WEBP or HEIC files are allowed.');
  const bytes = Utilities.base64Decode(String(b64 || ''));
  if (!bytes.length) throw new Error('The file is empty.');
  if (bytes.length > APP.MAX_FILE_MB * 1048576) throw new Error('File is larger than ' + APP.MAX_FILE_MB + ' MB.');
  if (!magicOk_(bytes, mime)) throw new Error('The file content does not match its type.');

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const vsh = sheet_(SH.V);
    const v = rows_(SH.V, V_HEAD).find(x => String(x.VehicleID) === String(vehicleId));
    if (!v) throw new Error('Vehicle not found. Save the vehicle first.');
    let folder;
    try { folder = DriveApp.getFolderById(v.FolderId); if (folder.isTrashed()) throw 0; }
    catch (e) {
      folder = docsFolder_().createFolder(folderName_(v));
      setCells_(vsh, v._row, V_HEAD, { FolderId: folder.getId() });
    }
    const safe = String(fileName || 'file').replace(/[^\w.\- ]+/g, '_').slice(-80);
    const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmm');
    const reg = String(v.RegNo).replace(/[^\w\-]+/g, '');
    const name = reg + '_' + docType.replace(/\s+/g, '') + '_' + stamp + '_' + safe;
    const file = folder.createFile(Utilities.newBlob(bytes, mime, name));
    file.setDescription('Uploaded by ' + s.u + ' via Vehicle Management');
    const docId = nextId_(rows_(SH.D, D_HEAD).map(d => d.DocID), 'DOC');
    sheet_(SH.D).appendRow([docId, v.VehicleID, docType, name, mime, Math.ceil(bytes.length / 1024), file.getId(), new Date(), s.u]);
    audit_(s.u, 'DOC_UPLOADED', v.RegNo + ' ' + docType + ' ' + name);
    return { id: docId };
  } finally { lock.releaseLock(); }
}

function getDocument(token, docId) {
  const s = session_(token, ALL);
  const d = rows_(SH.D, D_HEAD).find(x => String(x.DocID) === String(docId));
  if (!d) throw new Error('Document not found.');
  let blob;
  try { blob = DriveApp.getFileById(d.FileId).getBlob(); }
  catch (e) { throw new Error('The file is no longer available in Google Drive.'); }
  audit_(s.u, 'DOC_VIEWED', d.DocID + ' ' + d.FileName);
  return { name: str_(d.FileName), mime: str_(d.MimeType) || blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

function deleteDocument(token, docId) {
  const s = session_(token, EDIT);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = rows_(SH.D, D_HEAD).find(x => String(x.DocID) === String(docId));
    if (!d) throw new Error('Document not found.');
    try { DriveApp.getFileById(d.FileId).setTrashed(true); } catch (e) { /* already gone */ }
    sheet_(SH.D).deleteRow(d._row);
    audit_(s.u, 'DOC_DELETED', d.DocID + ' ' + d.FileName + ' (moved to Drive trash)');
    return true;
  } finally { lock.releaseLock(); }
}

/* ===================================== ADMIN ===================================== */

function listUsers(token) {
  session_(token, ADMIN);
  const now = new Date();
  return rows_(SH.U, U_HEAD).map(u => ({
    username: str_(u.Username), fullName: str_(u.FullName), role: str_(u.Role), active: String(u.Active) === 'Yes',
    locked: u.LockedUntil instanceof Date && u.LockedUntil > now, mustChange: String(u.MustChange) === 'Yes',
    lastLogin: fmtDT_(u.LastLogin)
  }));
}

function saveUser(token, u) {
  const s = session_(token, ADMIN);
  u = u || {};
  const username = String(u.username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Username: 3–30 characters, letters, numbers, dot, dash or underscore.');
  if (ROLES.indexOf(u.role) < 0) throw new Error('Choose a valid role.');
  const active = u.active === false ? 'No' : 'Yes';
  const fullName = clean_(u.fullName, 80);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SH.U), users = rows_(SH.U, U_HEAD);
    const ex = users.find(x => String(x.Username).toLowerCase() === username);
    if (u.isNew) {
      if (ex) throw new Error('Username already exists.');
      const temp = tempPw_(), salt = Utilities.getUuid();
      sh.appendRow([username, fullName, u.role, active, hash_(temp, salt), salt, 'Yes', 0, '', '']);
      audit_(s.u, 'USER_ADDED', username + ' (' + u.role + ')');
      return { tempPassword: temp };
    }
    if (!ex) throw new Error('User not found.');
    if (ex.Username === s.u && (u.role !== 'Admin' || active !== 'Yes')) throw new Error('You cannot remove your own administrator access.');
    const admins = users.filter(x => x.Role === 'Admin' && String(x.Active) === 'Yes' && x.Username !== ex.Username).length;
    if (admins === 0 && (u.role !== 'Admin' || active !== 'Yes')) throw new Error('At least one active administrator is required.');
    setCells_(sh, ex._row, U_HEAD, { FullName: fullName, Role: u.role, Active: active });
    if (ex.Role !== u.role || String(ex.Active) !== active) bumpUser_(ex.Username);
    audit_(s.u, 'USER_UPDATED', username + ' role=' + u.role + ' active=' + active);
    return {};
  } finally { lock.releaseLock(); }
}

function resetUserPassword(token, username) {
  const s = session_(token, ADMIN);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const u = rows_(SH.U, U_HEAD).find(x => x.Username === username);
    if (!u) throw new Error('User not found.');
    const temp = tempPw_(), salt = Utilities.getUuid();
    setCells_(sheet_(SH.U), u._row, U_HEAD, { PasswordHash: hash_(temp, salt), Salt: salt, MustChange: 'Yes', FailedAttempts: 0, LockedUntil: '' });
    bumpUser_(username);
    audit_(s.u, 'PASSWORD_RESET', username);
    return { tempPassword: temp };
  } finally { lock.releaseLock(); }
}

function deleteUser(token, username) {
  const s = session_(token, ADMIN);
  if (username === s.u) throw new Error('You cannot delete your own account.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const users = rows_(SH.U, U_HEAD);
    const u = users.find(x => x.Username === username);
    if (!u) throw new Error('User not found.');
    if (u.Role === 'Admin' && users.filter(x => x.Role === 'Admin' && String(x.Active) === 'Yes').length <= 1) {
      throw new Error('At least one active administrator is required.');
    }
    sheet_(SH.U).deleteRow(u._row);
    bumpUser_(username);
    audit_(s.u, 'USER_DELETED', username);
    return true;
  } finally { lock.releaseLock(); }
}

function saveSettings(token, x) {
  const s = session_(token, ADMIN);
  x = x || {};
  const days = parseInt(x.nearDays, 10);
  if (!(days >= 1 && days <= 365)) throw new Error('Near-expiry window must be between 1 and 365 days.');
  const emails = String(x.alertEmails || '').split(/[,;\s]+/).map(e => e.trim()).filter(Boolean);
  emails.forEach(e => { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('Invalid email: ' + e); });
  setSetting_('NEAR_EXPIRY_DAYS', days);
  setSetting_('ALERT_EMAILS', emails.join(', '));
  audit_(s.u, 'SETTINGS_UPDATED', 'nearDays=' + days + ' emails=' + emails.length);
  return true;
}

function getAudit(token) {
  session_(token, ADMIN);
  const sh = sheet_(SH.A), last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(300, last - 1);
  return sh.getRange(last - n + 1, 1, n, A_HEAD.length).getValues().reverse()
    .map(r => ({ time: fmtDT_(r[0]), user: str_(r[1]), action: str_(r[2]), details: str_(r[3]) }));
}

function backupNow(token) {
  const s = session_(token, ADMIN);
  const name = backup_();
  audit_(s.u, 'BACKUP', name);
  return name;
}

function logEvent(token, action, details) {
  const s = session_(token, ALL);
  if (['EXPORT_CSV'].indexOf(action) < 0) return false;
  audit_(s.u, action, clean_(details, 200));
  return true;
}

/* ================================ DAILY JOB (trigger) ================================ */

function dailyJob() {
  try { backup_(); audit_('system', 'BACKUP', 'Daily backup'); } catch (e) { audit_('system', 'BACKUP_FAILED', String(e)); }
  try { sendAlertEmail_(); } catch (e) { audit_('system', 'ALERT_EMAIL_FAILED', String(e)); }
}

function backup_() {
  const folder = DriveApp.getFolderById(prop_('BACKUP_ID'));
  const name = 'Vehicle Data Backup ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HHmm');
  DriveApp.getFileById(ss_().getId()).makeCopy(name, folder);
  const files = [], it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(APP.BACKUP_KEEP).forEach(f => f.setTrashed(true));
  return name;
}

function sendAlertEmail_() {
  const st = settings_();
  if (!st.ALERT_EMAILS) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const near = [], expired = [];
  rows_(SH.V, V_HEAD).forEach(v => EXPIRY.forEach(f => {
    const d = v[f.key];
    if (!(d instanceof Date)) return;
    const days = Math.round((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - today) / 86400000);
    const row = { reg: str_(v.RegNo), owner: str_(v.Owner), dept: str_(v.Department), doc: f.label, date: fmtD_(d), days: days };
    if (days < 0) expired.push(row); else if (days <= st.NEAR_EXPIRY_DAYS) near.push(row);
  }));
  if (!near.length && !expired.length) return;
  near.sort((a, b) => a.days - b.days); expired.sort((a, b) => a.days - b.days);
  const tbl = (list, lbl) => '<h3 style="color:#002664;font-family:Verdana">' + lbl + ' (' + list.length + ')</h3>' +
    '<table cellpadding="6" style="border-collapse:collapse;font-family:Verdana;font-size:12px">' +
    '<tr style="background:#002664;color:#fff"><th>Reg. No.</th><th>Owner</th><th>Department</th><th>Document</th><th>Expiry</th><th>Days</th></tr>' +
    list.map(r => '<tr style="border-bottom:1px solid #ddd"><td>' + h_(r.reg) + '</td><td>' + h_(r.owner) + '</td><td>' + h_(r.dept) +
      '</td><td>' + r.doc + '</td><td>' + r.date + '</td><td>' + (r.days < 0 ? Math.abs(r.days) + ' overdue' : r.days + ' left') + '</td></tr>').join('') +
    '</table>';
  MailApp.sendEmail({
    to: st.ALERT_EMAILS,
    subject: 'Vehicle documents: ' + expired.length + ' expired, ' + near.length + ' near expiry',
    htmlBody: '<p style="font-family:Verdana;font-size:13px">Daily summary from ' + APP.TITLE + '.</p>' +
      (expired.length ? tbl(expired, 'Expired documents') : '') + (near.length ? tbl(near, 'Expiring within ' + st.NEAR_EXPIRY_DAYS + ' days') : '')
  });
}

/* ==================================== HELPERS ==================================== */

function session_(token, roles, allowMustChange) {
  if (!token || typeof token !== 'string' || token.length > 120) throw new Error('SESSION_EXPIRED');
  const c = CacheService.getScriptCache(), raw = c.get('s:' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  const s = JSON.parse(raw);
  if (s.v !== userVer_(s.u)) { c.remove('s:' + token); throw new Error('SESSION_EXPIRED'); }
  if (s.mc && !allowMustChange) throw new Error('PASSWORD_CHANGE_REQUIRED');
  c.put('s:' + token, raw, APP.SESSION_SECONDS);
  if (roles && roles.indexOf(s.r) < 0) throw new Error('You do not have permission for this action.');
  return s;
}
function newSession_(u, mc) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('s:' + token,
    JSON.stringify({ u: u.Username, n: str_(u.FullName), r: u.Role, v: userVer_(u.Username), mc: !!mc }), APP.SESSION_SECONDS);
  return token;
}
function userVer_(u) { return PropertiesService.getScriptProperties().getProperty('uv:' + String(u).toLowerCase()) || '0'; }
function bumpUser_(u) {
  const p = PropertiesService.getScriptProperties(), k = 'uv:' + String(u).toLowerCase();
  p.setProperty(k, String(Number(p.getProperty(k) || 0) + 1));
}
function hash_(pw, salt) {
  let h = String(salt) + '|' + String(pw);
  for (let i = 0; i < APP.HASH_ROUNDS; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + salt, Utilities.Charset.UTF_8));
  }
  return h;
}
function policy_(pw) {
  pw = String(pw || '');
  if (pw.length < 8) return 'Password must be at least 8 characters.';
  if (!/[A-Za-z]/.test(pw) || !/\d/.test(pw)) return 'Password must contain letters and numbers.';
  return '';
}
function tempPw_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  for (;;) {
    const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
    let s = '';
    for (let i = 0; i < 10; i++) s += chars[(b[i] & 0xff) % chars.length];
    if (!policy_(s)) return s;
  }
}
function magicOk_(b, mime) {
  const x = i => b[i] & 0xff;
  if (mime === 'application/pdf') return x(0) === 0x25 && x(1) === 0x50 && x(2) === 0x44 && x(3) === 0x46;
  if (mime === 'image/jpeg') return x(0) === 0xff && x(1) === 0xd8;
  if (mime === 'image/png') return x(0) === 0x89 && x(1) === 0x50 && x(2) === 0x4e && x(3) === 0x47;
  if (mime === 'image/webp') return x(0) === 0x52 && x(1) === 0x49 && x(2) === 0x46 && x(3) === 0x46;
  return b.length > 12; // HEIC/HEIF
}
function ss_() {
  const id = prop_('SS_ID');
  if (!id) throw new Error('Setup has not been run. Open the script editor and run setup().');
  return SpreadsheetApp.openById(id);
}
function sheet_(n) {
  const s = ss_().getSheetByName(n);
  if (!s) throw new Error('Missing sheet "' + n + '". Run setup() again.');
  return s;
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function folderProp_(key, create) {
  const p = PropertiesService.getScriptProperties(), id = p.getProperty(key);
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* recreate */ } }
  const f = create(); p.setProperty(key, f.getId()); return f;
}
function docsFolder_() { return DriveApp.getFolderById(prop_('DOCS_ID')); }
function folderName_(r) { return String(r.RegNo).replace(/^'/, '') + ' - ' + String(r.Owner).replace(/^'/, ''); }
function ensureSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold')
    .setBackground('#002664').setFontColor('#FFFFFF').setFontFamily('Verdana');
  sh.setFrozenRows(1);
  return sh;
}
function rows_(name, head) {
  const sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, head.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => { o[h] = r[j]; });
    return o;
  }).filter(o => String(o[head[0]]).trim() !== '');
}
function setCells_(sh, row, head, obj) {
  Object.keys(obj).forEach(k => sh.getRange(row, head.indexOf(k) + 1).setValue(obj[k]));
}
function settings_() {
  const out = { NEAR_EXPIRY_DAYS: 30, ALERT_EMAILS: '' };
  const sh = ss_().getSheetByName(SH.S);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  out.NEAR_EXPIRY_DAYS = parseInt(out.NEAR_EXPIRY_DAYS, 10) || 30;
  out.ALERT_EMAILS = String(out.ALERT_EMAILS || '');
  return out;
}
function setSetting_(k, v) {
  const sh = sheet_(SH.S), last = sh.getLastRow();
  const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]) : [];
  const i = keys.indexOf(k);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(v); else sh.appendRow([k, v, '']);
}
function audit_(user, action, details) {
  try {
    const sh = sheet_(SH.A);
    sh.appendRow([new Date(), String(user || ''), action, clean_(details, 500)]);
    if (sh.getLastRow() > 6000) sh.deleteRows(2, 1000);
  } catch (e) { /* never block the main action */ }
}
function nextId_(ids, prefix) {
  let max = 0;
  ids.forEach(id => { const n = parseInt(String(id).replace(prefix, ''), 10); if (n > max) max = n; });
  return prefix + String(max + 1).padStart(prefix === 'VH' ? 4 : 5, '0');
}
function tz_() { return Session.getScriptTimeZone(); }
function fmtD_(v) { return (v instanceof Date && !isNaN(v)) ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : ''; }
function fmtDT_(v) { return (v instanceof Date && !isNaN(v)) ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd HH:mm') : ''; }
function str_(v) { return v instanceof Date ? fmtD_(v) : String(v == null ? '' : v); }
function h_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function mask_(v) { const s = String(v || ''); return s ? '••••••••' + s.slice(-4) : ''; }
function parseD_(s, label) {
  s = String(s || '').trim();
  if (!s) return '';
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Invalid date for ' + label + '.');
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  if (d.getMonth() !== +m[2] - 1 || +m[1] < 1990 || +m[1] > 2100) throw new Error('Invalid date for ' + label + '.');
  return d;
}
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, ' ').trim();
  if (max) s = s.slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // block spreadsheet formula injection
  return s;
}
