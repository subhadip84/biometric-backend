// core.js — Phase 1: the essential day-to-day functions, ported from Code.gs.
// (User management, imports, and announcements come in Phase 2/3.)

const sheetsApi = require('./sheets');

const HEADER_CANDIDATES = {
  name: ['studentname', 'name'],
  appNo: ['applicationno', 'appno', 'application', 'rollnumber', 'rollno', 'roll'],
  regNo: ['registrationnumber', 'regno', 'registration'],
  machineCode: ['machinecode', 'machine'],
  siteCode: ['sitecode'],
  studentType: ['studenttype', 'type']
};

function normalize(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function detectColumns(headers) {
  const col = {};
  Object.keys(HEADER_CANDIDATES).forEach(field => {
    let found = -1;
    for (let i = 0; i < headers.length; i++) {
      const norm = normalize(headers[i]);
      if (HEADER_CANDIDATES[field].some(cand => norm.indexOf(cand) !== -1)) {
        found = i;
        break;
      }
    }
    col[field] = found;
  });

  let statusCol = -1;
  for (let i = 0; i < headers.length; i++) {
    const norm = normalize(headers[i]);
    if (norm.indexOf('biometric') !== -1 || norm.indexOf('fingerprint') !== -1 ||
      (norm.indexOf('status') !== -1 && norm.indexOf('registration') === -1)) {
      statusCol = i;
      break;
    }
  }
  col.status = statusCol;

  let lockCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (normalize(headers[i]) === 'locked') { lockCol = i; break; }
  }
  col.lock = lockCol;

  let verifiedByCol = -1, verifiedAtCol = -1;
  let firstVerifiedByCol = -1, firstVerifiedAtCol = -1;
  for (let i = 0; i < headers.length; i++) {
    const norm = normalize(headers[i]);
    if (norm === 'firstverifiedby') { firstVerifiedByCol = i; continue; }
    if (norm === 'firstverifiedat') { firstVerifiedAtCol = i; continue; }
    if (verifiedByCol === -1 && norm.indexOf('verifiedby') !== -1) verifiedByCol = i;
    if (verifiedAtCol === -1 && norm.indexOf('verifiedat') !== -1) verifiedAtCol = i;
  }
  col.verifiedBy = verifiedByCol;
  col.verifiedAt = verifiedAtCol;
  col.firstVerifiedBy = firstVerifiedByCol;
  col.firstVerifiedAt = firstVerifiedAtCol;

  return col;
}

function colToLetter(idx) {
  let letter = '';
  let n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

async function ensureExtraColumns(sheetName, headers, col) {
  let changed = false;
  const additions = [];
  if (col.status === -1) { col.status = headers.length + additions.length; additions.push('Biometric Status'); changed = true; }
  if (col.lock === -1) { col.lock = headers.length + additions.length; additions.push('Locked'); changed = true; }
  if (col.verifiedBy === -1) { col.verifiedBy = headers.length + additions.length; additions.push('Verified By'); changed = true; }
  if (col.verifiedAt === -1) { col.verifiedAt = headers.length + additions.length; additions.push('Verified At'); changed = true; }
  if (col.firstVerifiedBy === -1) { col.firstVerifiedBy = headers.length + additions.length; additions.push('First Verified By'); changed = true; }
  if (col.firstVerifiedAt === -1) { col.firstVerifiedAt = headers.length + additions.length; additions.push('First Verified At'); changed = true; }

  if (changed) {
    const startCol = colToLetter(headers.length);
    await sheetsApi.writeRange(`${sheetName}!${startCol}1`, [additions]);
  }
  return col;
}

// ---------- User Accounts ----------

const USER_HEADERS = ['User ID', 'Password', 'Role', 'Name', 'Mobile No', 'Email', 'School', 'Permissions (JSON)', 'Must Change Password'];
const DEFAULT_USERS = {
  admin: { password: 'Akc@123', role: 'admin', name: 'Admin' },
  staff1: { password: 'Adamas@123', role: 'staff', name: 'Staff1' }
};

async function getAllUsers() {
  await sheetsApi.ensureSheet(sheetsApi.USER_SHEET_NAME, USER_HEADERS);
  const rows = await sheetsApi.readRange(`${sheetsApi.USER_SHEET_NAME}!A2:I`);
  if (!rows.length) {
    await writeAllUsers(DEFAULT_USERS);
    return JSON.parse(JSON.stringify(DEFAULT_USERS));
  }
  const users = {};
  rows.forEach(row => {
    const key = String(row[0] || '').trim();
    if (!key) return;
    const record = {
      password: String(row[1] || ''),
      role: String(row[2] || 'staff'),
      name: String(row[3] || ''),
      mobile: String(row[4] || ''),
      email: String(row[5] || ''),
      school: String(row[6] || '')
    };
    const permsJson = String(row[7] || '').trim();
    if (permsJson) {
      try { record.permissions = JSON.parse(permsJson); } catch (e) { /* ignore malformed */ }
    }
    if (String(row[8] || '').trim().toLowerCase() === 'yes') record.mustChangePassword = true;
    users[key] = record;
  });
  return users;
}

async function writeAllUsers(usersObj) {
  await sheetsApi.ensureSheet(sheetsApi.USER_SHEET_NAME, USER_HEADERS);
  await sheetsApi.clearRange(`${sheetsApi.USER_SHEET_NAME}!A2:I`);
  const keys = Object.keys(usersObj);
  if (!keys.length) return;
  const rows = keys.map(key => {
    const u = usersObj[key];
    return [
      key, u.password, u.role, u.name || '', u.mobile || '', u.email || '', u.school || '',
      u.permissions ? JSON.stringify(u.permissions) : '',
      u.mustChangePassword ? 'Yes' : ''
    ];
  });
  await sheetsApi.writeRange(`${sheetsApi.USER_SHEET_NAME}!A2`, rows);
}

const ALL_PERMISSION_KEYS = [
  'viewSummary', 'downloadCsv', 'unlockRecords', 'lockAll', 'viewActivityLog',
  'composeMessage', 'importStudents', 'importVerification', 'resetPasswords',
  'manageUsers', 'deleteStudent'
];

function effectivePermissions(userRecord) {
  if (userRecord.role !== 'admin') return {};
  if (!userRecord.permissions) {
    const all = {};
    ALL_PERMISSION_KEYS.forEach(k => { all[k] = true; });
    return all;
  }
  return userRecord.permissions;
}

function capitalizeFirst(s) {
  s = String(s || '');
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

// ---------- Admin password (stored in a small Settings sheet) ----------

const DEFAULT_ADMIN_PASSWORD = 'Password#1';
const SETTINGS_SHEET_NAME = 'Settings';

async function getAdminPassword() {
  await sheetsApi.ensureSheet(SETTINGS_SHEET_NAME, ['Key', 'Value']);
  const rows = await sheetsApi.readRange(`${SETTINGS_SHEET_NAME}!A2:B`);
  const row = rows.find(r => r[0] === 'ADMIN_PASSWORD');
  return row && row[1] ? row[1] : DEFAULT_ADMIN_PASSWORD;
}

async function setAdminPassword(newPassword) {
  await sheetsApi.ensureSheet(SETTINGS_SHEET_NAME, ['Key', 'Value']);
  const rows = await sheetsApi.readRange(`${SETTINGS_SHEET_NAME}!A2:B`);
  const idx = rows.findIndex(r => r[0] === 'ADMIN_PASSWORD');
  if (idx === -1) {
    await sheetsApi.appendRows(SETTINGS_SHEET_NAME, [['ADMIN_PASSWORD', newPassword]]);
  } else {
    await sheetsApi.writeRange(`${SETTINGS_SHEET_NAME}!A${idx + 2}:B${idx + 2}`, [['ADMIN_PASSWORD', newPassword]]);
  }
}

// ---------- Activity Log ----------

async function logActivity(actor, action, details) {
  try {
    await sheetsApi.ensureSheet(sheetsApi.ACTIVITY_LOG_SHEET_NAME, ['Timestamp', 'Actor', 'Action', 'Details']);
    const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19);
    await sheetsApi.appendRows(sheetsApi.ACTIVITY_LOG_SHEET_NAME, [[timestamp, actor || 'unknown', action, details || '']]);
  } catch (e) { /* never let logging break the main action */ }
}

async function getActivityLog(limit) {
  await sheetsApi.ensureSheet(sheetsApi.ACTIVITY_LOG_SHEET_NAME, ['Timestamp', 'Actor', 'Action', 'Details']);
  const rows = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:D`);
  const maxRows = Math.min(limit || 200, rows.length);
  const recent = rows.slice(-maxRows).reverse();
  return {
    ok: true,
    entries: recent.map(r => ({ timestamp: r[0] || '', actor: r[1] || '', action: r[2] || '', details: r[3] || '' }))
  };
}

// ---------- Core roster functions ----------

async function checkLogin(userId, password, deviceInfo) {
  const key = String(userId || '').trim();
  const device = deviceInfo ? ` — ${deviceInfo}` : '';
  const users = await getAllUsers();
  if (!key || !users.hasOwnProperty(key)) {
    await logActivity(key || '(blank)', 'Login Failed', 'Unknown User ID' + device);
    return { ok: false, error: 'Invalid User ID or password.' };
  }
  if (users[key].password !== password) {
    await logActivity(users[key].name || key, 'Login Failed', 'Incorrect password' + device);
    return { ok: false, error: 'Invalid User ID or password.' };
  }
  const displayName = users[key].name || capitalizeFirst(key);
  await logActivity(displayName, 'Login Successful', `${key} (${users[key].role})${device}`);
  return {
    ok: true,
    userId: key,
    isAdmin: users[key].role === 'admin',
    name: displayName,
    permissions: effectivePermissions(users[key]),
    mustChangePassword: !!users[key].mustChangePassword
  };
}

async function getStudents() {
  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  if (!data.length) return { ok: false, error: 'The sheet is empty.' };

  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const users = await getAllUsers();
  const nameToRole = {};
  Object.keys(users).forEach(uid => {
    const displayName = (users[uid].name || uid).trim().toLowerCase();
    if (displayName) nameToRole[displayName] = users[uid].role;
  });
  function isAdminLabel(rawLabel) {
    if (!rawLabel) return false;
    if (rawLabel.indexOf('[Admin]') !== -1) return true;
    const cleanName = rawLabel.replace(' [Admin]', '').trim().toLowerCase();
    return nameToRole[cleanName] === 'admin';
  }

  const students = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const name = col.name > -1 ? row[col.name] : '';
    const appNo = col.appNo > -1 ? row[col.appNo] : '';
    const regNo = col.regNo > -1 ? row[col.regNo] : '';
    const machineCode = col.machineCode > -1 ? row[col.machineCode] : '';
    const siteCode = col.siteCode > -1 ? row[col.siteCode] : '';
    const studentType = col.studentType > -1 ? row[col.studentType] : '';
    if (!name && !appNo && !regNo && !machineCode) continue;

    const statusVal = col.status > -1 ? row[col.status] : '';
    const v = String(statusVal || '').trim().toLowerCase();
    const isDone = (v === 'done' || v === 'yes' || v === 'true' || v === 'completed' || v === 'verified');

    const lockVal = col.lock > -1 ? row[col.lock] : '';
    const isLocked = String(lockVal || '').trim().toLowerCase() === 'yes';

    const verifiedByVal = col.verifiedBy > -1 ? (row[col.verifiedBy] || '') : '';
    const verifiedAtVal = col.verifiedAt > -1 ? (row[col.verifiedAt] || '') : '';
    const firstVerifiedByVal = col.firstVerifiedBy > -1 ? (row[col.firstVerifiedBy] || '') : '';
    const firstVerifiedAtVal = col.firstVerifiedAt > -1 ? (row[col.firstVerifiedAt] || '') : '';

    students.push({
      id: 'row' + (r + 1),
      row: r + 1,
      name: String(name || ''),
      appNo: String(appNo || ''),
      regNo: String(regNo || ''),
      machineCode: String(machineCode || ''),
      siteCode: String(siteCode || ''),
      studentType: String(studentType || ''),
      status: isDone ? 'done' : 'pending',
      locked: isLocked,
      verifiedBy: String(verifiedByVal),
      verifiedAt: String(verifiedAtVal),
      verifiedByIsAdmin: isAdminLabel(String(verifiedByVal)),
      firstVerifiedBy: String(firstVerifiedByVal),
      firstVerifiedAt: String(firstVerifiedAtVal),
      firstVerifiedByIsAdmin: isAdminLabel(String(firstVerifiedByVal))
    });
  }

  return { ok: true, students, file: sheetName };
}

async function updateStatus(rowId, status, userId) {
  if (status !== 'done' && status !== 'pending') {
    return { ok: false, error: "status must be 'done' or 'pending'" };
  }
  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const rowNum = parseInt(String(rowId).replace('row', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid student id.' };

  const row = data[rowNum - 1] || [];
  const currentLock = String(row[col.lock] || '').trim().toLowerCase();
  if (currentLock === 'yes') {
    return { ok: false, error: 'This record is locked. An admin needs to unlock it before it can be changed.' };
  }

  const value = status === 'done' ? 'Done' : 'Not Done';
  const users = await getAllUsers();
  const whoLabel = userId ? String(userId) : 'unknown';
  const verifierRole = (users[whoLabel] && users[whoLabel].role === 'admin') ? 'admin' : 'staff';
  const verifierName = (users[whoLabel] && users[whoLabel].name) ? users[whoLabel].name : whoLabel;
  const storedVerifiedBy = verifierName + (verifierRole === 'admin' ? ' [Admin]' : '');
  const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19);

  const updates = [
    { col: col.status, val: value },
    { col: col.lock, val: 'Yes' },
    { col: col.verifiedBy, val: storedVerifiedBy },
    { col: col.verifiedAt, val: timestamp }
  ];

  const existingFirstVerifiedBy = String(row[col.firstVerifiedBy] || '').trim();
  if (!existingFirstVerifiedBy) {
    updates.push({ col: col.firstVerifiedBy, val: storedVerifiedBy });
    updates.push({ col: col.firstVerifiedAt, val: timestamp });
  }

  for (const u of updates) {
    await sheetsApi.writeRange(`${sheetName}!${colToLetter(u.col)}${rowNum}`, [[u.val]]);
  }

  const studentName = col.name > -1 ? String(row[col.name] || '') : '';
  await logActivity(verifierName, status === 'done' ? 'Verified' : 'Marked Pending', `${studentName} (row ${rowNum})`);

  return { ok: true, verifiedByRole: verifierRole };
}

async function adminUnlock(rowId, password, actor) {
  const adminPassword = await getAdminPassword();
  if (password !== adminPassword) return { ok: false, error: 'Incorrect admin password.' };

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const rowNum = parseInt(String(rowId).replace('row', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid student id.' };

  await sheetsApi.writeRange(`${sheetName}!${colToLetter(col.lock)}${rowNum}`, [['No']]);

  const row = data[rowNum - 1] || [];
  const studentName = col.name > -1 ? String(row[col.name] || '') : '';
  await logActivity(actor || 'Admin', 'Unlocked Record', `${studentName} (row ${rowNum})`);

  return { ok: true };
}

async function adminLockAllDone(password, actor) {
  const adminPassword = await getAdminPassword();
  if (password !== adminPassword) return { ok: false, error: 'Incorrect admin password.' };

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  let count = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const statusVal = String(row[col.status] || '').trim().toLowerCase();
    const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true' || statusVal === 'completed' || statusVal === 'verified');
    const isLocked = String(row[col.lock] || '').trim().toLowerCase() === 'yes';
    if (isDone && !isLocked) {
      await sheetsApi.writeRange(`${sheetName}!${colToLetter(col.lock)}${r + 1}`, [['Yes']]);
      count++;
    }
  }

  await logActivity(actor || 'Admin', 'Bulk Lock Verified', `Locked ${count} record(s)`);
  return { ok: true, count };
}

module.exports = {
  normalize, detectColumns, colToLetter, ensureExtraColumns,
  getAllUsers, writeAllUsers, effectivePermissions, capitalizeFirst, ALL_PERMISSION_KEYS,
  getAdminPassword, setAdminPassword,
  logActivity, getActivityLog,
  checkLogin, getStudents, updateStatus, adminUnlock, adminLockAllDone
};
