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

const USER_HEADERS = ['User ID', 'Password', 'Role', 'Name', 'Mobile No', 'Email', 'School', 'Permissions (JSON)', 'Must Change Password', 'Demo Welcome Shown'];
const DEFAULT_USERS = {
  admin: { password: 'Akc@123', role: 'admin', name: 'Admin' },
  staff1: { password: 'Adamas@123', role: 'staff', name: 'Staff1' }
};

async function getAllUsers() {
  await sheetsApi.ensureSheet(sheetsApi.USER_SHEET_NAME, USER_HEADERS);
  const rows = await sheetsApi.readRange(`${sheetsApi.USER_SHEET_NAME}!A2:J`);
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
    if (String(row[9] || '').trim().toLowerCase() === 'yes') record.demoWelcomeShown = true;
    users[key] = record;
  });
  return users;
}

async function writeAllUsers(usersObj) {
  await sheetsApi.ensureSheet(sheetsApi.USER_SHEET_NAME, USER_HEADERS);
  await sheetsApi.clearRange(`${sheetsApi.USER_SHEET_NAME}!A2:J`);
  const keys = Object.keys(usersObj);
  if (!keys.length) return;
  const rows = keys.map(key => {
    const u = usersObj[key];
    return [
      key, u.password, u.role, u.name || '', u.mobile || '', u.email || '', u.school || '',
      u.permissions ? JSON.stringify(u.permissions) : '',
      u.mustChangePassword ? 'Yes' : '',
      u.demoWelcomeShown ? 'Yes' : ''
    ];
  });
  await sheetsApi.writeRange(`${sheetsApi.USER_SHEET_NAME}!A2`, rows);
}

const ALL_PERMISSION_KEYS = [
  'viewSummary', 'downloadCsv', 'unlockRecords', 'lockAll', 'viewActivityLog',
  'composeMessage', 'importStudents', 'importVerification', 'resetPasswords',
  'manageUsers', 'deleteStudent', 'hostelAccess'
];

function effectivePermissions(userRecord) {
  if (userRecord.role === 'demo') {
    if (userRecord.permissions) return userRecord.permissions;
    const all = {};
    ALL_PERMISSION_KEYS.forEach(k => { all[k] = true; });
    all.deleteStudent = false; // hard restriction, also enforced server-side below
    return all;
  }
  if (userRecord.role !== 'admin') {
    // Staff normally has no permissions, but hostelAccess can be individually
    // delegated to a staff account without promoting them to admin/demo.
    return { hostelAccess: !!(userRecord.permissions && userRecord.permissions.hostelAccess) };
  }
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

  let showDemoWelcome = false;
  if (users[key].role === 'demo' && !users[key].demoWelcomeShown) {
    showDemoWelcome = true;
    users[key].demoWelcomeShown = true;
    await writeAllUsers(users);
  }

  return {
    ok: true,
    userId: key,
    isAdmin: users[key].role === 'admin' || users[key].role === 'demo',
    isDemo: users[key].role === 'demo',
    name: displayName,
    permissions: effectivePermissions(users[key]),
    mustChangePassword: !!users[key].mustChangePassword,
    showDemoWelcome
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

  const usersForRoleCheck = await getAllUsers();
  const callerKey = userId ? String(userId) : '';
  if (usersForRoleCheck[callerKey] && usersForRoleCheck[callerKey].role === 'demo') {
    return { ok: false, error: 'Demo accounts cannot mark biometric verification status.' };
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
  const users = usersForRoleCheck;
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

  await sheetsApi.batchWriteRanges(
    updates.map(u => ({ range: `${sheetName}!${colToLetter(u.col)}${rowNum}`, values: [[u.val]] }))
  );

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

  const pendingUpdates = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const statusVal = String(row[col.status] || '').trim().toLowerCase();
    const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true' || statusVal === 'completed' || statusVal === 'verified');
    const isLocked = String(row[col.lock] || '').trim().toLowerCase() === 'yes';
    if (isDone && !isLocked) {
      pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.lock)}${r + 1}`, values: [['Yes']] });
    }
  }
  await sheetsApi.batchWriteRanges(pendingUpdates);
  const count = pendingUpdates.length;

  await logActivity(actor || 'Admin', 'Bulk Lock Verified', `Locked ${count} record(s)`);
  return { ok: true, count };
}

// ---------- User management ----------

function generateUserId(name, mobile) {
  const lettersOnly = String(name || '').replace(/[^a-zA-Z]/g, '').toLowerCase();
  let namePart;
  if (lettersOnly.length >= 5) namePart = lettersOnly.slice(0, 5);
  else if (lettersOnly.length > 0) namePart = lettersOnly;
  else namePart = 'user';
  const digitsOnly = String(mobile || '').replace(/[^0-9]/g, '');
  const mobilePart = digitsOnly.length >= 2 ? digitsOnly.slice(-2) : (digitsOnly || '00').padStart(2, '0');
  return namePart + mobilePart;
}

function generateRandomPassword() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const symbols = '!@#$%&*';
  let pwd = '';
  for (let i = 0; i < 8; i++) pwd += chars.charAt(Math.floor(Math.random() * chars.length));
  pwd += symbols.charAt(Math.floor(Math.random() * symbols.length));
  pwd += String(Math.floor(Math.random() * 10));
  return pwd;
}

async function checkUserIdAvailability(userId) {
  const key = String(userId || '').trim().toLowerCase();
  if (!key) return { ok: true, available: false, reason: 'User ID cannot be blank.' };
  const users = await getAllUsers();
  const taken = Object.keys(users).some(k => k.toLowerCase() === key);
  return { ok: true, available: !taken };
}

async function findContactConflicts(mobile, email, excludeUserId) {
  const users = await getAllUsers();
  const normMobile = String(mobile || '').replace(/[^0-9]/g, '');
  const normEmail = String(email || '').trim().toLowerCase();
  let mobileConflict = false, emailConflict = false;
  Object.keys(users).forEach(key => {
    if (excludeUserId && key === excludeUserId) return;
    const u = users[key];
    if (normMobile && String(u.mobile || '').replace(/[^0-9]/g, '') === normMobile) mobileConflict = true;
    if (normEmail && String(u.email || '').trim().toLowerCase() === normEmail) emailConflict = true;
  });
  return { mobileConflict, emailConflict };
}

async function checkContactAvailability(mobile, email, excludeUserId) {
  const conflicts = await findContactConflicts(mobile, email, excludeUserId);
  return { ok: true, mobileTaken: conflicts.mobileConflict, emailTaken: conflicts.emailConflict };
}

async function createUser(newUserId, newPassword, role, adminPassword, displayName, actor, permissions, autoGeneratePassword, mobile, email, school, autoGenerateUserId) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  role = (role === 'admin' || role === 'demo') ? role : 'staff';
  let name = (displayName && String(displayName).trim()) || '';

  const users = await getAllUsers();
  let key;

  if (autoGenerateUserId) {
    if (!name) return { ok: false, error: 'Name is required to auto-generate a User ID.' };
    const baseId = generateUserId(name, mobile);
    key = baseId;
    let suffix = 1;
    while (users.hasOwnProperty(key)) { key = baseId + suffix; suffix++; }
  } else {
    key = String(newUserId || '').trim();
    if (!key || !/^[a-zA-Z0-9_.-]+$/.test(key)) {
      return { ok: false, error: 'User ID must contain only letters, numbers, dots, dashes, or underscores.' };
    }
    if (users.hasOwnProperty(key)) return { ok: false, error: 'That User ID already exists.' };
  }

  let finalPassword;
  if (autoGeneratePassword) {
    finalPassword = generateRandomPassword();
  } else {
    if (!newPassword || newPassword.length < 4) return { ok: false, error: 'Password must be at least 4 characters.' };
    finalPassword = newPassword;
  }

  if (!name) name = key;

  const contactCheck = await findContactConflicts(mobile, email, null);
  if (contactCheck.mobileConflict) return { ok: false, error: 'That mobile number is already used by another account.' };
  if (contactCheck.emailConflict) return { ok: false, error: 'That email is already used by another account.' };

  const userRecord = {
    password: finalPassword, role, name,
    mobile: mobile ? String(mobile).trim() : '',
    email: email ? String(email).trim() : '',
    school: school ? String(school).trim() : ''
  };
  if (autoGeneratePassword) userRecord.mustChangePassword = true;
  if ((role === 'admin' || role === 'demo') && permissions && typeof permissions === 'object') {
    const cleanPerms = {};
    ALL_PERMISSION_KEYS.forEach(k => { cleanPerms[k] = !!permissions[k]; });
    if (role === 'demo') cleanPerms.deleteStudent = false; // hard restriction, also enforced in deleteStudent()
    userRecord.permissions = cleanPerms;
  } else if (role === 'staff' && permissions && typeof permissions === 'object') {
    userRecord.permissions = { hostelAccess: !!permissions.hostelAccess };
  }

  users[key] = userRecord;
  await writeAllUsers(users);
  await logActivity(actor || 'Admin', 'User Created', `${key} (${role})${autoGenerateUserId ? ' — auto-generated ID' : ''}${autoGeneratePassword ? ' — auto-generated password' : ''}`);
  return { ok: true, userId: key, generatedPassword: autoGeneratePassword ? finalPassword : null };
}

async function deleteUser(targetUserId, adminPassword, actor, actorUserId) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };
  const key = String(targetUserId || '').trim();
  const users = await getAllUsers();
  if (!key || !users.hasOwnProperty(key)) return { ok: false, error: 'Unknown account.' };

  if (actorUserId && users[actorUserId] && users[actorUserId].role === 'demo') {
    return { ok: false, error: 'Demo accounts cannot delete user accounts.' };
  }

  delete users[key];
  await writeAllUsers(users);
  await logActivity(actor || 'Admin', 'User Deleted', key);
  return { ok: true };
}

async function updateUserDetails(targetUserId, newName, newRole, permissions, adminPassword, actor, mobile, email, school) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };
  const key = String(targetUserId || '').trim();
  const users = await getAllUsers();
  if (!key || !users.hasOwnProperty(key)) return { ok: false, error: 'Unknown account.' };

  const role = (newRole === 'admin' || newRole === 'demo') ? newRole : 'staff';
  if (users[key].role === 'admin' && role === 'staff') {
    const adminCount = Object.keys(users).filter(k => users[k].role === 'admin').length;
    if (adminCount <= 1) return { ok: false, error: 'Cannot demote the only remaining admin account to staff.' };
  }

  const contactCheck = await findContactConflicts(mobile, email, key);
  if (contactCheck.mobileConflict) return { ok: false, error: 'That mobile number is already used by another account.' };
  if (contactCheck.emailConflict) return { ok: false, error: 'That email is already used by another account.' };

  const name = (newName && String(newName).trim()) || key;
  users[key].name = name;
  users[key].role = role;
  users[key].mobile = mobile ? String(mobile).trim() : '';
  users[key].email = email ? String(email).trim() : '';
  users[key].school = school ? String(school).trim() : '';

  if (role === 'admin' || role === 'demo') {
    const cleanPerms = {};
    ALL_PERMISSION_KEYS.forEach(k => { cleanPerms[k] = !!(permissions && permissions[k]); });
    if (role === 'demo') cleanPerms.deleteStudent = false; // hard restriction, also enforced in deleteStudent()
    users[key].permissions = cleanPerms;
  } else if (role === 'staff') {
    users[key].permissions = { hostelAccess: !!(permissions && permissions.hostelAccess) };
  } else {
    delete users[key].permissions;
  }

  await writeAllUsers(users);
  await logActivity(actor || 'Admin', 'User Updated', `${key} (${role})`);
  return { ok: true };
}

async function getUserList() {
  const users = await getAllUsers();
  const list = Object.keys(users).map(k => ({
    userId: k,
    role: users[k].role,
    name: users[k].name || capitalizeFirst(k),
    mobile: users[k].mobile || '',
    email: users[k].email || '',
    school: users[k].school || '',
    permissions: effectivePermissions(users[k])
  }));
  return { ok: true, users: list };
}

async function changeOwnPassword(userId, oldPassword, newPassword) {
  const key = String(userId || '').trim();
  const users = await getAllUsers();
  if (!key || !users.hasOwnProperty(key)) return { ok: false, error: 'Invalid account.' };
  if (users[key].password !== oldPassword) return { ok: false, error: 'Current password is incorrect.' };
  if (!newPassword || newPassword.length < 4) return { ok: false, error: 'New password must be at least 4 characters.' };
  users[key].password = newPassword;
  delete users[key].mustChangePassword;
  await writeAllUsers(users);
  return { ok: true };
}

async function adminResetPassword(targetUserId, newPassword, adminPassword) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };
  const key = String(targetUserId || '').trim();
  const users = await getAllUsers();
  if (!key || !users.hasOwnProperty(key)) return { ok: false, error: 'Unknown account.' };
  if (!newPassword || newPassword.length < 4) return { ok: false, error: 'New password must be at least 4 characters.' };
  users[key].password = newPassword;
  await writeAllUsers(users);
  return { ok: true };
}

async function changeAdminPassword(currentPassword, newPassword, actor) {
  const existing = await getAdminPassword();
  if (currentPassword !== existing) return { ok: false, error: 'Current admin password is incorrect.' };
  if (!newPassword || newPassword.length < 6) return { ok: false, error: 'New admin password must be at least 6 characters.' };
  await setAdminPassword(newPassword);
  await logActivity(actor || 'Admin', 'Admin Password Changed', 'The shared admin authorization password was updated.');
  return { ok: true };
}

async function deleteStudent(rowId, adminPassword, actor, actorUserId) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  if (actorUserId) {
    const users = await getAllUsers();
    if (users[actorUserId] && users[actorUserId].role === 'demo') {
      return { ok: false, error: 'Demo accounts cannot delete student records.' };
    }
  }

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const rowNum = parseInt(String(rowId).replace('row', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid student id.' };

  const row = data[rowNum - 1] || [];
  const name = col.name > -1 ? String(row[col.name] || '') : '';
  const appNo = col.appNo > -1 ? String(row[col.appNo] || '') : '';
  const regNo = col.regNo > -1 ? String(row[col.regNo] || '') : '';

  await sheetsApi.deleteRow(sheetName, rowNum);
  await logActivity(actor || 'Admin', 'Student Deleted', `${name} (App No: ${appNo}, Reg No: ${regNo})`);
  return { ok: true };
}

async function getDistinctSiteCodes() {
  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  if (!data.length) return { ok: true, siteCodes: [] };
  const headers = data[0];
  const col = detectColumns(headers);
  if (col.siteCode === -1) return { ok: true, siteCodes: [] };
  const codes = new Set();
  for (let r = 1; r < data.length; r++) {
    const val = String(data[r][col.siteCode] || '').trim();
    if (val) codes.add(val);
  }
  return { ok: true, siteCodes: Array.from(codes).sort() };
}

// ---------- Imports ----------

function detectImportColumns(headers) {
  return detectColumns(headers);
}

async function validateImportRows(uploadedHeaders, uploadedRows) {
  const col = detectImportColumns(uploadedHeaders);
  let missingBothCount = 0, missingNameCount = 0;
  uploadedRows.forEach(row => {
    const appNo = col.appNo > -1 ? String(row[col.appNo] || '').trim() : '';
    const regNo = col.regNo > -1 ? String(row[col.regNo] || '').trim() : '';
    const name = col.name > -1 ? String(row[col.name] || '').trim() : '';
    if (!appNo && !regNo) missingBothCount++;
    if (!name) missingNameCount++;
  });
  return {
    ok: true,
    detectedName: col.name > -1 ? uploadedHeaders[col.name] : null,
    detectedAppNo: col.appNo > -1 ? uploadedHeaders[col.appNo] : null,
    detectedRegNo: col.regNo > -1 ? uploadedHeaders[col.regNo] : null,
    missingBothCount, missingNameCount
  };
}

async function importNewStudents(uploadedHeaders, uploadedRows, adminPassword, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  if (!data.length) return { ok: false, error: 'The master sheet is empty — cannot detect its column layout.' };
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const uploadCol = detectImportColumns(uploadedHeaders);

  const existingAppNos = new Set();
  const existingRegNos = new Set();
  for (let r = 1; r < data.length; r++) {
    if (col.appNo > -1) { const v = String(data[r][col.appNo] || '').trim().toLowerCase(); if (v) existingAppNos.add(v); }
    if (col.regNo > -1) { const v = String(data[r][col.regNo] || '').trim().toLowerCase(); if (v) existingRegNos.add(v); }
  }

  const newRows = [];
  const skipped = [];
  uploadedRows.forEach(uRow => {
    const name = uploadCol.name > -1 ? String(uRow[uploadCol.name] || '').trim() : '';
    const appNo = uploadCol.appNo > -1 ? String(uRow[uploadCol.appNo] || '').trim() : '';
    const regNo = uploadCol.regNo > -1 ? String(uRow[uploadCol.regNo] || '').trim() : '';
    const machineCode = uploadCol.machineCode > -1 ? String(uRow[uploadCol.machineCode] || '').trim() : '';
    const siteCode = uploadCol.siteCode > -1 ? String(uRow[uploadCol.siteCode] || '').trim() : '';
    const studentType = uploadCol.studentType > -1 ? String(uRow[uploadCol.studentType] || '').trim() : '';

    if ((regNo && existingRegNos.has(regNo.toLowerCase())) || (appNo && existingAppNos.has(appNo.toLowerCase()))) {
      skipped.push({ name, appNo, regNo, reason: 'Duplicate App No or Reg No' });
      return;
    }

    const newRow = new Array(headers.length).fill('');
    if (col.name > -1) newRow[col.name] = name;
    if (col.appNo > -1) newRow[col.appNo] = appNo;
    if (col.regNo > -1) newRow[col.regNo] = regNo;
    if (col.machineCode > -1) newRow[col.machineCode] = machineCode;
    if (col.siteCode > -1) newRow[col.siteCode] = siteCode;
    if (col.studentType > -1) newRow[col.studentType] = studentType;
    newRow[col.status] = 'Not Done';
    newRow[col.lock] = '';
    newRows.push(newRow);

    if (regNo) existingRegNos.add(regNo.toLowerCase());
    if (appNo) existingAppNos.add(appNo.toLowerCase());
  });

  if (newRows.length) await sheetsApi.appendRows(sheetName, newRows);

  await logActivity(actor || 'Admin', 'Imported New Students', `${newRows.length} added, ${skipped.length} skipped as duplicates`);
  return { ok: true, added: newRows.length, skippedRows: skipped };
}

async function importVerificationUpdates(uploadedHeaders, uploadedRows, adminPassword, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const uploadCol = detectImportColumns(uploadedHeaders);
  const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19);

  let updated = 0, alreadyDone = 0;
  const notFoundRows = [];
  const pendingUpdates = [];

  for (const uRow of uploadedRows) {
    const name = uploadCol.name > -1 ? String(uRow[uploadCol.name] || '').trim() : '';
    const appNo = uploadCol.appNo > -1 ? String(uRow[uploadCol.appNo] || '').trim() : '';
    const regNo = uploadCol.regNo > -1 ? String(uRow[uploadCol.regNo] || '').trim() : '';

    let matchedRowNum = -1;
    for (let r = 1; r < data.length; r++) {
      const rRegNo = col.regNo > -1 ? String(data[r][col.regNo] || '').trim().toLowerCase() : '';
      const rAppNo = col.appNo > -1 ? String(data[r][col.appNo] || '').trim().toLowerCase() : '';
      if (regNo && rRegNo === regNo.toLowerCase()) { matchedRowNum = r + 1; break; }
      if (appNo && rAppNo === appNo.toLowerCase()) { matchedRowNum = r + 1; break; }
    }

    if (matchedRowNum === -1) {
      notFoundRows.push({ name, appNo, regNo });
      continue;
    }

    const rowData = data[matchedRowNum - 1];
    const statusVal = String(rowData[col.status] || '').trim().toLowerCase();
    const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true' || statusVal === 'completed' || statusVal === 'verified');
    if (isDone) { alreadyDone++; continue; }

    pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.status)}${matchedRowNum}`, values: [['Done']] });
    pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.lock)}${matchedRowNum}`, values: [['Yes']] });
    pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.verifiedBy)}${matchedRowNum}`, values: [['Verification Import']] });
    pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.verifiedAt)}${matchedRowNum}`, values: [[timestamp]] });
    if (!String(rowData[col.firstVerifiedBy] || '').trim()) {
      pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.firstVerifiedBy)}${matchedRowNum}`, values: [['Verification Import']] });
      pendingUpdates.push({ range: `${sheetName}!${colToLetter(col.firstVerifiedAt)}${matchedRowNum}`, values: [[timestamp]] });
    }
    updated++;
  }

  await sheetsApi.batchWriteRanges(pendingUpdates);

  await logActivity(actor || 'Admin', 'Imported Verification Updates', `${updated} updated, ${alreadyDone} already done, ${notFoundRows.length} not found`);
  return { ok: true, updated, alreadyDone, notFoundRows };
}

// ---------- Generic Settings storage (JSON blobs, e.g. announcements/stats) ----------

async function getSetting(key, defaultValue) {
  await sheetsApi.ensureSheet(SETTINGS_SHEET_NAME, ['Key', 'Value']);
  const rows = await sheetsApi.readRange(`${SETTINGS_SHEET_NAME}!A2:B`);
  const row = rows.find(r => r[0] === key);
  if (!row || !row[1]) return defaultValue;
  try { return JSON.parse(row[1]); } catch (e) { return defaultValue; }
}

async function setSetting(key, value) {
  await sheetsApi.ensureSheet(SETTINGS_SHEET_NAME, ['Key', 'Value']);
  const rows = await sheetsApi.readRange(`${SETTINGS_SHEET_NAME}!A2:B`);
  const idx = rows.findIndex(r => r[0] === key);
  const jsonVal = JSON.stringify(value);
  if (idx === -1) {
    await sheetsApi.appendRows(SETTINGS_SHEET_NAME, [[key, jsonVal]]);
  } else {
    await sheetsApi.writeRange(`${SETTINGS_SHEET_NAME}!A${idx + 2}:B${idx + 2}`, [[key, jsonVal]]);
  }
}

// ---------- Announcements ----------

async function getAnnouncements() {
  const list = await getSetting('ANNOUNCEMENTS', []);
  return { ok: true, announcements: list };
}

async function publishAnnouncement(message, adminPassword, publishedBy, style) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };
  const list = await getSetting('ANNOUNCEMENTS', []);
  const now = new Date();
  const entry = {
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 7),
    text: message,
    publishedBy: publishedBy || 'Admin',
    publishedAt: now.toISOString().replace('T', ' ').slice(0, 19),
    publishedAtRaw: now.toISOString(),
    style: style || 'gold'
  };
  list.push(entry);
  await setSetting('ANNOUNCEMENTS', list);
  await logActivity(publishedBy || 'Admin', 'Message Published', message.slice(0, 80));
  return { ok: true };
}

async function updateAnnouncement(id, message, adminPassword, style, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };
  const list = await getSetting('ANNOUNCEMENTS', []);
  const idx = list.findIndex(a => a.id === id);
  if (idx === -1) return { ok: false, error: 'Message not found.' };
  list[idx].text = message;
  list[idx].style = style || list[idx].style;
  await setSetting('ANNOUNCEMENTS', list);
  await logActivity(actor || 'Admin', 'Message Updated', message.slice(0, 80));
  return { ok: true };
}

async function deleteAnnouncement(id, adminPassword, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };
  const list = await getSetting('ANNOUNCEMENTS', []);
  const filtered = list.filter(a => a.id !== id);
  await setSetting('ANNOUNCEMENTS', filtered);
  await logActivity(actor || 'Admin', 'Message Deleted', id);
  return { ok: true };
}

// ---------- Help Assistant question stats ----------

async function logHelpQuestion(questionText) {
  const key = String(questionText || '').trim();
  if (!key) return { ok: true };
  const counts = await getSetting('HELP_QUESTION_COUNTS', {});
  counts[key] = (counts[key] || 0) + 1;
  await setSetting('HELP_QUESTION_COUNTS', counts);
  return { ok: true };
}

async function getHelpQuestionStats() {
  const counts = await getSetting('HELP_QUESTION_COUNTS', {});
  const list = Object.keys(counts).map(q => ({ question: q, count: counts[q] }));
  list.sort((a, b) => b.count - a.count);
  return { ok: true, stats: list };
}

// ---------- Session logging ----------

async function logSessionIp(actor, ipAddress, failureReason) {
  if (ipAddress) {
    await logActivity(actor || 'unknown', 'Session IP', ipAddress);
  } else {
    await logActivity(actor || 'unknown', 'Session IP Failed', failureReason || 'unknown reason');
  }
  return { ok: true };
}

async function logSessionEnd(actor, durationText) {
  await logActivity(actor || 'unknown', 'Session Ended', 'Duration: ' + (durationText || 'unknown'));
  return { ok: true };
}

// ---------- CSV export ----------

function csvEscape(val) {
  const str = String(val ?? '');
  return (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1)
    ? '"' + str.replace(/"/g, '""') + '"'
    : str;
}

async function exportRosterAsCsv(statusFilter, siteFilter) {
  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  const col = detectColumns(headers);
  const wantStatus = statusFilter && statusFilter !== 'all' ? statusFilter : null;
  const wantSite = siteFilter && siteFilter !== 'all' ? String(siteFilter).trim().toLowerCase() : null;

  let serialColIndex = -1;
  for (let h = 0; h < headers.length; h++) {
    const normHeader = normalize(headers[h]);
    if (['s', 'slno', 'srno', 'sno', 'serialno', 'serialnumber'].includes(normHeader)) { serialColIndex = h; break; }
  }

  const filteredRows = [headers];
  let serialCounter = 1;
  for (let r = 1; r < data.length; r++) {
    let row = data[r];
    if (wantStatus) {
      const statusVal = col.status > -1 ? String(row[col.status] || '').trim().toLowerCase() : '';
      const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true' || statusVal === 'completed' || statusVal === 'verified');
      if (wantStatus === 'done' && !isDone) continue;
      if (wantStatus === 'pending' && isDone) continue;
    }
    if (wantSite) {
      const siteVal = col.siteCode > -1 ? String(row[col.siteCode] || '').trim().toLowerCase() : '';
      if (siteVal !== wantSite) continue;
    }
    if (serialColIndex > -1) {
      row = row.slice();
      row[serialColIndex] = serialCounter;
    }
    serialCounter++;
    filteredRows.push(row);
  }

  const csvText = filteredRows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  return { ok: true, csv: csvText, rowCount: filteredRows.length - 1, filename: `roster_export_${Date.now()}.csv` };
}

// ---------- Import stats ----------

async function getLastImportInfo() {
  const data = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:D`);
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][2] === 'Imported New Students') {
      const match = String(data[i][3] || '').match(/^(\d+)\s+added/);
      const count = match ? parseInt(match[1], 10) : 0;
      return { ok: true, count, timestamp: data[i][0] || null };
    }
  }
  return { ok: true, count: 0, timestamp: null };
}

// The Activity Log stores timestamps in UTC (matching every other timestamp
// in this app); this reformats one as a readable India-time string, for
// the hostel "fresh upload" display specifically.
function formatUtcTimestampAsIndiaTime(utcTimestampStr) {
  if (!utcTimestampStr) return null;
  const utcDate = new Date(utcTimestampStr.replace(' ', 'T') + 'Z');
  if (isNaN(utcDate.getTime())) return utcTimestampStr;
  return new Intl.DateTimeFormat('en-IN', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: 'short', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hour12: true
  }).format(utcDate);
}

async function getLastHostelImportInfo() {
  const data = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:D`);
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][2] === 'Imported New Hostel Data') {
      const match = String(data[i][3] || '').match(/^(\d+)\s+added/);
      const count = match ? parseInt(match[1], 10) : 0;
      return { ok: true, count, timestamp: formatUtcTimestampAsIndiaTime(data[i][0] || null) };
    }
  }
  return { ok: true, count: 0, timestamp: null };
}

async function getTodayImportCount() {
  const info = await getLastImportInfo();
  const todayStr = new Date().toISOString().slice(0, 10);
  if (info.timestamp && String(info.timestamp).indexOf(todayStr) === 0) {
    return { ok: true, count: info.count };
  }
  return { ok: true, count: 0 };
}

async function getLastImportTimestamp() {
  const info = await getLastImportInfo();
  return { ok: true, timestamp: info.timestamp };
}

// ---------- Hostel Data: Face Capture verification system ----------

const HOSTEL_SHEET_NAME = 'Hostel Data';
const HOSTEL_BASE_HEADERS = [
  'S#', 'RegistrationNo', 'ApplicationNo', 'MachineCode', 'SiteCode', 'StudentName',
  'Course', 'AdmissionType', 'Gender', 'HostelName', 'RoomNo', 'FoodCoupon'
];

function detectHostelColumns(headers) {
  const col = {};
  headers.forEach((h, i) => { col[normalize(h)] = i; });
  let statusCol = -1, lockCol = -1, verifiedAtCol = -1;
  headers.forEach((h, i) => {
    const n = normalize(h);
    if (n.indexOf('facecapture') !== -1 || n === 'status') statusCol = i;
    if (n === 'locked') lockCol = i;
    if (n.indexOf('verifiedat') !== -1) verifiedAtCol = i;
  });
  col.status = statusCol;
  col.lock = lockCol;
  col.verifiedAt = verifiedAtCol;
  return col;
}

async function ensureHostelExtraColumns(headers, col) {
  const additions = [];
  if (col.status === -1) { col.status = headers.length + additions.length; additions.push('Face Capture Status'); }
  if (col.lock === -1) { col.lock = headers.length + additions.length; additions.push('Locked'); }
  if (col.verifiedAt === -1) { col.verifiedAt = headers.length + additions.length; additions.push('Verified At'); }
  if (additions.length) {
    await sheetsApi.writeRange(`${HOSTEL_SHEET_NAME}!${colToLetter(headers.length)}1`, [additions]);
  }
  return col;
}

async function getHostelData() {
  await sheetsApi.ensureSheet(HOSTEL_SHEET_NAME, HOSTEL_BASE_HEADERS);
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  if (!data.length) return { ok: true, students: [] };

  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const students = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const studentName = row[col['studentname']] || '';
    const regNo = row[col['registrationno']] || '';
    const appNo = row[col['applicationno']] || '';
    if (!studentName && !regNo && !appNo) continue;

    const statusVal = String(row[col.status] || '').trim().toLowerCase();
    const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true');
    const isLocked = String(row[col.lock] || '').trim().toLowerCase() === 'yes';

    students.push({
      id: 'hrow' + (r + 1),
      registrationNo: String(regNo || ''),
      applicationNo: String(appNo || ''),
      machineCode: String(row[col['machinecode']] || ''),
      siteCode: String(row[col['sitecode']] || ''),
      studentName: String(studentName || ''),
      course: String(row[col['course']] || ''),
      admissionType: String(row[col['admissiontype']] || ''),
      gender: String(row[col['gender']] || ''),
      hostelName: String(row[col['hostelname']] || ''),
      roomNo: String(row[col['roomno']] || ''),
      foodCoupon: String(row[col['foodcoupon']] || ''),
      status: isDone ? 'done' : 'pending',
      locked: isLocked,
      verifiedAt: String(row[col.verifiedAt] || '')
    });
  }
  return { ok: true, students };
}

async function updateHostelStatus(rowId, status, userId) {
  if (status !== 'done' && status !== 'pending') {
    return { ok: false, error: "status must be 'done' or 'pending'" };
  }
  const users = await getAllUsers();
  const callerKey = userId ? String(userId) : '';
  if (users[callerKey] && users[callerKey].role === 'demo') {
    return { ok: false, error: 'Demo accounts cannot mark face capture status.' };
  }

  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const rowNum = parseInt(String(rowId).replace('hrow', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid record id.' };

  const row = data[rowNum - 1] || [];
  if (String(row[col.lock] || '').trim().toLowerCase() === 'yes') {
    return { ok: false, error: 'This record is locked. An admin needs to unlock it before it can be changed.' };
  }

  const value = status === 'done' ? 'Done' : 'Not Done';
  const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19);
  await sheetsApi.batchWriteRanges([
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.status)}${rowNum}`, values: [[value]] },
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${rowNum}`, values: [['Yes']] },
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.verifiedAt)}${rowNum}`, values: [[timestamp]] }
  ]);

  const name = col['studentname'] > -1 ? String(row[col['studentname']] || '') : '';
  const actorName = (users[callerKey] && users[callerKey].name) || callerKey || 'unknown';
  await logActivity(actorName, status === 'done' ? 'Hostel Face Capture Verified' : 'Hostel Marked Pending', `${name} (row ${rowNum})`);
  return { ok: true };
}

async function adminUnlockHostel(rowId, password, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (password !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const rowNum = parseInt(String(rowId).replace('hrow', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid record id.' };

  await sheetsApi.writeRange(`${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${rowNum}`, [['No']]);
  await logActivity(actor || 'Admin', 'Hostel Record Unlocked', `row ${rowNum}`);
  return { ok: true };
}

async function exportHostelAsCsv(statusFilter) {
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  if (!data.length) return { ok: false, error: 'Hostel data is empty.' };
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const wantStatus = statusFilter && statusFilter !== 'all' ? statusFilter : null;
  const rows = [headers];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (wantStatus) {
      const statusVal = String(row[col.status] || '').trim().toLowerCase();
      const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true');
      if (wantStatus === 'done' && !isDone) continue;
      if (wantStatus === 'pending' && isDone) continue;
    }
    rows.push(row);
  }
  const csvText = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  return { ok: true, csv: csvText, rowCount: rows.length - 1, filename: `hostel_facecapture_${Date.now()}.csv` };
}

async function exportHostelVerifiedTodayAsCsv() {
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  if (!data.length) return { ok: false, error: 'Hostel data is empty.' };
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const todayStr = new Date().toISOString().slice(0, 10);
  const rows = [headers];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const verifiedAt = String(row[col.verifiedAt] || '');
    if (verifiedAt.indexOf(todayStr) === 0) rows.push(row);
  }
  const csvText = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  return { ok: true, csv: csvText, rowCount: rows.length - 1, filename: `hostel_verified_today_${Date.now()}.csv` };
}

async function importNewHostelData(uploadedHeaders, uploadedRows, adminPassword, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  await sheetsApi.ensureSheet(HOSTEL_SHEET_NAME, HOSTEL_BASE_HEADERS);
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const uCol = {};
  uploadedHeaders.forEach((h, i) => { uCol[normalize(h)] = i; });

  const existingAppNos = new Set(), existingRegNos = new Set();
  for (let r = 1; r < data.length; r++) {
    const a = String(data[r][col['applicationno']] || '').trim().toLowerCase();
    const g = String(data[r][col['registrationno']] || '').trim().toLowerCase();
    if (a) existingAppNos.add(a);
    if (g) existingRegNos.add(g);
  }

  const fieldMap = ['registrationno', 'applicationno', 'machinecode', 'sitecode', 'studentname', 'course', 'admissiontype', 'gender', 'hostelname', 'roomno', 'foodcoupon'];
  const newRows = [];
  const skipped = [];
  uploadedRows.forEach(uRow => {
    const appNo = uCol['applicationno'] > -1 ? String(uRow[uCol['applicationno']] || '').trim() : '';
    const regNo = uCol['registrationno'] > -1 ? String(uRow[uCol['registrationno']] || '').trim() : '';
    const name = uCol['studentname'] > -1 ? String(uRow[uCol['studentname']] || '').trim() : '';

    if ((regNo && existingRegNos.has(regNo.toLowerCase())) || (appNo && existingAppNos.has(appNo.toLowerCase()))) {
      skipped.push({ name, appNo, regNo, reason: 'Duplicate App No or Reg No' });
      return;
    }

    const newRow = new Array(headers.length).fill('');
    fieldMap.forEach(field => {
      if (col[field] > -1 && uCol[field] > -1) newRow[col[field]] = String(uRow[uCol[field]] || '');
    });
    newRow[col.status] = 'Not Done';
    newRow[col.lock] = '';
    newRows.push(newRow);

    if (regNo) existingRegNos.add(regNo.toLowerCase());
    if (appNo) existingAppNos.add(appNo.toLowerCase());
  });

  if (newRows.length) await sheetsApi.appendRows(HOSTEL_SHEET_NAME, newRows);
  await logActivity(actor || 'Admin', 'Imported New Hostel Data', `${newRows.length} added, ${skipped.length} skipped as duplicates`);
  return { ok: true, added: newRows.length, skippedRows: skipped };
}

async function importHostelVerificationUpdates(uploadedHeaders, uploadedRows, adminPassword, actor) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const uCol = {};
  uploadedHeaders.forEach((h, i) => { uCol[normalize(h)] = i; });

  let updated = 0, alreadyDone = 0;
  const notFoundRows = [];
  const pendingUpdates = [];

  uploadedRows.forEach(uRow => {
    const appNo = uCol['applicationno'] > -1 ? String(uRow[uCol['applicationno']] || '').trim() : '';
    const regNo = uCol['registrationno'] > -1 ? String(uRow[uCol['registrationno']] || '').trim() : '';
    const name = uCol['studentname'] > -1 ? String(uRow[uCol['studentname']] || '').trim() : '';

    let matchedRowNum = -1;
    for (let r = 1; r < data.length; r++) {
      const rReg = String(data[r][col['registrationno']] || '').trim().toLowerCase();
      const rApp = String(data[r][col['applicationno']] || '').trim().toLowerCase();
      if (regNo && rReg === regNo.toLowerCase()) { matchedRowNum = r + 1; break; }
      if (appNo && rApp === appNo.toLowerCase()) { matchedRowNum = r + 1; break; }
    }
    if (matchedRowNum === -1) { notFoundRows.push({ name, appNo, regNo }); return; }

    const rowData = data[matchedRowNum - 1];
    const statusVal = String(rowData[col.status] || '').trim().toLowerCase();
    const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true');
    if (isDone) { alreadyDone++; return; }

    pendingUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.status)}${matchedRowNum}`, values: [['Done']] });
    pendingUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${matchedRowNum}`, values: [['Yes']] });
    updated++;
  });

  await sheetsApi.batchWriteRanges(pendingUpdates);
  await logActivity(actor || 'Admin', 'Imported Hostel Face Capture Updates', `${updated} updated, ${alreadyDone} already done, ${notFoundRows.length} not found`);
  return { ok: true, updated, alreadyDone, notFoundRows };
}

module.exports = {
  normalize, detectColumns, colToLetter, ensureExtraColumns,
  getAllUsers, writeAllUsers, effectivePermissions, capitalizeFirst, ALL_PERMISSION_KEYS,
  getAdminPassword, setAdminPassword,
  logActivity, getActivityLog,
  checkLogin, getStudents, updateStatus, adminUnlock, adminLockAllDone,
  checkUserIdAvailability, checkContactAvailability,
  createUser, deleteUser, updateUserDetails, getUserList,
  changeOwnPassword, adminResetPassword, changeAdminPassword,
  deleteStudent, getDistinctSiteCodes,
  validateImportRows, importNewStudents, importVerificationUpdates,
  getAnnouncements, publishAnnouncement, updateAnnouncement, deleteAnnouncement,
  logHelpQuestion, getHelpQuestionStats,
  logSessionIp, logSessionEnd,
  exportRosterAsCsv,
  getLastImportInfo, getTodayImportCount, getLastImportTimestamp,
  getHostelData, updateHostelStatus, adminUnlockHostel, exportHostelAsCsv, exportHostelVerifiedTodayAsCsv,
  importNewHostelData, importHostelVerificationUpdates, getLastHostelImportInfo
};
