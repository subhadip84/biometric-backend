// core.js — Phase 1: the essential day-to-day functions, ported from Code.gs.
// (User management, imports, and announcements come in Phase 2/3.)

const sheetsApi = require('./sheets');
const { AsyncLocalStorage } = require('async_hooks');

// Request-scoped context (currently just the client IP) so logActivity()
// can record it without every function in the call chain needing to pass
// it through explicitly. AsyncLocalStorage correctly scopes this per
// request even when multiple requests are in flight concurrently - a
// plain module-level variable would risk one request's IP leaking into
// another's log entry.
const requestContext = new AsyncLocalStorage();

const HEADER_CANDIDATES = {
  name: ['studentname', 'name'],
  appNo: ['applicationno', 'appno', 'application', 'rollnumber', 'rollno', 'roll'],
  regNo: ['registrationnumber', 'regno', 'registration'],
  machineCode: ['machinecode', 'machine'],
  siteCode: ['sitecode'],
  studentType: ['studenttype', 'type']
};

// Generates the current moment as an India-time "YYYY-MM-DD HH:MM:SS" string,
// prefixed with ' so Google Sheets stores it as literal text (preventing
// auto-conversion to a native date, which silently reformats on read-back).
// Used everywhere a timestamp gets written to a sheet, so the database
// itself always shows IST directly - no UTC-to-IST conversion needed when
// displaying or reading these values back.
function getISTTimestampForStorage() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  }).formatToParts(now);
  const get = type => parts.find(p => p.type === type).value;
  return "'" + `${get('year')}-${get('month')}-${get('day')} ${get('hour')}:${get('minute')}:${get('second')}`;
}

// Today's date in India time as "YYYY-MM-DD", for comparing against the
// IST-based timestamps now stored directly in the database.
function getISTTodayDateString() {
  const now = new Date();
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(now);
  const get = type => parts.find(p => p.type === type).value;
  return `${get('year')}-${get('month')}-${get('day')}`;
}

function getISTLastDayDateString() {
  const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(yesterday);
  const get = type => parts.find(p => p.type === type).value;
  return `${get('year')}-${get('month')}-${get('day')}`;
}

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

  let notesCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (normalize(headers[i]) === 'notes') { notesCol = i; break; }
  }
  col.notes = notesCol;

  let photoUrlCol = -1;
  for (let i = 0; i < headers.length; i++) {
    const norm = normalize(headers[i]);
    if (norm === 'photourl' || norm === 'photo') { photoUrlCol = i; break; }
  }
  col.photoUrl = photoUrlCol;

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
  if (col.notes === -1) { col.notes = headers.length + additions.length; additions.push('Notes'); changed = true; }
  if (col.photoUrl === -1) { col.photoUrl = headers.length + additions.length; additions.push('Photo URL'); changed = true; }

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
  'manageUsers', 'deleteStudent', 'hostelAccess', 'allowUndo'
];

function effectivePermissions(userRecord) {
  if (userRecord.role === 'demo') {
    if (userRecord.permissions) return userRecord.permissions;
    const all = {};
    ALL_PERMISSION_KEYS.forEach(k => { all[k] = true; });
    all.deleteStudent = false; // hard restriction, also enforced server-side below
    return all;
  }
  if (userRecord.role === 'viewer') {
    // Read-only: can see the dashboard, activity log, and export data, but
    // has zero write permissions - no verify/unverify, no user management,
    // no imports, no locking, no messaging. Enforced here and, for the one
    // action that matters most (marking verification status), also
    // enforced directly in updateStatus/updateHostelStatus server-side.
    return {
      viewSummary: true,
      viewActivityLog: true,
      downloadCsv: true
    };
  }
  if (userRecord.role !== 'admin') {
    // Staff normally has no permissions, but hostelAccess and allowUndo can
    // each be individually delegated to a staff account without promoting
    // them to admin/demo.
    return {
      hostelAccess: !!(userRecord.permissions && userRecord.permissions.hostelAccess),
      allowUndo: !!(userRecord.permissions && userRecord.permissions.allowUndo)
    };
  }
  // Admin always gets every permission, regardless of what's stored -
  // a stored permissions object missing a newer key (like hostelAccess)
  // should never accidentally strip access from an admin account.
  const all = {};
  ALL_PERMISSION_KEYS.forEach(k => { all[k] = true; });
  return all;
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

// The Activity Log sheet already exists with 4 columns from all prior
// logging - ensureSheet only creates a sheet if missing entirely, it won't
// add a column to one that already exists. This lazily adds the IP header
// once (checked via a flag, not on every call, to avoid extra reads).
let activityLogIpColumnEnsured = false;
async function ensureActivityLogIpColumn() {
  if (activityLogIpColumnEnsured) return;
  try {
    const headerCell = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!E1`);
    if (!headerCell.length || headerCell[0][0] !== 'IP') {
      await sheetsApi.writeRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!E1`, [['IP']]);
    }
  } catch (e) { /* best-effort - don't block logging if this fails */ }
  activityLogIpColumnEnsured = true;
}

async function logActivity(actor, action, details) {
  try {
    await sheetsApi.ensureSheet(sheetsApi.ACTIVITY_LOG_SHEET_NAME, ['Timestamp', 'Actor', 'Action', 'Details', 'IP']);
    await ensureActivityLogIpColumn();
    const timestamp = getISTTimestampForStorage();
    const store = requestContext.getStore();
    const ip = (store && store.ip) || '';
    await sheetsApi.appendRows(sheetsApi.ACTIVITY_LOG_SHEET_NAME, [[timestamp, actor || 'unknown', action, details || '', ip]]);
  } catch (e) { /* never let logging break the main action */ }
}

async function getActivityLog(limit) {
  await sheetsApi.ensureSheet(sheetsApi.ACTIVITY_LOG_SHEET_NAME, ['Timestamp', 'Actor', 'Action', 'Details', 'IP']);
  const rows = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
  const maxRows = Math.min(limit || 200, rows.length);
  const recent = rows.slice(-maxRows).reverse();
  return {
    ok: true,
    entries: recent.map(r => ({ timestamp: r[0] || '', actor: r[1] || '', action: r[2] || '', details: r[3] || '', ip: r[4] || '' }))
  };
}

// ---------- Unusual activity detection ----------
// Deliberately rule-based rather than AI-judged: statistical thresholds are
// more reliable and predictable for this than an LLM's subjective read of
// "unusual", and false positives here would just be noise for an admin to
// scan past - simple, explainable rules keep that noise low.

const BURST_THRESHOLD_COUNT = 15;      // this many verifications...
const BURST_THRESHOLD_MINUTES = 10;    // ...within this many minutes = a burst worth a glance
const UNUSUAL_HOUR_START = 0;          // midnight
const UNUSUAL_HOUR_END = 5;            // 5 AM - outside typical working hours

async function getUnusualActivityFlags() {
  await sheetsApi.ensureSheet(sheetsApi.ACTIVITY_LOG_SHEET_NAME, ['Timestamp', 'Actor', 'Action', 'Details']);
  const rows = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
  // Only look at the last 7 days' worth of entries, not the entire history
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const recent = rows.filter(r => {
    const t = parseTimestampLoose(r[0]);
    return t && t.getTime() >= cutoff;
  });

  const flags = [];

  // --- Burst detection: same actor, many verification-type actions, tight time window ---
  const verifyActions = ['Verified', 'Hostel Face Capture Verified'];
  const byActor = {};
  recent.forEach(r => {
    const action = String(r[2] || '');
    if (!verifyActions.some(a => action.indexOf(a) !== -1)) return;
    const actor = String(r[1] || 'unknown');
    const t = parseTimestampLoose(r[0]);
    if (!t) return;
    if (!byActor[actor]) byActor[actor] = [];
    byActor[actor].push(t.getTime());
  });
  for (const actor of Object.keys(byActor)) {
    const times = byActor[actor].sort((a, b) => a - b);
    for (let i = 0; i + BURST_THRESHOLD_COUNT - 1 < times.length; i++) {
      const windowStart = times[i];
      const windowEnd = times[i + BURST_THRESHOLD_COUNT - 1];
      if ((windowEnd - windowStart) <= BURST_THRESHOLD_MINUTES * 60 * 1000) {
        flags.push({
          type: 'burst',
          actor,
          detail: `${BURST_THRESHOLD_COUNT} verifications within ${Math.round((windowEnd - windowStart) / 60000)} minute(s)`,
          timestamp: new Date(windowEnd).toISOString()
        });
        break; // one flag per actor is enough, don't spam duplicates from overlapping windows
      }
    }
  }

  // --- Unusual-hour detection: activity between midnight and 5 AM India time ---
  // Timestamps are already stored as IST directly, so the parsed hour
  // component below is already the correct India-time hour - no further
  // timezone conversion needed (or wanted).
  const unusualHourActors = {};
  recent.forEach(r => {
    const raw = String(r[0] || '').replace(/^'/, '').trim();
    const match = raw.match(/^\d{4}-\d{1,2}-\d{1,2}[ T](\d{1,2}):/);
    if (!match) return;
    const istHour = Number(match[1]);
    if (istHour >= UNUSUAL_HOUR_START && istHour < UNUSUAL_HOUR_END) {
      const actor = String(r[1] || 'unknown');
      if (actor === 'System' || actor === 'unknown') return; // scheduled jobs run overnight - not worth flagging
      if (!unusualHourActors[actor]) unusualHourActors[actor] = 0;
      unusualHourActors[actor]++;
    }
  });
  for (const actor of Object.keys(unusualHourActors)) {
    flags.push({
      type: 'unusual_hour',
      actor,
      detail: `${unusualHourActors[actor]} action(s) between midnight and 5 AM`,
      timestamp: null
    });
  }

  return { ok: true, flags };
}

// Parses a timestamp that may or may not have the apostrophe prefix used to
// force text storage, tolerating both padded and unpadded values.
function parseTimestampLoose(raw) {
  const str = String(raw || '').replace(/^'/, '').trim();
  const match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2}):(\d{1,2})/);
  if (!match) return null;
  const [, y, mo, d, h, mi, s] = match.map(Number);
  const date = new Date(y, mo - 1, d, h, mi, s);
  return isNaN(date.getTime()) ? null : date;
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
    isViewer: users[key].role === 'viewer',
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
    const notesVal = col.notes > -1 ? (row[col.notes] || '') : '';
    const photoUrlVal = col.photoUrl > -1 ? (row[col.photoUrl] || '') : '';

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
      firstVerifiedByIsAdmin: isAdminLabel(String(firstVerifiedByVal)),
      notes: String(notesVal),
      photoUrl: String(photoUrlVal)
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
  if (usersForRoleCheck[callerKey] && usersForRoleCheck[callerKey].role === 'viewer') {
    return { ok: false, error: 'Viewer accounts have read-only access and cannot mark verification status.' };
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
  // Prefixed with ' so Google Sheets stores this as literal text rather than
  // auto-converting to a native date, which silently strips zero-padding
  // from single-digit hours on read-back and breaks our IST conversion.
  const timestamp = getISTTimestampForStorage();

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

async function bulkUpdateStatus(rowIds, status, userId) {
  if (status !== 'done' && status !== 'pending') {
    return { ok: false, error: "status must be 'done' or 'pending'" };
  }
  const users = await getAllUsers();
  const callerKey = userId ? String(userId) : '';
  if (users[callerKey] && users[callerKey].role === 'demo') {
    return { ok: false, error: 'Demo accounts cannot mark biometric verification status.' };
  }
  if (users[callerKey] && users[callerKey].role === 'viewer') {
    return { ok: false, error: 'Viewer accounts have read-only access and cannot mark verification status.' };
  }

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const value = status === 'done' ? 'Done' : 'Not Done';
  const whoLabel = userId ? String(userId) : 'unknown';
  const verifierRole = (users[whoLabel] && users[whoLabel].role === 'admin') ? 'admin' : 'staff';
  const verifierName = (users[whoLabel] && users[whoLabel].name) ? users[whoLabel].name : whoLabel;
  const storedVerifiedBy = verifierName + (verifierRole === 'admin' ? ' [Admin]' : '');
  const timestamp = getISTTimestampForStorage();

  const allUpdates = [];
  let updatedCount = 0;
  let skippedLocked = 0;
  const updatedNames = [];

  rowIds.forEach(rowId => {
    const rowNum = parseInt(String(rowId).replace('row', ''), 10);
    if (!rowNum || rowNum < 2) return;
    const row = data[rowNum - 1] || [];
    const currentLock = String(row[col.lock] || '').trim().toLowerCase();
    if (currentLock === 'yes') { skippedLocked++; return; }

    allUpdates.push({ range: `${sheetName}!${colToLetter(col.status)}${rowNum}`, values: [[value]] });
    allUpdates.push({ range: `${sheetName}!${colToLetter(col.lock)}${rowNum}`, values: [['Yes']] });
    allUpdates.push({ range: `${sheetName}!${colToLetter(col.verifiedBy)}${rowNum}`, values: [[storedVerifiedBy]] });
    allUpdates.push({ range: `${sheetName}!${colToLetter(col.verifiedAt)}${rowNum}`, values: [[timestamp]] });

    const existingFirstVerifiedBy = String(row[col.firstVerifiedBy] || '').trim();
    if (!existingFirstVerifiedBy) {
      allUpdates.push({ range: `${sheetName}!${colToLetter(col.firstVerifiedBy)}${rowNum}`, values: [[storedVerifiedBy]] });
      allUpdates.push({ range: `${sheetName}!${colToLetter(col.firstVerifiedAt)}${rowNum}`, values: [[timestamp]] });
    }

    updatedCount++;
    const studentName = col.name > -1 ? String(row[col.name] || '') : '';
    if (studentName) updatedNames.push(studentName);
  });

  if (allUpdates.length) await sheetsApi.batchWriteRanges(allUpdates);

  if (updatedCount) {
    const summary = updatedNames.slice(0, 5).join(', ') + (updatedNames.length > 5 ? ` and ${updatedNames.length - 5} more` : '');
    await logActivity(verifierName, status === 'done' ? 'Bulk Verified' : 'Bulk Marked Pending', `${updatedCount} student(s): ${summary}`);
  }

  return { ok: true, updatedCount, skippedLocked };
}

async function updateStudentNote(rowId, note, actor) {
  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  let col = detectColumns(headers);
  col = await ensureExtraColumns(sheetName, headers, col);

  const rowNum = parseInt(String(rowId).replace('row', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid student id.' };

  await sheetsApi.writeRange(`${sheetName}!${colToLetter(col.notes)}${rowNum}`, [[String(note || '').slice(0, 2000)]]);

  const row = data[rowNum - 1] || [];
  const studentName = col.name > -1 ? String(row[col.name] || '') : '';
  await logActivity(actor || 'unknown', 'Updated Note', `${studentName} (row ${rowNum})`);

  return { ok: true };
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

  role = (role === 'admin' || role === 'demo' || role === 'viewer') ? role : 'staff';
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
    userRecord.permissions = { hostelAccess: !!permissions.hostelAccess, allowUndo: !!permissions.allowUndo };
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

  const role = (newRole === 'admin' || newRole === 'demo' || newRole === 'viewer') ? newRole : 'staff';
  if (users[key].role === 'admin' && role !== 'admin') {
    const adminCount = Object.keys(users).filter(k => users[k].role === 'admin').length;
    if (adminCount <= 1) return { ok: false, error: 'Cannot demote the only remaining admin account.' };
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
    users[key].permissions = { hostelAccess: !!(permissions && permissions.hostelAccess), allowUndo: !!(permissions && permissions.allowUndo) };
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
  const timestamp = getISTTimestampForStorage();

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

async function logSessionEnd(actor, durationText, userId) {
  await logActivity(actor || 'unknown', 'Session Ended', 'Duration: ' + (durationText || 'unknown'));
  if (userId) await clearActiveSession(userId);
  return { ok: true };
}

// ---------- Heartbeat-based unexpected-close detection ----------
// Client pings this periodically while logged in. If pings stop arriving
// (e.g. a browser crash, force-quit, or power loss - none of which give
// any JavaScript a chance to run and report a clean close), a periodic
// server-side check notices the gap and logs a proper "Session Ended"
// entry anyway, rather than leaving that session's end silently unrecorded.

const ACTIVE_SESSIONS_KEY = 'activeSessions';
const HEARTBEAT_STALE_MINUTES = 10; // no heartbeat for this long = presumed unexpectedly closed

async function recordHeartbeat(userId, actorName) {
  if (!userId) return { ok: false, error: 'Missing userId.' };
  const sessions = await getSetting(ACTIVE_SESSIONS_KEY, {});
  const existing = sessions[userId];
  const store = requestContext.getStore();
  const ip = (store && store.ip) || (existing && existing.ip) || '';
  sessions[userId] = {
    actorName: actorName || (existing && existing.actorName) || 'unknown',
    loginTime: (existing && existing.loginTime) || Date.now(),
    lastHeartbeat: Date.now(),
    ip
  };
  await setSetting(ACTIVE_SESSIONS_KEY, sessions);
  return { ok: true };
}

async function clearActiveSession(userId) {
  if (!userId) return { ok: true };
  const sessions = await getSetting(ACTIVE_SESSIONS_KEY, {});
  if (sessions[userId]) {
    delete sessions[userId];
    await setSetting(ACTIVE_SESSIONS_KEY, sessions);
  }
  return { ok: true };
}

// ---------- Public QR lookup (no login required) ----------
// Deliberately returns only what's needed for a quick public check - name,
// status, and site/hostel context - never the full roster, never who
// verified them, never any other student's data. A QR code only encodes
// one specific identifier, so this only ever matches and returns one record.

async function publicLookupStudent(identifier) {
  if (!identifier) return { ok: false, error: 'No identifier provided.' };
  const data = await getStudents();
  if (!data.ok) return { ok: false, error: 'Could not look this up right now.' };
  const q = String(identifier).trim().toLowerCase();
  const match = data.students.find(s =>
    (s.appNo && s.appNo.toLowerCase() === q) ||
    (s.regNo && s.regNo.toLowerCase() === q) ||
    (s.machineCode && s.machineCode.toLowerCase() === q)
  );
  if (!match) return { ok: false, error: 'No matching record found.' };
  return {
    ok: true,
    name: match.name,
    appNo: match.appNo,
    regNo: match.regNo,
    siteCode: match.siteCode,
    status: match.status
  };
}

async function publicLookupHostelStudent(identifier) {
  if (!identifier) return { ok: false, error: 'No identifier provided.' };
  const data = await getHostelData();
  if (!data.ok) return { ok: false, error: 'Could not look this up right now.' };
  const q = String(identifier).trim().toLowerCase();
  const match = data.students.find(s =>
    (s.applicationNo && s.applicationNo.toLowerCase() === q) ||
    (s.registrationNo && s.registrationNo.toLowerCase() === q) ||
    (s.machineCode && s.machineCode.toLowerCase() === q)
  );
  if (!match) return { ok: false, error: 'No matching record found.' };
  return {
    ok: true,
    name: match.studentName,
    appNo: match.applicationNo,
    regNo: match.registrationNo,
    hostelName: match.hostelName,
    status: match.status
  };
}

async function getOnlineUsers() {
  const sessions = await getSetting(ACTIVE_SESSIONS_KEY, {});
  const cutoff = Date.now() - HEARTBEAT_STALE_MINUTES * 60 * 1000;
  const online = Object.values(sessions)
    .filter(s => s.lastHeartbeat >= cutoff)
    .map(s => ({ actorName: s.actorName, loginTime: s.loginTime, lastHeartbeat: s.lastHeartbeat, ip: s.ip || '' }))
    .sort((a, b) => b.loginTime - a.loginTime);
  return { ok: true, users: online };
}

async function getStaffLeaderboard() {
  const rows = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const verifyActions = ['Verified', 'Hostel Face Capture Verified'];
  const counts = {};
  rows.forEach(r => {
    const action = String(r[2] || '');
    if (!verifyActions.some(a => action.indexOf(a) !== -1)) return;
    const t = parseTimestampLoose(r[0]);
    if (!t || t.getTime() < cutoff) return;
    const actor = String(r[1] || 'unknown');
    counts[actor] = (counts[actor] || 0) + 1;
  });
  const leaderboard = Object.keys(counts)
    .map(actor => ({ actor, count: counts[actor] }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);
  return { ok: true, leaderboard };
}

// ---------- "What changed since you were last here" login digest ----------
async function getLoginDigest(actorName) {
  const rows = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
  // Find this actor's second-most-recent "Login Successful" - the most recent
  // one is the login that just happened, so the one before that is the
  // actual "since you were last here" reference point.
  const logins = rows.filter(r => r[2] === 'Login Successful' && r[1] === actorName);
  if (logins.length < 2) return { ok: true, hasPrevious: false };
  const previousLoginTime = parseTimestampLoose(logins[logins.length - 2][0]);
  if (!previousLoginTime) return { ok: true, hasPrevious: false };

  const cutoffMs = previousLoginTime.getTime();
  let verifiedCount = 0, newStaffCount = 0, unusualFlagCount = 0;
  const verifyActions = ['Verified', 'Hostel Face Capture Verified'];
  rows.forEach(r => {
    const t = parseTimestampLoose(r[0]);
    if (!t || t.getTime() <= cutoffMs) return;
    const action = String(r[2] || '');
    if (verifyActions.some(a => action.indexOf(a) !== -1)) verifiedCount++;
    if (action === 'Created User') newStaffCount++;
  });
  const flags = await getUnusualActivityFlags();
  unusualFlagCount = (flags.ok && flags.flags) ? flags.flags.length : 0;

  return {
    ok: true, hasPrevious: true,
    previousLoginTime: logins[logins.length - 2][0],
    verifiedCount, newStaffCount, unusualFlagCount
  };
}

// ---------- Undo verification (per-user permission, set via Manage Users) ----------
const UNDO_WINDOW_SECONDS = 15;

async function undoRecentVerification(rowId, actor, actorUserId) {
  const users = await getAllUsers();
  const caller = users[String(actorUserId || '')];
  const callerPerms = caller ? effectivePermissions(caller) : {};
  if (!callerPerms.allowUndo) return { ok: false, error: 'Your account is not allowed to undo verifications. Ask your admin to enable this under Manage Users.' };

  const sheetName = await sheetsApi.getMasterSheetName();
  const data = await sheetsApi.readRange(`${sheetName}!A1:ZZ`);
  const headers = data[0];
  const col = detectColumns(headers);
  const rowNum = parseInt(String(rowId).replace('row', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid student id.' };
  const row = data[rowNum - 1] || [];

  const verifiedAtRaw = String(row[col.verifiedAt] || '').replace(/^'/, '').trim();
  const verifiedAt = parseTimestampLoose(verifiedAtRaw);
  if (!verifiedAt) return { ok: false, error: 'No recent verification found to undo.' };
  const secondsSince = (Date.now() - verifiedAt.getTime()) / 1000;
  if (secondsSince > UNDO_WINDOW_SECONDS + 5) { // small buffer for clock/network variance
    return { ok: false, error: 'The undo window has expired.' };
  }

  await sheetsApi.batchWriteRanges([
    { range: `${sheetName}!${colToLetter(col.status)}${rowNum}`, values: [['Not Done']] },
    { range: `${sheetName}!${colToLetter(col.lock)}${rowNum}`, values: [['']] }
  ]);
  const studentName = col.name > -1 ? String(row[col.name] || '') : '';
  await logActivity(actor || 'unknown', 'Undid Verification', `${studentName} (row ${rowNum})`);
  return { ok: true };
}

async function undoRecentHostelVerification(rowId, actor, actorUserId) {
  const users = await getAllUsers();
  const caller = users[String(actorUserId || '')];
  const callerPerms = caller ? effectivePermissions(caller) : {};
  if (!callerPerms.allowUndo) return { ok: false, error: 'Your account is not allowed to undo verifications. Ask your admin to enable this under Manage Users.' };

  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  const col = detectHostelColumns(headers);
  const rowNum = parseInt(String(rowId).replace('hrow', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid record id.' };
  const row = data[rowNum - 1] || [];

  const verifiedAtRaw = String(row[col.verifiedAt] || '').replace(/^'/, '').trim();
  const verifiedAt = parseTimestampLoose(verifiedAtRaw);
  if (!verifiedAt) return { ok: false, error: 'No recent verification found to undo.' };
  const secondsSince = (Date.now() - verifiedAt.getTime()) / 1000;
  if (secondsSince > UNDO_WINDOW_SECONDS + 5) {
    return { ok: false, error: 'The undo window has expired.' };
  }

  await sheetsApi.batchWriteRanges([
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.status)}${rowNum}`, values: [['Not Done']] },
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${rowNum}`, values: [['']] }
  ]);
  const studentName = col['studentname'] > -1 ? String(row[col['studentname']] || '') : '';
  await logActivity(actor || 'unknown', 'Undid Hostel Verification', `${studentName} (row ${rowNum})`);
  return { ok: true };
}

async function checkStaleSessions() {
  const sessions = await getSetting(ACTIVE_SESSIONS_KEY, {});
  const cutoff = Date.now() - HEARTBEAT_STALE_MINUTES * 60 * 1000;
  let changed = false;

  for (const userId of Object.keys(sessions)) {
    const session = sessions[userId];
    if (session.lastHeartbeat < cutoff) {
      const durationMs = session.lastHeartbeat - session.loginTime;
      const totalSeconds = Math.max(0, Math.floor(durationMs / 1000));
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const durationText = hours > 0 ? `${hours}h ${minutes}m` : `${minutes}m`;
      await logActivity(session.actorName, 'Session Ended', `Duration: ${durationText} (connection lost unexpectedly)`);
      delete sessions[userId];
      changed = true;
    }
  }

  if (changed) await setSetting(ACTIVE_SESSIONS_KEY, sessions);
  return { ok: true, checked: Object.keys(sessions).length };
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

  const filteredRows = [serialColIndex > -1 ? headers : ['Sl. No.', ...headers]];
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
      filteredRows.push(row);
    } else {
      filteredRows.push([serialCounter, ...row]);
    }
    serialCounter++;
  }

  const csvText = filteredRows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  return { ok: true, csv: csvText, rowCount: filteredRows.length - 1, filename: `roster_export_${Date.now()}.csv` };
}

// ---------- Import stats ----------

async function getLastImportInfo() {
  const data = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][2] === 'Imported New Students') {
      const match = String(data[i][3] || '').match(/^(\d+)\s+added/);
      const count = match ? parseInt(match[1], 10) : 0;
      if (count === 0) continue; // an import that added nothing shouldn't overwrite the last genuine "fresh data" moment
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
  const match = String(utcTimestampStr).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2}):(\d{1,2})/);
  if (!match) return utcTimestampStr;
  const [, y, mo, d, h, mi, s] = match.map(Number);
  const utcDate = new Date(Date.UTC(y, mo - 1, d, h, mi, s));
  if (isNaN(utcDate.getTime())) return utcTimestampStr;
  return new Intl.DateTimeFormat('en-IN', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: 'short', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hour12: true
  }).format(utcDate);
}

// Compact variant matching the Activity Log's dense "YYYY-MM-DD HH:MM:SS"
// style, just with the correct India-time offset applied.
function formatUtcTimestampAsIndiaTimeCompact(utcTimestampStr) {
  if (!utcTimestampStr) return '';
  // Parses components manually (rather than string-based ISO parsing) so
  // this tolerates unpadded values too, e.g. "2026-8-6 4:16:57" - which
  // can happen for older entries logged before dates were forced to text.
  const match = String(utcTimestampStr).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{1,2}):(\d{1,2})/);
  if (!match) return utcTimestampStr;
  const [, y, mo, d, h, mi, s] = match.map(Number);
  const utcDate = new Date(Date.UTC(y, mo - 1, d, h, mi, s));
  if (isNaN(utcDate.getTime())) return utcTimestampStr;
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  }).formatToParts(utcDate);
  const get = type => parts.find(p => p.type === type).value;
  return `${get('year')}-${get('month')}-${get('day')} ${get('hour')}:${get('minute')}:${get('second')}`;
}

async function getLastHostelImportInfo() {
  const data = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][2] === 'Imported New Hostel Data') {
      const match = String(data[i][3] || '').match(/^(\d+)\s+added/);
      const count = match ? parseInt(match[1], 10) : 0;
      if (count === 0) continue; // an import that added nothing shouldn't overwrite the last genuine "fresh data" moment
      return { ok: true, count, timestamp: data[i][0] || null };
    }
  }
  return { ok: true, count: 0, timestamp: null };
}

async function getTodayImportCount() {
  const info = await getLastImportInfo();
  const todayStr = getISTTodayDateString();
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
  let statusCol = -1, lockCol = -1, verifiedAtCol = -1, verifiedByCol = -1, firstVerifiedByCol = -1, firstVerifiedAtCol = -1;
  headers.forEach((h, i) => {
    const n = normalize(h);
    if (n.indexOf('facecapture') !== -1 || n === 'status') statusCol = i;
    if (n === 'locked') lockCol = i;
    if (n === 'firstverifiedby') firstVerifiedByCol = i;
    else if (n === 'firstverifiedat') firstVerifiedAtCol = i;
    else if (n.indexOf('verifiedby') !== -1) verifiedByCol = i;
    else if (n.indexOf('verifiedat') !== -1) verifiedAtCol = i;
  });
  col.status = statusCol;
  col.lock = lockCol;
  col.verifiedAt = verifiedAtCol;
  col.verifiedBy = verifiedByCol;
  col.firstVerifiedBy = firstVerifiedByCol;
  col.firstVerifiedAt = firstVerifiedAtCol;
  if (col.notes === undefined) col.notes = -1;
  if (col.photoUrl === undefined) col.photoUrl = (col['photourl'] !== undefined ? col['photourl'] : (col['photo'] !== undefined ? col['photo'] : -1));
  return col;
}

async function ensureHostelExtraColumns(headers, col) {
  const additions = [];
  if (col.status === -1) { col.status = headers.length + additions.length; additions.push('Face Capture Status'); }
  if (col.lock === -1) { col.lock = headers.length + additions.length; additions.push('Locked'); }
  if (col.verifiedAt === -1) { col.verifiedAt = headers.length + additions.length; additions.push('Verified At'); }
  if (col.verifiedBy === -1) { col.verifiedBy = headers.length + additions.length; additions.push('Verified By'); }
  if (col.firstVerifiedBy === -1) { col.firstVerifiedBy = headers.length + additions.length; additions.push('First Verified By'); }
  if (col.firstVerifiedAt === -1) { col.firstVerifiedAt = headers.length + additions.length; additions.push('First Verified At'); }
  if (col.notes === -1) { col.notes = headers.length + additions.length; additions.push('Notes'); }
  if (col.photoUrl === -1) { col.photoUrl = headers.length + additions.length; additions.push('Photo URL'); }
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
      verifiedAt: String(row[col.verifiedAt] || ''),
      verifiedBy: String(row[col.verifiedBy] || ''),
      firstVerifiedBy: String(row[col.firstVerifiedBy] || ''),
      firstVerifiedAt: String(row[col.firstVerifiedAt] || ''),
      notes: String(row[col.notes] || ''),
      photoUrl: String(row[col.photoUrl] || '')
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
  if (users[callerKey] && users[callerKey].role === 'viewer') {
    return { ok: false, error: 'Viewer accounts have read-only access and cannot mark face capture status.' };
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
  const timestamp = getISTTimestampForStorage();

  const caller = users[callerKey] || {};
  const callerName = caller.name || callerKey || 'unknown';
  let roleTag = '';
  if (caller.role === 'admin') roleTag = ' [Admin]';
  else if (caller.role === 'demo') roleTag = ' [Demo]';
  else if (caller.role === 'staff' && caller.permissions && caller.permissions.hostelAccess) roleTag = ' [Hosteller]';

  const existingFirstVerifiedBy = String(row[col.firstVerifiedBy] || '').trim();
  const isRevision = !!existingFirstVerifiedBy; // a prior verification already exists, so this is an overturn/change, not the first-ever mark
  const verifiedByLabel = callerName + roleTag + (isRevision ? ' (Revised)' : '');

  const writes = [
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.status)}${rowNum}`, values: [[value]] },
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${rowNum}`, values: [['Yes']] },
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.verifiedAt)}${rowNum}`, values: [[timestamp]] },
    { range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.verifiedBy)}${rowNum}`, values: [[verifiedByLabel]] }
  ];
  if (!existingFirstVerifiedBy) {
    writes.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.firstVerifiedBy)}${rowNum}`, values: [[callerName + roleTag]] });
    writes.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.firstVerifiedAt)}${rowNum}`, values: [[timestamp]] });
  }
  await sheetsApi.batchWriteRanges(writes);

  const name = col['studentname'] > -1 ? String(row[col['studentname']] || '') : '';
  await logActivity(callerName, status === 'done' ? 'Hostel Face Capture Verified' : 'Hostel Marked Pending', `${name} (row ${rowNum})`);
  return { ok: true };
}

async function bulkUpdateHostelStatus(rowIds, status, userId) {
  if (status !== 'done' && status !== 'pending') {
    return { ok: false, error: "status must be 'done' or 'pending'" };
  }
  const users = await getAllUsers();
  const callerKey = userId ? String(userId) : '';
  if (users[callerKey] && users[callerKey].role === 'demo') {
    return { ok: false, error: 'Demo accounts cannot mark face capture status.' };
  }
  if (users[callerKey] && users[callerKey].role === 'viewer') {
    return { ok: false, error: 'Viewer accounts have read-only access and cannot mark face capture status.' };
  }

  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const value = status === 'done' ? 'Done' : 'Not Done';
  const timestamp = getISTTimestampForStorage();
  const caller = users[callerKey] || {};
  const callerName = caller.name || callerKey || 'unknown';
  let roleTag = '';
  if (caller.role === 'admin') roleTag = ' [Admin]';
  else if (caller.role === 'demo') roleTag = ' [Demo]';
  else if (caller.role === 'staff' && caller.permissions && caller.permissions.hostelAccess) roleTag = ' [Hosteller]';

  const allUpdates = [];
  let updatedCount = 0;
  let skippedLocked = 0;
  const updatedNames = [];

  rowIds.forEach(rowId => {
    const rowNum = parseInt(String(rowId).replace('hrow', ''), 10);
    if (!rowNum || rowNum < 2) return;
    const row = data[rowNum - 1] || [];
    if (String(row[col.lock] || '').trim().toLowerCase() === 'yes') { skippedLocked++; return; }

    const existingFirstVerifiedBy = String(row[col.firstVerifiedBy] || '').trim();
    const isRevision = !!existingFirstVerifiedBy;
    const verifiedByLabel = callerName + roleTag + (isRevision ? ' (Revised)' : '');

    allUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.status)}${rowNum}`, values: [[value]] });
    allUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${rowNum}`, values: [['Yes']] });
    allUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.verifiedAt)}${rowNum}`, values: [[timestamp]] });
    allUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.verifiedBy)}${rowNum}`, values: [[verifiedByLabel]] });
    if (!existingFirstVerifiedBy) {
      allUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.firstVerifiedBy)}${rowNum}`, values: [[callerName + roleTag]] });
      allUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.firstVerifiedAt)}${rowNum}`, values: [[timestamp]] });
    }

    updatedCount++;
    const studentName = col['studentname'] > -1 ? String(row[col['studentname']] || '') : '';
    if (studentName) updatedNames.push(studentName);
  });

  if (allUpdates.length) await sheetsApi.batchWriteRanges(allUpdates);

  if (updatedCount) {
    const summary = updatedNames.slice(0, 5).join(', ') + (updatedNames.length > 5 ? ` and ${updatedNames.length - 5} more` : '');
    await logActivity(callerName, status === 'done' ? 'Bulk Hostel Verified' : 'Bulk Hostel Marked Pending', `${updatedCount} student(s): ${summary}`);
  }

  return { ok: true, updatedCount, skippedLocked };
}

async function updateHostelStudentNote(rowId, note, actor) {
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const rowNum = parseInt(String(rowId).replace('hrow', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid record id.' };

  await sheetsApi.writeRange(`${HOSTEL_SHEET_NAME}!${colToLetter(col.notes)}${rowNum}`, [[String(note || '').slice(0, 2000)]]);

  const row = data[rowNum - 1] || [];
  const studentName = col['studentname'] > -1 ? String(row[col['studentname']] || '') : '';
  await logActivity(actor || 'unknown', 'Updated Note', `${studentName} (row ${rowNum})`);

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

async function deleteHostelStudent(rowId, adminPassword, actor, actorUserId) {
  const currentAdminPassword = await getAdminPassword();
  if (adminPassword !== currentAdminPassword) return { ok: false, error: 'Incorrect admin password.' };

  if (actorUserId) {
    const users = await getAllUsers();
    if (users[actorUserId] && users[actorUserId].role === 'demo') {
      return { ok: false, error: 'Demo accounts cannot delete student records.' };
    }
  }

  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  const rowNum = parseInt(String(rowId).replace('hrow', ''), 10);
  if (!rowNum || rowNum < 2) return { ok: false, error: 'Invalid record id.' };

  const row = data[rowNum - 1] || [];
  const name = col['studentname'] > -1 ? String(row[col['studentname']] || '') : '';
  const appNo = col['applicationno'] > -1 ? String(row[col['applicationno']] || '') : '';

  await sheetsApi.deleteRow(HOSTEL_SHEET_NAME, rowNum);
  await logActivity(actor || 'Admin', 'Hostel Student Deleted', `${name} (App No: ${appNo})`);
  return { ok: true };
}

async function exportHostelAsCsv(statusFilter) {
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  if (!data.length) return { ok: false, error: 'Hostel data is empty.' };
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  let serialColIndex = -1;
  for (let h = 0; h < headers.length; h++) {
    const normHeader = normalize(headers[h]);
    if (['s', 'slno', 'srno', 'sno', 'serialno', 'serialnumber'].includes(normHeader)) { serialColIndex = h; break; }
  }

  const wantStatus = statusFilter && statusFilter !== 'all' ? statusFilter : null;
  const rows = [serialColIndex > -1 ? headers : ['Sl. No.', ...headers]];
  let serialCounter = 1;
  for (let r = 1; r < data.length; r++) {
    let row = data[r];
    if (wantStatus) {
      const statusVal = String(row[col.status] || '').trim().toLowerCase();
      const isDone = (statusVal === 'done' || statusVal === 'yes' || statusVal === 'true');
      if (wantStatus === 'done' && !isDone) continue;
      if (wantStatus === 'pending' && isDone) continue;
    }
    if (serialColIndex > -1) {
      row = row.slice();
      row[serialColIndex] = serialCounter;
      rows.push(row);
    } else {
      rows.push([serialCounter, ...row]);
    }
    serialCounter++;
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

  let serialColIndex = -1;
  for (let h = 0; h < headers.length; h++) {
    const normHeader = normalize(headers[h]);
    if (['s', 'slno', 'srno', 'sno', 'serialno', 'serialnumber'].includes(normHeader)) { serialColIndex = h; break; }
  }

  const todayStr = getISTTodayDateString();
  const rows = [serialColIndex > -1 ? headers : ['Sl. No.', ...headers]];
  let serialCounter = 1;
  for (let r = 1; r < data.length; r++) {
    let row = data[r];
    const verifiedAt = String(row[col.verifiedAt] || '');
    if (verifiedAt.indexOf(todayStr) === 0) {
      if (serialColIndex > -1) {
        row = row.slice();
        row[serialColIndex] = serialCounter;
        rows.push(row);
      } else {
        rows.push([serialCounter, ...row]);
      }
      serialCounter++;
    }
  }
  const csvText = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  return { ok: true, csv: csvText, rowCount: rows.length - 1, filename: `hostel_verified_today_${Date.now()}.csv` };
}

async function exportHostelVerifiedLastDayAsCsv() {
  const data = await sheetsApi.readRange(`${HOSTEL_SHEET_NAME}!A1:ZZ`);
  if (!data.length) return { ok: false, error: 'Hostel data is empty.' };
  const headers = data[0];
  let col = detectHostelColumns(headers);
  col = await ensureHostelExtraColumns(headers, col);

  let serialColIndex = -1;
  for (let h = 0; h < headers.length; h++) {
    const normHeader = normalize(headers[h]);
    if (['s', 'slno', 'srno', 'sno', 'serialno', 'serialnumber'].includes(normHeader)) { serialColIndex = h; break; }
  }

  const lastDayStr = getISTLastDayDateString();
  const rows = [serialColIndex > -1 ? headers : ['Sl. No.', ...headers]];
  let serialCounter = 1;
  for (let r = 1; r < data.length; r++) {
    let row = data[r];
    const verifiedAt = String(row[col.verifiedAt] || '');
    if (verifiedAt.indexOf(lastDayStr) === 0) {
      if (serialColIndex > -1) {
        row = row.slice();
        row[serialColIndex] = serialCounter;
        rows.push(row);
      } else {
        rows.push([serialCounter, ...row]);
      }
      serialCounter++;
    }
  }
  const csvText = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  return { ok: true, csv: csvText, rowCount: rows.length - 1, filename: `hostel_verified_last_day_${Date.now()}.csv` };
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

    const timestamp = getISTTimestampForStorage();
    pendingUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.status)}${matchedRowNum}`, values: [['Done']] });
    pendingUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.lock)}${matchedRowNum}`, values: [['Yes']] });
    pendingUpdates.push({ range: `${HOSTEL_SHEET_NAME}!${colToLetter(col.verifiedAt)}${matchedRowNum}`, values: [[timestamp]] });
    updated++;
  });

  await sheetsApi.batchWriteRanges(pendingUpdates);
  await logActivity(actor || 'Admin', 'Imported Hostel Face Capture Updates', `${updated} updated, ${alreadyDone} already done, ${notFoundRows.length} not found`);
  return { ok: true, updated, alreadyDone, notFoundRows };
}

// ---------- AI-powered Help Assistant ----------

const HELP_ASSISTANT_SYSTEM_PROMPT = `You are the in-app Help Assistant for "Biometric Verification Desk", an internal tool for Adamas University IT staff (AKC IT Support) tracking student biometric verification and hostel face-capture verification.

Key things staff and admins can do:
- Search for a student by App No, Reg No, Machine Code, or name; mark them Verified or Pending
- Once verified, a record locks - only an Admin can unlock it (enter admin password) before changing it again
- Admins can: bulk import new students, bulk import verification updates from a file, download CSV reports (filtered by Verified/Pending/All), manage user accounts and permissions, view the Activity Log, download a PDF of the activity log for a date range
- There's a separate "Hostel Lookup" page for Face Capture verification of hostel students, working the same way but with its own roster
- "Demo" accounts have full admin-like access but cannot verify/delete anything (hard-restricted)
- Staff accounts can be individually granted just "Hostel Data Access" without becoming admin

Answer questions clearly and concisely, in plain language, focused on how to actually do the thing they're asking about. If a question is about something outside this app's scope, say so briefly and suggest contacting AKC IT Support. Keep answers short - a few sentences at most, this is a small in-app chat widget, not a long document.`;

const ALLOWED_HELP_CHAT_EVENTS = ['Started New Chat', 'Ended Chat'];
async function logHelpChatEvent(actor, eventType) {
  if (!ALLOWED_HELP_CHAT_EVENTS.includes(eventType)) return { ok: false, error: 'Invalid event type.' };
  await logActivity(actor || 'unknown', `Help Chat: ${eventType}`, '');
  return { ok: true };
}

const VOICE_COMMAND_SYSTEM_PROMPT = `You parse a spoken voice command into a structured action for a student verification app. Respond with ONLY a JSON object, no other text, no markdown fences.

Recognized actions:
- "mark_verified": mark a student as verified/done. Needs an "identifier" (App No, Reg No, Machine Code, or student name, exactly as spoken).
- "mark_pending": mark a student back as pending/not done. Needs an "identifier".
- "search": the person just said a name, App No, Reg No, or Machine Code with no explicit instruction attached - they want to find/look up that student, not necessarily change their status. This is the right choice whenever no clear verb like "mark", "set", or "change" is present - default to this over "unknown" whenever an identifier-like value is present.
- "unknown": the command doesn't match a recognized action, or contains no identifiable name/number at all.

Respond with exactly this shape:
{"action": "mark_verified" | "mark_pending" | "search" | "unknown", "identifier": "<string or null>"}

Examples:
"mark APP-2026-1234 as verified" -> {"action":"mark_verified","identifier":"APP-2026-1234"}
"set Soubhagya Chatterjee to pending" -> {"action":"mark_pending","identifier":"Soubhagya Chatterjee"}
"Soubhagya Chatterjee" -> {"action":"search","identifier":"Soubhagya Chatterjee"}
"find APP-2026-1234" -> {"action":"search","identifier":"APP-2026-1234"}
"what's the weather" -> {"action":"unknown","identifier":null}`;

async function parseVoiceCommand(spokenText) {
  const apiKey = (process.env.GROQ_API_KEY || '').trim();
  if (!apiKey) return { ok: false, error: 'Voice commands are not configured yet.' };
  try {
    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'openai/gpt-oss-120b',
        max_tokens: 600,
        reasoning_effort: 'low',
        temperature: 0,
        messages: [
          { role: 'system', content: VOICE_COMMAND_SYSTEM_PROMPT },
          { role: 'user', content: String(spokenText || '').slice(0, 300) }
        ]
      })
    });
    if (!response.ok) { const errText = await response.text(); throw new Error(`Groq API error (${response.status}): ${errText}`); }
    const data = await response.json();
    const raw = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '{}';
    const cleaned = raw.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(cleaned);
    if (!['mark_verified', 'mark_pending', 'search', 'unknown'].includes(parsed.action)) {
      return { ok: true, action: 'unknown', identifier: null };
    }
    return { ok: true, action: parsed.action, identifier: parsed.identifier || null };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

async function askAiHelpAssistant(question) {
  const apiKey = (process.env.GROQ_API_KEY || '').trim();
  if (!apiKey) {
    return { ok: false, error: 'AI Help Assistant is not configured yet.' };
  }
  try {
    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'openai/gpt-oss-120b',
        max_tokens: 700,
        reasoning_effort: 'low',
        messages: [
          { role: 'system', content: HELP_ASSISTANT_SYSTEM_PROMPT },
          { role: 'user', content: String(question || '').slice(0, 500) }
        ]
      })
    });
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Groq API error (${response.status}): ${errText}`);
    }
    const data = await response.json();
    const answer = (data.choices && data.choices[0] && data.choices[0].message &&
      data.choices[0].message.content) || "I couldn't come up with an answer for that.";
    return { ok: true, answer };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ---------- Natural-language roster search (via Help Assistant) ----------
const ROSTER_FILTER_SYSTEM_PROMPT = `You parse a natural-language request into structured filter criteria for a student roster search tool. Respond with ONLY a JSON object, no other text, no markdown fences.

Recognized fields (all optional, omit or use null if not mentioned):
- "siteCode": a site/location code mentioned, exactly as spoken (e.g. "Site B" -> "B").
- "status": "done" if they want verified/completed students, "pending" if they want not-yet-verified students, null if status isn't mentioned.
- "nameContains": a name or partial name mentioned to search for.
- "isFilterRequest": true if this message is asking to find/filter/show students, false if it's an unrelated help question (in which case the other fields should be null).

Respond with exactly this shape:
{"isFilterRequest": true|false, "siteCode": "<string or null>", "status": "done"|"pending"|null, "nameContains": "<string or null>"}

Examples:
"everyone from site B still pending" -> {"isFilterRequest":true,"siteCode":"B","status":"pending","nameContains":null}
"show me who's done at site A" -> {"isFilterRequest":true,"siteCode":"A","status":"done","nameContains":null}
"find students named sharma" -> {"isFilterRequest":true,"siteCode":null,"status":null,"nameContains":"sharma"}
"how do I unlock a record" -> {"isFilterRequest":false,"siteCode":null,"status":null,"nameContains":null}`;

async function parseRosterFilterQuery(query) {
  const apiKey = (process.env.GROQ_API_KEY || '').trim();
  if (!apiKey) return { ok: false, error: 'This feature is not configured yet.' };
  try {
    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'openai/gpt-oss-120b',
        max_tokens: 600,
        reasoning_effort: 'low',
        temperature: 0,
        messages: [
          { role: 'system', content: ROSTER_FILTER_SYSTEM_PROMPT },
          { role: 'user', content: String(query || '').slice(0, 300) }
        ]
      })
    });
    if (!response.ok) { const errText = await response.text(); throw new Error(`Groq API error (${response.status}): ${errText}`); }
    const data = await response.json();
    const raw = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '{}';
    const cleaned = raw.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(cleaned);
    return { ok: true, isFilterRequest: !!parsed.isFilterRequest, siteCode: parsed.siteCode || null, status: parsed.status || null, nameContains: parsed.nameContains || null };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ---------- AI anomaly explanations ----------
async function explainUnusualActivity(flags) {
  const apiKey = (process.env.GROQ_API_KEY || '').trim();
  if (!apiKey || !flags || !flags.length) return { ok: true, explanation: '' };
  try {
    const flagsSummary = flags.map(f => `${f.type}: ${f.detail}`).join('; ');
    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'openai/gpt-oss-120b',
        max_tokens: 500,
        reasoning_effort: 'low',
        temperature: 0.3,
        messages: [
          { role: 'system', content: 'You write one short, plain-English sentence (max 30 words) summarizing unusual activity flags for a non-technical staff member at a student verification desk. Be calm and factual, not alarming - these are patterns worth a glance, not confirmed problems. No preamble, just the sentence.' },
          { role: 'user', content: flagsSummary.slice(0, 800) }
        ]
      })
    });
    if (!response.ok) { const errText = await response.text(); throw new Error(`Groq API error (${response.status}): ${errText}`); }
    const data = await response.json();
    const explanation = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
    return { ok: true, explanation: explanation.trim() };
  } catch (err) {
    return { ok: true, explanation: '' }; // non-critical - fail silently, raw flags still show
  }
}

// ---------- AI shift handoff notes ----------
async function generateShiftHandoffNote(actorName, sinceTimestampMs) {
  const apiKey = (process.env.GROQ_API_KEY || '').trim();
  if (!apiKey) return { ok: false, error: 'This feature is not configured yet.' };
  try {
    const rows = await sheetsApi.readRange(`${sheetsApi.ACTIVITY_LOG_SHEET_NAME}!A2:E`);
    const sinceMs = Number(sinceTimestampMs) || (Date.now() - 8 * 60 * 60 * 1000);
    const relevant = rows.filter(r => {
      if (r[1] !== actorName) return false;
      const t = parseTimestampLoose(r[0]);
      return t && t.getTime() >= sinceMs;
    });
    if (!relevant.length) return { ok: true, note: `No activity recorded for ${actorName} this session.` };
    const actionsSummary = relevant.map(r => `${r[2]}: ${r[3] || ''}`).join('\n').slice(0, 1500);
    const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'openai/gpt-oss-120b',
        max_tokens: 600,
        reasoning_effort: 'low',
        temperature: 0.3,
        messages: [
          { role: 'system', content: 'You write an extremely concise shift handoff note for the next staff member taking over a student verification desk, based on a list of actions the outgoing staff member performed this session. STRICT LIMIT: exactly 2 short lines of plain text, no bullet points, no headers, no line breaks within a line. Line 1: what was done (totals only - imports, verifications, new students). Line 2: anything worth flagging, or "Nothing unusual to flag" if there is nothing. No preamble, no markdown, just the 2 lines separated by one newline.' },
          { role: 'user', content: actionsSummary }
        ]
      })
    });
    if (!response.ok) { const errText = await response.text(); throw new Error(`Groq API error (${response.status}): ${errText}`); }
    const data = await response.json();
    const note = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
    return { ok: true, note: note.trim() };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ---------- Duplicate/typo detection (algorithmic, not AI - fast and
// reliable for this specific task; comparing every pair with an AI call
// would be slow and expensive for a large roster) ----------
function levenshteinDistance(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[m][n];
}

async function findDuplicateStudents() {
  const data = await getStudents();
  if (!data.ok) return { ok: false, error: 'Could not load the roster.' };
  const named = data.students
    .filter(s => s.name && s.name.trim().length > 2)
    .map(s => ({ id: s.id, name: s.name.trim(), appNo: s.appNo, normalized: s.name.trim().toLowerCase().replace(/[^a-z\s]/g, '') }));

  // Bucket by first letter to avoid full O(n^2) comparison across the whole roster
  const buckets = {};
  named.forEach(s => {
    const key = s.normalized[0] || '?';
    if (!buckets[key]) buckets[key] = [];
    buckets[key].push(s);
  });

  const pairs = [];
  Object.values(buckets).forEach(bucket => {
    for (let i = 0; i < bucket.length; i++) {
      for (let j = i + 1; j < bucket.length; j++) {
        const a = bucket[i], b = bucket[j];
        if (a.normalized === b.normalized) continue; // exact match isn't a typo case
        const dist = levenshteinDistance(a.normalized, b.normalized);
        const maxLen = Math.max(a.normalized.length, b.normalized.length);
        const similarity = 1 - dist / maxLen;
        if (similarity >= 0.82 && dist <= 3) {
          pairs.push({ nameA: a.name, appNoA: a.appNo, idA: a.id, nameB: b.name, appNoB: b.appNo, idB: b.id, similarity: Math.round(similarity * 100) });
        }
      }
    }
  });
  pairs.sort((x, y) => y.similarity - x.similarity);
  return { ok: true, pairs: pairs.slice(0, 50) };
}

async function findDuplicateHostelStudents() {
  const data = await getHostelData();
  if (!data.ok) return { ok: false, error: 'Could not load the hostel roster.' };
  const named = data.students
    .filter(s => s.studentName && s.studentName.trim().length > 2)
    .map(s => ({ id: s.id, name: s.studentName.trim(), appNo: s.applicationNo, normalized: s.studentName.trim().toLowerCase().replace(/[^a-z\s]/g, '') }));

  const buckets = {};
  named.forEach(s => {
    const key = s.normalized[0] || '?';
    if (!buckets[key]) buckets[key] = [];
    buckets[key].push(s);
  });

  const pairs = [];
  Object.values(buckets).forEach(bucket => {
    for (let i = 0; i < bucket.length; i++) {
      for (let j = i + 1; j < bucket.length; j++) {
        const a = bucket[i], b = bucket[j];
        if (a.normalized === b.normalized) continue;
        const dist = levenshteinDistance(a.normalized, b.normalized);
        const maxLen = Math.max(a.normalized.length, b.normalized.length);
        const similarity = 1 - dist / maxLen;
        if (similarity >= 0.82 && dist <= 3) {
          pairs.push({ nameA: a.name, appNoA: a.appNo, idA: a.id, nameB: b.name, appNoB: b.appNo, idB: b.id, similarity: Math.round(similarity * 100) });
        }
      }
    }
  });
  pairs.sort((x, y) => y.similarity - x.similarity);
  return { ok: true, pairs: pairs.slice(0, 50) };
}

// ---------- Custom report templates ----------
// Templates are shared/global (not per-user) - a report like "Weekly Site B
// Status" should be runnable by any admin, not just whoever created it.
const REPORT_TEMPLATES_KEY = 'customReportTemplates';

async function getReportTemplates() {
  const templates = await getSetting(REPORT_TEMPLATES_KEY, []);
  return { ok: true, templates };
}

async function saveReportTemplate(name, config, actor) {
  const trimmedName = String(name || '').trim();
  if (!trimmedName) return { ok: false, error: 'Give the report a name.' };
  const templates = await getSetting(REPORT_TEMPLATES_KEY, []);
  const template = {
    id: 'report_' + Date.now(),
    name: trimmedName,
    config: config || {},
    createdBy: actor || 'unknown',
    createdAt: Date.now()
  };
  templates.push(template);
  await setSetting(REPORT_TEMPLATES_KEY, templates);
  await logActivity(actor || 'unknown', 'Saved Report Template', trimmedName);
  return { ok: true, template };
}

async function deleteReportTemplate(id, actor) {
  const templates = await getSetting(REPORT_TEMPLATES_KEY, []);
  const target = templates.find(t => t.id === id);
  const filtered = templates.filter(t => t.id !== id);
  await setSetting(REPORT_TEMPLATES_KEY, filtered);
  if (target) await logActivity(actor || 'unknown', 'Deleted Report Template', target.name);
  return { ok: true };
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
  logSessionIp, logSessionEnd, recordHeartbeat, clearActiveSession, checkStaleSessions,
  exportRosterAsCsv,
  getLastImportInfo, getTodayImportCount, getLastImportTimestamp,
  getHostelData, updateHostelStatus, adminUnlockHostel, deleteHostelStudent, exportHostelAsCsv, exportHostelVerifiedTodayAsCsv, exportHostelVerifiedLastDayAsCsv,
  importNewHostelData, importHostelVerificationUpdates, getLastHostelImportInfo,
  askAiHelpAssistant, logHelpChatEvent, getUnusualActivityFlags, parseVoiceCommand, getOnlineUsers,
  getStaffLeaderboard, getLoginDigest, undoRecentVerification, undoRecentHostelVerification,
  publicLookupStudent, publicLookupHostelStudent,
  requestContext,
  parseRosterFilterQuery, explainUnusualActivity, generateShiftHandoffNote, findDuplicateStudents, findDuplicateHostelStudents,
  getReportTemplates, saveReportTemplate, deleteReportTemplate,
  bulkUpdateStatus, updateStudentNote, bulkUpdateHostelStatus, updateHostelStudentNote
};
