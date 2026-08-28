// Backs up any chosen sheet(s) from the spreadsheet into a Neon (Postgres)
// database as a secondary copy. Google Sheets remains the live, primary
// data source for the app at all times - this is a one-way, on-demand
// snapshot copy, not a live sync. Each run fully replaces the backup
// table's contents for each sheet with whatever is currently in that
// sheet, so the backup always reflects the most recent snapshot taken,
// not a running history.
//
// This is intentionally generic rather than hardcoded to specific sheets:
// each sheet gets its own Postgres table (name derived from the sheet's
// own name), storing every row as a JSONB object keyed by that sheet's
// own column headers. This means it works for any sheet - roster, hostel,
// user accounts, activity log, or anything added later - without needing
// to know its column structure in advance.

const { Pool } = require('pg');

let pool = null;

function getPool() {
  const connectionString = (process.env.NEON_DATABASE_URL || '').trim();
  if (!connectionString) return null;
  if (!pool) {
    pool = new Pool({ connectionString, ssl: { rejectUnauthorized: false } });
  }
  return pool;
}

// Turns a sheet's own name into a safe, predictable Postgres table name -
// lowercase, non-alphanumeric characters replaced with underscores, and
// prefixed so it can never collide with an unrelated table name.
function tableNameForSheet(sheetName) {
  const safe = String(sheetName || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 50);
  return `sheet_backup_${safe || 'untitled'}`;
}

// Turns a sheet's own header row into safe, unique, valid Postgres column
// names - handles blank headers (falls back to a positional name) and
// duplicate headers (appends _2, _3, etc. so they don't collide), and
// guarantees the result never starts with a digit (Postgres requires
// identifiers to start with a letter or underscore).
function sanitizeColumnNames(headers) {
  const seen = {};
  return headers.map((h, i) => {
    let base = String(h || '').toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    if (!base || /^[0-9]/.test(base)) base = `col_${i + 1}_${base}`.replace(/_+$/, '');
    base = `f_${base}`;
    if (seen[base] === undefined) {
      seen[base] = 0;
      return base;
    }
    seen[base]++;
    return `${base}_${seen[base]}`;
  });
}

async function ensureTableForSheet(client, tableName, columnNames) {
  await client.query(`DROP TABLE IF EXISTS ${tableName}`);
  const columnDefs = columnNames.map(c => `${c} TEXT`).join(', ');
  await client.query(`
    CREATE TABLE ${tableName} (
      row_index INTEGER PRIMARY KEY,
      ${columnDefs},
      backed_up_at TIMESTAMPTZ DEFAULT now()
    )
  `);
}

// User Accounts gets a proper table with real, named columns instead of
// the generic JSONB approach - this is the one sheet people actually want
// to browse directly in Neon's table view, so it's worth a dedicated shape
// rather than making everyone write a JSON-extraction query for it.
async function ensureUserAccountsTable(client) {
  await client.query(`DROP TABLE IF EXISTS sheet_backup_user_accounts`);
  await client.query(`
    CREATE TABLE sheet_backup_user_accounts (
      user_id TEXT PRIMARY KEY,
      password TEXT,
      role TEXT,
      name TEXT,
      mobile_no TEXT,
      email TEXT,
      school TEXT,
      permissions JSONB,
      must_change_password BOOLEAN,
      demo_welcome_shown BOOLEAN,
      backed_up_at TIMESTAMPTZ DEFAULT now()
    )
  `);
}

function findColumnIndex(headers, targetName){
  const normalize = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const target = normalize(targetName);
  return headers.findIndex(h => normalize(h) === target);
}

// The main roster also gets a proper table with real columns, same
// reasoning as User Accounts - this is core data people want to browse
// and query directly, not unwrap from JSON every time. Takes already-clean
// student objects (as getStudents() returns them) rather than raw sheet
// rows, since the app's own fuzzy header-matching logic (which handles
// things like "App No" vs "Application Number" vs "Roll No" all meaning
// the same field) already lives in core.js - reusing it here avoids a
// second, separate implementation of that same matching logic drifting
// out of sync with the real one.
async function ensureStudentRecordTable(client) {
  await client.query(`DROP TABLE IF EXISTS sheet_backup_student_record`);
  await client.query(`
    CREATE TABLE sheet_backup_student_record (
      id TEXT PRIMARY KEY,
      name TEXT,
      app_no TEXT,
      reg_no TEXT,
      machine_code TEXT,
      site_code TEXT,
      student_type TEXT,
      status TEXT,
      locked BOOLEAN,
      verified_by TEXT,
      verified_at TEXT,
      first_verified_by TEXT,
      first_verified_at TEXT,
      notes TEXT,
      photo_url TEXT,
      academic_year TEXT,
      backed_up_at TIMESTAMPTZ DEFAULT now()
    )
  `);
}

async function backupStudentRecord(client, students) {
  await ensureStudentRecordTable(client);
  if (!students.length) return 0;

  const CHUNK_SIZE = 500;
  let inserted = 0;
  for (let i = 0; i < students.length; i += CHUNK_SIZE) {
    const chunk = students.slice(i, i + CHUNK_SIZE);
    const values = [];
    const placeholders = chunk.map((s, idx) => {
      const base = idx * 16;
      values.push(
        s.id, s.name || '', s.appNo || '', s.regNo || '', s.machineCode || '',
        s.siteCode || '', s.studentType || '', s.status || '', !!s.locked,
        s.verifiedBy || '', s.verifiedAt || '', s.firstVerifiedBy || '',
        s.firstVerifiedAt || '', s.notes || '', s.photoUrl || '', s.academicYear || ''
      );
      return `($${base+1},$${base+2},$${base+3},$${base+4},$${base+5},$${base+6},$${base+7},$${base+8},$${base+9},$${base+10},$${base+11},$${base+12},$${base+13},$${base+14},$${base+15},$${base+16})`;
    }).join(',');
    await client.query(
      `INSERT INTO sheet_backup_student_record
        (id, name, app_no, reg_no, machine_code, site_code, student_type, status, locked, verified_by, verified_at, first_verified_by, first_verified_at, notes, photo_url, academic_year)
       VALUES ${placeholders}
       ON CONFLICT (id) DO UPDATE SET
         name = EXCLUDED.name, app_no = EXCLUDED.app_no, reg_no = EXCLUDED.reg_no,
         machine_code = EXCLUDED.machine_code, site_code = EXCLUDED.site_code, student_type = EXCLUDED.student_type,
         status = EXCLUDED.status, locked = EXCLUDED.locked, verified_by = EXCLUDED.verified_by,
         verified_at = EXCLUDED.verified_at, first_verified_by = EXCLUDED.first_verified_by,
         first_verified_at = EXCLUDED.first_verified_at, notes = EXCLUDED.notes, photo_url = EXCLUDED.photo_url,
         academic_year = EXCLUDED.academic_year, backed_up_at = now()`,
      values
    );
    inserted += chunk.length;
  }
  return inserted;
}

async function backupUserAccounts(client, headers, rows) {
  await ensureUserAccountsTable(client);
  if (!rows.length) return 0;

  const idx = {
    userId: findColumnIndex(headers, 'User ID'),
    password: findColumnIndex(headers, 'Password'),
    role: findColumnIndex(headers, 'Role'),
    name: findColumnIndex(headers, 'Name'),
    mobile: findColumnIndex(headers, 'Mobile No'),
    email: findColumnIndex(headers, 'Email'),
    school: findColumnIndex(headers, 'School'),
    permissions: findColumnIndex(headers, 'Permissions (JSON)'),
    mustChange: findColumnIndex(headers, 'Must Change Password'),
    demoShown: findColumnIndex(headers, 'Demo Welcome Shown')
  };

  const CHUNK_SIZE = 500;
  let inserted = 0;
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE).filter(row => idx.userId > -1 && String(row[idx.userId] || '').trim());
    if (!chunk.length) continue;
    const values = [];
    const placeholders = chunk.map((row, rowIdx) => {
      const base = rowIdx * 10;
      let permsJson = null;
      const rawPerms = idx.permissions > -1 ? String(row[idx.permissions] || '').trim() : '';
      if (rawPerms) { try { JSON.parse(rawPerms); permsJson = rawPerms; } catch (e) { /* leave null if malformed */ } }
      values.push(
        String(row[idx.userId] || '').trim(),
        idx.password > -1 ? String(row[idx.password] || '') : '',
        idx.role > -1 ? String(row[idx.role] || 'staff') : 'staff',
        idx.name > -1 ? String(row[idx.name] || '') : '',
        idx.mobile > -1 ? String(row[idx.mobile] || '') : '',
        idx.email > -1 ? String(row[idx.email] || '') : '',
        idx.school > -1 ? String(row[idx.school] || '') : '',
        permsJson,
        idx.mustChange > -1 ? String(row[idx.mustChange] || '').trim().toLowerCase() === 'yes' : false,
        idx.demoShown > -1 ? String(row[idx.demoShown] || '').trim().toLowerCase() === 'yes' : false
      );
      return `($${base+1},$${base+2},$${base+3},$${base+4},$${base+5},$${base+6},$${base+7},$${base+8}::jsonb,$${base+9},$${base+10})`;
    }).join(',');
    if (!placeholders) continue;
    await client.query(
      `INSERT INTO sheet_backup_user_accounts
        (user_id, password, role, name, mobile_no, email, school, permissions, must_change_password, demo_welcome_shown)
       VALUES ${placeholders}
       ON CONFLICT (user_id) DO UPDATE SET
         password = EXCLUDED.password, role = EXCLUDED.role, name = EXCLUDED.name,
         mobile_no = EXCLUDED.mobile_no, email = EXCLUDED.email, school = EXCLUDED.school,
         permissions = EXCLUDED.permissions, must_change_password = EXCLUDED.must_change_password,
         demo_welcome_shown = EXCLUDED.demo_welcome_shown, backed_up_at = now()`,
      values
    );
    inserted += chunk.length;
  }
  return inserted;
}

async function insertRowsBatch(client, tableName, columnNames, rows) {
  const CHUNK_SIZE = 500;
  const colCount = columnNames.length;
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const values = [];
    const placeholders = chunk.map((row, idx) => {
      const base = idx * (colCount + 1);
      values.push(i + idx + 1);
      for (let c = 0; c < colCount; c++) {
        values.push(row[c] ?? '');
      }
      const placeholderNums = [];
      for (let c = 0; c <= colCount; c++) placeholderNums.push(`$${base + c + 1}`);
      return `(${placeholderNums.join(',')})`;
    }).join(',');
    await client.query(
      `INSERT INTO ${tableName} (row_index, ${columnNames.join(',')}) VALUES ${placeholders}`,
      values
    );
  }
}

// sheetsData: array of { sheetName, headers, rows } - one entry per sheet
// the person chose to back up. Returns per-sheet counts so the caller can
// report exactly what was copied.
async function backupSheetsToNeon(sheetsData, studentRecordData) {
  const dbPool = getPool();
  if (!dbPool) {
    return { ok: false, error: 'Database backup is not configured yet. Add NEON_DATABASE_URL to enable this.' };
  }

  const client = await dbPool.connect();
  const results = [];
  try {
    await client.query('BEGIN');

    if (studentRecordData) {
      const count = await backupStudentRecord(client, studentRecordData.students);
      results.push({ sheetName: studentRecordData.sheetName, rowCount: count });
    }

    for (const sheet of sheetsData) {
      const isUserAccounts = String(sheet.sheetName || '').trim().toLowerCase() === 'user accounts';
      if (isUserAccounts) {
        const count = await backupUserAccounts(client, sheet.headers, sheet.rows);
        results.push({ sheetName: sheet.sheetName, rowCount: count });
        continue;
      }
      const tableName = tableNameForSheet(sheet.sheetName);
      const columnNames = sanitizeColumnNames(sheet.headers);
      await ensureTableForSheet(client, tableName, columnNames);
      if (sheet.rows.length) {
        await insertRowsBatch(client, tableName, columnNames, sheet.rows);
      }
      results.push({ sheetName: sheet.sheetName, rowCount: sheet.rows.length });
    }
    await client.query('COMMIT');
    return { ok: true, results };
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    return { ok: false, error: err.message };
  } finally {
    client.release();
  }
}

module.exports = { backupSheetsToNeon };
