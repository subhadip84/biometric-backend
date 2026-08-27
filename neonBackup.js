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

async function ensureTableForSheet(client, tableName) {
  await client.query(`
    CREATE TABLE IF NOT EXISTS ${tableName} (
      row_index INTEGER PRIMARY KEY,
      row_data JSONB,
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

async function insertRowsBatch(client, tableName, headers, rows) {
  const CHUNK_SIZE = 500;
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const values = [];
    const placeholders = chunk.map((row, idx) => {
      const rowObj = {};
      // Iterate over whichever is longer - headers or the row itself - so a
      // row with more populated cells than there are header columns still
      // has every one of its values captured, not silently dropped.
      const colCount = Math.max(headers.length, row.length);
      for (let colIdx = 0; colIdx < colCount; colIdx++) {
        const key = headers[colIdx] || `col_${colIdx}`;
        rowObj[key] = row[colIdx] ?? '';
      }
      const base = idx * 2;
      values.push(i + idx + 1, JSON.stringify(rowObj));
      return `($${base + 1}, $${base + 2}::jsonb)`;
    }).join(',');
    await client.query(
      `INSERT INTO ${tableName} (row_index, row_data) VALUES ${placeholders}`,
      values
    );
  }
}

// sheetsData: array of { sheetName, headers, rows } - one entry per sheet
// the person chose to back up. Returns per-sheet counts so the caller can
// report exactly what was copied.
async function backupSheetsToNeon(sheetsData) {
  const dbPool = getPool();
  if (!dbPool) {
    return { ok: false, error: 'Database backup is not configured yet. Add NEON_DATABASE_URL to enable this.' };
  }

  const client = await dbPool.connect();
  const results = [];
  try {
    await client.query('BEGIN');
    for (const sheet of sheetsData) {
      const isUserAccounts = String(sheet.sheetName || '').trim().toLowerCase() === 'user accounts';
      if (isUserAccounts) {
        const count = await backupUserAccounts(client, sheet.headers, sheet.rows);
        results.push({ sheetName: sheet.sheetName, rowCount: count });
        continue;
      }
      const tableName = tableNameForSheet(sheet.sheetName);
      await ensureTableForSheet(client, tableName);
      await client.query(`TRUNCATE ${tableName}`);
      if (sheet.rows.length) {
        await insertRowsBatch(client, tableName, sheet.headers, sheet.rows);
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
