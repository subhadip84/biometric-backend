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

async function insertRowsBatch(client, tableName, headers, rows) {
  const CHUNK_SIZE = 500;
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const values = [];
    const placeholders = chunk.map((row, idx) => {
      const rowObj = {};
      headers.forEach((h, colIdx) => { rowObj[h || `col_${colIdx}`] = row[colIdx] ?? ''; });
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
