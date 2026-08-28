// Backs up any chosen sheet(s) from the spreadsheet into a Neon (Postgres)
// database as a secondary copy. Google Sheets remains the live, primary
// data source for the app at all times - this is a one-way, on-demand
// snapshot copy, not a live sync. Each run fully replaces the backup
// table's contents for each sheet with whatever is currently in that
// sheet, so the backup always reflects the most recent snapshot taken,
// not a running history.
//
// Every sheet - including Student Record and User Accounts - is backed up
// through the same generic path: a table whose columns are derived
// directly from that sheet's own header row at the moment of backup. This
// means any column added to any sheet in the future is picked up
// automatically the next time a backup runs, with no code change ever
// needed here. The trade-off, deliberately accepted, is that every column
// is stored as plain text rather than a more specific type (e.g. a real
// boolean or JSON column) - genuinely automatic sync was prioritized over
// precise typing.

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
// the person chose to back up. Every sheet goes through the same generic
// path now, so any column present in a sheet's header row at backup time
// automatically becomes a column in its table - no special-casing, no
// separate schema to remember to update.
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
