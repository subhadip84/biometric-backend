// Backs up the current roster and hostel data into a Neon (Postgres) database
// as a secondary copy. Google Sheets remains the live, primary data source
// for the app at all times - this is a one-way, on-demand snapshot copy,
// not a live sync. Each run fully replaces the backup tables' contents with
// whatever is currently in the sheets, so the backup always reflects the
// most recent snapshot taken, not a running history.

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

async function ensureTables(client) {
  await client.query(`
    CREATE TABLE IF NOT EXISTS roster_backup (
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
      backed_up_at TIMESTAMPTZ DEFAULT now()
    )
  `);
  await client.query(`
    CREATE TABLE IF NOT EXISTS hostel_backup (
      id TEXT PRIMARY KEY,
      student_name TEXT,
      application_no TEXT,
      registration_no TEXT,
      machine_code TEXT,
      hostel_name TEXT,
      room_no TEXT,
      status TEXT,
      locked BOOLEAN,
      verified_by TEXT,
      verified_at TEXT,
      first_verified_by TEXT,
      first_verified_at TEXT,
      notes TEXT,
      photo_url TEXT,
      backed_up_at TIMESTAMPTZ DEFAULT now()
    )
  `);
}

async function insertBatch(client, tableName, columns, rows, mapRow) {
  const CHUNK_SIZE = 500;
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const values = [];
    const placeholders = chunk.map((row, idx) => {
      const mapped = mapRow(row);
      const base = idx * columns.length;
      values.push(...mapped);
      return '(' + mapped.map((_, j) => `$${base + j + 1}`).join(',') + ')';
    }).join(',');
    await client.query(
      `INSERT INTO ${tableName} (${columns.join(',')}) VALUES ${placeholders}`,
      values
    );
  }
}

async function backupToNeon(rosterStudents, hostelStudents) {
  const dbPool = getPool();
  if (!dbPool) {
    return { ok: false, error: 'Database backup is not configured yet. Add NEON_DATABASE_URL to enable this.' };
  }

  const client = await dbPool.connect();
  try {
    await ensureTables(client);
    await client.query('BEGIN');

    // Full snapshot replace - clear out the previous backup and insert the
    // current state, rather than trying to reconcile row-by-row changes.
    await client.query('TRUNCATE roster_backup');
    if (rosterStudents.length) {
      await insertBatch(
        client, 'roster_backup',
        ['id','name','app_no','reg_no','machine_code','site_code','student_type','status','locked','verified_by','verified_at','first_verified_by','first_verified_at','notes','photo_url'],
        rosterStudents,
        r => [r.id, r.name, r.appNo, r.regNo, r.machineCode, r.siteCode, r.studentType, r.status, !!r.locked, r.verifiedBy, r.verifiedAt, r.firstVerifiedBy, r.firstVerifiedAt, r.notes || '', r.photoUrl || '']
      );
    }

    await client.query('TRUNCATE hostel_backup');
    if (hostelStudents.length) {
      await insertBatch(
        client, 'hostel_backup',
        ['id','student_name','application_no','registration_no','machine_code','hostel_name','room_no','status','locked','verified_by','verified_at','first_verified_by','first_verified_at','notes','photo_url'],
        hostelStudents,
        h => [h.id, h.studentName, h.applicationNo, h.registrationNo, h.machineCode, h.hostelName, h.roomNo, h.status, !!h.locked, h.verifiedBy, h.verifiedAt, h.firstVerifiedBy, h.firstVerifiedAt, h.notes || '', h.photoUrl || '']
      );
    }

    await client.query('COMMIT');
    return { ok: true, rosterCount: rosterStudents.length, hostelCount: hostelStudents.length };
  } catch (err) {
    await client.query('ROLLBACK').catch(() => {});
    return { ok: false, error: err.message };
  } finally {
    client.release();
  }
}

module.exports = { backupToNeon };
