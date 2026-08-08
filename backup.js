// backup.js — daily automatic backup of the Google Sheet, replacing what
// used to run via Apps Script's time-driven trigger + DriveApp.makeCopy().

const { google } = require('googleapis');
const sheetsApi = require('./sheets');
const core = require('./core');

const BACKUP_FOLDER_NAME = 'Biometric Verification Backups';
const BACKUP_RETENTION_DAYS = 7;

function getIndiaTimestamp(date) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hour12: false
  }).formatToParts(date);
  const get = type => parts.find(p => p.type === type).value;
  return `${get('year')}-${get('month')}-${get('day')}_${get('hour')}-${get('minute')}`;
}

async function getDriveClient() {
  const authClient = await sheetsApi.getAuthClient();
  return google.drive({ version: 'v3', auth: authClient });
}

async function getOrCreateBackupFolder(drive) {
  const res = await drive.files.list({
    q: `name='${BACKUP_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id, name)',
    spaces: 'drive'
  });
  if (res.data.files && res.data.files.length) {
    return res.data.files[0].id;
  }
  const folder = await drive.files.create({
    requestBody: { name: BACKUP_FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder' },
    fields: 'id'
  });
  return folder.data.id;
}

async function cleanupOldBackups(drive, folderId) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_RETENTION_DAYS);

  const res = await drive.files.list({
    q: `'${folderId}' in parents and trashed=false`,
    fields: 'files(id, name, createdTime)',
    spaces: 'drive'
  });

  const files = res.data.files || [];
  if (!files.length) return 0;

  // Sort newest first, so we can always identify and protect the single
  // most recent backup - it must never be deleted, even if it happens to
  // be older than the retention window (e.g. if backups were interrupted
  // for a while, we never want to end up with zero backups at all).
  files.sort((a, b) => new Date(b.createdTime) - new Date(a.createdTime));
  const mostRecent = files[0];

  let deletedCount = 0;
  for (const file of files) {
    if (file.id === mostRecent.id) continue; // never delete the latest one
    const created = new Date(file.createdTime);
    if (created < cutoff) {
      await drive.files.delete({ fileId: file.id }); // permanent delete, not trash - trashed files still count against storage quota
      deletedCount++;
    }
  }
  return deletedCount;
}

async function emptyTrashedBackups(drive, folderId) {
  // One-time recovery helper: permanently deletes any already-trashed files
  // in the backup folder, needed because the old cleanup logic only moved
  // files to trash (which still counts against storage quota) instead of
  // permanently deleting them. This reclaims that space immediately.
  const res = await drive.files.list({
    q: `'${folderId}' in parents and trashed=true`,
    fields: 'files(id, name)',
    spaces: 'drive'
  });
  const files = res.data.files || [];
  for (const file of files) {
    await drive.files.delete({ fileId: file.id });
  }
  return files.length;
}

async function runAllTrashCleanup() {
  const drive = await getDriveClient();
  const folderId = await getOrCreateBackupFolder(drive);
  const count = await emptyTrashedBackups(drive, folderId);
  await core.logActivity('System', 'Emptied Trashed Backups', `${count} file(s) permanently deleted`);
  return { ok: true, count };
}

async function resetAllBackups() {
  // Deletes every existing backup file (not just trashed ones), needed to
  // fully reclaim the Service Account's exhausted storage quota. Use this
  // once, then let ownership-transfer handle future backups properly.
  const drive = await getDriveClient();
  const folderId = await getOrCreateBackupFolder(drive);
  const res = await drive.files.list({
    q: `'${folderId}' in parents and trashed=false`,
    fields: 'files(id, name)',
    spaces: 'drive'
  });
  const files = res.data.files || [];
  for (const file of files) {
    await drive.files.delete({ fileId: file.id });
  }
  await core.logActivity('System', 'Reset All Backups', `${files.length} file(s) permanently deleted`);
  return { ok: true, count: files.length };
}

async function runDailyBackup() {
  try {
    const drive = await getDriveClient();
    const folderId = await getOrCreateBackupFolder(drive);

    const now = new Date();
    const dateStr = getIndiaTimestamp(now);
    const spreadsheetMeta = await drive.files.get({ fileId: sheetsApi.SPREADSHEET_ID, fields: 'name' });
    const backupName = `${spreadsheetMeta.data.name} - Backup ${dateStr}`;

    await drive.files.copy({
      fileId: sheetsApi.SPREADSHEET_ID,
      requestBody: { name: backupName, parents: [folderId] }
    });

    await cleanupOldBackups(drive, folderId);
    await core.logActivity('System', 'Daily Backup Completed', backupName);
    console.log('Daily backup completed:', backupName);
  } catch (err) {
    console.error('Daily backup failed:', err.message);
    try { await core.logActivity('System', 'Daily Backup Failed', err.message); } catch (e) { /* ignore */ }
  }
}

module.exports = { runDailyBackup, runAllTrashCleanup, resetAllBackups };
