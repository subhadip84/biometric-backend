// backup.js — daily automatic backup of the Google Sheet, delivered as an
// email attachment via Resend's HTTPS API (not raw SMTP - Render's free
// tier blocks outbound SMTP ports like 587, so a plain SMTP connection
// never completes regardless of how correct the credentials are. Resend
// sends over regular HTTPS, same as any website request, which isn't
// blocked).

const { google } = require('googleapis');
const sheetsApi = require('./sheets');
const core = require('./core');

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

async function sendBackupEmail({ apiKey, to, subject, text, fileName, fileBuffer }) {
  const response = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      from: 'Biometric Verification Desk <onboarding@resend.dev>',
      to: [to],
      subject,
      text,
      attachments: [{ filename: fileName, content: fileBuffer.toString('base64') }]
    })
  });
  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`Resend API error (${response.status}): ${errorBody}`);
  }
  return response.json();
}

async function runDailyBackup() {
  try {
    const apiKey = (process.env.RESEND_API_KEY || '').trim();
    const recipient = (process.env.BACKUP_EMAIL_TO || '').trim();
    if (!apiKey) throw new Error('RESEND_API_KEY is not set in Render.');
    if (!recipient) throw new Error('BACKUP_EMAIL_TO is not set in Render - no address to send the backup to.');

    const drive = await getDriveClient();
    const spreadsheetMeta = await sheetsApi.withQuotaRetry(() =>
      drive.files.get({ fileId: sheetsApi.SPREADSHEET_ID, fields: 'name' })
    );
    const sheetName = spreadsheetMeta.data.name;

    // Exports the spreadsheet as XLSX bytes directly - a read-only operation
    // that never creates any file in Drive, so no storage quota is touched.
    const exportResponse = await sheetsApi.withQuotaRetry(() =>
      drive.files.export(
        { fileId: sheetsApi.SPREADSHEET_ID, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
        { responseType: 'arraybuffer' }
      )
    );
    const fileBuffer = Buffer.from(exportResponse.data);

    const now = new Date();
    const dateStr = getIndiaTimestamp(now);
    const fileName = `${sheetName} - Backup ${dateStr}.xlsx`;

    await sendBackupEmail({
      apiKey,
      to: recipient,
      subject: `Daily Backup - ${sheetName} - ${dateStr}`,
      text: `Attached is the daily backup of "${sheetName}", generated ${dateStr} (India time).`,
      fileName,
      fileBuffer
    });

    await core.logActivity('System', 'Daily Backup Completed', `${fileName} emailed to ${recipient}`);
    console.log('Daily backup emailed:', fileName);
  } catch (err) {
    console.error('Daily backup failed:', err.message);
    try { await core.logActivity('System', 'Daily Backup Failed', err.message); } catch (e) { /* ignore */ }
  }
}

module.exports = { runDailyBackup };
