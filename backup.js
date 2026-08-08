// backup.js — daily automatic backup of the Google Sheet, delivered as an
// email attachment rather than a Drive file copy. This avoids the Service
// Account's own (effectively zero) Drive storage quota entirely, since
// exporting a spreadsheet as bytes is a read-only operation - no new file
// is ever created in Drive, so there's nothing to run out of space for.

const { google } = require('googleapis');
const nodemailer = require('nodemailer');
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

function getMailTransporter() {
  const host = process.env.SMTP_HOST;
  const port = Number(process.env.SMTP_PORT || 587);
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  if (!host || !user || !pass) {
    throw new Error('Email backup is not configured - SMTP_HOST, SMTP_USER, and SMTP_PASS must be set in Render.');
  }
  return nodemailer.createTransport({
    host, port, secure: port === 465,
    auth: { user, pass }
  });
}

async function runDailyBackup() {
  try {
    const recipient = process.env.BACKUP_EMAIL_TO;
    if (!recipient) {
      throw new Error('BACKUP_EMAIL_TO is not set in Render - no address to send the backup to.');
    }

    const drive = await getDriveClient();
    const spreadsheetMeta = await drive.files.get({ fileId: sheetsApi.SPREADSHEET_ID, fields: 'name' });
    const sheetName = spreadsheetMeta.data.name;

    // Exports the spreadsheet as XLSX bytes directly - a read-only operation
    // that never creates any file in Drive, so no storage quota is touched.
    const exportResponse = await drive.files.export(
      { fileId: sheetsApi.SPREADSHEET_ID, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      { responseType: 'arraybuffer' }
    );
    const fileBuffer = Buffer.from(exportResponse.data);

    const now = new Date();
    const dateStr = getIndiaTimestamp(now);
    const fileName = `${sheetName} - Backup ${dateStr}.xlsx`;

    const transporter = getMailTransporter();
    await transporter.sendMail({
      from: process.env.SMTP_USER,
      to: recipient,
      subject: `Daily Backup - ${sheetName} - ${dateStr}`,
      text: `Attached is the daily backup of "${sheetName}", generated ${dateStr} (India time).`,
      attachments: [{ filename: fileName, content: fileBuffer }]
    });

    await core.logActivity('System', 'Daily Backup Completed', `${fileName} emailed to ${recipient}`);
    console.log('Daily backup emailed:', fileName);
  } catch (err) {
    console.error('Daily backup failed:', err.message);
    try { await core.logActivity('System', 'Daily Backup Failed', err.message); } catch (e) { /* ignore */ }
  }
}

module.exports = { runDailyBackup };
