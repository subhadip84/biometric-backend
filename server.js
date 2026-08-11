// server.js — the main backend server, replacing Google Apps Script.
// Talks to your Google Sheet via a Service Account (no per-user OAuth needed).

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const core = require('./core');
const cron = require('node-cron');
const { runDailyBackup } = require('./backup');

// Runs every day at 2:00 AM India time, matching the schedule the original
// Apps Script version used. node-cron handles the timezone conversion.
cron.schedule('0 2 * * *', () => {
  console.log('Running scheduled daily backup...');
  runDailyBackup();
}, { timezone: 'Asia/Kolkata' });

// Runs every 5 minutes, checking for sessions that stopped sending
// heartbeats (browser crash, force-quit, power loss, etc.) and logs their
// end properly, rather than leaving it silently unrecorded.
cron.schedule('*/5 * * * *', () => {
  core.checkStaleSessions().catch(err => console.error('Stale session check failed:', err.message));
});

const app = express();
const PORT = process.env.PORT || 3000;

// CORS: allow your GitHub Pages frontend (and localhost for testing) to call this API.
// Set ALLOWED_ORIGIN in your environment to your actual GitHub Pages URL for safety;
// falls back to allowing any origin if not set, so it works immediately either way.
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || '*';
app.use(cors({ origin: ALLOWED_ORIGIN }));

app.use(express.text({ type: '*/*', limit: '50mb' })); // frontend sends text/plain to avoid CORS preflight; 50mb accommodates large bulk imports

// ---------- Phase 1, 2, and 3 all wired in now ----------
const API_FUNCTIONS = {
  // Phase 1
  checkLogin: core.checkLogin,
  getStudents: core.getStudents,
  updateStatus: core.updateStatus,
  adminUnlock: core.adminUnlock,
  adminLockAllDone: core.adminLockAllDone,
  getActivityLog: core.getActivityLog,
  // Phase 2
  checkUserIdAvailability: core.checkUserIdAvailability,
  checkContactAvailability: core.checkContactAvailability,
  createUser: core.createUser,
  deleteUser: core.deleteUser,
  updateUserDetails: core.updateUserDetails,
  getUserList: core.getUserList,
  changeOwnPassword: core.changeOwnPassword,
  adminResetPassword: core.adminResetPassword,
  changeAdminPassword: core.changeAdminPassword,
  deleteStudent: core.deleteStudent,
  getDistinctSiteCodes: core.getDistinctSiteCodes,
  validateImportRows: core.validateImportRows,
  importNewStudents: core.importNewStudents,
  importVerificationUpdates: core.importVerificationUpdates,
  // Phase 3
  getAnnouncements: core.getAnnouncements,
  publishAnnouncement: core.publishAnnouncement,
  updateAnnouncement: core.updateAnnouncement,
  deleteAnnouncement: core.deleteAnnouncement,
  logHelpQuestion: core.logHelpQuestion,
  getHelpQuestionStats: core.getHelpQuestionStats,
  logSessionIp: core.logSessionIp,
  logSessionEnd: core.logSessionEnd,
  recordHeartbeat: core.recordHeartbeat,
  clearActiveSession: core.clearActiveSession,
  exportRosterAsCsv: core.exportRosterAsCsv,
  getLastImportInfo: core.getLastImportInfo,
  getTodayImportCount: core.getTodayImportCount,
  getLastImportTimestamp: core.getLastImportTimestamp,
  getHostelData: core.getHostelData,
  updateHostelStatus: core.updateHostelStatus,
  adminUnlockHostel: core.adminUnlockHostel,
  deleteHostelStudent: core.deleteHostelStudent,
  exportHostelAsCsv: core.exportHostelAsCsv,
  exportHostelVerifiedTodayAsCsv: core.exportHostelVerifiedTodayAsCsv,
  importNewHostelData: core.importNewHostelData,
  importHostelVerificationUpdates: core.importHostelVerificationUpdates,
  getLastHostelImportInfo: core.getLastHostelImportInfo,
  triggerBackupNow: async () => { await runDailyBackup(); return { ok: true, message: 'Backup triggered.' }; }
};

app.post('/', async (req, res) => {
  try {
    const body = JSON.parse(req.body || '{}');
    const fn = API_FUNCTIONS[body.fn];
    if (!fn) {
      return res.json({ ok: false, error: `Not yet available in the new backend: ${body.fn} (coming in a follow-up update)` });
    }
    const result = await fn.apply(null, body.args || []);
    res.json(result);
  } catch (err) {
    res.json({ ok: false, error: 'Server error: ' + err.message });
  }
});

app.get('/', (req, res) => {
  res.send('Biometric Verification Desk backend is running. Send POST requests with { fn, args }.');
});

// Catch-all for any route/method not otherwise handled - JSON, not HTML
app.use((req, res) => {
  res.status(404).json({ ok: false, error: 'Not found. This backend only responds to POST requests.' });
});

// Global error handler - guarantees JSON even if something fails before
// reaching our own try/catch (e.g. a body-parsing error in middleware).
app.use((err, req, res, next) => {
  res.status(500).json({ ok: false, error: 'Unexpected server error: ' + (err && err.message ? err.message : String(err)) });
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
