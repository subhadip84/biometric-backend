// server.js — the main backend server, replacing Google Apps Script.
// Talks to your Google Sheet via a Service Account (no per-user OAuth needed).

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const core = require('./core');

const app = express();
const PORT = process.env.PORT || 3000;

// CORS: allow your GitHub Pages frontend (and localhost for testing) to call this API.
// Set ALLOWED_ORIGIN in your environment to your actual GitHub Pages URL for safety;
// falls back to allowing any origin if not set, so it works immediately either way.
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || '*';
app.use(cors({ origin: ALLOWED_ORIGIN }));

app.use(express.text({ type: '*/*' })); // frontend sends text/plain to avoid CORS preflight

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
  exportRosterAsCsv: core.exportRosterAsCsv,
  getLastImportInfo: core.getLastImportInfo,
  getTodayImportCount: core.getTodayImportCount,
  getLastImportTimestamp: core.getLastImportTimestamp
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

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
