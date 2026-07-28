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

// ---------- Phase 1 functions are wired in now; Phase 2/3 will add more here ----------
const API_FUNCTIONS = {
  checkLogin: core.checkLogin,
  getStudents: core.getStudents,
  updateStatus: core.updateStatus,
  adminUnlock: core.adminUnlock,
  adminLockAllDone: core.adminLockAllDone,
  getActivityLog: core.getActivityLog
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
