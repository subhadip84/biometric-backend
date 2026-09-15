// server.js — the main backend server, replacing Google Apps Script.
// Talks to your Google Sheet via a Service Account (no per-user OAuth needed).

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const core = require('./core');
const cron = require('node-cron');
const { runDailyBackup } = require('./backup');
const http = require('http');
const { WebSocketServer } = require('ws');

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
app.set('trust proxy', true);
const PORT = process.env.PORT || 3000;

// CORS: allow your GitHub Pages frontend (and localhost for testing) to call this API.
// Set ALLOWED_ORIGIN in your environment to your actual GitHub Pages URL for safety;
// falls back to allowing any origin if not set, so it works immediately either way.
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || '*';
app.use(cors({ origin: ALLOWED_ORIGIN }));

app.use(express.text({ type: '*/*', limit: '50mb' })); // frontend sends text/plain to avoid CORS preflight; 50mb accommodates large bulk imports

// Maps a userId to the set of their currently-open WebSocket connections
// (a person could have more than one tab/device open). Used so a chat
// message can be pushed to the recipient the instant it's sent, with no
// polling delay, if they're currently connected - falls back to the
// existing polling-based unread badge/conversation list if they're not.
const chatConnections = new Map();

function pushChatMessageToUser(userId, message){
  const sockets = chatConnections.get(String(userId));
  if(!sockets || !sockets.size) return;
  const payload = JSON.stringify({ type: 'chatMessage', message });
  sockets.forEach(ws => {
    if(ws.readyState === ws.OPEN){
      try{ ws.send(payload); }catch(e){ /* connection likely stale; cleanup happens on close */ }
    }
  });
}

function pushReadReceiptToUser(userId, readByUserId){
  const sockets = chatConnections.get(String(userId));
  if(!sockets || !sockets.size) return;
  const payload = JSON.stringify({ type: 'read', from: readByUserId });
  sockets.forEach(ws => {
    if(ws.readyState === ws.OPEN){
      try{ ws.send(payload); }catch(e){ /* connection likely stale; cleanup happens on close */ }
    }
  });
}

// ---------- Phase 1, 2, and 3 all wired in now ----------
const API_FUNCTIONS = {
  // Phase 1
  checkLogin: core.checkLogin,
  logoutSession: core.logoutSession,
  getStudents: core.getStudents,
  updateStatus: core.updateStatus,
  adminUnlock: core.adminUnlock,
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
  updateStudentDetails: core.updateStudentDetails,
  setStudentActiveStatus: core.setStudentActiveStatus,
  bulkSetActiveStatus: core.bulkSetActiveStatus,
  getDistinctSiteCodes: core.getDistinctSiteCodes,
  validateImportRows: core.validateImportRows,
  importNewStudents: core.importNewStudents,
  importVerificationUpdates: core.importVerificationUpdates,
  importInactiveList: core.importInactiveList,
  // Phase 3
  getAnnouncements: core.getAnnouncements,
  publishAnnouncement: core.publishAnnouncement,
  updateAnnouncement: core.updateAnnouncement,
  deleteAnnouncement: core.deleteAnnouncement,
  logHelpQuestion: core.logHelpQuestion,
  askAiHelpAssistant: core.askAiHelpAssistant,
  getHelpFaqList: core.getHelpFaqList,
  logHelpChatEvent: core.logHelpChatEvent,
  getUnusualActivityFlags: core.getUnusualActivityFlags,
  parseVoiceCommand: core.parseVoiceCommand,
  getOnlineUsers: core.getOnlineUsers,
  sendChatMessage: async (toUserId, text, sessionToken) => {
    const result = await core.sendChatMessage(toUserId, text, sessionToken);
    if(result.ok) pushChatMessageToUser(toUserId, result.message);
    return result;
  },
  getChatMessages: async (otherUserId, sessionToken) => {
    const result = await core.getChatMessages(otherUserId, sessionToken);
    if(result.ok && result.didMarkRead){
      const session = core.validateSessionToken(sessionToken);
      if(session) pushReadReceiptToUser(otherUserId, session.userId);
    }
    return result;
  },
  getChatConversations: core.getChatConversations,
  getStaffLeaderboard: core.getStaffLeaderboard,
  getLoginDigest: core.getLoginDigest,
  undoRecentVerification: core.undoRecentVerification,
  undoRecentHostelVerification: core.undoRecentHostelVerification,
  publicLookupStudent: core.publicLookupStudent,
  publicLookupHostelStudent: core.publicLookupHostelStudent,
  parseRosterFilterQuery: core.parseRosterFilterQuery,
  explainUnusualActivity: core.explainUnusualActivity,
  generateShiftHandoffNote: core.generateShiftHandoffNote,
  findDuplicateStudents: core.findDuplicateStudents,
  findDuplicateHostelStudents: core.findDuplicateHostelStudents,
  getReportTemplates: core.getReportTemplates,
  saveReportTemplate: core.saveReportTemplate,
  deleteReportTemplate: core.deleteReportTemplate,
  backupToNeonDatabase: core.backupToNeonDatabase,
  getAvailableSheetsForBackup: core.getAvailableSheetsForBackup,
  getStudentAuditTrail: core.getStudentAuditTrail,
  findInconsistentLockStatus: core.findInconsistentLockStatus,
  fixInconsistentLockStatus: core.fixInconsistentLockStatus,
  findInconsistentHostelLockStatus: core.findInconsistentHostelLockStatus,
  fixInconsistentHostelLockStatus: core.fixInconsistentHostelLockStatus,
  bulkUpdateStatus: core.bulkUpdateStatus,
  updateStudentNote: core.updateStudentNote,
  bulkUpdateHostelStatus: core.bulkUpdateHostelStatus,
  updateHostelStudentNote: core.updateHostelStudentNote,
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
  updateHostelStudentDetails: core.updateHostelStudentDetails,
  setHostelStudentActiveStatus: core.setHostelStudentActiveStatus,
  bulkSetHostelActiveStatus: core.bulkSetHostelActiveStatus,
  exportHostelAsCsv: core.exportHostelAsCsv,
  exportHostelVerifiedTodayAsCsv: core.exportHostelVerifiedTodayAsCsv,
  exportHostelVerifiedLastDayAsCsv: core.exportHostelVerifiedLastDayAsCsv,
  importNewHostelData: core.importNewHostelData,
  importHostelVerificationUpdates: core.importHostelVerificationUpdates,
  importHostelInactiveList: core.importHostelInactiveList,
  getLastHostelImportInfo: core.getLastHostelImportInfo,
  triggerBackupNow: async () => { await runDailyBackup(); return { ok: true, message: 'Backup triggered.' }; }
};

app.post('/', async (req, res) => {
  const clientIp = req.ip || 'unknown';
  core.requestContext.run({ ip: clientIp }, async () => {
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

const server = http.createServer(app);

// WebSocket server for instant chat delivery, on its own path so it never
// interferes with the existing POST '/' dispatcher the rest of the app uses.
const wss = new WebSocketServer({ server, path: '/ws' });

wss.on('connection', (ws, req) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const token = url.searchParams.get('token');
  const session = token ? core.validateSessionToken(token) : null;
  if(!session){
    ws.close(4001, 'Invalid or expired session');
    return;
  }
  const userId = session.userId;
  if(!chatConnections.has(userId)) chatConnections.set(userId, new Set());
  chatConnections.get(userId).add(ws);

  // Heartbeat: a network intermediary (proxy, load balancer) can silently
  // drop an idle connection without either side's normal close/error events
  // firing - it looks perfectly "open" on both ends while actually being
  // dead, which is exactly what would make a push fail invisibly and fall
  // all the way back to the slow poll. isAlive gets reset to true only when
  // a pong is actually received; if a connection didn't answer the previous
  // ping by the time the next one is due, it's genuinely gone and gets
  // dropped immediately, so a real reconnect (and real delivery) can happen
  // right away instead of the client believing a dead connection is fine.
  ws.isAlive = true;
  ws.on('pong', () => { ws.isAlive = true; });

  ws.on('close', () => {
    const sockets = chatConnections.get(userId);
    if(sockets){
      sockets.delete(ws);
      if(!sockets.size) chatConnections.delete(userId);
    }
  });

  // The client can send small signals of its own now (currently just
  // "typing"), relayed straight to the other person's live connection if
  // they have one - this never touches Sheets, since a typing indicator is
  // purely ephemeral and has no reason to be persisted or polled for.
  ws.on('message', (data) => {
    let parsed;
    try{ parsed = JSON.parse(data.toString()); }catch(e){ return; }
    if(parsed && parsed.type === 'typing' && parsed.to){
      const sockets = chatConnections.get(String(parsed.to));
      if(!sockets) return;
      const payload = JSON.stringify({ type: 'typing', from: userId });
      sockets.forEach(recipientWs => {
        if(recipientWs.readyState === recipientWs.OPEN){
          try{ recipientWs.send(payload); }catch(e){ /* stale connection; heartbeat will clean it up */ }
        }
      });
    }
  });

  ws.on('error', () => { /* connection will close and clean itself up */ });
});

// Runs the ping/pong check across every connected client on a fixed
// interval - independent of any individual connection's own state.
setInterval(() => {
  wss.clients.forEach(ws => {
    if(ws.isAlive === false) return ws.terminate();
    ws.isAlive = false;
    ws.ping();
  });
}, 25000);

(async () => {
  // Restores sessions that were still active before this restart, so
  // redeploying (or a free-tier spin-down/up) doesn't force everyone
  // using the app right now to log back in.
  await core.rehydrateSessionsFromStorage();
  server.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
})();
