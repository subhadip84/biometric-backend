// sheets.js — Google Sheets API connection layer, using a service account.
// Replaces Apps Script's SpreadsheetApp with equivalent Sheets API v4 calls.

const { google } = require('googleapis');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const MASTER_SHEET_NAME = process.env.MASTER_SHEET_NAME || null; // null = use first sheet
const USER_SHEET_NAME = 'User Accounts';
const ACTIVITY_LOG_SHEET_NAME = 'Activity Log';

function getCredentials() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set.');
  return JSON.parse(raw);
}

let authClientPromise = null;
function getAuthClient() {
  if (!authClientPromise) {
    const credentials = getCredentials();
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
      ]
    });
    authClientPromise = auth.getClient();
  }
  return authClientPromise;
}

let sheetsClientPromise = null;
function getSheetsClient() {
  if (!sheetsClientPromise) {
    sheetsClientPromise = getAuthClient().then(authClient => google.sheets({ version: 'v4', auth: authClient }));
  }
  return sheetsClientPromise;
}

// Always fetches fresh sheet titles - no caching, so tab renames (like the
// master sheet name) are picked up immediately without needing a server
// restart. This is a lightweight metadata call, not a performance concern
// for a low-traffic internal tool.
async function getSheetTitles() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  return meta.data.sheets.map(s => s.properties.title);
}

async function getMasterSheetName() {
  if (MASTER_SHEET_NAME) return MASTER_SHEET_NAME;
  const titles = await getSheetTitles();
  return titles[0];
}

async function sheetExists(name) {
  const titles = await getSheetTitles();
  return titles.indexOf(name) !== -1;
}

async function createSheet(name, headerRow) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: name } } }] }
  });
  if (headerRow && headerRow.length) {
    await writeRange(`${name}!A1`, [headerRow]);
  }
}

async function ensureSheet(name, headerRow) {
  if (!(await sheetExists(name))) {
    await createSheet(name, headerRow);
  }
}

async function readRange(rangeA1) {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: rangeA1
  });
  return res.data.values || [];
}

async function writeRange(rangeA1, values) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: rangeA1,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values }
  });
}

// Combines many individual cell/range writes into ONE API call, instead of
// one call per write. This is essential to avoid hitting Google Sheets API's
// per-minute write quota when updating several cells at once (e.g. marking
// a student verified touches 4-6 cells; bulk operations touch many rows).
async function batchWriteRanges(updates) {
  if (!updates.length) return;
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: updates.map(u => ({ range: u.range, values: u.values }))
    }
  });
}

async function appendRows(sheetName, values) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values }
  });
}

async function clearRange(rangeA1) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: rangeA1
  });
}

async function getSheetIdByName(name) {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.title === name);
  return sheet ? sheet.properties.sheetId : null;
}

async function deleteRow(sheetName, rowNumber1Indexed) {
  const sheetId = await getSheetIdByName(sheetName);
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId,
            dimension: 'ROWS',
            startIndex: rowNumber1Indexed - 1,
            endIndex: rowNumber1Indexed
          }
        }
      }]
    }
  });
}

module.exports = {
  getMasterSheetName,
  ensureSheet,
  readRange,
  writeRange,
  batchWriteRanges,
  appendRows,
  clearRange,
  deleteRow,
  getAuthClient,
  SPREADSHEET_ID,
  USER_SHEET_NAME,
  ACTIVITY_LOG_SHEET_NAME
};
