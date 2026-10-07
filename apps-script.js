/**
 * Google Apps Script — Sadhana Tracker (canonical backend)
 *
 * SETUP / REDEPLOY (do this whenever you change this code):
 *  1. Open your Apps Script project (script.google.com).
 *  2. Replace ALL the code in Code.gs with this entire file.
 *  3. Run setupSheet() once  (function dropdown → setupSheet → Run → allow).
 *  4. Deploy → Manage deployments → ✏️ edit the active deployment →
 *       Version: "New version" → Deploy.
 *     ⚠️ Just saving (Ctrl+S) does NOT update the live /exec URL.
 *        You MUST pick "New version" or the old code keeps running.
 *  5. The /exec URL stays the same after a "New version" redeploy — no need
 *     to touch index.html. (Only a brand-new deployment changes the URL.)
 *
 * Sheet "111" columns:  A=Date(text)  B=Cycles  C=Shambhavi  D=Bhuta  E=Devi  F=Restored
 *   Restored = comma list of practices repaired with a ticket, e.g. "shambhavi,devi"
 * Sheet "Tickets" columns:  A=Id  B=Practice  C=BoughtOn  D=UsedFor  E=UsedAt
 */

const SHEET_ID       = '1c3-6c4X9d5wWWe0d9rnwz7R_e-WYDJmpbDN6pPGgN1U';
const SHEET_NAME     = '111';
const TICKETS_SHEET  = 'Tickets';

/* ---------- READ: every logged day + every ticket ---------- */
function doGet(e) {
  try {
    const sheet = getSheet_();
    const last  = sheet.getLastRow();
    let data = [];

    if (last > 1) {
      // getDisplayValues keeps the date column as the exact text shown in the
      // cell ("2026-06-10") so there is never any timezone shifting.
      const rows = sheet.getRange(2, 1, last - 1, 6).getDisplayValues();
      data = rows
        .filter(r => String(r[0]).trim() !== '')
        .map(r => [
          String(r[0]).trim(),                                    // date  "YYYY-MM-DD"
          Number(r[1]) || 0,                                      // cycles 0–3
          r[2] === true || String(r[2]).toUpperCase() === 'TRUE', // shambhavi
          r[3] === true || String(r[3]).toUpperCase() === 'TRUE', // bhuta
          r[4] === true || String(r[4]).toUpperCase() === 'TRUE', // devi
          String(r[5] || '').trim(),                              // restored (csv)
        ]);
    }

    return json({ success: true, data, tickets: readTickets_() });
  } catch (err) {
    return json({ success: false, error: err.toString(), data: [], tickets: [] });
  }
}

/* ---------- WRITE: upsert one day, or one ticket ---------- */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'ticket') return saveTicket_(p);
    return saveDay_(p);
  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

function saveDay_(p) {
  const date      = String(p.date || '').trim();
  const cycles    = Number(p.cycles != null ? p.cycles : p.surya_cycles) || 0;
  const shambhavi = p.shambhavi === true || String(p.shambhavi).toUpperCase() === 'TRUE';
  const bhuta     = p.bhuta     === true || String(p.bhuta).toUpperCase()     === 'TRUE';
  const devi      = p.devi      === true || String(p.devi).toUpperCase()      === 'TRUE';
  const restored  = Array.isArray(p.restored) ? p.restored.join(',') : String(p.restored || '');

  if (!date) return json({ success: false, error: 'Missing date' });

  const sheet = getSheet_();
  sheet.getRange('A:A').setNumberFormat('@');  // keep dates as plain text

  const rowIdx = findRow_(sheet, date);
  const row = [date, cycles, shambhavi, bhuta, devi, restored];
  if (rowIdx === -1) sheet.appendRow(row);
  else sheet.getRange(rowIdx, 1, 1, 6).setValues([row]);

  return json({ success: true, date: date });
}

function saveTicket_(p) {
  const id = String(p.id || '').trim();
  if (!id) return json({ success: false, error: 'Missing ticket id' });

  const sheet = getTicketsSheet_();
  const rowIdx = findRow_(sheet, id);
  const row = [id, p.practice || '', p.boughtOn || '', p.usedFor || '', p.usedAt || ''];
  if (rowIdx === -1) sheet.appendRow(row);
  else sheet.getRange(rowIdx, 1, 1, 5).setValues([row]);

  return json({ success: true, id: id });
}

function readTickets_() {
  const sheet = getTicketsSheet_();
  const last  = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange(2, 1, last - 1, 5).getDisplayValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      id:       String(r[0]).trim(),
      practice: String(r[1]).trim(),
      boughtOn: String(r[2]).trim(),
      usedFor:  String(r[3]).trim() || null,
      usedAt:   String(r[4]).trim() || null,
    }));
}

/* ---------- helpers ---------- */
function findRow_(sheet, key) {
  const last = sheet.getLastRow();
  if (last <= 1) return -1;
  const keys = sheet.getRange(2, 1, last - 1, 1).getDisplayValues().flat();
  const found = keys.findIndex(k => String(k).trim() === key);
  return found === -1 ? -1 : found + 2;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function getTicketsSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(TICKETS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TICKETS_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([['Id', 'Practice', 'Bought On', 'Used For', 'Used At']]);
    sheet.getRange('A:E').setNumberFormat('@');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run once: create headers and set the date column to plain text. */
function setupSheet() {
  const sheet = getSheet_();
  sheet.getRange(1, 1, 1, 6).setValues([[
    'Date', 'Surya Kriya Cycles', 'Shambhavi', 'Bhuta Shuddhi', 'Devi Sadhana', 'Restored'
  ]]);
  sheet.getRange(1, 1, 1, 6)
    .setFontWeight('bold').setBackground('#EC671B').setFontColor('#FFFFFF');
  sheet.getRange('A:A').setNumberFormat('@');   // dates stored as text
  sheet.setColumnWidth(1, 120);
  sheet.setFrozenRows(1);
  getTicketsSheet_();
  Logger.log('Setup complete. Now Deploy → Manage deployments → New version.');
}
