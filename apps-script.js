/**
 * Google Apps Script — 111 Days Sadhana Tracker (canonical backend)
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
 * Sheet columns:  A=Date(text)  B=Cycles  C=Shambhavi  D=Bhuta  E=Devi
 */

const SHEET_ID   = '1c3-6c4X9d5wWWe0d9rnwz7R_e-WYDJmpbDN6pPGgN1U';
const SHEET_NAME = '111';

/* ---------- READ: return every logged day as [date, cycles, sh, bh, devi] ---------- */
function doGet(e) {
  try {
    const sheet = getSheet_();
    const last  = sheet.getLastRow();
    if (last <= 1) return json({ success: true, data: [] });

    // getDisplayValues keeps the date column as the exact text shown in the
    // cell ("2026-06-10") so there is never any timezone shifting.
    const rows = sheet.getRange(2, 1, last - 1, 5).getDisplayValues();
    const data = rows
      .filter(r => String(r[0]).trim() !== '')
      .map(r => [
        String(r[0]).trim(),                                   // date  "YYYY-MM-DD"
        Number(r[1]) || 0,                                     // cycles 0–3
        r[2] === true || String(r[2]).toUpperCase() === 'TRUE',// shambhavi
        r[3] === true || String(r[3]).toUpperCase() === 'TRUE',// bhuta
        r[4] === true || String(r[4]).toUpperCase() === 'TRUE',// devi
      ]);

    return json({ success: true, data });
  } catch (err) {
    return json({ success: false, error: err.toString(), data: [] });
  }
}

/* ---------- WRITE: upsert one day by date ---------- */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);

    // Forgiving input: accept either {date,cycles,...} (canonical) or any
    // close variant, so a stale client can't break the write.
    const date      = String(p.date || '').trim();
    const cycles    = Number(p.cycles != null ? p.cycles : p.surya_cycles) || 0;
    const shambhavi = p.shambhavi === true || String(p.shambhavi).toUpperCase() === 'TRUE';
    const bhuta     = p.bhuta     === true || String(p.bhuta).toUpperCase()     === 'TRUE';
    const devi      = p.devi      === true || String(p.devi).toUpperCase()      === 'TRUE';

    if (!date) return json({ success: false, error: 'Missing date' });

    const sheet = getSheet_();
    sheet.getRange('A:A').setNumberFormat('@');  // keep dates as plain text

    // Find an existing row for this date.
    const last  = sheet.getLastRow();
    let rowIdx  = -1;
    if (last > 1) {
      const dates = sheet.getRange(2, 1, last - 1, 1).getDisplayValues().flat();
      const found = dates.findIndex(d => String(d).trim() === date);
      if (found !== -1) rowIdx = found + 2;
    }

    const row = [date, cycles, shambhavi, bhuta, devi];
    if (rowIdx === -1) sheet.appendRow(row);
    else sheet.getRange(rowIdx, 1, 1, 5).setValues([row]);

    return json({ success: true, date: date });
  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

/* ---------- helpers ---------- */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
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
  sheet.getRange(1, 1, 1, 5).setValues([[
    'Date', 'Surya Kriya Cycles', 'Shambhavi', 'Bhuta Shuddhi', 'Devi Sadhana'
  ]]);
  sheet.getRange(1, 1, 1, 5)
    .setFontWeight('bold').setBackground('#EC671B').setFontColor('#FFFFFF');
  sheet.getRange('A:A').setNumberFormat('@');   // dates stored as text
  sheet.setColumnWidth(1, 120);
  sheet.setFrozenRows(1);
  Logger.log('Setup complete. Now Deploy → Manage deployments → New version.');
}
