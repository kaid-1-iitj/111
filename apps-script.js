/**
 * Google Apps Script — 111 Days Sadhana Tracker
 *
 * SETUP:
 *  1. Open script.google.com → New Project
 *  2. Paste this entire file into Code.gs
 *  3. Replace SHEET_ID below with your Google Sheet ID
 *  4. Run setupSheet() once (Functions menu → setupSheet → Run)
 *  5. Deploy → New Deployment → Web App
 *       Execute as: Me
 *       Who has access: Anyone
 *  6. Copy the Web App URL and paste it in index.html as CFG.SCRIPT_URL
 */

const SHEET_ID   = '1c3-6c4X9d5wWWe0d9rnwz7R_e-WYDJmpbDN6pPGgN1U';   // ← replace this
const SHEET_NAME = '111';

// ─── Sheet columns: A=Date, B=Cycles, C=Shambhavi, D=Bhuta, E=Devi

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const last  = sheet.getLastRow();

    if (last <= 1) return json({ data: [] });

    const rows = sheet.getRange(2, 1, last - 1, 5).getValues();
    const data = rows
      .filter(r => r[0])
      .map(r => [
        r[0],               // date string  e.g. "2026-06-10"
        Number(r[1]) || 0,  // Surya cycles 0-3
        r[2] === true || r[2] === 'TRUE',  // Shambhavi
        r[3] === true || r[3] === 'TRUE',  // Bhuta
        r[4] === true || r[4] === 'TRUE',  // Devi
      ]);

    return json({ data });
  } catch (err) {
    return json({ error: err.toString(), data: [] });
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const { date, cycles, shambhavi, bhuta, devi } = p;

    // Find existing row for this date
    const last = sheet.getLastRow();
    let rowIdx  = -1;
    if (last > 1) {
      const dates = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      const found = dates.indexOf(date);
      if (found !== -1) rowIdx = found + 2; // 1-indexed + header offset
    }

    const row = [date, cycles, shambhavi, bhuta, devi];
    if (rowIdx === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIdx, 1, 1, 5).setValues([row]);
    }

    return json({ success: true });
  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run this once to create headers in the sheet. */
function setupSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.getRange(1, 1, 1, 5).setValues([[
    'Date', 'Surya Kriya Cycles', 'Shambhavi', 'Bhuta Shuddhi', 'Devi Sadhana'
  ]]);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#EC671B').setFontColor('#FFFFFF');
  sheet.setColumnWidth(1, 120);
  sheet.setFrozenRows(1);
  Logger.log('Sheet setup complete.');
}
