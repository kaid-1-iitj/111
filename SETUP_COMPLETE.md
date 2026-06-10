# 111 Days Sadhana Tracker — Complete Setup Guide

Your app is live at: **https://kaid-1-iitj.github.io/111/**

## The Problem

Data doesn't persist when you refresh because **the Google Apps Script Web App hasn't been deployed yet**. This guide fixes it in 10 minutes.

---

## ✓ What's Already Done

- ✓ `index.html` — Complete tracker UI (GitHub Pages)
- ✓ `apps-script.js` — Backend code ready to deploy
- ✓ Google Sheet created with headers
- ✓ SHEET_ID is correctly set in code

---

## ⚠️ What's Missing: Deploy the Web App

### Step 1: Go to Apps Script

1. Open **script.google.com**
2. Sign in with your Google account
3. Click **New Project**
4. Name it: `111 Days Sadhana`

### Step 2: Paste the Code

1. Delete the default `function myFunction() {}` code
2. Copy everything from the bottom of this file (under `--- CODE TO PASTE ---`)
3. Paste into `Code.gs` in Apps Script
4. **Replace ONLY this line** with your actual Sheet ID:
   ```js
   const SHEET_ID = '1c3-6c4X9d5wWWe0d9rnwz7R_e-WYDJmpbDN6pPGgN1U';
   ```
   (Keep it exactly as is — it's your Sheet ID)

### Step 3: Run Setup (One-time only)

1. In Apps Script, click the **function dropdown** (shows "myFunction")
2. Select **setupSheet**
3. Click **Run**
4. Click **Review Permissions** → **Allow**
5. Wait 5 seconds
6. Check your Google Sheet — you should see orange headers in row 1

### Step 4: Deploy as Web App

1. Click **Deploy** (top right)
2. Click **New Deployment** (or find existing and click the pencil to **Edit**)
3. Click the **gear icon** → select **Web app**
4. Set these exactly:
   - **Description:** `111 Days Sadhana Tracker`
   - **Execute as:** Your email (auto-filled)
   - **Who has access:** `Anyone`
5. Click **Deploy**
6. A popup shows your Web App URL — **COPY IT**
   - Looks like: `https://script.google.com/macros/s/ABC123.../exec`

### Step 5: Update Your Tracker

1. Go to your GitHub repo: **https://github.com/kaid-1-iitj/111**
2. Open **index.html** in the editor
3. Find this line (around line 804):
   ```js
   SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbzZDXZzoJxqRMBK3hfYphg07ThIyNKclR5jyRpP7wLmxK7A6PzCuKMJDsdRFVuAoHwl/exec',
   ```
4. Replace the URL with the one you copied in Step 4
5. Scroll down, click **Commit changes**
6. Wait 30 seconds for GitHub Pages to rebuild

### Step 6: Test It

1. Open **https://kaid-1-iitj.github.io/111/** in your browser
2. Log some practices (toggle switches, set Surya cycles)
3. Click **Save Today's Practice**
4. Check the message — should say "Saved ✔"
5. **Refresh the page** — data should still be there
6. **Check your Google Sheet** — row 2 should have your data

---

## 🐛 Troubleshooting

### "Save failed" or no message
- Open browser **DevTools** (F12 or Cmd+Option+I)
- Go to **Console** tab
- Try saving again
- Look for red error messages
- Screenshot the error and send it

### Data saves but doesn't reload
- Check Google Sheet has the "111" tab at the bottom
- Make sure SHEET_NAME in Code.gs is `'111'` (not "Sheet1" or "Sadhana")

### Web App URL not working
- In Apps Script, go to **Deployments** (left menu)
- Make sure you see a "Web app" deployment
- If it shows red warning, delete and deploy again

---

## Code to Paste into Code.gs

```javascript
/**
 * Google Apps Script — 111 Days Sadhana Tracker
 * Paste this entire code into Code.gs in script.google.com
 */

const SHEET_ID   = '1c3-6c4X9d5wWWe0d9rnwz7R_e-WYDJmpbDN6pPGgN1U';
const SHEET_NAME = '111';

// Columns: A=Date, B=Cycles, C=Shambhavi, D=Bhuta, E=Devi

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const last  = sheet.getLastRow();

    if (last <= 1) return json({ data: [] });

    const rows = sheet.getRange(2, 1, last - 1, 5).getValues();
    const data = rows
      .filter(r => r[0])
      .map(r => [
        r[0],
        Number(r[1]) || 0,
        r[2] === true || r[2] === 'TRUE',
        r[3] === true || r[3] === 'TRUE',
        r[4] === true || r[4] === 'TRUE',
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

    const last = sheet.getLastRow();
    let rowIdx  = -1;
    if (last > 1) {
      const dates = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      const found = dates.indexOf(date);
      if (found !== -1) rowIdx = found + 2;
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
  Logger.log('Setup complete. Go back to GitHub and update the SCRIPT_URL.');
}
```

---

## Timeline

- **Today:** Deploy Web App (10 min)
- **Tomorrow:** Your 111-day journey begins!

Good luck with your sadhana! 🔥
