// ─────────────────────────────────────────────────────────────────────────────
// Desert Bloom — Profit First | Google Apps Script
// ─────────────────────────────────────────────────────────────────────────────
// SETUP INSTRUCTIONS:
//   1. Go to https://sheets.google.com and create a new spreadsheet
//   2. Name it "Desert Bloom - Profit First"
//   3. Go to Extensions → Apps Script
//   4. Delete everything in the editor and paste this entire file
//   5. Click Save (Ctrl+S)
//   6. Click Deploy → New Deployment
//   7. Type = Web App
//   8. Execute as: Me
//   9. Who has access: Anyone
//  10. Click Deploy → Copy the Web App URL
//  11. Paste that URL into your Profit First app under Settings → Google Sheets Sync
// ─────────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet(data.sheet);

    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      addHeaders(sheet, data.sheet);
    }

    sheet.appendRow(data.row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Desert Bloom Profit First connected ✓' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function addHeaders(sheet, sheetName) {
  const headers = {
    'Allocations':     [['Date', 'Total Income', "Owner's Pay", 'Tax', 'Profit', 'OpEx']],
    'Expenses':        [['Date', 'Amount', 'Description', 'Category', 'Account', 'Receipt']],
    'AccountBalances': [['Date', 'Type', "Owner's Pay", 'Tax', 'Profit', 'OpEx']],
    'MonthlySummary':  [['Month', 'Year', 'Total Revenue', 'Total Expenses', 'Net Profit', 'Profit Allocated']]
  };

  if (headers[sheetName]) {
    sheet.getRange(1, 1, 1, headers[sheetName][0].length).setValues(headers[sheetName]);
    sheet.getRange(1, 1, 1, headers[sheetName][0].length)
      .setBackground('#0a1628')
      .setFontColor('#3b82f6')
      .setFontWeight('bold');
  }
}
