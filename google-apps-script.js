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
  const action = e.parameter.action;

  // Return all data for cross-device sync
  if (action === 'getData') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = { allocations: [], expenses: [] };

    // Load allocations
    const allocSheet = ss.getSheetByName('Allocations');
    if (allocSheet && allocSheet.getLastRow() > 1) {
      const rows = allocSheet.getRange(2, 1, allocSheet.getLastRow() - 1, 6).getValues();
      result.allocations = rows.map((row, i) => ({
        id: i + 1,
        date:     row[0] ? String(row[0]) : '',
        income:   parseFloat(row[1]) || 0,
        ownerPay: parseFloat(row[2]) || 0,
        tax:      parseFloat(row[3]) || 0,
        profit:   parseFloat(row[4]) || 0,
        opex:     parseFloat(row[5]) || 0,
      })).filter(r => r.income > 0);
    }

    // Load expenses
    const expSheet = ss.getSheetByName('Expenses');
    if (expSheet && expSheet.getLastRow() > 1) {
      const rows = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 6).getValues();
      result.expenses = rows.map((row, i) => ({
        id:          i + 1,
        date:        row[0] ? String(row[0]) : '',
        amount:      parseFloat(row[1]) || 0,
        description: row[2] ? String(row[2]) : '',
        category:    row[3] ? String(row[3]) : '',
        account:     row[4] ? String(row[4]) : '',
        photo:       null,
      })).filter(r => r.amount > 0);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default status check
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
