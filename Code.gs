// Paste into Extensions > Apps Script of a Google Sheet. Change KEY before deploying.
const KEY = 'CHANGE_ME_TO_A_SECRET';

function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }

// Survey -> saves one row: timestamp, language, q1..q12 (answers stored as option numbers)
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const row = [new Date(), d.lang];
  for (let i = 1; i <= 12; i++) row.push("'" + (d.a[i] || ''));
  sheet_().appendRow(row);
  return ContentService.createTextOutput('ok');
}

// Dashboard -> returns all rows as JSON, only with the right key.
// A row counts as a response if column B holds a valid language (en / fr / ar).
// Column A (timestamp) is no longer required, so rows typed in by hand also work.
function doGet(e) {
  let out = [];
  if (e.parameter.key === KEY) {
    out = sheet_().getDataRange().getValues()
      .filter(r => ['en', 'fr', 'ar'].includes(String(r[1]).trim().toLowerCase()))
      .map(r => {
        const x = r.slice(1).map(String);
        x[0] = x[0].trim().toLowerCase();
        return x;
      });
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
