// Paste into Extensions > Apps Script of a Google Sheet.
// IMPORTANT: replace the value below with YOUR secret key (the same one you type in the dashboard).
const KEY = 'CHANGE_ME_TO_A_SECRET';

// Questions that have an "Other" text box. Must match `other` in questions.js and OC in dashboard.html.
const OTHER_Q = [1, 2, 6];
// Columns returned to the dashboard: language + q1..q12 + one text per OTHER_Q
const COLS = 1 + 12 + OTHER_Q.length;

function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }

// Survey -> saves one row:
// A timestamp | B language | C..N q1..q12 (option numbers, multi answers like 2|5) | O,P,Q "Other" text for q1, q2, q6
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const a = d.a || {}, o = d.o || {};
  const row = [new Date(), d.lang];
  for (let i = 1; i <= 12; i++) row.push("'" + (a[i] || ''));
  OTHER_Q.forEach(i => row.push("'" + String(o[i] || '').slice(0, 200)));
  sheet_().appendRow(row);
  return ContentService.createTextOutput('ok');
}

// Dashboard -> returns all rows as JSON, only with the right key.
// A row counts as a response if column B holds a valid language (en / fr / ar).
function doGet(e) {
  let out;
  if (e.parameter.key === KEY) {
    out = sheet_().getDataRange().getValues()
      .filter(r => ['en', 'fr', 'ar'].includes(String(r[1]).trim().toLowerCase()))
      .map(r => {
        const x = r.slice(1).map(String);
        x[0] = x[0].trim().toLowerCase();
        while (x.length < COLS) x.push('');
        return x;
      });
  } else {
    out = { error: 'unauthorized' };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
