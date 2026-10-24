/**
 * Wedding invitation — RSVP backend (Google Sheets)
 *
 * Setup:
 * 1. Create a new Google Sheet.
 * 2. Extensions → Apps Script, paste this whole file, save.
 * 3. Deploy → New deployment → type "Web app":
 *      Execute as: Me    |    Who has access: Anyone
 * 4. Copy the Web app URL and paste it into CONFIG.sheetUrl in index.html.
 *
 * Replies land in an "RSVP" tab that is created automatically.
 * Messages are shown on the page in the wishes section; put any value in the
 * "Hide" column to hide one. After editing this file, redeploy:
 * Deploy → Manage deployments → Edit → Version: New version.
 */

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Trim, limit length, and stop values from being read as spreadsheet formulas
function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.type === 'rsvp') {
    sheet_('RSVP', ['Time', 'Name', 'Response', 'Contact', 'Message', 'Hide'])
      .appendRow([new Date(), clean_(d.name, 80), clean_(d.response, 10), clean_(d.contact, 120), clean_(d.message, 500)]);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Public list of wishes (name + message only, never contact details), newest first
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RSVP');
  const rows = sh ? sh.getDataRange().getValues().slice(1) : [];
  const wishes = rows
    .filter(r => String(r[4]).trim() && !String(r[5] || '').trim())
    .reverse()
    .slice(0, 200)
    .map(r => ({ name: String(r[1]), message: String(r[4]), time: r[0] }));
  return ContentService.createTextOutput(JSON.stringify({ wishes }))
    .setMimeType(ContentService.MimeType.JSON);
}
