/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  MBA Discovery Tool — Google Apps Script Web App
 *  Receives quiz responses and writes them to your Google Sheet.
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  SETUP (one-time, takes ~5 minutes):
 *
 *  1. Create a new Google Sheet:
 *       https://sheets.new
 *     Rename it: "MBA Discovery Tool — Responses"
 *
 *  2. In the sheet, add these headers in Row 1 (columns A–N):
 *       Timestamp | Name | Email | Contact | Course | Specialisation |
 *       Domain 1 | Score 1 | Domain 2 | Score 2 | Domain 3 | Score 3
 *
 *  3. Open Apps Script editor:
 *       Extensions → Apps Script
 *
 *  4. Delete any existing code in the editor.
 *     Copy & paste ALL of the code below into the editor.
 *
 *  5. Click "Save" (Ctrl/Cmd + S). Name the project "MBA Discovery Tool".
 *
 *  6. Deploy as a Web App:
 *       Click "Deploy" → "New deployment"
 *       → Type: Web App
 *       → Execute as: Me
 *       → Who has access: Anyone
 *       → Click "Deploy"
 *       → Copy the Web App URL (looks like: https://script.google.com/macros/s/AK.../exec)
 *
 *  7. Paste that URL into app.js:
 *       const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AK.../exec';
 *
 *  IMPORTANT: Every time you modify this script and redeploy,
 *  click "Deploy" → "Manage deployments" → edit → "New version" to update.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── Paste this code into the Apps Script editor ───────────────────────────

const SHEET_NAME = 'Sheet1'; // Change if your sheet tab has a different name

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

    // Auto-add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Name', 'Email', 'Contact', 'Course', 'Specialisation',
        'Domain 1', 'Score 1', 'Domain 2', 'Score 2', 'Domain 3', 'Score 3'
      ]);
    }

    sheet.appendRow([
      data.timestamp      || new Date().toISOString(),
      data.name           || '',
      data.email          || '',
      data.contact        || '',
      data.course         || '',
      data.specialisation || '',
      data.domain1        || '',
      data.score1         || 0,
      data.domain2        || '',
      data.score2         || 0,
      data.domain3        || '',
      data.score3         || 0,
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: test this function manually in the Apps Script editor
function testDoPost() {
  const mock = {
    postData: {
      contents: JSON.stringify({
        timestamp:      new Date().toISOString(),
        name:           'Test User',
        email:          'test@example.com',
        contact:        '+91 99999 00000',
        course:         'MBA',
        specialisation: 'Marketing',
        domain1:        'Marketing & Branding', score1: 80,
        domain2:        'Content & Creative',   score2: 70,
        domain3:        'Product & Growth',     score3: 60,
      })
    }
  };
  Logger.log(doPost(mock));
}
