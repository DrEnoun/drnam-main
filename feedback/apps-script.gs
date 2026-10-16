/**
 * "Help me improve": receives feedback from the Main Hub and saves it in a Google Sheet.
 *
 * Setup (about 3 minutes):
 *  1. Make a new Google Sheet (any name, for example "Hub feedback").
 *  2. In the Sheet: Extensions > Apps Script. Delete the sample code and paste this whole file. Save.
 *  3. Deploy > New deployment > type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy and allow the permissions.
 *  4. Copy the Web app URL (it ends with /exec).
 *  5. Open index.html in drnam-main, find  var FEEDBACK_URL = '';  and paste the URL between the quotes.
 *
 * Each reply becomes a new row in a tab called "Feedback".
 * If you change this code later: Deploy > Manage deployments > edit > New version.
 */
var SHEET_NAME = 'Feedback';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    if (d.website) return reply_(true); // hidden trap field: a real person leaves it empty
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Time', 'Rating (1-5)', 'I am', 'What I liked', 'What can be better', 'Contact (optional)', 'Page']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), clean_(d.rating, 2), clean_(d.role, 40), clean_(d.liked, 1500), clean_(d.better, 1500), clean_(d.contact, 200), clean_(d.page, 200)]);
    return reply_(true);
  } catch (err) {
    return reply_(false);
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() { return reply_(true); } // opening the link in a browser just says ok

// Keep text short and stop spreadsheet formulas ("=...", "+...") from running.
function clean_(v, max) {
  v = String(v == null ? '' : v).slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply_(ok) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok })).setMimeType(ContentService.MimeType.JSON);
}
