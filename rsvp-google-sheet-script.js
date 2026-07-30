/**
 * ══════════════════════════════════════════════════════════════
 *  Ram & Sweatha Wedding — RSVP → Google Sheets
 *  Saves directly into: "Big Day Invite Response"
 *  https://docs.google.com/spreadsheets/d/1W4z9u4Y4yi0BZ5F6qqB_xqfVZQtYg081qatHzxvg4B8
 * ══════════════════════════════════════════════════════════════
 *
 *  ── SETUP (do this once) ──────────────────────────────────────
 *  1. Go to https://script.google.com  ->  open your project
 *  2. Delete ALL old code, paste THIS entire file, click Save
 *  3. In the function dropdown (top toolbar) select:  setup
 *  4. Click Run  ->  approve the permissions popup
 *  5. Open "Execution log" — it confirms the sheet is accessible
 *  6. Deploy -> Manage deployments -> Edit (pencil) ->
 *        Version: "New version"  ->  Deploy   (keeps same /exec URL)
 *  7. Done — submit a test RSVP on the invite and check your sheet!
 * ══════════════════════════════════════════════════════════════
 */

// ── Your existing "Big Day Invite Response" sheet ──────────────
const SPREADSHEET_ID = '1W4z9u4Y4yi0BZ5F6qqB_xqfVZQtYg081qatHzxvg4B8';
const TAB_NAME       = 'RSVP Responses';

// ── Get the spreadsheet directly by ID ─────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ── Build a styled header row ───────────────────────────────────
function buildHeaders_(sheet) {
  const headers = ['Timestamp (IST)', 'Name', 'Phone', 'Guests',
    'Events Attending', 'Attending?', 'Message / Blessings', 'Submitted At (UTC)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground('#7b1c2e').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);
  const widths = [160, 150, 130, 70, 150, 100, 260, 180];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
}

// ── Ensure the RSVP tab + headers exist ────────────────────────
function ensureTab_(ss) {
  let sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    buildHeaders_(sheet);
  }
  return sheet;
}

// ── RUN THIS ONCE: confirms the sheet is accessible ────────────
function setup() {
  const ss = getSpreadsheet();
  const sheet = ensureTab_(ss);
  Logger.log('════════════════════════════════════════');
  Logger.log('✅ Connected to: ' + ss.getName());
  Logger.log(ss.getUrl());
  Logger.log('Tab ready: ' + sheet.getName());
  Logger.log('════════════════════════════════════════');
}

// ── Shared helper: save one RSVP row ───────────────────────────
function saveRow_(p) {
  const ss    = getSpreadsheet();
  const sheet = ensureTab_(ss);
  sheet.appendRow([
    p.timestamp   || new Date().toLocaleString(),
    p.name        || '',
    p.phone       || '',
    p.guests      || '',
    p.events      || '',
    p.attending   || '',
    p.message     || '',
    p.submittedAt || new Date().toISOString()
  ]);
  const row = sheet.getLastRow();
  const att = (p.attending || '').toLowerCase();
  if (att.indexOf('yes') > -1) sheet.getRange(row, 1, 1, 8).setBackground('#e8f5e9');
  else if (att.indexOf('no') > -1) sheet.getRange(row, 1, 1, 8).setBackground('#fce4ec');
}

// ── doGet: handles XHR/Image GET requests from desktop ─────────
function doGet(e) {
  if (e.parameter && e.parameter.name) {
    try {
      saveRow_(e.parameter);
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    } catch (err) {
      return ContentService.createTextOutput('ERR:' + err.toString());
    }
  }
  const url = getSpreadsheet().getUrl();
  return HtmlService.createHtmlOutput(
    '<h2>Ram &amp; Sweatha RSVP is running! 🎉</h2>' +
    '<p>Responses saved here: <a href="' + url + '" target="_blank">' + url + '</a></p>'
  );
}

// ── doPost: handles sendBeacon FormData from mobile ────────────
function doPost(e) {
  try {
    // sendBeacon sends as multipart/form-data — read via e.parameter
    const p = (e.parameter && e.parameter.name) ? e.parameter
            : (e.postData ? JSON.parse(e.postData.contents) : {});
    saveRow_(p);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('ERR:' + err.toString());
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
