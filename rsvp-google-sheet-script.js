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

// ── Handle RSVP POST from the invite ───────────────────────────
function doPost(e) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ensureTab_(ss);

    // Read individual form fields sent by the hidden form
    const p = e.parameter || {};
    let data = {};
    if (p.name) {
      // direct field submission from hidden form
      data = p;
    } else if (p.payload) {
      data = JSON.parse(p.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    sheet.appendRow([
      data.timestamp   || new Date().toLocaleString(),
      data.name        || '',
      data.phone       || '',
      data.guests      || '',
      data.events      || '',
      data.attending   || '',
      data.message     || '',
      data.submittedAt || new Date().toISOString()
    ]);

    const row = sheet.getLastRow();
    const att = (data.attending || '').toLowerCase();
    if (att.indexOf('yes') > -1) sheet.getRange(row, 1, 1, 8).setBackground('#e8f5e9');
    else if (att.indexOf('no') > -1) sheet.getRange(row, 1, 1, 8).setBackground('#fce4ec');

    return json_({ status: 'success' });
  } catch (err) {
    return json_({ status: 'error', message: err.toString() });
  }
}

// ── Visiting the /exec URL shows a link to your sheet ──────────
function doGet() {
  const ss  = getSpreadsheet();
  const url = ss.getUrl();
  return HtmlService.createHtmlOutput(
    '<h2>Ram &amp; Sweatha RSVP is running! 🎉</h2>' +
    '<p>Responses are saved in your Google Sheet:</p>' +
    '<p><a href="' + url + '" target="_blank" style="font-size:16px">' + url + '</a></p>'
  );
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
