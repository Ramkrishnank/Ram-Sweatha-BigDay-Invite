/**
 * ══════════════════════════════════════════════════════════════
 *  Ram & Sweatha Wedding — RSVP → Google Sheets (STANDALONE SAFE)
 * ══════════════════════════════════════════════════════════════
 *
 *  ⚡ WHY YOU COULDN'T FIND YOUR RESPONSES:
 *  A standalone Apps Script has NO attached spreadsheet, so the old
 *  getActiveSpreadsheet() returned null and nothing was ever saved.
 *  This version AUTO-CREATES a spreadsheet in YOUR Google Drive and
 *  remembers it — so every RSVP is stored and easy to find.
 *
 *  ── SETUP (do this once) ──────────────────────────────────────
 *  1. Go to https://script.google.com  ->  open your project
 *  2. Delete ALL old code, paste THIS entire file, click Save
 *  3. In the function dropdown (top toolbar) select:  setup
 *  4. Click Run  ->  approve the permissions popup
 *  5. Open "Execution log" — it prints your NEW SHEET URL. Open it!
 *  6. Deploy -> Manage deployments -> Edit (pencil) ->
 *        Version: "New version"  ->  Deploy   (keeps same /exec URL)
 *     (If first time: Deploy -> New deployment -> Web app ->
 *        Execute as: Me · Access: Anyone -> Deploy)
 *  7. Done — submit a test RSVP on the invite and refresh the sheet.
 * ══════════════════════════════════════════════════════════════
 */

const SHEET_NAME = 'Ram & Sweatha — RSVP Responses';
const TAB_NAME   = 'RSVP Responses';
const PROP_KEY   = 'RSVP_SPREADSHEET_ID';

// ── Get (or create) the spreadsheet, remembering its ID ────────
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_KEY);

  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (e) { /* deleted — fall through and make a new one */ }
  }

  const ss = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty(PROP_KEY, ss.getId());
  buildHeaders_(ss);
  return ss;
}

// ── Build a styled header row + tab ────────────────────────────
function buildHeaders_(ss) {
  let sheet = ss.getSheets()[0];
  sheet.setName(TAB_NAME);
  const headers = ['Timestamp (IST)', 'Name', 'Phone', 'Guests',
    'Events Attending', 'Attending?', 'Message / Blessings', 'Submitted At (UTC)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setBackground('#7b1c2e').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);
  const widths = [160, 150, 130, 70, 150, 100, 260, 180];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
  return sheet;
}

// ── RUN THIS ONCE: prints your sheet URL in the log ────────────
function setup() {
  const ss = getSpreadsheet();
  Logger.log('════════════════════════════════════════');
  Logger.log('Your RSVP responses are saved here:');
  Logger.log(ss.getUrl());
  Logger.log('════════════════════════════════════════');
  return ss.getUrl();
}

// ── Handle RSVP POST from the invite ───────────────────────────
function doPost(e) {
  try {
    const ss   = getSpreadsheet();
    const sheet = ss.getSheetByName(TAB_NAME) || buildHeaders_(ss);
    const data = JSON.parse(e.postData.contents);

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
  const url = getSpreadsheet().getUrl();
  return HtmlService.createHtmlOutput(
    '<h2>Ram &amp; Sweatha RSVP is running!</h2>' +
    '<p>Your responses are saved in this Google Sheet:</p>' +
    '<p><a href="' + url + '" target="_blank">' + url + '</a></p>'
  );
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
