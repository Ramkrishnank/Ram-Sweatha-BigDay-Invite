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

// ── Handle RSVP via GET (survives Google's 302 redirect) ───────
function doGet(e) {
  // If query params present, it's an RSVP submission
  if (e.parameter && e.parameter.name) {
    try {
      const ss    = getSpreadsheet();
      const sheet = ensureTab_(ss);
      const p     = e.parameter;

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

      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;color:green;padding:20px">✅ RSVP saved! Thank you ' + p.name + '.</p>');
    } catch (err) {
      return HtmlService.createHtmlOutput('<p style="color:red">Error: ' + err.toString() + '</p>');
    }
  }

  // No params — just show the sheet link
  const url = getSpreadsheet().getUrl();
  return HtmlService.createHtmlOutput(
    '<h2>Ram &amp; Sweatha RSVP is running! 🎉</h2>' +
    '<p>Responses saved here: <a href="' + url + '" target="_blank">' + url + '</a></p>'
  );
}

// ── doPost kept as fallback ─────────────────────────────────────
function doPost(e) {
  try {
    const ss    = getSpreadsheet();
    const sheet = ensureTab_(ss);
    const p = e.parameter || {};
    let data = p.name ? p : (p.payload ? JSON.parse(p.payload) : JSON.parse(e.postData.contents));

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

    return json_({ status: 'success' });
  } catch (err) {
    return json_({ status: 'error', message: err.toString() });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
