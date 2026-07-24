/**
 * ══════════════════════════════════════════════════════════════
 *  Ram & Sweatha Wedding — RSVP Google Sheets Script
 *  Paste this entire file into Google Apps Script and deploy
 * ══════════════════════════════════════════════════════════════
 *
 * SETUP STEPS:
 *  1. Go to https://script.google.com  → New Project
 *  2. Delete any existing code and paste this entire file
 *  3. Click "Save" (Ctrl+S), name it "Wedding RSVP"
 *  4. Click "Deploy" → "New Deployment"
 *  5. Type = "Web App"
 *  6. Execute as = "Me"
 *  7. Who has access = "Anyone"
 *  8. Click "Deploy" → copy the Web App URL
 *  9. Paste that URL into index.html at:
 *         const SHEET_URL = 'PASTE_URL_HERE';
 * 10. Redeploy index.html to GitHub Pages
 * ══════════════════════════════════════════════════════════════
 */

// ── Handle RSVP POST from wedding invite ──────────────────────
function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName('RSVP Responses');

    // Create sheet with headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('RSVP Responses');
      const headers = [
        'Timestamp (IST)', 'Name', 'Phone', 'Guests',
        'Events Attending', 'Attending?', 'Message / Blessings', 'Submitted At (UTC)'
      ];
      sheet.appendRow(headers);

      // Style the header row
      const hdrRange = sheet.getRange(1, 1, 1, headers.length);
      hdrRange.setBackground('#7b1c2e')
              .setFontColor('#ffffff')
              .setFontWeight('bold')
              .setFontSize(11);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 140);
      sheet.setColumnWidth(3, 130);
      sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 170);
      sheet.setColumnWidth(6, 100);
      sheet.setColumnWidth(7, 250);
      sheet.setColumnWidth(8, 180);
    }

    // Parse incoming data
    const data = JSON.parse(e.postData.contents);

    // Append new RSVP row
    sheet.appendRow([
      data.timestamp    || new Date().toLocaleString(),
      data.name         || '',
      data.phone        || '',
      data.guests       || '',
      data.events       || '',
      data.attending    || '',
      data.message      || '',
      data.submittedAt  || new Date().toISOString()
    ]);

    // Color YES rows green, NO rows red
    const lastRow  = sheet.getLastRow();
    const attending = (data.attending || '').toLowerCase();
    if (attending.includes('yes')) {
      sheet.getRange(lastRow, 1, 1, 8).setBackground('#e8f5e9');
    } else if (attending.includes('no')) {
      sheet.getRange(lastRow, 1, 1, 8).setBackground('#fce4ec');
    }

    // Auto-resize columns
    sheet.autoResizeColumns(1, 8);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'RSVP saved!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Health check GET ──────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'running',
      message: 'Ram & Sweatha Wedding RSVP endpoint is active! 🎊'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Optional: Email notification on new RSVP ─────────────────
// Uncomment and set your email to get notified on every RSVP
/*
function sendNotification(data) {
  const email   = 'YOUR_EMAIL@gmail.com';
  const subject = `💍 New RSVP — ${data.name} (${data.attending})`;
  const body    = `New RSVP received!\n\nName: ${data.name}\nPhone: ${data.phone}\nGuests: ${data.guests}\nEvents: ${data.events}\nAttending: ${data.attending}\nMessage: ${data.message}\nTime: ${data.timestamp}`;
  MailApp.sendEmail(email, subject, body);
}
*/
