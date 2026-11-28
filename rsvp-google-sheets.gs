const SPREADSHEET_ID = '1faTrjdlPBAc_lp6LixmatYjHkUGoKXPuhvW_WytdeEc';
const RSVP_SHEET_NAME = 'RSVP Responses';
const RSVP_HEADERS = [
  'Saved At (UTC)',
  'Guest Name',
  'RSVP Response',
  'Number of Guests',
  'Phone / Email',
  'Blessing',
  'Couple',
  'Submitted At (IST)',
  'Submitted At (UTC)'
];

function doGet() {
  return jsonResponse({ success: true, message: 'RSVP sheet endpoint is ready.' });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing RSVP submission.');
    }

    const rsvp = JSON.parse(e.postData.contents);
    const responseOptions = [
      'Yes — Joyfully Attending',
      'Maybe — Still Confirming',
      'No — Regretfully Declines (Sending Love)'
    ];
    const name = safeCellText(rsvp.name, 200);
    const attendance = safeCellText(rsvp.attendance, 100);

    if (!name || !responseOptions.includes(attendance)) {
      throw new Error('Invalid RSVP name or response.');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = spreadsheet.getSheetByName(RSVP_SHEET_NAME)
        || spreadsheet.insertSheet(RSVP_SHEET_NAME);

      if (sheet.getLastRow() === 0) {
        sheet.appendRow(RSVP_HEADERS);
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        new Date(),
        name,
        attendance,
        safeCellText(rsvp.guests, 30),
        safeCellText(rsvp.contact, 500),
        safeCellText(rsvp.blessing, 2000),
        safeCellText(rsvp.couple, 200),
        safeCellText(rsvp.submittedAtIST, 100),
        safeCellText(rsvp.submittedAt, 100)
      ]);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({ success: true });
  } catch (error) {
    console.error('RSVP sheet save failed:', error);
    return jsonResponse({ success: false, error: error.message });
  }
}

function safeCellText(value, maxLength) {
  const text = String(value == null ? '' : value).trim().slice(0, maxLength);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function jsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
