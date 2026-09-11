/**
 * Sérénité Ô Chalet — booking intake script
 *
 * SETUP
 * 1. Open the Google Sheet (the native Sheet, not the .xlsx copy):
 *    https://docs.google.com/spreadsheets/d/13T9NxlREdYwcCxQVlxQambJ7-8XsghVaggKzq8dvPGQ/edit
 * 2. Extensions > Apps Script.
 * 3. Delete any starter code, paste this whole file in, and save.
 * 4. Deploy > New deployment > select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy, authorize when prompted, then copy the Web app URL
 *    (ends in /exec).
 * 6. Send that URL back — it goes into the WEB_APP_URL constant in the
 *    website's booking form so submissions land here automatically.
 *
 * WHAT IT DOES
 * Every request adds one row to the "Reservations" tab, filling the
 * same columns the sheet already uses: BookingID, Guest, Phone, Email,
 * CheckIn, CheckOut, Nights, NightlyRate, Cleaning, Total, DepositDue,
 * DepositPaid, Balance, Status, Invoice — plus two extra columns for
 * Guests (headcount) and Notes, appended after Invoice, since the form
 * collects those and the sheet didn't have a place for them yet.
 */

var SHEET_NAME = 'Reservations';
var STANDARD_RATE = 285;
var PEAK_RATE = 340;
var CLEANING_FEE = 95;
var DEPOSIT_PCT = 0.30;

function doGet(e) {
  // Visiting the deployed URL directly (a GET request) lands here — just a
  // health check to confirm the deployment is live. Real form submissions
  // use doPost below.
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Booking endpoint is live. Submit via POST to add a reservation.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Could not find a "' + SHEET_NAME + '" tab.');

    // Make sure the two extra columns exist (Guests, Notes) so nothing overwrites Invoice.
    var headerRange = sheet.getRange(1, 1, 1, 17).getValues()[0];
    if (headerRange[15] !== 'Guests') sheet.getRange(1, 16).setValue('Guests');
    if (headerRange[16] !== 'Notes') sheet.getRange(1, 17).setValue('Notes');

    var checkin = new Date(data.checkin);
    var checkout = new Date(data.checkout);
    var nights = Math.max(1, Math.round((checkout - checkin) / (1000 * 60 * 60 * 24)));
    var isPeak = (checkin.getMonth() === 6 || checkin.getMonth() === 7); // Jul=6, Aug=7
    var rate = isPeak ? PEAK_RATE : STANDARD_RATE;
    var total = nights * rate + CLEANING_FEE;
    var depositDue = Math.round(total * DEPOSIT_PCT);
    var balance = total;

    // Find the next free row: first row below the header with no Guest name.
    var lastRow = 1;
    var guestCol = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).getValues();
    for (var i = 0; i < guestCol.length; i++) {
      if (guestCol[i][0] !== '' && guestCol[i][0] !== null) lastRow = i + 2;
    }
    var targetRow = lastRow + 1;

    var bookingId = lastRow > 1 ? (Number(sheet.getRange(lastRow, 1).getValue()) || (lastRow - 1)) + 1 : 1;

    // Clear any stray leftover formulas/values in that row first.
    sheet.getRange(targetRow, 1, 1, 17).clearContent();

    sheet.getRange(targetRow, 1, 1, 17).setValues([[
      bookingId,
      data.name || '',
      data.phone || '',
      data.email || '',
      data.checkin || '',
      data.checkout || '',
      nights,
      rate,
      CLEANING_FEE,
      total,
      depositDue,
      '',
      balance,
      'Requested',
      '',
      data.guests || '',
      data.notes || ''
    ]]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, bookingId: bookingId, row: targetRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
