/**
 * Sérénité Ô Chalet — Guest-to-Host Messaging System
 * 
 * SETUP
 * 1. Open the Google Sheet:
 *    https://docs.google.com/spreadsheets/d/13T9NxlREdYwcCxQVlxQambJ7-8XsghVaggKzq8dvPGQ/edit
 * 2. Create two new tabs: "Messages" and "MessageThreads"
 * 3. In the "Messages" tab, add headers:
 *    MessageID | ThreadID | GuestName | GuestEmail | GuestPhone | Subject | Message | SentAt | Status | HostResponse | ResponseSentAt
 * 4. In "MessageThreads" tab, add headers:
 *    ThreadID | BookingID | GuestName | GuestEmail | GuestPhone | Subject | CreatedAt | LastMessageAt | Status
 * 5. Extensions > Apps Script.
 * 6. Keep the existing doPost() for bookings, add this code to the file.
 * 7. Deploy > New deployment > select type "Web app".
 * 8. Copy the new Web app URL and update MESSAGING_URL in index.html
 */

var MESSAGES_SHEET = 'Messages';
var THREADS_SHEET = 'MessageThreads';
var HOST_EMAIL = 'serenite.o.chalet@gmail.com'; // Change to your email

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    var data = JSON.parse(e.postData.contents);
    
    // Route to appropriate handler based on action type
    if (data.action === 'booking') {
      return handleBooking(data);
    } else if (data.action === 'message') {
      return handleMessage(data);
    } else {
      return sendError('Unknown action: ' + data.action);
    }
  } catch (err) {
    return sendError(String(err));
  } finally {
    lock.releaseLock();
  }
}

function handleBooking(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reservations');
  if (!sheet) throw new Error('Could not find "Reservations" tab.');
  
  var checkin = new Date(data.checkin);
  var checkout = new Date(data.checkout);
  var nights = Math.max(1, Math.round((checkout - checkin) / (1000 * 60 * 60 * 24)));
  var isPeak = (checkin.getMonth() === 6 || checkin.getMonth() === 7);
  var rate = isPeak ? 340 : 250;
  var CLEANING_FEE = 95;
  var DEPOSIT_PCT = 0.30;
  var total = nights * rate + CLEANING_FEE;
  var depositDue = Math.round(total * DEPOSIT_PCT);
  var balance = total;

  var lastRow = 1;
  var guestCol = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).getValues();
  for (var i = 0; i < guestCol.length; i++) {
    if (guestCol[i][0] !== '' && guestCol[i][0] !== null) lastRow = i + 2;
  }
  var targetRow = lastRow + 1;
  var bookingId = lastRow > 1 ? (Number(sheet.getRange(lastRow, 1).getValue()) || (lastRow - 1)) + 1 : 1;

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
    data.partySize || '',
    data.message || ''
  ]]);

  return sendSuccess({ ok: true, bookingId: bookingId, row: targetRow });
}

function handleMessage(data) {
  var threadsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(THREADS_SHEET);
  var messagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MESSAGES_SHEET);
  
  if (!threadsSheet) throw new Error('Could not find "' + THREADS_SHEET + '" tab.');
  if (!messagesSheet) throw new Error('Could not find "' + MESSAGES_SHEET + '" tab.');

  // Generate IDs
  var threadId = data.threadId || Utilities.getUuid();
  var messageId = Utilities.getUuid();
  var now = new Date();

  // Check if thread exists
  var threadExists = false;
  var allThreads = threadsSheet.getDataRange().getValues();
  for (var i = 1; i < allThreads.length; i++) {
    if (allThreads[i][0] === threadId) {
      threadExists = true;
      // Update last message time
      threadsSheet.getRange(i + 1, 8).setValue(now);
      break;
    }
  }

  // Create new thread if needed
  if (!threadExists && !data.threadId) {
    var nextThreadRow = threadsSheet.getLastRow() + 1;
    threadsSheet.getRange(nextThreadRow, 1, 1, 9).setValues([[
      threadId,
      data.bookingId || '',
      data.guestName || '',
      data.guestEmail || '',
      data.guestPhone || '',
      data.subject || 'General Inquiry',
      now,
      now,
      'Open'
    ]]);
  }

  // Add message
  var nextMessageRow = messagesSheet.getLastRow() + 1;
  messagesSheet.getRange(nextMessageRow, 1, 1, 11).setValues([[
    messageId,
    threadId,
    data.guestName || '',
    data.guestEmail || '',
    data.guestPhone || '',
    data.subject || 'General Inquiry',
    data.message || '',
    now,
    'Unread',
    '',
    ''
  ]]);

  // Send email notification to host
  sendHostNotification(data, threadId);

  return sendSuccess({
    ok: true,
    messageId: messageId,
    threadId: threadId,
    timestamp: now.toISOString()
  });
}

function sendHostNotification(data, threadId) {
  try {
    var subject = '[New Message] ' + (data.subject || 'Guest Inquiry') + ' from ' + data.guestName;
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px;">' +
      '<h2 style="color: #1E3B34; margin-top: 0;">New Message from ' + (data.guestName || 'Guest') + '</h2>' +
      '<p><strong>Subject:</strong> ' + (data.subject || 'General Inquiry') + '</p>' +
      '<p><strong>From:</strong> ' + data.guestName + ' (' + data.guestEmail + ')</p>' +
      '<p><strong>Phone:</strong> ' + data.guestPhone + '</p>' +
      (data.bookingId ? '<p><strong>Booking ID:</strong> ' + data.bookingId + '</p>' : '') +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="white-space: pre-wrap; line-height: 1.6;">' + (data.message || '') + '</p>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="color: #666; font-size: 12px;">Thread ID: ' + threadId + '</p>' +
      '<p style="color: #666; font-size: 12px;">Sent: ' + new Date().toLocaleString() + '</p>' +
      '</div>';

    MailApp.sendEmail(HOST_EMAIL, subject, 'New message from ' + data.guestName + ': ' + data.message, {
      htmlBody: htmlBody
    });
  } catch (err) {
    Logger.log('Failed to send notification email: ' + err);
  }
}

function sendSuccess(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendError(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
