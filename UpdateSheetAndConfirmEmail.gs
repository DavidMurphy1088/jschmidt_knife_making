// Paste this into Google Apps Script opened from within the Google Sheet
// (Extensions > Apps Script), then deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
// Paste the deployment URL into index.html as the APPS_SCRIPT_URL value.

const VERSION = '1.0.1';

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ version: VERSION }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var step = 'parsing order data';
  var data = {};
  try {
    data  = JSON.parse(e.postData.contents);
    const knives = data.knives || [];

    const now       = new Date();
    const day       = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd');
    const month     = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
    const year      = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy');
    const time      = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
    const timestamp = day + ' ' + month + ' ' + year;

    const knifeLines = knives.map(function(k, i) {
      return '  Knife ' + (i + 1) + ': ' + k.width + ' | ' + k.grind + ' | ' + k.profile;
    }).join('\n');

    step = 'accessing Google Sheet';
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName('Online Orders');
    if (!sheet) sheet = ss.insertSheet('Online Orders');

    step = 'writing header row';
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Name', 'Address', 'Phone', 'Email',
        'Knife #', 'Width', 'Grind', 'Profile'
      ]);
    }

    step = 'writing customer row to sheet';
    sheet.appendRow([
      timestamp,
      data.name    || '',
      data.address || '',
      data.phone   || '',
      data.email   || '',
      '', '', '', ''
    ]);

    step = 'writing knife rows to sheet';
    knives.forEach(function(knife, i) {
      sheet.appendRow([
        '', '', '', '', '',
        i + 1,
        knife.width   || '',
        knife.grind   || '',
        knife.profile || ''
      ]);
    });

    step = 'writing separator row to sheet';
    sheet.appendRow(['', '', '', '', '', '', '', '', '']);

    const firstName = (data.name || '').split(' ')[0];
    const knifeWord = knives.length === 1 ? 'knife' : 'knives';

    step = 'sending confirmation email to customer';
    MailApp.sendEmail({
      to:      data.email,
      subject: 'Your Schmidt Knives order — ' + (data.name || ''),
      body:
        'Dear ' + firstName + ',\n\n'
        + 'Thank you so much for your order and your interest in our knife selection.\n\n'
        + 'Your order:\n\n'
        + knifeLines + '\n\n'
        + 'Total: ' + knives.length + ' ' + knifeWord + ' at USD $70 each — no payment required now.\n'
        + 'I will be in touch once 150 orders are received and production begins,\n'
        + 'and again when your knives are ready.\n\n'
        + 'If you have any questions at all please don\'t hesitate to get in touch.\n\n'
        + 'Warm regards,\n'
        + 'John Schmidt\n'
        + 'Schmidt Knives\n'
        + 'jpschmidt44@gmail.com\n'
        + 'Wellington, New Zealand',
    });

    step = 'sending notification email to John';
    MailApp.sendEmail({
      to:      'jpschmidt44@gmail.com',
      subject: 'New knife order, ' + day + ' ' + month + ' ' + time + ' — ' + (data.name || ''),
      body:
        'New knife order received ' + day + ' ' + month + ' at ' + time + '\n\n'
        + 'Name:     ' + (data.name    || '') + '\n'
        + 'Email:    ' + (data.email   || '') + '\n'
        + 'Phone:    ' + (data.phone   || '') + '\n'
        + 'Address:  ' + (data.address || '') + '\n\n'
        + 'Knives ordered (' + (data.knives || []).length + '):\n'
        + knifeLines + '\n\n'
        + '▶️ Added to the Google Sheet.',
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Build order summary from whatever data was parsed before the failure
    var knives   = data.knives || [];
    var knifeSum = knives.length > 0
      ? knives.map(function(k, i) {
          return '  Knife ' + (i + 1) + ': ' + (k.width || '?') + ' | ' + (k.grind || '?') + ' | ' + (k.profile || '?');
        }).join('\n')
      : '  (no knife data available)';

    try {
      MailApp.sendEmail({
        to:      'davidmurphy1088@gmail.com',
        subject: 'Schmidt Knives — ORDER PROCESSING ERROR',
        body:
          'An error occurred while processing a knife order.\n\n'
          + 'Failed at step: ' + step + '\n'
          + 'Error: ' + err.toString() + '\n\n'
          + '── Order details ──────────────────────\n'
          + 'Name:     ' + (data.name    || '(not parsed)') + '\n'
          + 'Email:    ' + (data.email   || '(not parsed)') + '\n'
          + 'Phone:    ' + (data.phone   || '(not parsed)') + '\n'
          + 'Address:  ' + (data.address || '(not parsed)') + '\n\n'
          + 'Knives:\n' + knifeSum + '\n\n'
          + 'Check the Apps Script Executions log for a full stack trace.',
      });
    } catch (mailErr) {
      // If the error email itself fails, nothing more we can do — check Executions log
      Logger.log('Failed to send error notification: ' + mailErr.toString());
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString(), step: step }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run testAll() to test the full process end-to-end.
// Both emails go to davidmurphy1088@gmail.com so you can verify without involving John or a real customer.
function testAll() {
  const TEST_EMAIL = 'davidmurphy1088@gmail.com';

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('Online Orders');
  if (!sheet) sheet = ss.insertSheet('Online Orders');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp', 'Name', 'Address', 'Phone', 'Email',
      'Knife #', 'Width', 'Grind', 'Profile'
    ]);
  }

  const now       = new Date();
  const day       = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd');
  const month     = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
  const year      = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy');
  const time      = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
  const timestamp = day + ' ' + month + ' ' + year;

  const name    = 'Test Customer';
  const address = '123 Test Street, Wellington, NZ';
  const phone   = '+64 21 000 0000';
  const knives  = [
    { width: '9 mm',  grind: 'Double Bevel',                 profile: 'Curved'   },
    { width: '12 mm', grind: 'RH (Right-hand single bevel)', profile: 'Straight' }
  ];

  // Sheet update
  sheet.appendRow([timestamp, name, address, phone, TEST_EMAIL, '', '', '', '']);
  knives.forEach(function(knife, i) {
    sheet.appendRow(['', '', '', '', '', i + 1, knife.width, knife.grind, knife.profile]);
  });
  sheet.appendRow(['', '', '', '', '', '', '', '', '']);

  const knifeLines = knives.map(function(k, i) {
    return '  Knife ' + (i + 1) + ': ' + k.width + ' | ' + k.grind + ' | ' + k.profile;
  }).join('\n');

  // Confirmation email (to David instead of real customer)
  MailApp.sendEmail({
    to:      TEST_EMAIL,
    subject: '[TEST1] Your Schmidt Knives order — ' + name,
    body:
      'Dear Test,\n\n'
      + 'Thank you so much for your order and your interest in our knife selection.\n\n'
      + 'Your order:\n\n'
      + knifeLines + '\n\n'
      + 'Total: 2 knives at USD $70 each — no payment required now.\n'
      + 'I will be in touch once 150 orders are received and production begins,\n'
      + 'and again when your knives are ready.\n\n'
      + 'Warm regards,\n'
      + 'John Schmidt\n'
      + 'Schmidt Knives\n'
      + 'jpschmidt44@gmail.com\n'
      + 'Wellington, New Zealand',
  });

  // Notification email (to David instead of John)
  MailApp.sendEmail({
    to:      TEST_EMAIL,
    subject: '[TEST1] New knife order, ' + day + ' ' + month + ' ' + time + ' — ' + name,
    body:
      'New knife order received ' + day + ' ' + month + ' at ' + time + '\n\n'
      + 'Name:     ' + name    + '\n'
      + 'Email:    ' + TEST_EMAIL + '\n'
      + 'Phone:    ' + phone   + '\n'
      + 'Address:  ' + address + '\n\n'
      + 'Knives ordered (2):\n'
      + knifeLines + '\n\n'
      + '▶️ Added to the Google Sheet.',
  });

  Logger.log('testAll complete — sheet updated, 2 emails sent to ' + TEST_EMAIL);
}
