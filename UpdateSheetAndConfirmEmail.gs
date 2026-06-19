// Paste this into Google Apps Script opened from within the Google Sheet
// (Extensions > Apps Script), then deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
// Paste the deployment URL into index.html as the APPS_SCRIPT_URL value.

function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
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
    const knives    = data.knives || [];

    // Customer details row
    sheet.appendRow([
      timestamp,
      data.name    || '',
      data.address || '',
      data.phone   || '',
      data.email   || '',
      '', '', '', ''
    ]);

    // One row per knife
    knives.forEach(function(knife, i) {
      sheet.appendRow([
        '', '', '', '', '',
        i + 1,
        knife.width   || '',
        knife.grind   || '',
        knife.profile || ''
      ]);
    });

    // Blank separator row between orders
    sheet.appendRow(['', '', '', '', '', '', '', '', '']);

    const knifeLines = knives.map(function(k, i) {
      return '  Knife ' + (i + 1) + ': ' + k.width + ' | ' + k.grind + ' | ' + k.profile;
    }).join('\n');

    const firstName = (data.name || '').split(' ')[0];
    const knifeWord = knives.length === 1 ? 'knife' : 'knives';

    // Confirmation email to customer
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

    // Notification email to John
    MailApp.sendEmail({
      to:      'jpschmidt44@gmail.com',
      subject: 'New knife order, ' + day + ' ' + month + ' ' + time + ' — ' + (data.name || ''),
      body:
        'New knife order received ' + day + ' ' + month + ' at ' + time + '\n\n'
        + 'Name:     ' + (data.name    || '') + '\n'
        + 'Email:    ' + (data.email   || '') + '\n'
        + 'Phone:    ' + (data.phone   || '') + '\n'
        + 'Address:  ' + (data.address || '') + '\n\n'
        + 'Knives ordered (' + knives.length + '):\n'
        + knifeLines + '\n\n'
        + '▶️ Added to the Google Sheet.',
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run testAll() once from the editor to grant MailApp + Sheets permissions before deploying.
// Select testAll from the function dropdown and click Run.
function testAll() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Online Orders') || ss.insertSheet('Online Orders');
  Logger.log('Sheet found: ' + sheet.getName() + ', rows: ' + sheet.getLastRow());

  MailApp.sendEmail({
    to:      'davidmurphy1088@gmail.com',
    subject: 'Apps Script authorisation test — Schmidt Knives',
    body:    'MailApp and SpreadsheetApp are both authorised. Script is ready.',
  });
}
