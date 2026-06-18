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
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
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

    // Build knife list for confirmation email
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

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run this once from the editor to grant MailApp permission before deploying:
// select testEmail from the dropdown and click Run
function testEmail() {
  MailApp.sendEmail({
    to:      'jpschmidt44@gmail.com',
    subject: 'Apps Script authorisation test — Schmidt Knives',
    body:    'If you received this, MailApp is authorised and the script is ready to deploy.',
  });
}
