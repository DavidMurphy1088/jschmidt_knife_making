// David's script — runs under davidmurphy1088@gmail.com
// Sends organiser notification to John when someone places an order.
// Deploy as Web App: Execute as: Me, Who has access: Anyone
// Paste the deployment URL into index.html as NOTIFY_ORGANISER_URL.
//
// NOTE: currently sends to davidmurphy1088@gmail.com for testing.
// Change the 'to' address to jpschmidt44@gmail.com for production.

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const now    = new Date();
    const day    = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd');
    const month  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
    const time   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
    const knives = data.knives || [];

    const knifeLines = knives.map(function(k, i) {
      return '  Knife ' + (i + 1) + ': ' + k.width + ' | ' + k.grind + ' | ' + k.profile;
    }).join('\n');

    MailApp.sendEmail({
      to:      'davidmurphy1088@gmail.com',   // ← change to jpschmidt44@gmail.com for production
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

// Run this once from the editor to grant MailApp permission before deploying.
function testEmail() {
  MailApp.sendEmail({
    to:      'davidmurphy1088@gmail.com',
    subject: 'Organiser notification test — Schmidt Knives',
    body:    'If you received this, the organiser notification script is working correctly.',
  });
}
