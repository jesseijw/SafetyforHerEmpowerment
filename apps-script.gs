/**
 * S.H.E. Safety Map — review backend
 * ------------------------------------------------------------
 * What this does:
 *  1. Receives new map submissions from the website (doPost)
 *  2. Logs each one as "Pending" in a Google Sheet
 *  3. Emails safetyforherempowerment@gmail.com with the details
 *     and two links: Approve / Deny
 *  4. When you click one of those links (doGet), it updates the
 *     sheet — no login needed, the link itself is the approval
 *  5. Serves the list of Approved submissions as JSON (doGet
 *     with ?action=list) so the live website can load them
 *
 * SETUP (about 10 minutes, all free):
 *  1. Go to https://sheets.google.com and create a new blank sheet.
 *     Name it whatever you like (e.g. "SHE Map Submissions").
 *  2. In that sheet: Extensions → Apps Script.
 *  3. Delete anything in the editor and paste this whole file in.
 *  4. Click the disk icon to save. Name the project anything.
 *  5. Click "Deploy" → "New deployment".
 *     - Click the gear icon next to "Select type" → choose "Web app".
 *     - Description: anything.
 *     - Execute as: "Me".
 *     - Who has access: "Anyone".
 *     - Click "Deploy". Authorize it with your Google account when asked
 *       (you'll see an "unverified app" warning — click Advanced →
 *       Go to [project name] → Allow. This is normal for scripts you
 *       write yourself.)
 *  6. Copy the Web App URL it gives you.
 *  7. In index.html, find the line:
 *         const REPORT_ENDPOINT = '';
 *     and paste your Web App URL between the quotes.
 *  8. Re-upload index.html to wherever your site is hosted. Done —
 *     new submissions will now email you with Approve/Deny links,
 *     and approved ones will show up on the live map automatically.
 *
 * NOTE: Every time you change this script's code, you must create a
 * NEW deployment (Deploy → Manage deployments → Edit → New version)
 * for the changes to take effect on the same URL.
 * ------------------------------------------------------------
 */

const SHEET_NAME = 'Submissions';
const NOTIFY_EMAIL = 'safetyforherempowerment@gmail.com';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','type','level','name','desc','short','lat','lng','submittedAt','status','token']);
  }
  return sheet;
}

// Called by the website when someone submits the "Add to the Safety Map" form.
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const token = Utilities.getUuid();

    sheet.appendRow([
      data.id || ('user-' + Date.now()),
      data.type || 'concern',
      data.level || '',
      data.name || '',
      data.desc || '',
      data.short || '',
      data.lat,
      data.lng,
      data.submittedAt || new Date().toISOString(),
      'Pending',
      token
    ]);

    sendReviewEmail_(data, token);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handles: the Approve/Deny links clicked from the email, and
// ?action=list requests from the website to load approved pins.
function doGet(e) {
  const params = e.parameter;

  if (params.action === 'list') {
    return ContentService.createTextOutput(JSON.stringify(getApproved_()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.action === 'approve' || params.action === 'deny') {
    const result = setStatus_(params.id, params.token, params.action === 'approve' ? 'Approved' : 'Denied');
    return HtmlService.createHtmlOutput(
      `<html><body style="font-family:sans-serif;text-align:center;padding:3rem;">
        <h2>${result}</h2>
        <p>You can close this tab.</p>
      </body></html>`
    );
  }

  return ContentService.createTextOutput('S.H.E. Safety Map backend is running.');
}

function sendReviewEmail_(data, token) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const approveUrl = `${webAppUrl}?action=approve&id=${encodeURIComponent(data.id)}&token=${token}`;
  const denyUrl    = `${webAppUrl}?action=deny&id=${encodeURIComponent(data.id)}&token=${token}`;
  const typeLabel = data.type === 'concern' ? `Concern (${data.level} level)` : data.type;

  const body =
    `A new submission came in for the S.H.E. Safety Map:\n\n` +
    `Type: ${typeLabel}\n` +
    `Name: ${data.name}\n` +
    `Description: ${data.desc}\n` +
    `Location: ${data.lat}, ${data.lng}\n` +
    `Map: https://maps.google.com/?q=${data.lat},${data.lng}\n\n` +
    `Approve (adds it to the live map): ${approveUrl}\n\n` +
    `Deny (discards it): ${denyUrl}\n`;

  MailApp.sendEmail(NOTIFY_EMAIL, `S.H.E. Map Submission: ${data.name}`, body);
}

function setStatus_(id, token, status) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id && values[i][10] === token) {
      sheet.getRange(i + 1, 10).setValue(status); // "status" column
      return `Marked as ${status}: ${values[i][3]}`;
    }
  }
  return 'Submission not found or link already used.';
}

function getApproved_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row[9] === 'Approved') {
      rows.push({
        id: row[0],
        type: row[1],
        level: row[2] || null,
        name: row[3],
        desc: row[4],
        short: row[5],
        lat: Number(row[6]),
        lng: Number(row[7]),
        userSubmitted: true
      });
    }
  }
  return rows;
}
