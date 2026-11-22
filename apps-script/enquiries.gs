/**
 * Smart System of Systems — briefing enquiries → Google Sheet
 *
 * Lives inside the "SSoS Briefing Enquiries" sheet (Extensions → Apps Script),
 * deployed as a Web app: Execute as = Me, Who has access = Anyone.
 * The website's briefing form POSTs to the web app URL; each enquiry becomes
 * one row, and (optionally) an email alert goes to the sheet owner.
 */

const NOTIFY_BY_EMAIL = true;   // set to false to stop the email alert
// Extra people who get every alert, comma-separated. Kept empty in the public
// repo; the live copy in the Google sheet has the real addresses.
const EXTRA_RECIPIENTS = '';
const MAX_LEN = 5000;           // cap any single field

const FIELDS = ['first_name', 'last_name', 'organisation', 'role', 'country',
                'email', 'enquiry_type', 'context', 'page'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  // Honeypot: real visitors never see or fill the "website" field.
  if (p.website) return ok_();
  if (!p.first_name || !p.email) return ok_();

  const clean = (v) => {
    v = String(v || '').slice(0, MAX_LEN).trim();
    // Stop anything a visitor types from being run as a spreadsheet formula.
    return /^[=+\-@]/.test(v) ? "'" + v : v;
  };
  const received = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');
  const row = [received].concat(FIELDS.map((f) => clean(p[f])));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow(row);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_BY_EMAIL) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      MailApp.sendEmail({
        to: [Session.getEffectiveUser().getEmail()].concat(EXTRA_RECIPIENTS ? [EXTRA_RECIPIENTS] : []).join(','),
        subject: 'New briefing enquiry: ' + row[1] + ' ' + row[2] + ' (' + row[3] + ')',
        body: [
          'Received: ' + received + ' IST',
          'Name: ' + row[1] + ' ' + row[2],
          'Organisation: ' + row[3],
          'Role: ' + row[4],
          'Country: ' + row[5],
          'Email: ' + row[6],
          'Enquiry: ' + row[7],
          '',
          row[8] || '(no context given)',
          '',
          'All enquiries: ' + ss.getUrl(),
        ].join('\n'),
      });
    } catch (err) { /* the row is saved even if the email fails */ }
  }
  return ok_();
}

function doGet() { return ContentService.createTextOutput('SSoS enquiries endpoint is running.'); }

function ok_() { return ContentService.createTextOutput('ok'); }
