/**
 * VIQI 75 Years' Celebration in AZ — registration backend.
 *
 * Lives inside the private Google Sheet (Extensions → Apps Script), so the
 * responses never touch the public GitHub repo. See apps-script/SETUP.md.
 *
 * Each registration is stored under the ID vnsc_75_az_<email>. Submitting
 * again with the same email updates that row instead of adding a new one.
 */

const ORGANIZER_EMAIL = 'viqis.in.az@gmail.com';
const EVENT_NAME = 'VIQI 75 Years’ Celebration in AZ';
const SHEET_NAME = 'Registrations';

// Keep in sync with the prices shown in index.html. Totals are recalculated
// here so an edited page can't change what gets recorded.
const CATEGORIES = [
  { field: 'Viqi Alumni Count',   label: 'Viqi Alumni',                         price: 150 },
  { field: 'Student/Guest Count', label: 'Student / Guest (above 10 yrs)',      price: 75 },
  { field: 'Child Count',         label: 'Child (above 2 yrs, under 10 yrs)',   price: 50 },
];
const MAX_PER_CATEGORY = 10;

const HEADERS = [
  'Registration ID', 'Submitted At', 'Last Updated At',
  'Full Name', 'SSC Batch', 'HSC Batch', 'Email',
  ...CATEGORIES.map(c => c.label),
  'Total Due ($)',
];

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};

    const name  = clean(p['Full Name']);
    const ssc   = clean(p['SSC Batch']);
    const hsc   = clean(p['HSC Batch']);
    const email = clean(p['Email']).toLowerCase();

    if (!name || !ssc || !hsc) return reply({ ok: false, error: 'Please fill in all required fields.' });
    if (!/^[^\s@<>,;"]+@[^\s@<>,;"]+\.[^\s@<>,;"]+$/.test(email)) {
      return reply({ ok: false, error: 'Please enter a valid email address.' });
    }

    const counts = CATEGORIES.map(c => {
      const n = parseInt(p[c.field], 10);
      return Number.isFinite(n) ? Math.max(0, Math.min(MAX_PER_CATEGORY, n)) : 0;
    });
    if (counts.every(n => n === 0)) return reply({ ok: false, error: 'Please add at least one attendee.' });

    const total = counts.reduce((sum, n, i) => sum + n * CATEGORIES[i].price, 0);
    const id = 'vnsc_75_az_' + email;
    const now = new Date();

    // Serialize writes so two submissions at once can't clobber each other.
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let updated = false;
    try {
      const sheet = getSheet();
      const rowValues = [id, now, now, name, ssc, hsc, email, ...counts, total].map(safeCell);

      const ids = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0])
        : [];
      const existing = ids.indexOf(id);

      if (existing === -1) {
        sheet.appendRow(rowValues);
      } else {
        const row = existing + 2;
        rowValues[1] = sheet.getRange(row, 2).getValue(); // keep original submission time
        sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
        updated = true;
      }
    } finally {
      lock.releaseLock();
    }

    // The row is already saved, so an email failure (e.g. daily quota) shouldn't fail the registration.
    let emailSent = true;
    try {
      sendConfirmation({ id, name, ssc, hsc, email, counts, total, updated });
    } catch (mailErr) {
      console.error('Confirmation email failed for ' + id + ': ' + mailErr);
      emailSent = false;
    }

    return reply({ ok: true, id, total, updated, emailSent });
  } catch (err) {
    console.error(err);
    return reply({ ok: false, error: 'Something went wrong on our side. Please try again.' });
  }
}

function sendConfirmation(r) {
  const lines = CATEGORIES
    .map((c, i) => ({ c, n: r.counts[i] }))
    .filter(x => x.n > 0)
    .map(x => `<tr><td>${esc(x.c.label)} (${x.n} × $${x.c.price})</td><td align="right">$${(x.n * x.c.price).toLocaleString()}</td></tr>`)
    .join('');

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a18;max-width:560px">
      <h2 style="margin:0 0 8px">${esc(EVENT_NAME)}</h2>
      <p>Dear ${esc(r.name)},</p>
      <p>Thank you for registering! ${r.updated ? 'Your registration has been <strong>updated</strong> with the details below.' : 'Here are your registration details.'}</p>
      <p><strong>Registration ID:</strong> ${esc(r.id)}</p>
      <table cellpadding="6" style="border-collapse:collapse;width:100%;border:1px solid #e3e3de">
        <tr><td>Name</td><td align="right">${esc(r.name)}</td></tr>
        <tr><td>SSC Batch</td><td align="right">${esc(r.ssc)}</td></tr>
        <tr><td>HSC Batch</td><td align="right">${esc(r.hsc)}</td></tr>
        <tr><td>Email</td><td align="right">${esc(r.email)}</td></tr>
        ${lines}
        <tr style="border-top:2px solid #e3e3de"><td><strong>Total Due</strong></td><td align="right"><strong>$${r.total.toLocaleString()}</strong></td></tr>
      </table>
      <p>Children under 2 attend free of charge.</p>
      <p>Payment details will follow. If anything above is wrong, simply submit the form again with the same email address and it will replace this registration, or reply to this email.</p>
      <p>Warmly,<br>VIQI 75 Years’ Celebration in AZ Planning Committee</p>
    </div>`;

  MailApp.sendEmail({
    to: r.email,
    cc: ORGANIZER_EMAIL,
    replyTo: ORGANIZER_EMAIL,
    name: EVENT_NAME,
    subject: `${r.updated ? 'Updated registration' : 'Registration confirmed'}: ${EVENT_NAME} (${r.id})`,
    htmlBody: html,
  });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/** Run once from the editor to create the sheet and grant permissions. */
function setup() {
  getSheet();
  console.log('Ready. Remaining email quota today: ' + MailApp.getRemainingDailyQuota());
}

function clean(v) {
  return String(v == null ? '' : v).trim().slice(0, 200);
}

// Stop typed text like "=HYPERLINK(...)" from being run as a spreadsheet formula.
function safeCell(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function esc(v) {
  return String(v).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
