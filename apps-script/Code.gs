/**
 * VIQI 75 Years' Celebration in AZ — registration backend.
 *
 * Lives inside the private Google Sheet (Extensions → Apps Script), so the
 * responses never touch the public GitHub repo. See apps-script/SETUP.md.
 *
 * Each registration is stored under the ID vnsc_az_75_<name>_<ssc>_<hsc>_<phone>
 * (see makeId). Submitting again with the same name, batches and phone number
 * updates that row instead of adding a new one.
 *
 * Payment check: fill in "Zelle Confirmation #" (and optionally "Amount
 * Received ($)") for a row, then tick "Payment Confirmed". That sends the
 * "Registration Confirmed" email once and stamps "Confirmation Email Sent At".
 */

const ORGANIZER_EMAIL = 'viqis.in.az@gmail.com';
const ZELLE_RECIPIENT = 'Sabira Enayet';
const ZELLE_PHONE = '(480) 543-9295';
const EVENT_NAME = 'VIQI 75 Years’ Celebration in AZ';
const SHEET_NAME = 'Registrations';

// Keep in sync with the prices shown in index.html. Totals are recalculated
// here so an edited page can't change what gets recorded.
const CATEGORIES = [
  { field: 'Viqi Alumni Count',   label: 'VIQI Alumni',                               price: 150 },
  { field: 'Student/Guest Count', label: 'VIQI Alumni Students / Guests over age 10', price: 75 },
  { field: 'Child Count',         label: 'Children ages 2–10',                        price: 50 },
];
const MAX_PER_CATEGORY = 10;
const MAX_CONTRIBUTION = 100000;

// Sheet columns, in the same order as the form. Rows are written by header
// name, and getSheet() rearranges an older sheet to match this list, so the
// order can change without misaligning anyone's answers.
// `was` lists earlier names for the same column.
const COLUMNS = [
  { header: 'Registration ID' },
  { header: 'Submitted At' },
  { header: 'Last Updated At' },
  { header: 'Full Name' },
  { header: 'Email' },
  { header: 'Phone Number' },
  { header: 'SSC Batch' },
  { header: 'HSC Batch' },
  { header: 'Meet & Greet Headcount (Jan 30)' },
  { header: 'Gala Lunch Headcount (Jan 31)' },
  { header: 'Category 1: VIQI Alumni ($150)',                          was: ['Viqi Alumni'] },
  { header: 'Category 2: VIQI Alumni Students / Guests over 10 ($75)', was: ['Student / Guest (above 10 yrs)'] },
  { header: 'Category 3: Children ages 2–10 ($50)',                    was: ['Child (above 2 yrs, under 10 yrs)'] },
  { header: 'Category 4: Children under 2 (Free)',                     was: ['Children under 2 (free)'] },
  { header: 'Additional Contribution ($)' },
  { header: 'Total Due ($)' },
  { header: 'Zelle Account Name' },
  // Filled in by the committee when checking Zelle (see onPaymentEdit).
  // Earlier versions of the form collected the confirmation # directly.
  { header: 'Zelle Confirmation #', was: ['Zelle Confirmation # (no longer collected)'], admin: true },
  { header: 'Amount Received ($)', admin: true },
  { header: 'Payment Confirmed', admin: true },
  { header: 'Confirmation Email Sent At', admin: true },
  // No longer on the form; kept at the end so older answers aren't lost
  { header: 'Zelle Phone Number (no longer collected)', was: ['Zelle Phone Number'], retired: true },
];
const HEADERS = COLUMNS.map(c => c.header);
const MAX_EVENT_HEADCOUNT = 40;

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};

    const name  = clean(p['Full Name']);
    const ssc   = clean(p['SSC Batch']);
    const hsc   = clean(p['HSC Batch']);
    const email = clean(p['Email']).toLowerCase();
    const phone = clean(p['Phone Number']);
    const zelleName = clean(p['Zelle Account Name']);

    if (!name || !ssc || !hsc || !phone || !zelleName) return reply({ ok: false, error: 'Please fill in all required fields.' });
    if (!/^[^\s@<>,;"]+@[^\s@<>,;"]+\.[^\s@<>,;"]+$/.test(email)) {
      return reply({ ok: false, error: 'Please enter a valid email address.' });
    }

    const counts = CATEGORIES.map(c => {
      const n = parseInt(p[c.field], 10);
      return Number.isFinite(n) ? Math.max(0, Math.min(MAX_PER_CATEGORY, n)) : 0;
    });
    if (counts.every(n => n === 0)) return reply({ ok: false, error: 'Please add at least one attendee.' });

    const headcount = field => {
      const n = parseInt(p[field], 10);
      return Number.isFinite(n) ? Math.max(0, Math.min(MAX_EVENT_HEADCOUNT, n)) : 0;
    };
    const meet = headcount('Meet and Greet Headcount');
    const gala = headcount('Gala Lunch Headcount');
    if (meet + gala === 0) {
      return reply({ ok: false, error: 'Please tell us how many people will join at least one of the two events.' });
    }

    const extra = Math.round(parseFloat(p['Additional Contribution']));
    const contribution = Number.isFinite(extra) ? Math.max(0, Math.min(MAX_CONTRIBUTION, extra)) : 0;
    const infantsRaw = parseInt(p['Children Under 2 Count'], 10);
    const infants = Number.isFinite(infantsRaw) ? Math.max(0, Math.min(MAX_PER_CATEGORY, infantsRaw)) : 0;
    const fees = counts.reduce((sum, n, i) => sum + n * CATEGORIES[i].price, 0);
    const total = fees + contribution;
    const id = makeId(name, ssc, hsc, phone);
    const now = new Date();

    // Serialize writes so two submissions at once can't clobber each other.
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let updated = false;
    try {
      const sheet = getSheet();
      const record = {
        'Registration ID': id,
        'Submitted At': now,
        'Last Updated At': now,
        'Full Name': name,
        'Email': email,
        'Phone Number': phone,
        'SSC Batch': ssc,
        'HSC Batch': hsc,
        'Meet & Greet Headcount (Jan 30)': meet,
        'Gala Lunch Headcount (Jan 31)': gala,
        'Category 1: VIQI Alumni ($150)': counts[0],
        'Category 2: VIQI Alumni Students / Guests over 10 ($75)': counts[1],
        'Category 3: Children ages 2–10 ($50)': counts[2],
        'Category 4: Children under 2 (Free)': infants,
        'Additional Contribution ($)': contribution,
        'Total Due ($)': total,
        'Zelle Account Name': zelleName,
        'Payment Confirmed': false,
      };
      const rowValues = HEADERS.map(h => (h in record ? safeCell(record[h]) : ''));

      const lastRow = lastRegistrationRow(sheet);
      const ids = lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]))
        : [];
      const existing = ids.indexOf(id);

      if (existing === -1) {
        // Write right under the last registration. (appendRow() skips past
        // empty rows that have checkbox formatting, landing at row 1000+.)
        const row = lastRow + 1;
        sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
        sheet.getRange(row, HEADERS.indexOf('Payment Confirmed') + 1).setDataValidation(checkboxRule());
      } else {
        const row = existing + 2;
        // Keep the original submission time, the committee's payment columns,
        // and any answers to retired questions
        const before = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
        COLUMNS.forEach((c, i) => {
          if (c.retired || c.admin || c.header === 'Submitted At') rowValues[i] = before[i];
        });
        sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
        updated = true;
      }
    } finally {
      lock.releaseLock();
    }

    // The row is already saved, so an email failure (e.g. daily quota) shouldn't fail the registration.
    let emailSent = true;
    try {
      sendConfirmation({ id, name, ssc, hsc, email, phone, meet, gala, zelleName, counts, infants, contribution, total, updated });
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
    .join('') + (r.infants > 0
      ? `<tr><td>Children under 2 (${r.infants})</td><td align="right">Free</td></tr>`
      : '') + (r.contribution > 0
      ? `<tr><td>Additional contribution — thank you!</td><td align="right">$${r.contribution.toLocaleString()}</td></tr>`
      : '');

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a18;max-width:560px">
      <h2 style="margin:0 0 8px">${esc(EVENT_NAME)}</h2>
      <p>Dear ${esc(r.name)},</p>
      <p>Thank you for registering! ${r.updated ? 'Your registration has been <strong>updated</strong> with the details below.' : 'We have received your registration details below.'}</p>
      <p style="background:#fef3c7;border-left:4px solid #a16207;padding:10px 14px"><strong>Your registration is in progress — it is not confirmed yet.</strong></p>
      <p><strong>Registration ID:</strong> ${esc(r.id)}</p>
      <table cellpadding="6" style="border-collapse:collapse;width:100%;border:1px solid #e3e3de">
        <tr><td>Name</td><td align="right">${esc(r.name)}</td></tr>
        <tr><td>Email</td><td align="right">${esc(r.email)}</td></tr>
        <tr><td>Phone Number</td><td align="right">${esc(r.phone)}</td></tr>
        <tr><td>SSC Batch</td><td align="right">${esc(r.ssc)}</td></tr>
        <tr><td>HSC Batch</td><td align="right">${esc(r.hsc)}</td></tr>
        <tr><td>Meet, Greet &amp; Reminisce (Jan 30)</td><td align="right">${r.meet} ${r.meet === 1 ? 'person' : 'people'}</td></tr>
        <tr><td>Gala Lunch (Jan 31)</td><td align="right">${r.gala} ${r.gala === 1 ? 'person' : 'people'}</td></tr>
        <tr><td>Zelle Account Name</td><td align="right">${esc(r.zelleName)}</td></tr>
        ${lines}
        <tr style="border-top:2px solid #e3e3de"><td><strong>Total Due</strong></td><td align="right"><strong>$${r.total.toLocaleString()}</strong></td></tr>
      </table>
      <h3 style="margin:20px 0 6px">Next step: complete your Zelle payment</h3>
      <p>Please send your Total Due of <strong>$${r.total.toLocaleString()}</strong> via Zelle to:</p>
      <p style="margin-left:14px">Zelle Recipient: <strong>${esc(ZELLE_RECIPIENT)}</strong><br>Zelle Phone Number: <strong>${esc(ZELLE_PHONE)}</strong></p>
      <div style="background:#e8eefb;border:2px solid #1d4fa8;border-radius:8px;padding:12px 16px;margin:12px 0">
        <p style="margin:0 0 6px;color:#1d4fa8;font-weight:bold">📝 Important: add a message in Zelle</p>
        <p style="margin:0 0 8px">Please make sure to add your <strong>name, SSC batch and HSC batch</strong> in the optional message field while sending the money through Zelle. This will help identify and confirm your registration. Thank you!</p>
        <p style="margin:0">Your message: <strong>${esc(r.name)}, SSC ${esc(r.ssc)}, HSC ${esc(r.hsc)}</strong></p>
      </div>
      <p><strong>Your Zelle payment is the official confirmation of your registration and attendance.</strong> Please save your Zelle confirmation number, as it will serve as your registration record.</p>
      <p><strong>We will send you a confirmation email once we have received and confirmed your Zelle payment.</strong> Until then, your registration remains in progress.</p>
      <p>If anything above is wrong, simply submit the form again with the same full name, SSC batch, HSC batch and phone number and it will replace this registration, or reply to this email.</p>
      <p>Warmly,<br>VIQI 75 Years’ Celebration in AZ Planning Committee</p>
    </div>`;

  MailApp.sendEmail({
    to: r.email,
    cc: ORGANIZER_EMAIL,
    replyTo: ORGANIZER_EMAIL,
    name: EVENT_NAME,
    subject: `Registration in progress${r.updated ? ' (updated)' : ''}: ${EVENT_NAME} (${r.id})`,
    htmlBody: html,
  });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  migrateLayout(sheet);
  tidyRows(sheet);
  migrateIds(sheet);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  addCheckboxes(sheet);
  return sheet;
}

/** Row number of the last registration (by Registration ID), or 1 if none. */
function lastRegistrationRow(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

function checkboxRule() {
  return SpreadsheetApp.newDataValidation().requireCheckbox().build();
}

// Checkboxes go only on rows that hold a registration. Formatting empty rows
// (or using insertCheckboxes(), which writes FALSE into every cell) makes
// Google treat them as used, which pushed new registrations to row 1000+.
function addCheckboxes(sheet) {
  const col = HEADERS.indexOf('Payment Confirmed') + 1;
  const lastRow = lastRegistrationRow(sheet);
  if (sheet.getMaxRows() > lastRow) {
    sheet.getRange(lastRow + 1, col, sheet.getMaxRows() - lastRow, 1).clearDataValidations();
  }
  if (lastRow >= 2) sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(checkboxRule());
}

/**
 * Clears rows that hold only leftovers (a blank ID like "vnsc_az_75", an
 * unticked FALSE checkbox, or nothing) and moves the real registrations up so
 * they sit together under the header. Saves a backup tab first. Does nothing
 * once the sheet is tidy.
 */
function tidyRows(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return;
  const width = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, last - 1, width).getValues();
  const isJunk = r => r.every(v => v === '' || v === false || String(v).trim() === 'vnsc_az_75');
  const keep = data.filter(r => !isJunk(r));
  if (keep.length === data.length) return;

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.copyTo(sheet.getParent()).setName(SHEET_NAME + ' backup ' + stamp);

  sheet.getRange(2, 1, last - 1, width).clearContent().clearDataValidations();
  if (keep.length) sheet.getRange(2, 1, keep.length, width).setValues(keep);
  console.log('Removed ' + (data.length - keep.length) + ' empty rows; ' + keep.length + ' registrations kept. Backup: "' + SHEET_NAME + ' backup ' + stamp + '"');
}

/**
 * If the sheet's columns don't match HEADERS, rearrange every row to match,
 * moving each value by its column name. Saves a backup copy of the tab first.
 * Columns it doesn't recognize (e.g. your own "Paid?" notes) move to the end.
 * Only values move; rows stay put, so row colors and highlights stay with them.
 */
function migrateLayout(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return; // brand-new sheet

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const oldHeaders = data[0].map(h => String(h).trim());
  if (HEADERS.every((h, i) => oldHeaders[i] === h)) return; // already in order

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.copyTo(sheet.getParent()).setName(SHEET_NAME + ' backup ' + stamp);

  const source = COLUMNS.map(c => {
    const names = [c.header].concat(c.was || []);
    for (const n of names) {
      const i = oldHeaders.indexOf(n);
      if (i !== -1) return i;
    }
    return -1; // new column: starts empty
  });
  const used = new Set(source.filter(i => i !== -1));
  const extras = oldHeaders.map((h, i) => i).filter(i => !used.has(i) && oldHeaders[i] !== '');

  const headers = HEADERS.concat(extras.map(i => oldHeaders[i]));
  const rows = data.slice(1).map(r => source.map(i => (i === -1 ? '' : r[i])).concat(extras.map(i => r[i])));

  sheet.clearContents();
  // Checkboxes belong to a column name, not a position; getSheet() re-adds them.
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  console.log('Rearranged ' + rows.length + ' rows to the new column order. Backup: "' + SHEET_NAME + ' backup ' + stamp + '"');
}

/**
 * "Jane  Doe", "2010", "2012", "(480) 555-0123" →
 * "vnsc_az_75_jane-doe_2010_2012_4805550123". Case, extra spaces, punctuation
 * and phone formatting don't matter, so small typing differences still match.
 */
function makeId(name, ssc, hsc, phone) {
  const slug = v => String(v).trim().toLowerCase().replace(/[^\p{L}\p{M}\p{N}]+/gu, '-').replace(/^-+|-+$/g, '');
  // Digits only; "+1 480…" and "480…" are the same US number
  let digits = String(phone == null ? '' : phone).replace(/\D/g, '');
  if (digits.length === 11 && digits[0] === '1') digits = digits.slice(1);
  return ['vnsc_az_75', slug(name), slug(ssc), slug(hsc), digits].filter(Boolean).join('_');
}

/** Bring every row's ID up to the current format (earlier rows used other formats). */
function migrateIds(sheet) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const data = sheet.getRange(2, 1, n, HEADERS.length).getValues();
  const at = h => HEADERS.indexOf(h);
  let changed = false;
  const ids = data.map(r => {
    if (!String(r[at('Full Name')]).trim()) return [r[0]]; // not a registration
    const id = makeId(r[at('Full Name')], r[at('SSC Batch')], r[at('HSC Batch')], r[at('Phone Number')]);
    if (id !== String(r[0])) changed = true;
    return [id];
  });
  if (changed) sheet.getRange(2, 1, n, 1).setValues(ids);
}

/**
 * Run from the editor to create the sheet (or reorder its columns), add the
 * Payment Confirmed checkboxes, install the checkbox trigger, and grant
 * permissions. Safe to run again.
 */
function setup() {
  getSheet();

  // Sending email needs an "installable" edit trigger; a plain onEdit() can't.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onPaymentEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onPaymentEdit').forSpreadsheet(ss).onEdit().create();

  console.log('Ready. Remaining email quota today: ' + MailApp.getRemainingDailyQuota());
}

/**
 * Runs whenever someone edits the sheet. When "Payment Confirmed" is ticked
 * on a row that hasn't been emailed yet, sends the "Registration Confirmed"
 * email and records when it went out. Untick + clear the timestamp to resend.
 */
function onPaymentEdit(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SHEET_NAME) return;

  const col = HEADERS.indexOf('Payment Confirmed') + 1;
  if (col < range.getColumn() || col > range.getLastColumn()) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    for (let row = Math.max(2, range.getRow()); row <= range.getLastRow(); row++) {
      confirmRow(range.getSheet(), row);
    }
  } finally {
    lock.releaseLock();
  }
}

/** Adds a "VNSC 75" menu to the sheet each time it is opened. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('VNSC 75')
    .addItem('Send confirmation emails for ticked rows', 'sendPendingConfirmations')
    .addToUi();
}

/**
 * Menu item (or run from the editor): sends the "Registration Confirmed"
 * email for every ticked row that hasn't been emailed yet, then shows what
 * happened. A backup to the checkbox trigger.
 */
function sendPendingConfirmations() {
  const sheet = getSheet();
  const results = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    for (let row = 2; row <= lastRegistrationRow(sheet); row++) {
      const result = confirmRow(sheet, row);
      if (result) results.push('Row ' + row + ': ' + result);
    }
  } finally {
    lock.releaseLock();
  }
  const summary = results.length ? results.join('\n') : 'No ticked rows are waiting for an email.';
  console.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch (e) { /* run from the editor: see the log */ }
}

/**
 * Sends the confirmed email for one row if Payment Confirmed is ticked and it
 * hasn't been sent yet. Returns a short description, or '' if nothing to do.
 */
function confirmRow(sheet, row) {
  const col = HEADERS.indexOf('Payment Confirmed') + 1;
  const values = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const r = {};
  HEADERS.forEach((h, i) => { r[h] = values[i]; });

  const ticked = r['Payment Confirmed'] === true || String(r['Payment Confirmed']).toUpperCase() === 'TRUE';
  if (!ticked || r['Confirmation Email Sent At']) return '';

  const who = r['Full Name'] + ' <' + r['Email'] + '>';
  const box = sheet.getRange(row, col);
  if (!String(r['Zelle Confirmation #']).trim()) {
    box.setValue(false).setNote('Enter the Zelle Confirmation # for this row first, then tick again.');
    return 'skipped ' + who + ': no Zelle Confirmation # yet (box unticked).';
  }
  try {
    sendPaymentConfirmed(r);
    box.setNote(null);
    sheet.getRange(row, HEADERS.indexOf('Confirmation Email Sent At') + 1).setValue(new Date());
    return 'sent to ' + who + '.';
  } catch (err) {
    console.error('Payment confirmation email failed for row ' + row + ': ' + err);
    box.setValue(false).setNote('Email could not be sent (' + err.message + '). Tick again to retry.');
    return 'FAILED for ' + who + ': ' + err.message;
  }
}

function sendPaymentConfirmed(r) {
  const money = v => '$' + Number(v || 0).toLocaleString();
  const people = n => `${Number(n) || 0} ${Number(n) === 1 ? 'person' : 'people'}`;
  const received = r['Amount Received ($)'] !== '' ? r['Amount Received ($)'] : r['Total Due ($)'];
  const contribution = Number(r['Additional Contribution ($)']) || 0;

  const counts = [
    ['VIQI Alumni ($150 each)', r['Category 1: VIQI Alumni ($150)']],
    ['VIQI Alumni Students / Guests over 10 ($75 each)', r['Category 2: VIQI Alumni Students / Guests over 10 ($75)']],
    ['Children ages 2–10 ($50 each)', r['Category 3: Children ages 2–10 ($50)']],
    ['Children under 2 (free)', r['Category 4: Children under 2 (Free)']],
  ].filter(([, n]) => Number(n) > 0)
   .map(([label, n]) => `<tr><td>${esc(label)}</td><td align="right">${Number(n)}</td></tr>`)
   .join('');

  const row = (label, value) => `<tr><td>${label}</td><td align="right">${value}</td></tr>`;
  const table = inner => `<table cellpadding="6" style="border-collapse:collapse;width:100%;border:1px solid #e3e3de;margin:0 0 16px">${inner}</table>`;
  const heading = t => `<h3 style="margin:18px 0 6px">${t}</h3>`;

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a18;max-width:560px">
      <h2 style="margin:0 0 8px">${esc(EVENT_NAME)}</h2>
      <p>Dear ${esc(r['Full Name'])},</p>
      <p style="background:#dcfce7;border-left:4px solid #166534;padding:10px 14px"><strong>Great news! We have received your Zelle payment, and your registration is now CONFIRMED.</strong> Thank you!</p>
      ${heading('Registration details')}
      ${table(
        row('Registration ID', esc(r['Registration ID'])) +
        row('Name', esc(r['Full Name'])) +
        row('SSC / HSC Batch', `${esc(r['SSC Batch'])} / ${esc(r['HSC Batch'])}`)
      )}
      ${heading('Your gatherings')}
      ${table(
        row('January 30, 2027 — Meet, Greet &amp; Reminisce', people(r['Meet & Greet Headcount (Jan 30)'])) +
        row('January 31, 2027 — Gala Lunch', people(r['Gala Lunch Headcount (Jan 31)']))
      )}
      ${heading('Participation')}
      ${table(counts + (contribution > 0 ? row('Additional contribution', money(contribution)) : ''))}
      ${heading('Payment received')}
      ${table(
        row('Amount received', `<strong>${esc(money(received))}</strong>`) +
        row('Zelle confirmation #', esc(r['Zelle Confirmation #']))
      )}
      ${contribution > 0 ? '<p>Thank you so much for your generous additional contribution. It will go a long way toward making this 75-year celebration a successful and memorable event for everyone.</p>' : ''}
      <p>Please keep this email and your Zelle confirmation number as your registration record.</p>
      ${heading('What’s next')}
      <p>We will share the venue, timings and other event details closer to the date. If any of your plans change, or if anything above looks incorrect, simply reply to this email and let us know.</p>
      <p>We are so looking forward to reconnecting, reminiscing, and celebrating 75 years of VNSC together with you and your family!</p>
      <p>Warmly,<br>VIQI 75 Years’ Celebration in AZ Planning Committee</p>
    </div>`;

  MailApp.sendEmail({
    to: String(r['Email']),
    cc: ORGANIZER_EMAIL,
    replyTo: ORGANIZER_EMAIL,
    name: EVENT_NAME,
    subject: `Registration Confirmed: ${EVENT_NAME} (${r['Registration ID']})`,
    htmlBody: html,
  });
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
