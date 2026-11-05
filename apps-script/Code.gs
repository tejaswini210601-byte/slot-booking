/**
 * SLOT BOOKING – backend (Google Apps Script)
 *
 * This file is the "brain" of the system. It runs on Google's servers,
 * attached to a Google Sheet, which acts as our database.
 *
 *   Booking page (GitHub Pages)  --asks for slots / sends bookings-->  this script  --reads/writes-->  Google Sheet
 *                                                                        |
 *                                                                        +--> sends emails (confirmation, reminder, released)
 *
 * Main parts:
 *   doGet / doPost      – the "doors" the web page and email links knock on
 *   listSlots_          – works out which slots are free
 *   book_               – saves a booking (with a lock so two people can't take the same slot)
 *   checkReminders      – runs every 15 min: sends reminders and releases unconfirmed slots
 *   handleLink_         – what happens when someone clicks "Yes, I'll use it" / "Release my slot"
 */

// ---------- SETTINGS (change these to fit the building) ----------
const CONFIG = {
  BUILDING_NAME: 'Laundry room booking',
  RESOURCES: ['Laundry room'],   // add names here if several rooms/machines can be booked at the same time
  OPEN_HOUR: 7,                  // first slot starts 07:00
  CLOSE_HOUR: 21,                // last slot ends 21:00
  SLOT_HOURS: 1,                 // 07–08, 08–09, ... 20–21 (14 slots a day)
  DAYS_AHEAD: 7,                 // how many days residents can see and book
  MAX_PER_APARTMENT: 2,          // max upcoming bookings per apartment at once
  REMINDER_HOURS_BEFORE: 2,      // reminder email is sent this long before the slot
  CONFIRM_CUTOFF_MINUTES: 30,    // not confirmed by this many minutes before = slot released
};

const SHEET_NAME = 'Bookings';
const HEADERS = ['id', 'date', 'start', 'resource', 'name', 'apartment', 'email',
                 'status', 'token', 'createdAt', 'reminderSentAt'];
// status can be: booked → confirmed | released (by the resident) | expired (not confirmed in time)

// ---------- ONE-TIME SETUP: run this once from the editor ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange('A:K').setNumberFormat('@'); // store everything as plain text (stops Sheets turning dates into numbers)
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);

  // Create the 15-minute timer (delete an old one first so we never get two)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(15).create();
}

// ---------- DOORS: web requests come in here ----------
function doGet(e) {
  rememberUrl_();
  const p = (e && e.parameter) || {};
  if (p.action === 'slots') return json_(listSlots_());
  if (p.action === 'confirm' || p.action === 'release') return page_(handleLink_(p));
  return json_({ ok: true, service: 'slot-booking' });
}

function doPost(e) {
  rememberUrl_();
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  try { return json_(book_(data)); }
  catch (err) { return json_({ ok: false, error: 'Something went wrong. Please try again.' }); }
}

// ---------- FREE SLOTS ----------
function listSlots_() {
  const now = new Date();
  const taken = {};
  readBookings_().filter(isActive_).forEach(b => { taken[key_(b.date, b.start, b.resource)] = true; });

  // Anchor at noon today so adding whole days never skips a date around clock changes
  const todayNoon = slotDate_(Utilities.formatDate(now, tz_(), 'yyyy-MM-dd'), '12:00');
  const days = [];
  for (let i = 0; i < CONFIG.DAYS_AHEAD; i++) {
    const day = new Date(todayNoon.getTime() + i * 86400000);
    const date = Utilities.formatDate(day, tz_(), 'yyyy-MM-dd');
    const slots = [];
    for (let h = CONFIG.OPEN_HOUR; h + CONFIG.SLOT_HOURS <= CONFIG.CLOSE_HOUR; h += CONFIG.SLOT_HOURS) {
      const start = pad_(h) + ':00';
      if (slotDate_(date, start) <= now) continue; // already started
      CONFIG.RESOURCES.forEach(resource => {
        slots.push({ start, end: end_(start), resource, free: !taken[key_(date, start, resource)] });
      });
    }
    days.push({ date, label: Utilities.formatDate(day, tz_(), 'EEEE d MMM'), slots });
  }
  return { title: CONFIG.BUILDING_NAME, days };
}

// ---------- MAKE A BOOKING ----------
function book_(d) {
  const name = String(d.name || '').trim();
  const apartment = String(d.apartment || '').trim().toUpperCase();
  const email = String(d.email || '').trim().toLowerCase();
  const date = String(d.date || '');
  const start = String(d.start || '');
  const resource = String(d.resource || '');

  if (!name || !apartment) return { ok: false, error: 'Please fill in your name and apartment number.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Please enter a valid email address.' };

  // The slot must be one we actually offer right now
  const day = listSlots_().days.find(x => x.date === date);
  const slot = day && day.slots.find(s => s.start === start && s.resource === resource);
  if (!slot) return { ok: false, error: 'That slot is not available.' };

  // Lock: only one booking is processed at a time, so two people can't grab the same slot
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const active = readBookings_().filter(isActive_);
    if (active.some(b => key_(b.date, b.start, b.resource) === key_(date, start, resource))) {
      return { ok: false, error: 'Sorry, someone just booked that slot. Please pick another.' };
    }
    const now = new Date();
    const upcoming = active.filter(b => b.apartment === apartment && slotDate_(b.date, b.start) > now);
    if (upcoming.length >= CONFIG.MAX_PER_APARTMENT) {
      return { ok: false, error: `Apartment ${apartment} already has ${CONFIG.MAX_PER_APARTMENT} upcoming bookings.` };
    }

    // Booking very close to the start? Then they clearly mean it – no reminder needed.
    const minsUntil = (slotDate_(date, start) - now) / 60000;
    const status = minsUntil <= CONFIG.REMINDER_HOURS_BEFORE * 60 ? 'confirmed' : 'booked';

    const b = {
      id: Utilities.getUuid().slice(0, 8), date, start, resource, name, apartment, email, status,
      token: Utilities.getUuid(), createdAt: now.toISOString(), reminderSentAt: '',
    };
    sheet_().appendRow(HEADERS.map(h => b[h]));
    sendConfirmation_(b);
    return { ok: true, message: `Booked! A confirmation email was sent to ${email}.` };
  } finally {
    lock.releaseLock();
  }
}

// ---------- TIMER: runs every 15 minutes ----------
function checkReminders() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    readBookings_().forEach(b => {
      if (b.status !== 'booked') return;
      const mins = (slotDate_(b.date, b.start) - now) / 60000;

      if (!b.reminderSentAt && mins <= CONFIG.REMINDER_HOURS_BEFORE * 60 && mins > CONFIG.CONFIRM_CUTOFF_MINUTES) {
        sendReminder_(b);
        updateBooking_(b, { reminderSentAt: now.toISOString() });
      } else if (b.reminderSentAt && mins <= CONFIG.CONFIRM_CUTOFF_MINUTES) {
        updateBooking_(b, { status: 'expired' }); // slot shows as free on the page again
        MailApp.sendEmail({
          to: b.email,
          subject: `Slot released: ${when_(b)}`,
          htmlBody: `<p>Hi ${esc_(b.name)},</p><p>You didn't confirm, so your slot <b>${when_(b)}</b> was released for other residents.</p>`,
        });
      }
    });
  } finally {
    lock.releaseLock();
  }
}

// ---------- EMAIL BUTTON CLICKS ----------
function handleLink_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = readBookings_().find(x => x.id === p.id && x.token === p.token);
    if (!b) return '<p>Booking not found.</p>';
    if (!isActive_(b)) return '<p>This booking is no longer active.</p>';
    if (slotDate_(b.date, b.start) <= new Date()) return '<p>This slot has already started.</p>';

    // Step 1: show a button first. (Some email apps "pre-open" links to scan them –
    // this extra click stops a scanner from releasing a slot by accident.)
    if (!p.go) {
      const isConfirm = p.action === 'confirm';
      return `<p><b>${when_(b)}</b><br>Apartment ${esc_(b.apartment)}</p>` +
        button_(actionUrl_(p.action, b) + '&go=1',
                isConfirm ? "Yes, I'll use it" : 'Release my slot',
                isConfirm ? '#067647' : '#b42318');
    }

    // Step 2: do it
    if (p.action === 'confirm') {
      updateBooking_(b, { status: 'confirmed' });
      return `<p>Thanks! Your slot <b>${when_(b)}</b> is confirmed.</p>`;
    }
    updateBooking_(b, { status: 'released' });
    return `<p>Done – your slot <b>${when_(b)}</b> is now free for other residents. Thank you!</p>`;
  } finally {
    lock.releaseLock();
  }
}

// ---------- EMAILS ----------
function sendConfirmation_(b) {
  const note = b.status === 'confirmed' ? '' :
    `<p>You'll get a reminder ${CONFIG.REMINDER_HOURS_BEFORE} hours before. Please confirm it – if the slot isn't confirmed ${CONFIG.CONFIRM_CUTOFF_MINUTES} minutes before the start, it's released for other residents.</p>`;
  MailApp.sendEmail({
    to: b.email,
    subject: `Booking confirmed: ${when_(b)}`,
    htmlBody: `<p>Hi ${esc_(b.name)},</p>
      <p>Your booking is confirmed:<br><b>${when_(b)}</b><br>Apartment ${esc_(b.apartment)}</p>
      ${note}
      <p>Can't make it? Please free the slot for others:</p>
      ${button_(actionUrl_('release', b), 'Cancel my booking', '#b42318')}`,
  });
}

function sendReminder_(b) {
  const cutoff = Utilities.formatDate(
    new Date(slotDate_(b.date, b.start).getTime() - CONFIG.CONFIRM_CUTOFF_MINUTES * 60000), tz_(), 'HH:mm');
  MailApp.sendEmail({
    to: b.email,
    subject: `Reminder: ${when_(b)} – are you using it?`,
    htmlBody: `<p>Hi ${esc_(b.name)},</p>
      <p>Your slot starts soon: <b>${when_(b)}</b>.</p>
      <p>Please confirm by <b>${cutoff}</b>, or the slot will be released for other residents.</p>
      ${button_(actionUrl_('confirm', b), "Yes, I'll use it", '#067647')}
      ${button_(actionUrl_('release', b), 'Release my slot', '#b42318')}`,
  });
}

// ---------- SMALL HELPERS ----------
function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME); }

function readBookings_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues().map((row, i) => {
    const b = { _row: i + 2 };
    HEADERS.forEach((h, j) => { b[h] = String(row[j]); });
    return b;
  });
}

function updateBooking_(b, fields) {
  const sh = sheet_();
  Object.keys(fields).forEach(k => {
    sh.getRange(b._row, HEADERS.indexOf(k) + 1).setValue(fields[k]);
    b[k] = fields[k];
  });
}

function isActive_(b) { return b.status === 'booked' || b.status === 'confirmed'; }
function key_(date, start, resource) { return date + '|' + start + '|' + resource; }
function tz_() { return Session.getScriptTimeZone(); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function end_(start) { return pad_(Number(start.split(':')[0]) + CONFIG.SLOT_HOURS) + ':00'; }
function slotDate_(date, start) { return Utilities.parseDate(date + ' ' + start, tz_(), 'yyyy-MM-dd HH:mm'); }
function when_(b) { return `${b.date} ${b.start}–${end_(b.start)} (${b.resource})`; }

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Remember this web app's public address, so timer emails can link back to it
function rememberUrl_() {
  const url = ScriptApp.getService().getUrl();
  if (url && /\/exec$/.test(url)) PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', url);
}
function actionUrl_(action, b) {
  const base = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl();
  return `${base}?action=${action}&id=${b.id}&token=${b.token}`;
}

function button_(url, label, color) {
  return `<a href="${url}" target="_top" style="display:inline-block;padding:12px 18px;margin:6px 8px 6px 0;background:${color};color:#fff;text-decoration:none;border-radius:8px;font-weight:600">${label}</a>`;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function page_(inner) {
  return HtmlService.createHtmlOutput(
    `<div style="font-family:system-ui,sans-serif;max-width:480px;margin:40px auto;padding:0 16px;font-size:17px;line-height:1.5">${inner}</div>`)
    .setTitle('Slot booking')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
