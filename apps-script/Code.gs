/**
 * LATOLANKATU 9 – laundry & sauna booking – backend (Google Apps Script)
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
 *   listSlots_          – works out which hours are free, per building and facility
 *   book_               – saves a booking (with a lock so two people can't take the same slot)
 *   checkReminders      – runs every 15 min: sends reminders and releases unconfirmed slots
 *   handleLink_         – what happens when someone clicks "Yes, I'll use it" / "Release my slot"
 */

// ---------- SETTINGS (change these to fit the buildings) ----------
const CONFIG = {
  SITE_NAME: 'Latolankatu 9',
  CITY: 'Joensuu',
  // Each building and the facilities it has. Add/remove buildings or facilities here.
  BUILDINGS: {
    A: ['laundry', 'sauna'],
    B: ['laundry', 'sauna'],
    C: ['laundry', 'sauna'],
    D: ['laundry', 'sauna'],
  },
  FACILITIES: { laundry: 'Laundry room', sauna: 'Sauna' },
  OPEN_HOUR: 7,                  // first slot starts 07:00
  CLOSE_HOUR: 21,                // last slot ends 21:00
  SLOT_HOURS: 1,                 // 07–08, 08–09, ... 20–21 (14 slots a day)
  DAYS_AHEAD: 14,                // residents can book two weeks ahead
  MAX_PER_APARTMENT: 2,          // max upcoming bookings per apartment, per facility
  REMINDER_HOURS_BEFORE: 2,      // reminder email is sent this long before the slot
  CONFIRM_CUTOFF_MINUTES: 30,    // not confirmed by this many minutes before = slot released
};

const TIMEZONE = 'Europe/Helsinki';
const SHEET_NAME = 'Bookings';
const HEADERS = ['id', 'date', 'start', 'building', 'facility', 'name', 'apartment', 'email',
                 'status', 'token', 'createdAt', 'reminderSentAt'];
// status can be: booked → confirmed | released (by the resident) | expired (not confirmed in time)

// ---------- ONE-TIME SETUP: run this once from the editor ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);

  // If an older version of the Bookings tab exists with different columns, keep it as a backup
  if (sh) {
    const current = sh.getRange(1, 1, 1, sh.getLastColumn() || 1).getValues()[0].join(',');
    if (current !== HEADERS.join(',')) {
      sh.setName(SHEET_NAME + ' (old ' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH-mm') + ')');
      sh = null;
    }
  }
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@'); // plain text, so Sheets doesn't turn dates into numbers
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
  if (p.action === 'slots') return json_(listSlots_(String(p.building || firstBuilding_()).toUpperCase()));
  if (p.action === 'confirm' || p.action === 'release') return page_(handleLink_(p));
  return json_({ ok: true, service: 'latolankatu-booking' });
}

function doPost(e) {
  rememberUrl_();
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  try { return json_(book_(data)); }
  catch (err) { return json_({ ok: false, error: 'Something went wrong. Please try again.' }); }
}

// ---------- FREE SLOTS for one building ----------
function listSlots_(building) {
  const facilities = CONFIG.BUILDINGS[building];
  if (!facilities) return { ok: false, error: 'Unknown building.' };

  const now = new Date();
  const taken = {}; // slot key → who booked it (first name + apartment only, since the page is public)
  readBookings_().filter(isActive_).forEach(b => {
    taken[key_(b.date, b.start, b.building, b.facility)] = {
      name: String(b.name).split(/\s+/)[0],
      apartment: b.building + ' ' + b.apartment,
      confirmed: b.status === 'confirmed',
    };
  });

  // Anchor at noon today so adding whole days never skips a date around clock changes
  const todayNoon = slotDate_(Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd'), '12:00');
  const days = [];
  for (let i = 0; i < CONFIG.DAYS_AHEAD; i++) {
    const day = new Date(todayNoon.getTime() + i * 86400000);
    const date = Utilities.formatDate(day, TIMEZONE, 'yyyy-MM-dd');
    const hours = [];
    for (let h = CONFIG.OPEN_HOUR; h + CONFIG.SLOT_HOURS <= CONFIG.CLOSE_HOUR; h += CONFIG.SLOT_HOURS) {
      const start = pad_(h) + ':00';
      if (slotDate_(date, start) <= now) continue; // already started
      // For each facility: null = free, otherwise who booked it
      const booked = {};
      facilities.forEach(f => { booked[f] = taken[key_(date, start, building, f)] || null; });
      hours.push({ start, end: end_(start), booked });
    }
    days.push({
      date,
      weekday: Utilities.formatDate(day, TIMEZONE, 'EEE'),
      day: Utilities.formatDate(day, TIMEZONE, 'd'),
      month: Utilities.formatDate(day, TIMEZONE, 'MMM'),
      label: Utilities.formatDate(day, TIMEZONE, 'EEEE d MMMM'),
      hours,
    });
  }
  return {
    ok: true,
    site: CONFIG.SITE_NAME,
    city: CONFIG.CITY,
    building,
    buildings: Object.keys(CONFIG.BUILDINGS),
    facilities: facilities.map(id => ({ id, name: CONFIG.FACILITIES[id] })),
    rules: {
      reminderHours: CONFIG.REMINDER_HOURS_BEFORE,
      cutoffMinutes: CONFIG.CONFIRM_CUTOFF_MINUTES,
      maxPerApartment: CONFIG.MAX_PER_APARTMENT,
    },
    days,
  };
}

// ---------- MAKE A BOOKING ----------
function book_(d) {
  const building = String(d.building || '').trim().toUpperCase();
  const facility = String(d.facility || '').trim();
  const name = String(d.name || '').trim();
  const apartment = String(d.apartment || '').trim().toUpperCase();
  const email = String(d.email || '').trim().toLowerCase();
  const date = String(d.date || '');
  const start = String(d.start || '');

  if (!CONFIG.BUILDINGS[building]) return { ok: false, error: 'Choose your building.' };
  if (CONFIG.BUILDINGS[building].indexOf(facility) < 0) return { ok: false, error: 'Choose the laundry room or the sauna.' };
  if (!name || !apartment) return { ok: false, error: 'Fill in your name and apartment number.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return { ok: false, error: 'Enter a valid email address.' };

  // The hour must be one we actually offer right now
  const day = listSlots_(building).days.find(x => x.date === date);
  const hour = day && day.hours.find(h => h.start === start);
  if (!hour) return { ok: false, error: 'That time is no longer available.' };

  // Lock: only one booking is processed at a time, so two people can't grab the same slot
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const active = readBookings_().filter(isActive_);
    if (active.some(b => key_(b.date, b.start, b.building, b.facility) === key_(date, start, building, facility))) {
      return { ok: false, error: 'Someone just booked this slot. Pick another time.' };
    }
    const now = new Date();
    const upcoming = active.filter(b => b.building === building && b.apartment === apartment &&
                                        b.facility === facility && slotDate_(b.date, b.start) > now);
    if (upcoming.length >= CONFIG.MAX_PER_APARTMENT) {
      return { ok: false, error: `Apartment ${building} ${apartment} already has ${CONFIG.MAX_PER_APARTMENT} upcoming ${CONFIG.FACILITIES[facility].toLowerCase()} bookings.` };
    }

    // Booking very close to the start? Then they clearly mean it – no reminder needed.
    const minsUntil = (slotDate_(date, start) - now) / 60000;
    const status = minsUntil <= CONFIG.REMINDER_HOURS_BEFORE * 60 ? 'confirmed' : 'booked';

    const b = {
      id: Utilities.getUuid().slice(0, 8), date, start, building, facility, name, apartment, email, status,
      token: Utilities.getUuid(), createdAt: now.toISOString(), reminderSentAt: '',
    };
    // Write the new row as plain text, so Sheets keeps "2026-11-05" and "18:00" exactly as they are
    const sh = sheet_();
    const range = sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length);
    range.setNumberFormat('@').setValues([HEADERS.map(h => b[h])]);
    SpreadsheetApp.flush(); // save now, so the page sees the booking straight away
    sendConfirmation_(b);
    return { ok: true, message: `Confirmation sent to ${email}.` };
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
          htmlBody: email_(b, `<p>You didn't confirm, so your slot <b>${when_(b)}</b> was released for your neighbours.</p>`),
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
      return `<p style="font-size:20px;margin:0 0 4px"><b>${when_(b)}</b></p><p style="margin:0 0 16px;color:#5b6b7b">Apartment ${esc_(b.building)} ${esc_(b.apartment)}</p>` +
        button_(actionUrl_(p.action, b) + '&go=1',
                isConfirm ? "Yes, I'll use it" : 'Release my slot',
                isConfirm ? '#23466e' : '#b42318');
    }

    // Step 2: do it
    if (p.action === 'confirm') {
      updateBooking_(b, { status: 'confirmed' });
      return `<p>Confirmed. <b>${when_(b)}</b> is yours.</p>`;
    }
    updateBooking_(b, { status: 'released' });
    return `<p>Released. <b>${when_(b)}</b> is now free for your neighbours. Thank you!</p>`;
  } finally {
    lock.releaseLock();
  }
}

// ---------- EMAILS ----------
function sendConfirmation_(b) {
  const note = b.status === 'confirmed' ? '' :
    `<p>You'll get a reminder ${CONFIG.REMINDER_HOURS_BEFORE} hours before. Confirm it there – if the slot isn't confirmed ${CONFIG.CONFIRM_CUTOFF_MINUTES} minutes before it starts, it's released for your neighbours.</p>`;
  MailApp.sendEmail({
    to: b.email,
    subject: `Booked: ${when_(b)}`,
    htmlBody: email_(b, `<p>Your booking is confirmed:</p>
      <p style="font-size:18px"><b>${when_(b)}</b><br>Apartment ${esc_(b.building)} ${esc_(b.apartment)}</p>
      ${note}
      <p>Can't make it? Free the slot for someone else:</p>
      ${button_(actionUrl_('release', b), 'Cancel my booking', '#b42318')}`),
  });
}

function sendReminder_(b) {
  const cutoff = Utilities.formatDate(
    new Date(slotDate_(b.date, b.start).getTime() - CONFIG.CONFIRM_CUTOFF_MINUTES * 60000), TIMEZONE, 'HH:mm');
  MailApp.sendEmail({
    to: b.email,
    subject: `Reminder: ${when_(b)} – are you using it?`,
    htmlBody: email_(b, `<p>Your slot starts soon: <b>${when_(b)}</b>.</p>
      <p>Confirm by <b>${cutoff}</b>, or the slot is released for your neighbours.</p>
      ${button_(actionUrl_('confirm', b), "Yes, I'll use it", '#23466e')}
      ${button_(actionUrl_('release', b), 'Release my slot', '#b42318')}`),
  });
}

function email_(b, inner) {
  return `<div style="font-family:system-ui,sans-serif;font-size:16px;line-height:1.5;color:#16212e;max-width:520px">
    <p style="color:#5b6b7b;margin:0 0 12px">${CONFIG.SITE_NAME}, ${CONFIG.CITY}</p>
    <p>Hi ${esc_(b.name)},</p>${inner}</div>`;
}

// ---------- SMALL HELPERS ----------
function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME); }
function firstBuilding_() { return Object.keys(CONFIG.BUILDINGS)[0]; }

function readBookings_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues().map((row, i) => {
    const b = { _row: i + 2 };
    HEADERS.forEach((h, j) => {
      let v = row[j];
      // Google Sheets sometimes turns "2026-11-05" or "18:00" into its own date values.
      // Turn them back into our text format so bookings always match their slot.
      if (v instanceof Date) {
        if (h === 'date') v = Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
        else if (h === 'start') v = Utilities.formatDate(v, TIMEZONE, 'HH:mm');
        else v = v.toISOString();
      }
      b[h] = String(v).trim();
    });
    if (/^\d:\d\d$/.test(b.start)) b.start = '0' + b.start; // "7:00" → "07:00"
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
function key_(date, start, building, facility) { return [date, start, building, facility].join('|'); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function end_(start) { return pad_(Number(start.split(':')[0]) + CONFIG.SLOT_HOURS) + ':00'; }
function slotDate_(date, start) { return Utilities.parseDate(date + ' ' + start, TIMEZONE, 'yyyy-MM-dd HH:mm'); }
function when_(b) {
  const day = Utilities.formatDate(slotDate_(b.date, b.start), TIMEZONE, 'EEE d MMM');
  return `${CONFIG.FACILITIES[b.facility] || b.facility}, building ${b.building}, ${day} ${b.start}–${end_(b.start)}`;
}

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
    `<div style="font-family:system-ui,sans-serif;max-width:480px;margin:40px auto;padding:0 16px;font-size:17px;line-height:1.5;color:#16212e">
      <p style="color:#5b6b7b;margin:0 0 16px">${CONFIG.SITE_NAME}, ${CONFIG.CITY}</p>${inner}</div>`)
    .setTitle('Latolankatu 9 booking')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
