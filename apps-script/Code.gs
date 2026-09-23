/**
 * RSVP backend for the SaHi wedding site — runs as a Google Apps Script web app
 * bound to a Google Sheet. The sheet tab "RSVPs" is the guest list.
 *
 * Setup is in README.md. The admin password is NOT in this file: set it in
 * Project Settings → Script Properties as ADMIN_PASSWORD.
 *
 * Every request is a POST with a JSON body (sent as text/plain so browsers skip
 * the CORS preflight Apps Script can't answer):
 *   { action: "rsvp", name, contact, status, adults, kids, message, website }
 *   { action: "list",   password }            → { rsvps: [...] }
 *   { action: "delete", password, id }
 */

// The spreadsheet the replies go to (the long ID in its URL, between /d/ and /edit).
// Leave empty when the script is opened from the sheet via Extensions → Apps Script.
const SHEET_ID = '1e6gmML4902r1Cwv2XBySAriMzI0Xl_1cBF9HLeZ3zas';
const SHEET = 'RSVPs';
const HEADERS = ['ID', 'Name', 'Phone / Email', 'Status', 'Adults', 'Children',
                 'Message', 'First replied', 'Last updated', 'Key'];
const STATUSES = ['Attending', 'Might Attend', "Can't Make It"];

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ error: 'Bad request' }); }
  try {
    switch (body.action) {
      case 'rsvp':   return out(saveRsvp(body));
      case 'list':   checkPassword(body.password); return out({ rsvps: listRsvps() });
      case 'delete': checkPassword(body.password); return out(deleteRsvp(body.id));
      default:       return out({ error: 'Unknown action' });
    }
  } catch (err) {
    return out({ error: err.message });
  }
}

function doGet() {
  return out({ ok: true, service: 'SaHi RSVP' });
}

// ── guests ────────────────────────────────────────────────────────────────

function saveRsvp(b) {
  if (b.website) return { ok: true };                 // honeypot: bots fill it, people never see it

  const clean = (v, max) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  const name = clean(b.name, 80);
  const contact = clean(b.contact, 120);
  const message = String(b.message == null ? '' : b.message).trim().slice(0, 1000);
  if (!name) throw new Error('Please add your name.');
  if (STATUSES.indexOf(b.status) < 0) throw new Error('Please choose whether you can attend.');

  const coming = b.status !== "Can't Make It";
  const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, parseInt(v, 10) || 0));
  const adults = coming ? clamp(b.adults, 1, 15) : 0;
  const kids = coming ? clamp(b.kids, 0, 15) : 0;

  // Same person replying again updates their row: emails match case-insensitively,
  // phones by their last 10 digits (so +1 / +91 don't matter), otherwise by name.
  const digits = contact.replace(/\D/g, '');
  const key = contact.indexOf('@') >= 0 ? contact.toLowerCase()
            : digits.length >= 7 ? 'tel:' + digits.slice(-10)
            : 'name:' + name.toLowerCase();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const now = new Date();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][9] === key) {
        const keepMsg = message || data[i][6];
        sh.getRange(i + 1, 2, 1, 9).setValues([[
          text(name), text(contact), b.status, adults, kids, text(keepMsg), data[i][7], now, key]]);
        return { ok: true, updated: true };
      }
    }
    sh.appendRow([now.getTime(), text(name), text(contact), b.status, adults, kids, text(message), now, now, key]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── admin ─────────────────────────────────────────────────────────────────

function checkPassword(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 10) throw new Error('Too many wrong passwords — try again in 10 minutes.');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw new Error('ADMIN_PASSWORD is not set in Script Properties.');
  if (String(pw || '') !== real) {
    cache.put('fails', String(fails + 1), 600);
    Utilities.sleep(500);
    throw new Error('Wrong password');
  }
}

function listRsvps() {
  const data = sheet().getDataRange().getValues().slice(1);
  const iso = d => d instanceof Date ? d.toISOString() : String(d || '');
  const untext = v => String(v).replace(/^'/, '');
  return data.filter(r => r[0] !== '').map(r => ({
    id: Number(r[0]), name: untext(r[1]), contact: untext(r[2]), status: r[3],
    adults: Number(r[4]) || 0, kids: Number(r[5]) || 0, message: untext(r[6]),
    created_at: iso(r[7]), updated_at: iso(r[8]),
  })).sort((a, b) => b.updated_at.localeCompare(a.updated_at));
}

function deleteRsvp(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (Number(data[i][0]) === Number(id)) { sh.deleteRow(i + 1); return { ok: true }; }
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── helpers ───────────────────────────────────────────────────────────────

function sheet() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f5e8cc');
    sh.hideColumns(10);                                // Key: bookkeeping only
    sh.getRange('B:C').setNumberFormat('@');          // keep leading zeros / + in phone numbers
  }
  return sh;
}

// Guests type free text; a leading = + - @ would run as a formula in Sheets.
function text(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
