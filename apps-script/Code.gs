/**
 * RSVP + gift wishlist backend for the SaHi wedding site — a Google Apps Script
 * web app writing to a Google Sheet. Tab "RSVPs" is the guest list, tab "Gifts"
 * the wishlist.
 *
 * Setup is in README.md. The admin password is NOT in this file: set it in
 * Project Settings → Script Properties as ADMIN_PASSWORD.
 *
 * Every request is a POST with a JSON body (sent as text/plain so browsers skip
 * the CORS preflight Apps Script can't answer):
 *   guests  { action: "rsvp", name, contact, status, adults, kids, message, website }
 *           { action: "gifts" }                        → { gifts: [...] }  (no names of who reserved)
 *           { action: "claim", id, name, contact }     → { gifts: [...] }
 *   admin   { action: "list", password }               → { rsvps: [...], gifts: [...] }
 *           { action: "delete", password, id }
 *           { action: "addGift", password, item, link, price, note }
 *           { action: "deleteGift", password, id }
 *           { action: "releaseGift", password, id }
 */

// The spreadsheet the replies go to (the long ID in its URL, between /d/ and /edit).
// Leave empty when the script is opened from the sheet via Extensions → Apps Script.
const SHEET_ID = '1e6gmML4902r1Cwv2XBySAriMzI0Xl_1cBF9HLeZ3zas';
const RSVP_TAB = { name: 'RSVPs', headers: ['ID', 'Name', 'Phone / Email', 'Status', 'Adults', 'Children',
                   'Message', 'First replied', 'Last updated', 'Key'], textCols: 'B:C', hideCol: 10 };
const GIFT_TAB = { name: 'Gifts', headers: ['ID', 'Item', 'Link', 'Price', 'Note',
                   'Reserved by', 'Reserved contact', 'Reserved at', 'Added'], textCols: 'B:G' };
const STATUSES = ['Attending', 'Might Attend', "Can't Make It"];

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ error: 'Bad request' }); }
  try {
    const admin = () => checkPassword(body.password);
    switch (body.action) {
      case 'rsvp':        return out(saveRsvp(body));
      case 'gifts':       return out({ gifts: listGifts(false) });
      case 'claim':       return out(claimGift(body));
      case 'list':        admin(); return out({ rsvps: listRsvps(), gifts: listGifts(true) });
      case 'delete':      admin(); return out(deleteRow(RSVP_TAB, body.id));
      case 'addGift':     admin(); return out(addGift(body));
      case 'deleteGift':  admin(); return out(deleteRow(GIFT_TAB, body.id));
      case 'releaseGift': admin(); return out(releaseGift(body.id));
      default:            return out({ error: 'Unknown action' });
    }
  } catch (err) {
    return out({ error: err.message });
  }
}

function doGet() {
  return out({ ok: true, service: 'SaHi RSVP' });
}

// ── guests: RSVP ──────────────────────────────────────────────────────────

function saveRsvp(b) {
  if (b.website) return { ok: true };                 // honeypot: bots fill it, people never see it

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

  return locked(() => {
    const sh = tab(RSVP_TAB);
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
  });
}

// ── guests: gift wishlist ─────────────────────────────────────────────────

function listGifts(forAdmin) {
  return tab(GIFT_TAB).getDataRange().getValues().slice(1).filter(r => r[0] !== '').map(r => {
    const g = { id: Number(r[0]), item: untext(r[1]), link: untext(r[2]), price: untext(r[3]),
                note: untext(r[4]), reserved: String(r[5]) !== '' };
    if (forAdmin) Object.assign(g, { reservedBy: untext(r[5]), reservedContact: untext(r[6]), reservedAt: iso(r[7]) });
    return g;
  });
}

function claimGift(b) {
  const name = clean(b.name, 80);
  if (!name) throw new Error('Please RSVP first so we know who is gifting.');
  locked(() => {
    const sh = tab(GIFT_TAB);
    const i = rowIndex(sh, b.id);
    if (i < 0) throw new Error('That gift is no longer on the list.');
    if (String(sh.getRange(i, 6).getValue()) !== '') throw new Error('Someone has just reserved this gift — please pick another.');
    sh.getRange(i, 6, 1, 3).setValues([[text(name), text(clean(b.contact, 120)), new Date()]]);
  });
  return { ok: true, gifts: listGifts(false) };
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
  return tab(RSVP_TAB).getDataRange().getValues().slice(1).filter(r => r[0] !== '').map(r => ({
    id: Number(r[0]), name: untext(r[1]), contact: untext(r[2]), status: r[3],
    adults: Number(r[4]) || 0, kids: Number(r[5]) || 0, message: untext(r[6]),
    created_at: iso(r[7]), updated_at: iso(r[8]),
  })).sort((a, b) => b.updated_at.localeCompare(a.updated_at));
}

function addGift(b) {
  const item = clean(b.item, 120);
  if (!item) throw new Error('Please give the gift a name.');
  const link = clean(b.link, 500);
  if (link && !/^https?:\/\//i.test(link)) throw new Error('The link must start with http:// or https://');
  locked(() => tab(GIFT_TAB).appendRow([Date.now(), text(item), link, text(clean(b.price, 40)),
                                         text(clean(b.note, 300)), '', '', '', new Date()]));
  return { ok: true };
}

function releaseGift(id) {
  locked(() => {
    const sh = tab(GIFT_TAB), i = rowIndex(sh, id);
    if (i > 0) sh.getRange(i, 6, 1, 3).clearContent();
  });
  return { ok: true };
}

function deleteRow(spec, id) {
  locked(() => {
    const sh = tab(spec), i = rowIndex(sh, id);
    if (i > 0) sh.deleteRow(i);
  });
  return { ok: true };
}

// ── helpers ───────────────────────────────────────────────────────────────

function tab(spec) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(spec.name);
  if (!sh) {
    sh = ss.insertSheet(spec.name);
    sh.appendRow(spec.headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, spec.headers.length).setFontWeight('bold').setBackground('#f5e8cc');
    sh.getRange(spec.textCols).setNumberFormat('@');  // keep leading zeros / + in phone numbers
    if (spec.hideCol) sh.hideColumns(spec.hideCol);    // bookkeeping only
  }
  return sh;
}

// 1-based sheet row holding this ID, or -1.
function rowIndex(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) if (Number(ids[i][0]) === Number(id)) return i + 1;
  return -1;
}

function locked(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
}

// Guests type free text; a leading = + - @ would run as a formula in Sheets.
function text(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function untext(v) {
  return String(v).replace(/^'/, '');
}

function iso(d) {
  return d instanceof Date ? d.toISOString() : String(d || '');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
