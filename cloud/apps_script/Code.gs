/**
 * Warehouse app: the shared address book. A Google Apps Script web app bound to a Google Sheet (Extensions > Apps
 * Script in that sheet). Every computer running the app keeps its own address book and syncs with this one:
 *   push  its new and changed addresses and learned spellings (within seconds of a change)
 *   pull  what changed here since its last pull (on start, every 5 minutes, before a preview)
 * The app never waits on this: it works on its own copy, and syncs when it can.
 *
 * Which copy wins when two computers changed the same address: a verified one (confirmed by the courier portal)
 * over one that isn't; otherwise the latest change. Deleting is done on each computer only: nothing is ever deleted
 * here. Every change is written to a log, one log spreadsheet a month, with who (the name typed in the app) and
 * which computer.
 *
 * Setup: docs/CENTRAL_ADDRESS_BOOK.md. Script property SECRET = the key in each computer's .env (WAREHOUSE_SYNC_KEY).
 * Tests run this same file in node with stand-ins for Google's services (tests/cloud_server.js).
 */

var ADDRESS_SHEET = 'Addresses';
var SPELLING_SHEET = 'Spellings';
var LOGS_SHEET = 'Logs';
var ADDRESS_COLS = ['gid', 'receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country',
  'authority_to_leave', 'address_key', 'receiver_key', 'contact_key', 'verified_at', 'changed_at', 'changed_by',
  'changed_pc', 'rev'];
var FIELDS = ['receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country', 'authority_to_leave',
  'address_key', 'receiver_key', 'contact_key', 'verified_at'];
var SHOWN = ['receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country', 'authority_to_leave',
  'verified_at'];
var NUMBER_COLS = { authority_to_leave: true, verified_at: true, changed_at: true, rev: true };
var SPELLING_COLS = ['source_key', 'gid', 'changed_at', 'changed_by', 'changed_pc', 'rev'];
var MAX_PULL = 1000;
var MAX_PUSH = 500;
var MAX_TEXT = 200;

function doGet() {
  return ContentService.createTextOutput('Warehouse shared address book: running.');
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return reply({ error: 'Bad request.' });
  }
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
  if (!secret || !sameText(String(req.key || ''), secret)) {
    return reply({ error: 'Not allowed: the key in this computer\'s .env (WAREHOUSE_SYNC_KEY) is wrong.' });
  }
  // One request at a time, so a pull never reads half a push and two pushes never give out the same rev
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return reply({ error: 'The shared address book is busy: trying again shortly.' });
  try {
    var book = openBook();
    if (req.action === 'ping') return reply({ ok: true, rev: currentRev() });
    if (req.action === 'pull') return reply(pull(book, Number(req.since) || 0, Math.min(Number(req.limit) || MAX_PULL, MAX_PULL)));
    if (req.action === 'push') return reply(push(book, req));
    return reply({ error: 'Unknown action.' });
  } catch (err) {
    return reply({ error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------ the sheets

function openBook() {
  var id = PropertiesService.getScriptProperties().getProperty('BOOK_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return {
    ss: ss,
    addresses: sheet(ss, ADDRESS_SHEET, ADDRESS_COLS),
    spellings: sheet(ss, SPELLING_SHEET, SPELLING_COLS),
  };
}

function sheet(ss, name, cols) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Every text goes in with a leading apostrophe, so Sheets keeps it as typed: '0803' stays 0803 (not 803), '=SUM()' and
// '1-2' stay text (not a formula or a date). Sheets hides the apostrophe; reading takes off one if it's still there.
function toCell(col, value) {
  if (NUMBER_COLS[col]) return value === null || value === undefined || value === '' ? '' : Number(value) || 0;
  var text = String(value === null || value === undefined ? '' : value);
  return text ? "'" + text : '';
}

function fromCell(col, value) {
  if (NUMBER_COLS[col]) return value === '' || value === null || value === undefined ? null : Number(value);
  var text = String(value === null || value === undefined ? '' : value);
  return text.charAt(0) === "'" ? text.slice(1) : text;
}

function readTable(sh, cols) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map(function (v, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = fromCell(c, v[j]); });
    return o;
  });
}

function rowValues(o, cols) {
  return cols.map(function (c) { return toCell(c, o[c]); });
}

// Changed rows written back in one call when the Sheets API service is turned on (Services > Google Sheets API),
// else row by row (Apps Script sends those together too)
function writeRows(sh, objs, cols) {
  if (!objs.length) return;
  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
    var name = sh.getName();
    var data = objs.map(function (o) {
      return { range: "'" + name + "'!A" + o._row, values: [rowValues(o, cols)] };
    });
    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'USER_ENTERED', data: data }, sh.getParent().getId());
    return;
  }
  objs.forEach(function (o) { sh.getRange(o._row, 1, 1, cols.length).setValues([rowValues(o, cols)]); });
}

function appendRows(sh, objs, cols) {
  if (!objs.length) return;
  var start = sh.getLastRow() + 1;
  sh.getRange(start, 1, objs.length, cols.length).setValues(objs.map(function (o) { return rowValues(o, cols); }));
}

function currentRev() {
  return Number(PropertiesService.getScriptProperties().getProperty('REV')) || 0;
}

function saveRev(rev) {
  PropertiesService.getScriptProperties().setProperty('REV', String(rev));
}

// ------------------------------------------------------------------ pull

// What changed here after `since` (each change has the next number, its rev), oldest first, at most `limit` at a
// time. 'rev' is where the next pull starts; 'more' says there's more to fetch now.
function pull(book, since, limit) {
  var changes = [];
  readTable(book.addresses, ADDRESS_COLS).forEach(function (r) { if (r.rev > since) changes.push(['row', r]); });
  readTable(book.spellings, SPELLING_COLS).forEach(function (s) { if (s.rev > since) changes.push(['spelling', s]); });
  changes.sort(function (a, b) { return a[1].rev - b[1].rev; });
  var page = changes.slice(0, limit);
  var more = changes.length > page.length;
  return {
    rows: page.filter(function (c) { return c[0] === 'row'; }).map(function (c) { return publicRow(c[1]); }),
    spellings: page.filter(function (c) { return c[0] === 'spelling'; }).map(function (c) { return publicSpelling(c[1]); }),
    rev: more ? page[page.length - 1][1].rev : currentRev(),
    more: more,
  };
}

function publicRow(r) {
  var o = {};
  ADDRESS_COLS.forEach(function (c) { o[c] = r[c]; });
  return o;
}

function publicSpelling(s) {
  var o = {};
  SPELLING_COLS.forEach(function (c) { o[c] = s[c]; });
  return o;
}

// ------------------------------------------------------------------ push

function identity(r) {
  return [r.address_key, r.receiver_key, r.contact_key].join('|');
}

// The rule both sides follow: a verified copy beats one that isn't; otherwise the latest change wins
function incomingWins(inc, cur) {
  var iv = !!inc.verified_at, cv = !!cur.verified_at;
  if (iv !== cv) return iv;
  return (Number(inc.changed_at) || 0) >= (Number(cur.changed_at) || 0);
}

function tidyIncoming(r) {
  var o = { gid: String(r.gid || '').slice(0, 64) };
  FIELDS.forEach(function (f) {
    o[f] = NUMBER_COLS[f] ? (r[f] === null || r[f] === undefined || r[f] === '' ? null : Number(r[f]) || 0)
      : String(r[f] === null || r[f] === undefined ? '' : r[f]).slice(0, MAX_TEXT);
  });
  o.authority_to_leave = o.authority_to_leave ? 1 : 0;
  o.changed_at = Number(r.changed_at) || 0;
  return o;
}

function differences(cur, inc) {
  var out = [];
  SHOWN.forEach(function (f) {
    var a = cur[f] === null || cur[f] === undefined ? '' : String(cur[f]);
    var b = inc[f] === null || inc[f] === undefined ? '' : String(inc[f]);
    if (f === 'verified_at') {
      if (!!cur.verified_at !== !!inc.verified_at) out.push(inc.verified_at ? 'verified' : 'no longer verified');
    } else if (a !== b) {
      out.push(f + ': ' + (a || '(blank)') + ' -> ' + (b || '(blank)'));
    }
  });
  return out.join('; ');
}

function oneLine(r) {
  return [r.line1, r.line2, r.suburb, r.state, r.postcode].filter(function (x) { return x; }).join(', ');
}

function push(book, req) {
  var who = String(req.user || '').trim().slice(0, 60);
  var pc = String(req.pc || '').trim().slice(0, 60);
  if (!who) return { error: 'Type your name on the Address Book page of the app first: every change is logged with it.' };
  var now = new Date();
  var table = readTable(book.addresses, ADDRESS_COLS);
  var before = table.length;
  var byGid = {}, byKey = {};
  table.forEach(function (r, i) { byGid[r.gid] = i; byKey[identity(r)] = i; });
  var rev = currentRev();
  var changed = {}, results = [], log = [];

  (req.rows || []).slice(0, MAX_PUSH).forEach(function (raw) {
    var inc = tidyIncoming(raw);
    if (!inc.gid) return;
    var i = byGid[inc.gid];
    if (i === undefined) i = byKey[identity(inc)];  // the same address saved on another computer first
    if (i === undefined) {
      var row = Object.assign({}, inc, { changed_by: who, changed_pc: pc, rev: ++rev, _row: 0 });
      table.push(row);
      byGid[row.gid] = table.length - 1;
      byKey[identity(row)] = table.length - 1;
      results.push({ gid: inc.gid, canonical: inc.gid, status: 'saved' });
      log.push([now, who, pc, 'added', row.gid, row.receiver, oneLine(row) + (row.verified_at ? ' (verified)' : '')]);
      return;
    }
    var cur = table[i];
    var clash = byKey[identity(inc)];
    var diff = differences(cur, inc);
    if (!incomingWins(inc, cur) || (clash !== undefined && clash !== i)) {
      // This computer's copy loses: it gets the shared one back
      results.push({ gid: inc.gid, canonical: cur.gid, status: 'kept', row: publicRow(cur) });
      if (diff) {
        log.push([now, who, pc, 'kept the shared copy', cur.gid, cur.receiver,
          (clash !== undefined && clash !== i ? 'another entry already has that receiver at that address. ' :
            cur.verified_at && !inc.verified_at ? 'the shared copy is verified. ' : 'the shared copy is newer. ') +
          'Not taken: ' + diff]);
      }
      return;
    }
    results.push({ gid: inc.gid, canonical: cur.gid, status: 'saved' });
    if (!diff) return;  // nothing new: no change, no rev
    delete byKey[identity(cur)];
    FIELDS.forEach(function (f) { cur[f] = inc[f]; });
    cur.changed_at = inc.changed_at;
    cur.changed_by = who;
    cur.changed_pc = pc;
    cur.rev = ++rev;
    byKey[identity(cur)] = i;
    if (i < before) changed[i] = true;
    log.push([now, who, pc, 'changed', cur.gid, cur.receiver, diff]);
  });

  // Learned spellings: the raw Excel text of an address -> the entry it was saved as (the latest wins)
  var canonical = {};
  results.forEach(function (r) { canonical[r.gid] = r.canonical; });
  var spellTable = readTable(book.spellings, SPELLING_COLS);
  var spellBefore = spellTable.length;
  var bySource = {};
  spellTable.forEach(function (s, i) { bySource[s.source_key] = i; });
  var spellChanged = {}, spellResults = {};
  (req.spellings || []).slice(0, MAX_PUSH * 2).forEach(function (raw) {
    var key = String(raw.source_key || '').slice(0, 400);
    var gid = canonical[raw.gid] || String(raw.gid || '');
    var at = Number(raw.changed_at) || 0;
    if (!key) return;
    if (byGid[gid] === undefined) { spellResults[key] = { status: 'unknown' }; return; }
    var j = bySource[key];
    if (j !== undefined) {
      var s = spellTable[j];
      if (s.gid === gid) { spellResults[key] = { status: 'saved' }; return; }
      if ((Number(s.changed_at) || 0) > at) { spellResults[key] = { status: 'kept', gid: s.gid }; return; }
      s.gid = gid; s.changed_at = at; s.changed_by = who; s.changed_pc = pc; s.rev = ++rev;
      if (j < spellBefore) spellChanged[j] = true;
    } else {
      spellTable.push({ source_key: key, gid: gid, changed_at: at, changed_by: who, changed_pc: pc, rev: ++rev, _row: 0 });
      bySource[key] = spellTable.length - 1;
    }
    spellResults[key] = { status: 'saved' };
  });

  writeRows(book.addresses, Object.keys(changed).map(function (i) { return table[i]; }), ADDRESS_COLS);
  appendRows(book.addresses, table.slice(before), ADDRESS_COLS);
  writeRows(book.spellings, Object.keys(spellChanged).map(function (j) { return spellTable[j]; }), SPELLING_COLS);
  appendRows(book.spellings, spellTable.slice(spellBefore), SPELLING_COLS);
  saveRev(rev);
  writeLog(book.ss, log);
  return { results: results, spellings: spellResults, rev: rev };
}

// ------------------------------------------------------------------ the log

// One log spreadsheet a month (a busy month is hundreds of thousands of rows: one spreadsheet holds 10 million cells),
// listed with its link in the Logs sheet of the address book
function writeLog(ss, rows) {
  if (!rows.length) return;
  var props = PropertiesService.getScriptProperties();
  var month = Utilities.formatDate(rows[0][0], Session.getScriptTimeZone(), 'yyyy-MM');
  var id = props.getProperty('LOG_' + month);
  var logBook;
  if (id) {
    logBook = SpreadsheetApp.openById(id);
  } else {
    logBook = SpreadsheetApp.create('Warehouse address book log ' + month);
    props.setProperty('LOG_' + month, logBook.getId());
    var first = logBook.getSheets()[0];
    first.getRange(1, 1, 1, 7).setValues([['When', 'Who', 'Computer', 'What', 'Entry', 'Receiver', 'Details']]);
    first.setFrozenRows(1);
    var list = ss.getSheetByName(LOGS_SHEET) || ss.insertSheet(LOGS_SHEET);
    if (list.getLastRow() < 1) list.getRange(1, 1, 1, 2).setValues([['Month', 'Log']]);
    list.getRange(list.getLastRow() + 1, 1, 1, 2).setValues([["'" + month, logBook.getUrl()]]);
  }
  var sh = logBook.getSheets()[0];
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows.map(function (r) {
    return [r[0]].concat(r.slice(1).map(function (x) { return x ? "'" + String(x).slice(0, 5000) : ''; }));
  }));
}

// ------------------------------------------------------------------ helpers

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Compares the key without stopping at the first different letter (no timing hints)
function sameText(a, b) {
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
