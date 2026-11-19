// Runs cloud/apps_script/Code.gs (the shared address book's Apps Script) in node, behind a local web address, with
// stand-ins for Google's services, so tests/test_cloud_sync.py can sync real AddressBooks through the real script.
//   node tests/cloud_server.js SECRET   -> prints the port, then serves POST / (doPost) and GET /__state (the sheets)
// The stand-in sheet behaves like Sheets where it matters: a leading apostrophe marks text and is hidden, and text
// that looks like a number without one becomes a number ('0803' -> 803), so a missing apostrophe fails a test.
'use strict';
const fs = require('fs');
const http = require('http');
const path = require('path');
const vm = require('vm');

const secret = process.argv[2] || 'test-secret';
const props = { SECRET: secret };
const books = {};
let nextId = 1;

function cellIn(v) {
  if (typeof v === 'string') {
    if (v.startsWith("'")) return v.slice(1);
    if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);
  }
  return v;
}

function makeSheet(name, parent) {
  const rows = [];  // rows[0] is row 1
  const sheet = {
    name, rows,
    getName: () => name,
    getParent: () => parent,
    getLastRow: () => rows.length,
    setFrozenRows() {},
    getRange(r, c, nr = 1, nc = 1) {
      return {
        setValues(values) {
          for (let i = 0; i < nr; i++) {
            rows[r - 1 + i] = rows[r - 1 + i] || [];
            for (let j = 0; j < nc; j++) rows[r - 1 + i][c - 1 + j] = cellIn(values[i][j]);
          }
        },
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = rows[r - 1 + i] || [];
            out.push(Array.from({ length: nc }, (_, j) => (row[c - 1 + j] === undefined ? '' : row[c - 1 + j])));
          }
          return out;
        },
      };
    },
  };
  return sheet;
}

function makeBook(title) {
  const id = `book${nextId++}`;
  const book = {
    id, title, sheets: [],
    getId: () => id,
    getUrl: () => `https://docs.example/${id}`,
    getSheets: () => book.sheets,
    getSheetByName: (n) => book.sheets.find((s) => s.name === n) || null,
    insertSheet(n) { const s = makeSheet(n, book); book.sheets.push(s); return s; },
  };
  book.sheets.push(makeSheet('Sheet1', book));
  books[id] = book;
  return book;
}

const main = makeBook('Warehouse address book');
const context = {
  PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  SpreadsheetApp: { getActiveSpreadsheet: () => main, openById: (id) => books[id], create: (t) => makeBook(t) },
  ContentService: {
    MimeType: { JSON: 'json' },
    createTextOutput: (text) => ({ text, setMimeType() { return this; } }),
  },
  Utilities: { formatDate: (d) => d.toISOString().slice(0, 7) },
  Session: { getScriptTimeZone: () => 'UTC' },
  console,
};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'cloud', 'apps_script', 'Code.gs'), 'utf8'), context);

const server = http.createServer((req, res) => {
  let body = '';
  req.on('data', (c) => { body += c; });
  req.on('end', () => {
    if (req.method === 'GET' && req.url === '/__state') {
      const dump = Object.values(books).map((b) => ({ id: b.id, title: b.title, sheets: b.sheets.map((s) => ({ name: s.name, rows: s.rows })) }));
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ books: dump, props }));
      return;
    }
    const out = req.method === 'POST' ? context.doPost({ postData: { contents: body } }) : context.doGet();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(out.text);
  });
});
server.listen(0, '127.0.0.1', () => { process.stdout.write(`${server.address().port}\n`); });
