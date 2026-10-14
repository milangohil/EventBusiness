// Runs apps-script/Code.gs against an in-memory fake of Google Sheets.
// Usage: node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function fakeSheet(name) {
  const rows = [];
  const sheet = {
    getName: () => name,
    getLastRow: () => {
      for (let i = rows.length - 1; i >= 0; i--) if (rows[i] && rows[i].some((v) => v !== '')) return i + 1;
      return 0;
    },
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r ? r.length : 0), 0),
    getMaxRows: () => Math.max(rows.length, 1000),
    setFrozenRows: () => {},
    getRange: (r, c, nr = 1, nc = 1) => ({
      getValues: () => Array.from({ length: nr }, (_, i) =>
        Array.from({ length: nc }, (_, j) => ((rows[r - 1 + i] || [])[c - 1 + j] ?? ''))),
      setValues: (vals) => {
        vals.forEach((row, i) => {
          rows[r - 1 + i] = rows[r - 1 + i] || [];
          row.forEach((v, j) => { rows[r - 1 + i][c - 1 + j] = v; });
        });
        return sheet.getRange(r, c, nr, nc);
      },
      setFontWeight() { return this; },
      setNumberFormat() { return this; }
    }),
    _rows: rows
  };
  return sheet;
}

function load(user = 'staff@example.com') {
  const sheets = {};
  const ss = {
    getId: () => 'sheet-1',
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = fakeSheet(n))
  };
  const props = {};
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    Session: { getActiveUser: () => ({ getEmail: () => ctx._user }), getScriptTimeZone: () => 'Asia/Kolkata' },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 10) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    HtmlService: {},
    _user: user
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8'), ctx);
  ctx.setup();
  return { api: ctx, sheets };
}

const booking = (extra = {}) => Object.assign({
  status: 'Booked', customerName: 'Sharma', phone: '98765 43210', eventType: 'Shaadi',
  eventDate: '2026-12-20', slot: 'Evening', hall: 'Vrindavan Vatika', guests: '500', ratePerPlate: '900'
}, extra);

test('setup creates sheets and seeds halls', () => {
  const { api } = load();
  const d = api.apiBootstrap();
  assert.strictEqual(d.halls.length, 6);
  assert.strictEqual(d.records.length, 0);
});

test('creates a booking with ID, cleaned phone and computed total', () => {
  const { api } = load();
  const r = api.apiSaveRecord(booking());
  assert.strictEqual(r.id, 'EB-0001');
  assert.strictEqual(r.phone, '9876543210');
  assert.strictEqual(r.totalAmount, 450000);
  assert.strictEqual(r.createdBy, 'staff@example.com');
  assert.strictEqual(r.followUpDate, '');
  const d = api.apiBootstrap();
  assert.strictEqual(d.records.length, 1);
  assert.strictEqual(d.records[0].hall, 'Vrindavan Vatika');
});

test('blocks a double booking of the same hall, date and slot', () => {
  const { api } = load();
  api.apiSaveRecord(booking());
  assert.throws(() => api.apiSaveRecord(booking({ customerName: 'Patel', phone: '9123456789' })), /already booked/);
  // Different slot is fine, and an enquiry for the same slot is allowed.
  api.apiSaveRecord(booking({ customerName: 'Patel', phone: '9123456789', slot: 'Morning' }));
  api.apiSaveRecord(booking({ customerName: 'Shah', phone: '9000000000', status: 'Enquiry' }));
  assert.strictEqual(api.apiBootstrap().records.length, 3);
});

test('turning an enquiry into a clashing booking is rejected', () => {
  const { api } = load();
  api.apiSaveRecord(booking());
  const enq = api.apiSaveRecord(booking({ customerName: 'Shah', phone: '9000000000', status: 'Enquiry' }));
  assert.ok(enq.followUpDate, 'open enquiry gets a default follow-up date');
  assert.throws(() => api.apiSaveRecord({ id: enq.id, updatedAt: enq.updatedAt, status: 'Booked' }), /already booked/);
});

test('validates required fields', () => {
  const { api } = load();
  assert.throws(() => api.apiSaveRecord({ customerName: '', phone: '123' }), /name is required.*10 digits/);
  assert.throws(() => api.apiSaveRecord(booking({ guests: '' })), /guest count/);
  assert.throws(() => api.apiSaveRecord(booking({ status: 'Lost' })), /lost/);
});

test('updates keep created info and reject stale edits', () => {
  const { api } = load();
  const r = api.apiSaveRecord(booking());
  const u = api.apiSaveRecord({ id: r.id, updatedAt: r.updatedAt, guests: '550', totalAmount: '' });
  assert.strictEqual(u.guests, 550);
  assert.strictEqual(u.totalAmount, 495000);
  assert.strictEqual(u.createdAt, r.createdAt);
  assert.strictEqual(api.apiBootstrap().records.length, 1);
  assert.throws(() => api.apiSaveRecord({ id: r.id, updatedAt: 'old-version', guests: '600' }), /Someone else changed/);
});

test('records payments against a booking', () => {
  const { api } = load();
  const r = api.apiSaveRecord(booking());
  const p = api.apiAddPayment({ recordId: r.id, amount: '100000', mode: 'UPI', date: '2026-10-14' });
  assert.strictEqual(p.id, 'PAY-0001');
  assert.strictEqual(p.amount, 100000);
  api.apiAddPayment({ recordId: r.id, amount: '50000', mode: 'Cash' });
  assert.strictEqual(api.apiBootstrap().payments.length, 2);
  assert.throws(() => api.apiAddPayment({ recordId: r.id, amount: '0', mode: 'Cash' }), /greater than zero/);
  assert.throws(() => api.apiAddPayment({ recordId: 'EB-9999', amount: '10', mode: 'Cash' }), /not found/);
});

test('columns are matched by header name, so reordering in the sheet is safe', () => {
  const { api, sheets } = load();
  api.apiSaveRecord(booking());
  const rows = sheets.Records._rows;
  rows.forEach((row) => { const t = row[0]; row[0] = row[6]; row[6] = t; }); // swap ID and Customer Name columns
  const d = api.apiBootstrap();
  assert.strictEqual(d.records[0].id, 'EB-0001');
  assert.strictEqual(d.records[0].customerName, 'Sharma');
  assert.strictEqual(api.apiSaveRecord(booking({ slot: 'Morning' })).id, 'EB-0002');
});
