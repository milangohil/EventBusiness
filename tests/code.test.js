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
      clearContent: () => {
        for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (rows[r - 1 + i]) rows[r - 1 + i][c - 1 + j] = '';
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
    LanguageApp: { translate: (t, from, to) => { if (t === 'fail') throw new Error('quota'); return '[' + from + '>' + to + '] ' + t; } },
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

// ---------------------------------------------------------------- Stage 2: kitchen

function kitchenSetup() {
  const env = load();
  const { api } = env;
  const sup = api.apiSaveSupplier({ name: 'Dairy Supplier', phone: '98250 00000' });
  const paneer = api.apiSaveIngredient({ name: 'पनीर', nameEn: 'Paneer', unit: 'kg', category: 'Dairy', supplierId: sup.id });
  const butter = api.apiSaveIngredient({ name: 'मक्खन', nameEn: 'Butter', unit: 'kg', category: 'Dairy', supplierId: sup.id });
  const rice = api.apiSaveIngredient({ name: 'बासमती चावल', nameEn: 'Basmati rice', unit: 'kg', category: 'Grocery' });
  const pbm = api.apiSaveDish({ name: 'पनीर बटर मसाला', category: 'Main course', recipe: [{ ingredientId: paneer.id, qtyPer100: 8 }, { ingredientId: butter.id, qtyPer100: 1.5 }] });
  const jeera = api.apiSaveDish({ name: 'जीरा राइस', category: 'Rice', recipe: [{ ingredientId: rice.id, qtyPer100: 6 }, { ingredientId: butter.id, qtyPer100: 0.5 }] });
  const rec = api.apiSaveRecord(booking());
  return Object.assign(env, { sup, paneer, butter, rice, pbm, jeera, rec });
}

test('setup creates the kitchen sheets', () => {
  const { sheets } = load();
  ['Suppliers', 'Ingredients', 'Dishes', 'Recipes', 'Kitchen'].forEach((n) => assert.ok(sheets[n], n + ' sheet exists'));
});

test('saves suppliers, ingredients and dishes with recipes', () => {
  const { api, sup, paneer, pbm, jeera } = kitchenSetup();
  assert.strictEqual(sup.id, 'SUP-0001');
  assert.strictEqual(sup.phone, '9825000000');
  assert.strictEqual(paneer.id, 'ING-0001');
  assert.strictEqual(pbm.dish.id, 'DSH-0001');
  const d = api.apiBootstrap();
  assert.strictEqual(d.dishes.length, 2);
  assert.strictEqual(d.recipes.length, 4);
  assert.strictEqual(d.recipes.filter((r) => r.dishId === jeera.dish.id).length, 2);
});

test('editing a recipe replaces only that dish\'s lines', () => {
  const { api, pbm, paneer, jeera } = kitchenSetup();
  api.apiSaveDish({ id: pbm.dish.id, recipe: [{ ingredientId: paneer.id, qtyPer100: 9 }] });
  const d = api.apiBootstrap();
  const lines = d.recipes.filter((r) => r.dishId === pbm.dish.id);
  assert.strictEqual(lines.length, 1);
  assert.strictEqual(lines[0].qtyPer100, 9);
  assert.strictEqual(d.recipes.filter((r) => r.dishId === jeera.dish.id).length, 2);
  assert.strictEqual(d.dishes.find((x) => x.id === pbm.dish.id).name, 'पनीर बटर मसाला', 'name kept when not sent');
});

test('rejects bad kitchen master data', () => {
  const { api, paneer } = kitchenSetup();
  assert.throws(() => api.apiSaveSupplier({ name: '' }), /name is required/);
  assert.throws(() => api.apiSaveSupplier({ name: 'X', phone: '123' }), /10 digits/);
  assert.throws(() => api.apiSaveIngredient({ name: 'Salt', unit: 'bag' }), /unit/);
  assert.throws(() => api.apiSaveIngredient({ name: 'पनीर', unit: 'kg' }), /already exists/);
  assert.throws(() => api.apiSaveIngredient({ name: 'Salt', unit: 'kg', supplierId: 'SUP-9999' }), /not found/);
  assert.throws(() => api.apiSaveDish({ name: 'Dal', category: 'Soup-ish' }), /category/);
  assert.throws(() => api.apiSaveDish({ name: 'Dal', category: 'Dal', recipe: [{ ingredientId: paneer.id, qtyPer100: 0 }] }), /more than zero/);
  assert.throws(() => api.apiSaveDish({ name: 'Dal', category: 'Dal', recipe: [{ ingredientId: paneer.id, qtyPer100: 1 }, { ingredientId: paneer.id, qtyPer100: 2 }] }), /twice/);
  assert.throws(() => api.apiSaveDish({ name: 'Dal', category: 'Dal', recipe: [{ ingredientId: 'ING-9999', qtyPer100: 1 }] }), /not found/);
});

test('menu, quantities and approval flow for a booking', () => {
  const { api, rec, pbm, jeera, paneer, butter, rice } = kitchenSetup();
  assert.throws(() => api.apiSaveQuantities({ recordId: rec.id, items: [] }), /menu before/);
  let k = api.apiSaveMenu({ recordId: rec.id, dishIds: [pbm.dish.id, jeera.dish.id], menuStatus: 'Final', menuNotes: 'Less spicy' });
  assert.strictEqual(k.dishIds.length, 2);
  assert.strictEqual(k.menuStatus, 'Final');
  k = api.apiSaveQuantities({ recordId: rec.id, updatedAt: k.updatedAt, approve: true, items: [
    { ingredientId: paneer.id, calc: 40, qty: 42 }, { ingredientId: butter.id, calc: 10, qty: 10 }, { ingredientId: rice.id, calc: 30, qty: 30 }] });
  assert.strictEqual(k.qtyStatus, 'Approved');
  assert.strictEqual(k.approvedBy, 'staff@example.com');
  assert.strictEqual(k.guestsUsed, 500);
  const d = api.apiBootstrap();
  assert.strictEqual(d.kitchen.length, 1);
  assert.strictEqual(d.kitchen[0].items.length, 3);
  assert.strictEqual(d.kitchen[0].items[0].qty, 42);
  assert.deepStrictEqual(Array.from(d.kitchen[0].dishIds), [pbm.dish.id, jeera.dish.id]);
  // Re-saving the same menu keeps approval; changing dishes flags a re-check.
  k = api.apiSaveMenu({ recordId: rec.id, dishIds: [jeera.dish.id, pbm.dish.id], menuStatus: 'Final' });
  assert.strictEqual(k.qtyStatus, 'Approved');
  k = api.apiSaveMenu({ recordId: rec.id, dishIds: [jeera.dish.id], menuStatus: 'Final' });
  assert.strictEqual(k.qtyStatus, 'Recheck');
  assert.throws(() => api.apiSaveMenu({ recordId: rec.id, updatedAt: 'stale', dishIds: [] }), /Someone else/);
  assert.throws(() => api.apiSaveMenu({ recordId: 'EB-9999', dishIds: [] }), /not found/);
  assert.throws(() => api.apiSaveMenu({ recordId: rec.id, dishIds: ['DSH-9999'] }), /not found/);
  assert.throws(() => api.apiSaveQuantities({ recordId: rec.id, items: [{ ingredientId: paneer.id, calc: 1, qty: '' }] }), /quantity for every item/);
  k = api.apiSaveQuantities({ recordId: rec.id, items: [{ ingredientId: rice.id, calc: 30, qty: 30 }] });
  assert.strictEqual(k.qtyStatus, 'Draft');
  assert.strictEqual(k.approvedBy, '');
});

test('translates English names to Hindi and reports failures clearly', () => {
  const { api } = load();
  assert.strictEqual(api.apiTranslate('  Paneer Tikka '), '[en>hi] Paneer Tikka');
  assert.strictEqual(api.apiTranslate(''), '');
  assert.throws(() => api.apiTranslate('fail'), /type the Hindi name/);
});
