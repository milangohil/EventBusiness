/**
 * Event Business - Stage 1 (enquiries, bookings, availability, payments).
 *
 * Storage is the Google Sheet this script is attached to:
 *   Records  - one row per enquiry / booking
 *   Payments - one row per payment received
 *   Halls    - list of venues (edit this sheet to add / rename halls)
 *
 * Run setup() once from the Apps Script editor, then deploy as a web app.
 */

var SHEETS = {
  RECORDS: 'Records', PAYMENTS: 'Payments', HALLS: 'Halls',
  SUPPLIERS: 'Suppliers', INGREDIENTS: 'Ingredients', DISHES: 'Dishes', RECIPES: 'Recipes', KITCHEN: 'Kitchen'
};

var RECORD_FIELDS = [
  ['id', 'ID'],
  ['createdAt', 'Created At'],
  ['createdBy', 'Created By'],
  ['updatedAt', 'Updated At'],
  ['updatedBy', 'Updated By'],
  ['status', 'Status'],
  ['customerName', 'Customer Name'],
  ['phone', 'Phone'],
  ['address', 'Address'],
  ['eventType', 'Event Type'],
  ['eventDate', 'Event Date'],
  ['slot', 'Slot'],
  ['hall', 'Hall'],
  ['guests', 'Guests'],
  ['source', 'Source'],
  ['ratePerPlate', 'Rate Per Plate'],
  ['totalAmount', 'Total Amount'],
  ['followUpDate', 'Next Follow-up'],
  ['lostReason', 'Lost Reason'],
  ['notes', 'Notes']
];

var PAYMENT_FIELDS = [
  ['id', 'Payment ID'],
  ['recordId', 'Booking ID'],
  ['date', 'Date'],
  ['amount', 'Amount'],
  ['mode', 'Mode'],
  ['reference', 'Reference'],
  ['recordedBy', 'Recorded By'],
  ['recordedAt', 'Recorded At']
];

var HALL_FIELDS = [
  ['name', 'Hall'],
  ['capacity', 'Capacity'],
  ['active', 'Active']
];

var SUPPLIER_FIELDS = [
  ['id', 'Supplier ID'],
  ['name', 'Name'],
  ['phone', 'Phone'],
  ['supplies', 'Supplies'],
  ['active', 'Active']
];

var INGREDIENT_FIELDS = [
  ['id', 'Ingredient ID'],
  ['name', 'Name'],
  ['nameEn', 'English Name'],
  ['unit', 'Unit'],
  ['category', 'Category'],
  ['supplierId', 'Supplier ID'],
  ['active', 'Active']
];

var DISH_FIELDS = [
  ['id', 'Dish ID'],
  ['name', 'Name'],
  ['nameEn', 'English Name'],
  ['category', 'Category'],
  ['active', 'Active']
];

// One row per ingredient used in a dish. Quantity is in the ingredient's unit, for 100 guests.
var RECIPE_FIELDS = [
  ['dishId', 'Dish ID'],
  ['ingredientId', 'Ingredient ID'],
  ['qtyPer100', 'Qty per 100 Guests']
];

// One row per booking: its menu and the approved purchase list.
var KITCHEN_FIELDS = [
  ['recordId', 'Booking ID'],
  ['dishIds', 'Dish IDs'],
  ['menuStatus', 'Menu Status'],
  ['menuNotes', 'Menu Notes'],
  ['guestsUsed', 'Guests Used'],
  ['items', 'Items'],
  ['qtyStatus', 'Quantity Status'],
  ['approvedBy', 'Approved By'],
  ['approvedAt', 'Approved At'],
  ['kitchenNotes', 'Kitchen Notes'],
  ['updatedAt', 'Updated At'],
  ['updatedBy', 'Updated By']
];

var DEFAULT_HALLS = [
  ['Mahavir Banquet', '', 'Yes'],
  ['Nirmala Hall', '', 'Yes'],
  ['Vrindavan Vatika', '', 'Yes'],
  ['C.H. Hall', '', 'Yes'],
  ['Barfiwala Hall', '', 'Yes'],
  ['M.M. Public School', '', 'Yes']
];

var LISTS = {
  statuses: ['Enquiry', 'Visit Scheduled', 'Tasting Scheduled', 'Booked', 'Completed', 'Cancelled', 'Lost'],
  openStatuses: ['Enquiry', 'Visit Scheduled', 'Tasting Scheduled'],
  blockingStatuses: ['Booked', 'Completed'],
  eventTypes: ['Shaadi', 'Reception', 'Engagement', 'Other'],
  slots: ['Morning', 'Evening'],
  sources: ['Word of mouth', 'Google', 'Walk-in', 'Repeat customer', 'Other'],
  paymentModes: ['Cash', 'UPI', 'Bank transfer', 'Cheque', 'Online link'],
  units: ['kg', 'gram', 'litre', 'ml', 'piece', 'dozen', 'packet', 'box', 'bundle'],
  ingredientCategories: ['Grocery', 'Vegetables', 'Fruits', 'Dairy', 'Spices', 'Oil & Ghee', 'Dry fruits', 'Disposables', 'Other'],
  // [English, Hindi] - Hindi is used on the printed menu.
  dishCategories: [
    ['Welcome drink', 'वेलकम ड्रिंक'], ['Starter', 'स्टार्टर'], ['Soup', 'सूप'], ['Chaat', 'चाट'], ['Farsan', 'फरसाण'],
    ['Main course', 'मुख्य व्यंजन'], ['Dal', 'दाल'], ['Rice', 'चावल'], ['Roti', 'रोटी'], ['Salad', 'सलाद'],
    ['Raita', 'रायता'], ['Sweet', 'मिठाई'], ['Dessert', 'डेज़र्ट'], ['Other', 'अन्य']
  ],
  menuStatuses: ['Draft', 'Final'],
  qtyStatuses: ['Draft', 'Approved', 'Recheck']
};

var TEXT_COLUMNS = ['id', 'createdAt', 'updatedAt', 'phone', 'eventDate', 'followUpDate', 'recordId', 'date', 'recordedAt',
  'supplierId', 'dishId', 'ingredientId', 'dishIds', 'items', 'approvedAt'];
var FOLLOW_UP_DAYS = 3;

// ---------------------------------------------------------------- web app

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Event Business')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once from the editor: creates the sheets and remembers this spreadsheet. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the Google Sheet (Extensions > Apps Script) and run setup again.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ensureSheet_(ss, SHEETS.RECORDS, RECORD_FIELDS, []);
  ensureSheet_(ss, SHEETS.PAYMENTS, PAYMENT_FIELDS, []);
  ensureSheet_(ss, SHEETS.HALLS, HALL_FIELDS, DEFAULT_HALLS);
  ensureSheet_(ss, SHEETS.SUPPLIERS, SUPPLIER_FIELDS, []);
  ensureSheet_(ss, SHEETS.INGREDIENTS, INGREDIENT_FIELDS, []);
  ensureSheet_(ss, SHEETS.DISHES, DISH_FIELDS, []);
  ensureSheet_(ss, SHEETS.RECIPES, RECIPE_FIELDS, []);
  ensureSheet_(ss, SHEETS.KITCHEN, KITCHEN_FIELDS, []);
  return 'Setup complete';
}

// ---------------------------------------------------------------- API (called from Index.html)

function apiBootstrap() {
  var ss = getSpreadsheet_();
  return {
    user: currentUser_(),
    today: today_(),
    lists: LISTS,
    halls: readRows_(ss.getSheetByName(SHEETS.HALLS), HALL_FIELDS)
      .filter(function (h) { return h.name && String(h.active).toLowerCase() !== 'no'; }),
    records: readRows_(ss.getSheetByName(SHEETS.RECORDS), RECORD_FIELDS),
    payments: readRows_(ss.getSheetByName(SHEETS.PAYMENTS), PAYMENT_FIELDS),
    suppliers: readOptional_(ss, SHEETS.SUPPLIERS, SUPPLIER_FIELDS),
    ingredients: readOptional_(ss, SHEETS.INGREDIENTS, INGREDIENT_FIELDS),
    dishes: readOptional_(ss, SHEETS.DISHES, DISH_FIELDS),
    recipes: readOptional_(ss, SHEETS.RECIPES, RECIPE_FIELDS).filter(function (r) { return r.dishId && r.ingredientId; }),
    kitchen: readOptional_(ss, SHEETS.KITCHEN, KITCHEN_FIELDS).map(parseKitchen_)
  };
}

/** English to Hindi for dish and ingredient names, using Google Translate built into Apps Script. */
function apiTranslate(text) {
  text = String(text || '').trim().slice(0, 200);
  if (!text) return '';
  try {
    return LanguageApp.translate(text, 'en', 'hi');
  } catch (e) {
    throw new Error('Translation is not available right now. Please type the Hindi name.');
  }
}

// ---------------------------------------------------------------- kitchen API (Stage 2)

function apiSaveSupplier(input) {
  return withLock_(function () {
    var sheet = kitchenSheet_(SHEETS.SUPPLIERS);
    var rows = readRows_(sheet, SUPPLIER_FIELDS);
    var rec = mergeEntity_(SUPPLIER_FIELDS, input, findById_(rows, input.id, 'Supplier'));
    rec.phone = String(rec.phone).replace(/[^\d+]/g, '');
    rec.active = rec.active === 'No' ? 'No' : 'Yes';
    if (!rec.name) throw new Error('Supplier name is required.');
    if (rec.phone && !/^\+?\d{10,13}$/.test(rec.phone)) throw new Error('Supplier phone should have 10 digits.');
    return saveEntity_(sheet, SUPPLIER_FIELDS, 'SUP-', rows, rec);
  });
}

function apiSaveIngredient(input) {
  return withLock_(function () {
    var sheet = kitchenSheet_(SHEETS.INGREDIENTS);
    var rows = readRows_(sheet, INGREDIENT_FIELDS);
    var rec = mergeEntity_(INGREDIENT_FIELDS, input, findById_(rows, input.id, 'Ingredient'));
    rec.active = rec.active === 'No' ? 'No' : 'Yes';
    if (!rec.name) throw new Error('Ingredient name is required.');
    if (LISTS.units.indexOf(rec.unit) === -1) throw new Error('Choose a unit for ' + rec.name + '.');
    if (rec.category && LISTS.ingredientCategories.indexOf(rec.category) === -1) throw new Error('Unknown ingredient category.');
    if (rec.supplierId) findById_(readRows_(kitchenSheet_(SHEETS.SUPPLIERS), SUPPLIER_FIELDS), rec.supplierId, 'Supplier');
    var clash = rows.filter(function (r) { return r.id !== rec.id && String(r.name).trim().toLowerCase() === rec.name.toLowerCase(); })[0];
    if (clash) throw new Error('An ingredient called "' + rec.name + '" already exists (' + clash.id + ').');
    return saveEntity_(sheet, INGREDIENT_FIELDS, 'ING-', rows, rec);
  });
}

/** Saves a dish and replaces its recipe lines: input.recipe = [{ingredientId, qtyPer100}]. */
function apiSaveDish(input) {
  return withLock_(function () {
    var sheet = kitchenSheet_(SHEETS.DISHES);
    var rows = readRows_(sheet, DISH_FIELDS);
    var rec = mergeEntity_(DISH_FIELDS, input, findById_(rows, input.id, 'Dish'));
    rec.active = rec.active === 'No' ? 'No' : 'Yes';
    if (!rec.name) throw new Error('Dish name is required.');
    var cats = LISTS.dishCategories.map(function (c) { return c[0]; });
    if (cats.indexOf(rec.category) === -1) throw new Error('Choose a category for ' + rec.name + '.');
    var clash = rows.filter(function (r) { return r.id !== rec.id && String(r.name).trim().toLowerCase() === rec.name.toLowerCase(); })[0];
    if (clash) throw new Error('A dish called "' + rec.name + '" already exists (' + clash.id + ').');

    var ingredientIds = readRows_(kitchenSheet_(SHEETS.INGREDIENTS), INGREDIENT_FIELDS).map(function (i) { return i.id; });
    var seen = {};
    var lines = (input.recipe || []).map(function (l) {
      var qty = Number(l.qtyPer100);
      if (ingredientIds.indexOf(l.ingredientId) === -1) throw new Error('Ingredient ' + l.ingredientId + ' was not found.');
      if (seen[l.ingredientId]) throw new Error('An ingredient is listed twice in this recipe.');
      if (!(qty > 0)) throw new Error('Every recipe quantity must be more than zero.');
      seen[l.ingredientId] = true;
      return { ingredientId: l.ingredientId, qtyPer100: qty };
    });

    var dish = saveEntity_(sheet, DISH_FIELDS, 'DSH-', rows, rec);
    var recipeSheet = kitchenSheet_(SHEETS.RECIPES);
    var keep = readRows_(recipeSheet, RECIPE_FIELDS).filter(function (r) { return r.dishId !== dish.id; });
    lines.forEach(function (l) { keep.push({ dishId: dish.id, ingredientId: l.ingredientId, qtyPer100: l.qtyPer100 }); });
    replaceRows_(recipeSheet, RECIPE_FIELDS, keep);
    return { dish: dish, recipe: lines.map(function (l) { return { dishId: dish.id, ingredientId: l.ingredientId, qtyPer100: l.qtyPer100 }; }) };
  });
}

/** Saves the menu for a booking. Changing the dishes after approval flags quantities for a re-check. */
function apiSaveMenu(input) {
  return withLock_(function () {
    var ctx = kitchenContext_(input);
    var dishIds = readRows_(kitchenSheet_(SHEETS.DISHES), DISH_FIELDS).map(function (d) { return d.id; });
    var ids = [];
    (input.dishIds || []).forEach(function (id) {
      if (dishIds.indexOf(id) === -1) throw new Error('Dish ' + id + ' was not found.');
      if (ids.indexOf(id) === -1) ids.push(id);
    });
    var k = ctx.kitchen;
    var changed = k.dishIds.slice().sort().join(',') !== ids.slice().sort().join(',');
    k.dishIds = ids;
    k.menuStatus = input.menuStatus === 'Final' ? 'Final' : 'Draft';
    k.menuNotes = clean_(input.menuNotes);
    if (changed && k.qtyStatus === 'Approved') k.qtyStatus = 'Recheck';
    return writeKitchen_(ctx, k);
  });
}

/** Saves the purchase list: input.items = [{ingredientId, calc, qty, extra}]. approve=true marks it approved. */
function apiSaveQuantities(input) {
  return withLock_(function () {
    var ctx = kitchenContext_(input);
    var ingredientIds = readRows_(kitchenSheet_(SHEETS.INGREDIENTS), INGREDIENT_FIELDS).map(function (i) { return i.id; });
    var items = (input.items || []).map(function (it) {
      if (ingredientIds.indexOf(it.ingredientId) === -1) throw new Error('Ingredient ' + it.ingredientId + ' was not found.');
      var qty = Number(it.qty), calc = Number(it.calc) || 0;
      if (!(qty >= 0) || it.qty === '' || it.qty === null) throw new Error('Enter a quantity for every item (0 if not needed).');
      return { ingredientId: it.ingredientId, calc: round_(calc), qty: round_(qty), extra: !!it.extra };
    });
    var k = ctx.kitchen;
    if (!k.dishIds.length) throw new Error('Save a menu before the quantities.');
    k.items = items;
    k.guestsUsed = Number(ctx.record.guests) || '';
    k.kitchenNotes = clean_(input.kitchenNotes);
    if (input.approve) {
      k.qtyStatus = 'Approved';
      k.approvedBy = currentUser_();
      k.approvedAt = nowIso_();
    } else {
      k.qtyStatus = 'Draft';
      k.approvedBy = '';
      k.approvedAt = '';
    }
    return writeKitchen_(ctx, k);
  });
}

function kitchenContext_(input) {
  var record = readRows_(getSpreadsheet_().getSheetByName(SHEETS.RECORDS), RECORD_FIELDS)
    .filter(function (r) { return r.id === input.recordId; })[0];
  if (!record) throw new Error('Booking ' + input.recordId + ' was not found.');
  var sheet = kitchenSheet_(SHEETS.KITCHEN);
  var rows = readRows_(sheet, KITCHEN_FIELDS);
  var existing = rows.filter(function (r) { return r.recordId === record.id; })[0];
  if (existing && input.updatedAt && existing.updatedAt && input.updatedAt !== existing.updatedAt) {
    throw new Error('Someone else changed this kitchen sheet (by ' + existing.updatedBy + '). Please refresh and try again.');
  }
  var kitchen = existing ? parseKitchen_(existing) : {
    recordId: record.id, dishIds: [], menuStatus: 'Draft', menuNotes: '', guestsUsed: '', items: [],
    qtyStatus: 'Draft', approvedBy: '', approvedAt: '', kitchenNotes: '', updatedAt: '', updatedBy: ''
  };
  return { record: record, sheet: sheet, existing: existing, kitchen: kitchen };
}

function writeKitchen_(ctx, k) {
  k.updatedAt = nowIso_();
  k.updatedBy = currentUser_();
  var row = {};
  KITCHEN_FIELDS.forEach(function (f) { row[f[0]] = k[f[0]]; });
  row.dishIds = k.dishIds.join(',');
  row.items = JSON.stringify(k.items || []);
  if (ctx.existing) writeRow_(ctx.sheet, KITCHEN_FIELDS, ctx.existing._row, row);
  else appendRow_(ctx.sheet, KITCHEN_FIELDS, row);
  var out = parseKitchen_(row);
  delete out._row;
  return out;
}

function parseKitchen_(row) {
  var k = {};
  KITCHEN_FIELDS.forEach(function (f) { k[f[0]] = row[f[0]] === undefined ? '' : row[f[0]]; });
  k.dishIds = Array.isArray(row.dishIds) ? row.dishIds : String(row.dishIds || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  if (Array.isArray(row.items)) k.items = row.items;
  else {
    try { k.items = row.items ? JSON.parse(row.items) : []; } catch (e) { k.items = []; }
  }
  if (row._row) k._row = row._row;
  return k;
}

function kitchenSheet_(name) {
  var sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('The ' + name + ' sheet is missing. Run setup() again in the Apps Script editor.');
  return sheet;
}

function readOptional_(ss, name, fields) {
  var sheet = ss.getSheetByName(name);
  return sheet ? readRows_(sheet, fields) : [];
}

function findById_(rows, id, label) {
  if (!id) return null;
  var found = rows.filter(function (r) { return r.id === id; })[0];
  if (!found) throw new Error(label + ' ' + id + ' was not found. Please refresh.');
  return found;
}

function mergeEntity_(fields, input, existing) {
  var rec = {};
  fields.forEach(function (f) {
    rec[f[0]] = existing ? existing[f[0]] : '';
    if (input.hasOwnProperty(f[0]) && f[0] !== 'id') rec[f[0]] = clean_(input[f[0]]);
  });
  rec.name = String(rec.name || '').trim();
  if (existing) { rec.id = existing.id; rec._row = existing._row; }
  return rec;
}

function saveEntity_(sheet, fields, prefix, rows, rec) {
  if (rec._row) writeRow_(sheet, fields, rec._row, rec);
  else {
    rec.id = nextId_(prefix, rows);
    appendRow_(sheet, fields, rec);
  }
  delete rec._row;
  return rec;
}

function round_(n) {
  return Math.round(Number(n) * 1000) / 1000;
}

/**
 * Creates (no id) or updates (with id) an enquiry/booking.
 * Rejects double bookings and edits made on top of someone else's newer change.
 */
function apiSaveRecord(input) {
  return withLock_(function () {
    var sheet = getSpreadsheet_().getSheetByName(SHEETS.RECORDS);
    var records = readRows_(sheet, RECORD_FIELDS);
    var existing = null;
    if (input.id) {
      existing = records.filter(function (r) { return r.id === input.id; })[0];
      if (!existing) throw new Error('Record ' + input.id + ' was not found. Please refresh.');
      if (input.updatedAt && existing.updatedAt && input.updatedAt !== existing.updatedAt) {
        throw new Error('Someone else changed this record (by ' + existing.updatedBy + '). Please refresh and try again.');
      }
    }

    var rec = normaliseRecord_(input, existing);
    validateRecord_(rec);

    if (LISTS.blockingStatuses.indexOf(rec.status) !== -1) {
      var clash = records.filter(function (r) {
        return r.id !== rec.id &&
          LISTS.blockingStatuses.indexOf(r.status) !== -1 &&
          r.hall === rec.hall && r.eventDate === rec.eventDate && r.slot === rec.slot;
      })[0];
      if (clash) {
        throw new Error(rec.hall + ' is already booked on ' + rec.eventDate + ' (' + rec.slot + ') for ' +
          clash.customerName + ' [' + clash.id + '].');
      }
    }

    var now = nowIso_();
    var user = currentUser_();
    rec.updatedAt = now;
    rec.updatedBy = user;
    if (existing) {
      writeRow_(sheet, RECORD_FIELDS, existing._row, rec);
    } else {
      rec.id = nextId_('EB-', records);
      rec.createdAt = now;
      rec.createdBy = user;
      appendRow_(sheet, RECORD_FIELDS, rec);
    }
    delete rec._row;
    return rec;
  });
}

function apiAddPayment(input) {
  return withLock_(function () {
    var ss = getSpreadsheet_();
    var records = readRows_(ss.getSheetByName(SHEETS.RECORDS), RECORD_FIELDS);
    var rec = records.filter(function (r) { return r.id === input.recordId; })[0];
    if (!rec) throw new Error('Booking ' + input.recordId + ' was not found.');

    var amount = Number(input.amount);
    if (!(amount > 0)) throw new Error('Enter an amount greater than zero.');
    if (LISTS.paymentModes.indexOf(input.mode) === -1) throw new Error('Choose a payment mode.');
    if (input.date && !isIsoDate_(input.date)) throw new Error('Payment date is not valid.');

    var sheet = ss.getSheetByName(SHEETS.PAYMENTS);
    var payments = readRows_(sheet, PAYMENT_FIELDS);
    var payment = {
      id: nextId_('PAY-', payments),
      recordId: rec.id,
      date: input.date || today_(),
      amount: amount,
      mode: input.mode,
      reference: clean_(input.reference),
      recordedBy: currentUser_(),
      recordedAt: nowIso_()
    };
    appendRow_(sheet, PAYMENT_FIELDS, payment);
    return payment;
  });
}

// ---------------------------------------------------------------- record rules

function normaliseRecord_(input, existing) {
  var rec = {};
  RECORD_FIELDS.forEach(function (f) {
    var key = f[0];
    rec[key] = existing ? existing[key] : '';
    if (input.hasOwnProperty(key) && ['id', 'createdAt', 'createdBy', 'updatedAt', 'updatedBy'].indexOf(key) === -1) {
      rec[key] = clean_(input[key]);
    }
  });
  if (existing) rec._row = existing._row;
  if (!rec.status) rec.status = 'Enquiry';
  rec.phone = String(rec.phone).replace(/[^\d+]/g, '');
  rec.guests = rec.guests === '' ? '' : Number(rec.guests);
  rec.ratePerPlate = rec.ratePerPlate === '' ? '' : Number(rec.ratePerPlate);
  if (rec.totalAmount === '' && rec.guests && rec.ratePerPlate) rec.totalAmount = rec.guests * rec.ratePerPlate;
  rec.totalAmount = rec.totalAmount === '' ? '' : Number(rec.totalAmount);
  if (LISTS.openStatuses.indexOf(rec.status) !== -1 && !rec.followUpDate) rec.followUpDate = addDays_(today_(), FOLLOW_UP_DAYS);
  if (LISTS.openStatuses.indexOf(rec.status) === -1) rec.followUpDate = '';
  return rec;
}

function validateRecord_(rec) {
  var errors = [];
  if (!rec.customerName) errors.push('Customer name is required.');
  if (!/^\+?\d{10,13}$/.test(rec.phone)) errors.push('Phone number should have 10 digits.');
  if (LISTS.statuses.indexOf(rec.status) === -1) errors.push('Unknown status.');
  if (rec.eventType && LISTS.eventTypes.indexOf(rec.eventType) === -1) errors.push('Unknown event type.');
  if (rec.slot && LISTS.slots.indexOf(rec.slot) === -1) errors.push('Slot must be Morning or Evening.');
  if (rec.eventDate && !isIsoDate_(rec.eventDate)) errors.push('Event date is not valid.');
  if (rec.followUpDate && !isIsoDate_(rec.followUpDate)) errors.push('Follow-up date is not valid.');
  ['guests', 'ratePerPlate', 'totalAmount'].forEach(function (k) {
    if (rec[k] !== '' && !(rec[k] >= 0)) errors.push(k + ' must be a positive number.');
  });
  if (LISTS.blockingStatuses.indexOf(rec.status) !== -1) {
    if (!rec.hall) errors.push('Choose a hall for a booking.');
    if (!rec.eventDate) errors.push('Choose the event date for a booking.');
    if (!rec.slot) errors.push('Choose Morning or Evening for a booking.');
    if (!rec.guests) errors.push('Enter the guest count for a booking.');
  }
  if (rec.status === 'Lost' && !rec.lostReason) errors.push('Please note why the customer was lost.');
  if (errors.length) throw new Error(errors.join(' '));
}

// ---------------------------------------------------------------- sheet helpers

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || !ss.getSheetByName(SHEETS.RECORDS)) throw new Error('The app is not set up yet. Run setup() in the Apps Script editor.');
  return ss;
}

function ensureSheet_(ss, name, fields, seedRows) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, fields.length).setValues([fields.map(function (f) { return f[1]; })]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (seedRows.length) sheet.getRange(2, 1, seedRows.length, fields.length).setValues(seedRows);
  }
  fields.forEach(function (f, i) {
    if (TEXT_COLUMNS.indexOf(f[0]) !== -1) sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  });
  return sheet;
}

/** Maps columns by header text, so columns can be reordered in the sheet. */
function headerMap_(sheet, fields) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var map = {};
  fields.forEach(function (f) {
    var idx = headers.indexOf(f[1]);
    if (idx === -1) throw new Error('Column "' + f[1] + '" is missing from the ' + sheet.getName() + ' sheet.');
    map[f[0]] = idx;
  });
  return { map: map, width: headers.length };
}

function readRows_(sheet, fields) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var h = headerMap_(sheet, fields);
  var values = sheet.getRange(2, 1, last - 1, h.width).getValues();
  var tz = Session.getScriptTimeZone();
  var rows = [];
  values.forEach(function (row, i) {
    if (row.join('') === '') return;
    var obj = { _row: i + 2 };
    fields.forEach(function (f) {
      var v = row[h.map[f[0]]];
      if (v instanceof Date) {
        v = /At$/.test(f[0]) ? v.toISOString() : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      }
      obj[f[0]] = v;
    });
    rows.push(obj);
  });
  return rows;
}

function rowValues_(h, fields, obj, base) {
  var row = base ? base.slice() : new Array(h.width).fill('');
  fields.forEach(function (f) {
    var v = obj[f[0]];
    row[h.map[f[0]]] = v === undefined || v === null ? '' : v;
  });
  return row;
}

function appendRow_(sheet, fields, obj) {
  writeRow_(sheet, fields, sheet.getLastRow() + 1, obj, true);
}

function writeRow_(sheet, fields, rowNumber, obj, isNew) {
  var h = headerMap_(sheet, fields);
  // Keep dates, phones and IDs as plain text so Sheets does not reformat them.
  fields.forEach(function (f) {
    if (TEXT_COLUMNS.indexOf(f[0]) !== -1) sheet.getRange(rowNumber, h.map[f[0]] + 1).setNumberFormat('@');
  });
  var range = sheet.getRange(rowNumber, 1, 1, h.width);
  range.setValues([rowValues_(h, fields, obj, isNew ? null : range.getValues()[0])]);
}

/** Rewrites every data row of a sheet (used for recipe lines, which are replaced per dish). */
function replaceRows_(sheet, fields, rows) {
  var h = headerMap_(sheet, fields);
  var last = sheet.getLastRow();
  if (last >= 2) sheet.getRange(2, 1, last - 1, h.width).clearContent();
  if (!rows.length) return;
  fields.forEach(function (f) {
    if (TEXT_COLUMNS.indexOf(f[0]) !== -1) sheet.getRange(2, h.map[f[0]] + 1, rows.length, 1).setNumberFormat('@');
  });
  sheet.getRange(2, 1, rows.length, h.width).setValues(rows.map(function (r) { return rowValues_(h, fields, r); }));
}

function nextId_(prefix, rows) {
  var max = 0;
  rows.forEach(function (r) {
    var n = parseInt(String(r.id).replace(prefix, ''), 10);
    if (String(r.id).indexOf(prefix) === 0 && n > max) max = n;
  });
  return prefix + ('0000' + (max + 1)).slice(-4);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- small utilities

function currentUser_() {
  return Session.getActiveUser().getEmail() || 'unknown';
}

function clean_(v) {
  return typeof v === 'string' ? v.trim() : (v === undefined || v === null ? '' : v);
}

function isIsoDate_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s + 'T00:00:00Z').getTime());
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function nowIso_() {
  return new Date().toISOString();
}

function addDays_(isoDate, days) {
  var d = new Date(isoDate + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}
