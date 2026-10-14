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

var SHEETS = { RECORDS: 'Records', PAYMENTS: 'Payments', HALLS: 'Halls' };

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
  paymentModes: ['Cash', 'UPI', 'Bank transfer', 'Cheque', 'Online link']
};

var TEXT_COLUMNS = ['id', 'createdAt', 'updatedAt', 'phone', 'eventDate', 'followUpDate', 'recordId', 'date', 'recordedAt'];
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
    payments: readRows_(ss.getSheetByName(SHEETS.PAYMENTS), PAYMENT_FIELDS)
  };
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
