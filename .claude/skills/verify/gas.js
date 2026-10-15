// Loads the real apps-script/Code.gs in a VM with a fake SpreadsheetApp that mimics Sheets type coercion.
const fs = require('fs'), vm = require('vm');
function makeSheet(name) {
  const rows = [], fmt = {};
  const coerce = (v, r, c) => {
    if (fmt[r + ',' + c] === '@' || typeof v !== 'string') return v;
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return new Date(v + 'T00:00:00+05:30');          // Sheets turns ISO dates into Date cells
    if (/^-?\d+(\.\d+)?$/.test(v.trim()) && v.trim() !== '') return Number(v);              // and numeric text into numbers
    return v;
  };
  const sheet = {
    getName: () => name,
    getLastRow: () => { for (let i = rows.length - 1; i >= 0; i--) if (rows[i] && rows[i].some(v => v !== '' && v !== undefined)) return i + 1; return 0; },
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r ? r.length : 0), 0),
    getMaxRows: () => Math.max(rows.length, 1000),
    setFrozenRows() {},
    getRange: (r, c, nr = 1, nc = 1) => ({
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = (rows[r - 1 + i] || [])[c - 1 + j]; return v === undefined ? '' : (v instanceof Date ? new Date(v) : v); })),
      setValues(vals) { vals.forEach((row, i) => { rows[r - 1 + i] = rows[r - 1 + i] || []; row.forEach((v, j) => { rows[r - 1 + i][c - 1 + j] = coerce(v, r + i, c + j); }); }); return this; },
      clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (rows[r - 1 + i]) rows[r - 1 + i][c - 1 + j] = ''; },
      setNumberFormat(f) { for (let i = 0; i < Math.min(nr, 3000); i++) for (let j = 0; j < nc; j++) fmt[(r + i) + ',' + (c + j)] = f; return this; },
      setFontWeight() { return this; }
    }),
    _rows: rows
  };
  return sheet;
}
function createBackend(codePath) {
  const sheets = {}, props = {};
  const ss = { getId: () => 'ss-1', getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = makeSheet(n)) };
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    Session: { getActiveUser: () => ({ getEmail: () => ctx.__user }), getScriptTimeZone: () => 'Asia/Kolkata' },
    Utilities: { formatDate: (d, tz, f) => new Date(d.getTime() + 5.5 * 3600e3).toISOString().slice(0, 10) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    LanguageApp: { translate: (t) => ({ 'paneer tikka': 'पनीर टिक्का', 'paneer': 'पनीर', 'butter': 'मक्खन' }[t.toLowerCase()] || ('हिं:' + t)) },
    HtmlService: {}, __user: 'owner@example.com'
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(codePath, 'utf8'), ctx);
  ctx.setup();
  return { ctx, sheets,
    call(user, fn, arg) {
      ctx.__user = user;
      if (!/^api[A-Z]/.test(fn) || typeof ctx[fn] !== 'function') throw new Error('Script function not found: ' + fn);
      const out = ctx[fn](arg === undefined ? undefined : JSON.parse(JSON.stringify(arg)));
      // google.script.run cannot return Date objects; surface any leak as the failure the real client would see.
      const walk = (v, path) => { if (v instanceof Date) throw new Error('Date returned at ' + path + ' (google.script.run would fail)');
        if (v && typeof v === 'object') Object.keys(v).forEach(k => walk(v[k], path + '.' + k)); };
      walk(out, fn);
      return JSON.parse(JSON.stringify(out === undefined ? null : out));
    }
  };
}
module.exports = { createBackend };
