// Offline test harness for apps-script/Code.gs.
// Loads the backend into a Node VM with an in-memory Google Sheets/Drive mock,
// so scheduling and storage logic can be exercised without touching any real
// spreadsheet. Run tests with:  TZ=Asia/Riyadh node apps-script/tests/run.js
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function pad(n) { return String(n).padStart(2, '0'); }
const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function formatDate(d, _tz, fmt) {
  // The harness runs with TZ=Asia/Riyadh so local getters match APP.timezone.
  return fmt.replace('yyyy', d.getFullYear()).replace('MMM', MONTHS[d.getMonth()])
    .replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate()))
    .replace('EEE', DAYS[d.getDay()]).replace('HH', pad(d.getHours()))
    .replace('mm', pad(d.getMinutes())).replace('ss', pad(d.getSeconds()));
}

// ---- Minimal Sheets mock: a sheet is a 2D array of values (row 0 = headers) ----
class MockRange {
  constructor(sheet, row, col, nRows, nCols) { Object.assign(this, { sheet, row, col, nRows, nCols }); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nRows; r++) {
      const src = this.sheet.data[this.row - 1 + r] || [];
      const line = [];
      for (let c = 0; c < this.nCols; c++) line.push(src[this.col - 1 + c] === undefined ? '' : src[this.col - 1 + c]);
      out.push(line);
    }
    return out;
  }
  setValues(values) {
    this.sheet.ss.failWrites();
    values.forEach((line, r) => {
      const idx = this.row - 1 + r;
      while (this.sheet.data.length <= idx) this.sheet.data.push([]);
      line.forEach((v, c) => { this.sheet.data[idx][this.col - 1 + c] = v; });
    });
    return this;
  }
  clearContent() {
    this.sheet.ss.failWrites();
    for (let r = 0; r < this.nRows; r++) {
      const line = this.sheet.data[this.row - 1 + r];
      if (line) for (let c = 0; c < this.nCols; c++) line[this.col - 1 + c] = '';
    }
    return this;
  }
  clear() { return this.clearContent(); }
  getRow() { return this.row; }
  setFontWeight() { return this; } setBackground() { return this; } setFontColor() { return this; }
  createTextFinder(text) {
    const self = this;
    return { useRegularExpression() { return this; }, matchEntireCell() { return this; }, matchCase() { return this; },
      findNext() {
        for (let r = 0; r < self.nRows; r++) {
          const v = (self.sheet.data[self.row - 1 + r] || [])[self.col - 1];
          if (String(v) === String(text)) return new MockRange(self.sheet, self.row + r, self.col, 1, 1);
        }
        return null;
      } };
  }
}
class MockSheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.data = []; this.maxRows = 1000; }
  getName() { return this.name; }
  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) if ((this.data[i] || []).some(v => v !== '' && v !== null && v !== undefined)) return i + 1;
    return 0;
  }
  getLastColumn() { return Math.max(0, ...this.data.map(r => { let n = 0; (r || []).forEach((v, i) => { if (v !== '' && v !== undefined) n = i + 1; }); return n; })); }
  getMaxRows() { return Math.max(this.maxRows, this.data.length); }
  getRange(row, col, nRows, nCols) { return new MockRange(this, row, col, nRows || 1, nCols || 1); }
  getDataRange() { return new MockRange(this, 1, 1, this.getLastRow(), this.getLastColumn()); }
  deleteRows(start, count) {
    this.ss.failWrites();
    if (count >= this.getMaxRows() - 1) throw new Error('Cannot delete all non-frozen rows');
    this.data.splice(start - 1, count);
  }
  deleteRow(i) { this.deleteRows(i, 1); }
  insertRowAfter() { this.maxRows++; }
  setFrozenRows() {} autoResizeColumns() {}
  clear() { this.data = []; }
}
class MockSpreadsheet {
  constructor(id) { this.id = id; this.sheets = {}; this.writeBudget = Infinity; }
  failWrites() { if (this.writeBudget-- <= 0) throw new Error('Simulated Sheets service failure'); }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { return (this.sheets[n] = new MockSheet(this, n)); }
}

function createEnv() {
  const files = {};
  let fileSeq = 0;
  const db = new MockSpreadsheet('DB');
  files.DB = db;
  const store = () => { const m = {}; return {
    get: k => (k in m ? m[k] : null), put: (k, v) => { m[k] = v; }, remove: k => { delete m[k]; },
    getAll: ks => { const o = {}; ks.forEach(k => { if (k in m) o[k] = m[k]; }); return o; }, putAll: o => { Object.assign(m, o); }, removeAll: ks => { ks.forEach(k => { delete m[k]; }); },
    getProperty: k => (k in m ? m[k] : null), setProperty: (k, v) => { m[k] = String(v); },
    deleteProperty: k => { delete m[k]; }, getProperties: () => Object.assign({}, m) }; };
  const cache = store(), props = store();
  let uuid = 0;
  const sandbox = {
    console,
    Utilities: {
      formatDate, getUuid: () => (0xabc00000 + (++uuid)).toString(16).padStart(8, '0') + '-0000-4000-8000-000000000000', // unique in the first 8 chars, like makeId_ needs
      computeDigest: (_a, s) => Array.from(require('crypto').createHash('sha256').update(String(s)).digest()).map(b => (b > 127 ? b - 256 : b)),
      base64EncodeWebSafe: b => Buffer.from(b.map(x => (x + 256) % 256)).toString('base64url'),
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => db,
      openById: id => { if (sandbox.__openFails || !files[id]) throw new Error('Simulated openById failure'); return files[id]; },
      create: name => { const id = 'ARCH' + (++fileSeq); files[id] = new MockSpreadsheet(id); files[id].name = name; return files[id]; },
      flush: () => {}
    },
    DriveApp: { getFileById: id => ({ getParents: () => ({ hasNext: () => false }), setTrashed: () => { sandbox.__trashed.push(id); } }),
      getRootFolder: () => ({ removeFile() {} }) },
    CacheService: { getScriptCache: () => cache },
    PropertiesService: { getScriptProperties: () => props },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => '' }) },
    ContentService: { createTextOutput: t => ({ setMimeType() { return t; } }), MimeType: { JSON: 'json' } },
    __files: files, __db: db, __trashed: [], __openFails: false
  };
  vm.createContext(sandbox);
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  vm.runInContext(code + '\n;this.__HEADERS = HEADERS; this.__SHEETS = SHEETS;', sandbox, { filename: 'Code.gs' });
  return sandbox;
}

// Seeds a sheet from an array of objects using HEADERS order.
function seed(env, sheetName, objects) {
  const headers = env.__HEADERS[sheetName];
  const sh = env.__db.getSheetByName(sheetName) || env.__db.insertSheet(sheetName);
  sh.data = [headers.slice()].concat(objects.map(o => headers.map(h => (o[h] === undefined ? '' : o[h]))));
}

module.exports = { createEnv, seed };
