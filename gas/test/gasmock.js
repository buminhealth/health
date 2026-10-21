// Minimal in-memory stand-ins for the Apps Script services Code.gs uses, so the real
// Code.gs can run under Node for testing. Not a full emulation.
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');

const toSigned = buf => Array.from(buf, b => (b > 127 ? b - 256 : b));
const toBuf = arr => Buffer.from(arr.map(b => b & 0xff));
const asBuf = (v, charset) => typeof v === 'string' ? Buffer.from(v, 'utf8') : toBuf(v);

let idSeq = 0;
class Blob {
  constructor(bytes, mime, name) { this.buf = bytes; this.mime = mime || 'application/octet-stream'; this.name = name || ''; }
  getBytes() { return toSigned(this.buf); }
  getDataAsString() { return this.buf.toString('utf8'); }
}
class Iter { constructor(a) { this.a = a; this.i = 0; } hasNext() { return this.i < this.a.length; } next() { return this.a[this.i++]; } }
class DFile {
  constructor(parent, blob) { this.id = 'f' + (++idSeq); this.parent = parent; this.blob = blob; this.desc = ''; this.trashed = false; this.updated = new Date(Date.now() + idSeq); }
  getId() { return this.id; } getName() { return this.blob.name; } getBlob() { return this.blob; }
  getDescription() { return this.desc; } setDescription(d) { this.desc = d; return this; }
  setTrashed(t) { this.trashed = t; } getSize() { return this.blob.buf.length; } getMimeType() { return this.blob.mime; }
  getLastUpdated() { return this.updated; }
}
const folders = {};
class DFolder {
  constructor(name, parent) { this.id = 'd' + (++idSeq); this.name = name; this.parent = parent; this.files = []; this.folders = []; folders[this.id] = this; }
  getId() { return this.id; } getName() { return this.name; } getUrl() { return 'https://drive/' + this.id; }
  getFoldersByName(n) { return new Iter(this.folders.filter(f => f.name === n)); }
  getFolders() { return new Iter(this.folders.slice()); }
  createFolder(n) { const f = new DFolder(n, this); this.folders.push(f); return f; }
  getFilesByName(n) { return new Iter(this.files.filter(f => !f.trashed && f.getName() === n)); }
  getFiles() { return new Iter(this.files.filter(f => !f.trashed)); }
  createFile(a, content, mime) {
    const blob = a instanceof Blob ? a : new Blob(Buffer.from(content, 'utf8'), mime, a);
    const f = new DFile(this, blob); this.files.push(f); return f;
  }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr: nr || 1, nc: nc || 1 }); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.cell(this.r + i, this.c + j)); out.push(row); } return out; }
  getDisplayValues() { return this.getValues().map(r => r.map(v => v instanceof Date ? v.toISOString().slice(0, 10) : String(v))); }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  setValue(x) { this.sh.set(this.r, this.c, x); return this; }
  setNumberFormat() { return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.rows = []; }
  cell(r, c) { const row = this.rows[r - 1] || []; const v = row[c - 1]; return v === undefined ? '' : v; }
  set(r, c, v) { while (this.rows.length < r) this.rows.push([]); this.rows[r - 1][c - 1] = v; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); }
  appendRow(a) { this.rows.push(a.slice()); }
  insertRowBefore(i) { this.rows.splice(i - 1, 0, []); }
  deleteRow(i) { this.rows.splice(i - 1, 1); }
  getRange(a, b, c, d) { if (typeof a === 'string') return new Range(this, 1, 2, 1, 1); return new Range(this, a, b, c, d); }
  getDataRange() { return new Range(this, 1, 1, this.rows.length, this.getLastColumn()); }
  setFrozenRows() {}
}
function makeEnv() {
  const sheets = {};
  const props = {};
  const cache = {};
  const root = new DFolder('root', null);
  const env = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: n => sheets[n] || null,
        insertSheet: n => (sheets[n] = new Sheet(n)),
      }),
      getUi: () => { throw new Error('no ui in test'); },
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = String(v); }, remove: k => { delete cache[k]; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    DriveApp: { createFolder: n => root.createFolder(n), getFolderById: id => { if (!folders[id]) throw new Error('no folder'); return folders[id]; } },
    MimeType: { PLAIN_TEXT: 'text/plain' },
    MailApp: { sent: [], sendEmail(o) { this.sent.push(o); } },
    ScriptApp: { getProjectTriggers: () => [] },
    ContentService: {
      MimeType: { JAVASCRIPT: 'js', JSON: 'json' },
      createTextOutput: t => ({ text: t, setMimeType() { return this; } }),
    },
    Utilities: {
      Charset: { UTF_8: 'utf8' },
      DigestAlgorithm: { SHA_256: 'sha256', SHA_1: 'sha1' },
      getUuid: () => crypto.randomUUID(),
      base64Encode: (v) => asBuf(v).toString('base64'),
      base64Decode: s => toSigned(Buffer.from(s, 'base64')),
      base64EncodeWebSafe: (v) => asBuf(v).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      base64DecodeWebSafe: s => toSigned(Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64')),
      computeHmacSha256Signature: (v, k) => toSigned(crypto.createHmac('sha256', k).update(v).digest()),
      computeDigest: (alg, v) => toSigned(crypto.createHash(alg).update(asBuf(v)).digest()),
      newBlob: (v, mime, name) => new Blob(asBuf(v), mime, name),
      formatDate: (d, tz, fmt) => {
        const k = new Date(d.getTime() + 9 * 3600 * 1000).toISOString();
        return fmt.indexOf('HH') >= 0 ? k.slice(0, 19) : k.slice(0, 10);
      },
    },
  };
  return { env, sheets, props, root, Sheet };
}
function load(codePath) {
  const m = makeEnv();
  const ctx = vm.createContext(m.env);
  vm.runInContext(fs.readFileSync(codePath, 'utf8') + '\n;this.__api={doGet,doPost,sha256Hex_,sendReminders_,nhRows_,nhWrite_};', ctx);
  return Object.assign(m, { ctx, api: ctx.__api });
}
module.exports = { load };
