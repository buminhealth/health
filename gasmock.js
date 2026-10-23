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
  setName(n) { return new Blob(this.buf, this.mime, n); }   // 실제 getBlob()은 사본이라 파일 이름은 안 바뀜
}
class Iter { constructor(a) { this.a = a; this.i = 0; } hasNext() { return this.i < this.a.length; } next() { return this.a[this.i++]; } }
class DFile {
  constructor(parent, blob) { this.id = 'f' + (++idSeq); this.parent = parent; this.blob = blob; this.desc = ''; this.trashed = false; this.updated = new Date(Date.now() + idSeq); }
  getId() { return this.id; } getName() { return this.blob.name; } getBlob() { return this.blob; }
  getDescription() { return this.desc; } setDescription(d) { this.desc = d; return this; }
  setTrashed(t) { this.trashed = t; } getSize() { return this.blob.buf.length; } getMimeType() { return this.blob.mime; }
  getLastUpdated() { return this.updated; }
  getUrl() { return 'https://drive.google.com/file/d/' + this.id + '/view'; }
  setSharing(a, p) { this.sharing = a + '/' + p; }
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
  insertRowsAfter(i, n) { this.rows.splice(i, 0, ...Array.from({ length: n }, () => [])); }
  deleteRows(i, n) { this.rows.splice(i - 1, n); }
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
    DriveApp: { createFolder: n => root.createFolder(n), getFolderById: id => { if (!folders[id]) throw new Error('no folder'); return folders[id]; },
                Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' } },
    MimeType: { PLAIN_TEXT: 'text/plain' },
    MailApp: { sent: [], sendEmail(o) { this.sent.push(o); } },
    // 가짜 Firebase: signInWithPassword + Firestore 목록 (FIREBASE.users / FIREBASE.docs 로 데이터 지정)
    FIREBASE: { users: {}, docs: [], calls: [] },
    GEMINI: { status: 200, reply: {}, calls: [] },
    KOSHA: { okKey: 'koshaKey1234567890abcdefXYZ', calls: [] },
    LAW: { okOc: 'bumin3131', calls: [], byl: {
      '100': [
        { '별표번호': '0022', '별표구분': '별표', '별표제목': '특수건강진단 대상 유해인자(제201조 관련)', '별표내용': [['1. 화학적 인자', '가. 유기화합물', '1) 가솔린(Gasoline; 8006-61-9)', '2) 글루타르알데히드(Glutaraldehyde; 111-30-8)', '라. 산 및 알칼리류', '1) 염화수소(Hydrogen chloride; 7647-01-0)']] },
        { '별표번호': '0023', '별표구분': '별표', '별표제목': '특수건강진단의 시기 및 주기(제202조 관련)', '별표내용': [['1 N,N-디메틸아세트아미드, 디메틸포름아미드 … 6개월', '2 벤젠 … 6개월', '3 1,1,2,2-테트라클로로에탄, 사염화탄소, 아크릴로니트릴 … 6개월', '광물성 분진 … 24개월', '그 밖의 유해인자 … 12개월']] },
        { '별표번호': '0022', '별표구분': '서식', '별표제목': '서식 22', '별표내용': 'x' }] } },
    UrlFetchApp: {
      fetch(u, opt) {
        const fb = this.__fb;
        fb.calls.push(u);
        const resp = (code, obj) => ({ getResponseCode: () => code, getContentText: () => JSON.stringify(obj) });
        if (u.startsWith('https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword')) {
          const b = JSON.parse(opt.payload);
          if (fb.users[b.email] !== b.password) return resp(400, { error: { message: 'INVALID_LOGIN_CREDENTIALS' } });
          return resp(200, { idToken: 'fb-id-token' });
        }
        if (u === 'https://bumin-leave.web.app/__/firebase/init.json') {
          return resp(200, { apiKey: 'AIzaFromSite', projectId: 'bumin-leave', authDomain: 'bumin-leave.firebaseapp.com' });
        }
        // 가짜 국가법령정보 OPEN API (DRF): LAW.okOc 인 OC만 통과
        if (u.startsWith('https://www.law.go.kr/DRF/')) {
          const lw = env.LAW; const oc = decodeURIComponent((u.match(/OC=([^&]+)/) || [])[1] || '');
          lw.calls.push(u);
          if (oc !== lw.okOc) return resp(200, { result: '사용자 정보 검증에 실패하였습니다.' });
          if (u.indexOf('lawSearch.do') >= 0) {
            const q = decodeURIComponent((u.match(/query=([^&]+)/) || [])[1] || '');
            return resp(200, { LawSearch: { law: [{ '법령명한글': q, '법령일련번호': q.indexOf('기준') >= 0 ? '200' : '100', '시행일자': '20250101' }] } });
          }
          const mst = (u.match(/MST=(\d+)/) || [])[1];
          const jo = mst === '200' ? [{ '조문번호': '3', '조문여부': '조문', '조문제목': '전도의 방지', '조문내용': '제3조(전도의 방지) 사업주는 근로자가 작업장에서 넘어지거나 미끄러지는 등의 위험이 없도록 작업장 바닥 등을 안전하고 청결한 상태로 유지하여야 한다.' },
            { '조문번호': '1', '조문여부': '전문', '조문제목': '' },
            { '조문번호': '597', '조문여부': '조문', '조문제목': '혈액노출 예방 조치', '조문내용': '제597조(혈액노출 예방 조치) 사업주는 혈액노출 위험이 있는 작업을 하는 근로자에게 주사침 등 날카로운 도구를 안전하게 처리하도록 하여야 한다.' }]
            .concat(Array.from({ length: 60 }, (_, i) => ({ '조문번호': String(100 + i), '조문여부': '조문', '조문제목': '시험조문' + i, '조문내용': '시험 내용 ' + i }))) : [];
          return resp(200, { '법령': { '조문': { '조문단위': jo }, '별표': { '별표단위': (lw.byl[mst] || []) } } });
        }
        // 가짜 안전보건공단 OPEN API (공공데이터포털): KOSHA.okKey 인 serviceKey만 통과
        if (u.startsWith('https://apis.data.go.kr/B552468/')) {
          const ks = env.KOSHA; ks.calls.push(u);
          const key = decodeURIComponent((u.match(/serviceKey=([^&]+)/) || [])[1] || '');
          if (key !== ks.okKey) return { getResponseCode: () => 401, getContentText: () => '<OpenAPI_ServiceResponse><cmmMsgHeader><returnAuthMsg>SERVICE_KEY_IS_NOT_REGISTERED_ERROR</returnAuthMsg></cmmMsgHeader></OpenAPI_ServiceResponse>' };
          const kw = decodeURIComponent((u.match(/keyword=([^&]+)/) || [])[1] || '');
          return resp(200, { response: { header: { resultCode: '00' }, body: { items: { item: [
            { title: kw + ' 취급 안전 가이드', filepath: 'https://portal.kosha.or.kr/media/1', regDate: '2026-01-02' },
            { title: kw + ' 노출 재해 사례', contents: kw + '을(를) 다루다 눈에 튀어 각막 손상', filepath: 'https://portal.kosha.or.kr/media/2' }] }, totalCount: 2 } } });
        }
        // 가짜 Gemini: GEMINI.status / GEMINI.reply 로 응답 지정, 받은 요청은 GEMINI.calls
        if (u.startsWith('https://generativelanguage.googleapis.com/v1beta/models/')) {
          const gm = env.GEMINI;
          gm.calls.push({ url: u, key: opt.headers && opt.headers['x-goog-api-key'], body: JSON.parse(opt.payload) });
          if (gm.status !== 200) return resp(gm.status, { error: { message: 'quota' } });
          return resp(200, { candidates: [{ content: { parts: [{ text: '```json\n' + JSON.stringify(gm.reply) + '\n```' }] } }] });
        }
        if (u.startsWith('https://firestore.googleapis.com/v1/projects/')) {
          if (!opt || !opt.headers || opt.headers.Authorization !== 'Bearer fb-id-token') return resp(403, {});
          if (fb.quota) return resp(429, { error: { message: 'Quota exceeded.', status: 'RESOURCE_EXHAUSTED' } });
          // 조건 조회 (runQuery): fieldFilter / compositeFilter(AND) 의 >= · == 만 흉내
          if (u.endsWith(':runQuery')) {
            const q = JSON.parse(opt.payload).structuredQuery;
            const coll = q.from[0].collectionId;
            const docs = (fb.collections && fb.collections[coll]) || [];
            const val = v => v ? ('stringValue' in v ? v.stringValue : ('timestampValue' in v ? String(v.timestampValue).slice(0, 10) : '')) : undefined;
            const one = f => d => { const fv = val((d.fields || {})[f.field.fieldPath]); const want = val(f.value);
              if (fv === undefined) return false;
              return f.op === 'EQUAL' ? fv === want : f.op === 'GREATER_THAN_OR_EQUAL' ? fv >= want : fv <= want; };
            const filters = q.where.compositeFilter ? q.where.compositeFilter.filters.map(x => x.fieldFilter) : [q.where.fieldFilter];
            fb.queries = (fb.queries || []).concat([{ coll, filters: filters.map(f => f.field.fieldPath + f.op) }]);
            const hits = docs.filter(d => filters.every(f => one(f)(d)));
            return resp(200, hits.length ? hits.map(d => ({ document: d, readTime: 'x' })) : [{ readTime: 'x' }]);
          }
          const coll = decodeURIComponent(u.split('/documents/')[1].split('?')[0]);
          const ps = (u.match(/pageSize=(\d+)/) || [])[1];
          if (fb.collections && fb.collections[coll]) return resp(200, { documents: ps ? fb.collections[coll].slice(0, +ps) : fb.collections[coll] });
          if (fb.collections && fb.denied && fb.denied.indexOf(coll) >= 0) return resp(403, { error: { message: 'Missing or insufficient permissions.' } });
          return resp(200, { documents: fb.docs });
        }
        return resp(404, {});
      },
    },
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
  env.UrlFetchApp.__fb = env.FIREBASE;
  return { env, sheets, props, root, Sheet };
}
function load(codePath) {
  const m = makeEnv();
  const ctx = vm.createContext(m.env);
  vm.runInContext(fs.readFileSync(codePath, 'utf8') + '\n;this.__api={doGet,doPost,sha256Hex_,sendReminders_,nhRows_,nhWrite_,dailyCheck,fbConfigFromInput_};', ctx);
  return Object.assign(m, { ctx, api: ctx.__api });
}
module.exports = { load };
