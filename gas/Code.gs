// ════════════════════════════════════════════════════════════════
// 부민병원 MSDS 통합시스템 - GAS Code.gs v2 (1단계: 서버 인증 + 비공개 저장소)
//
// 변경 요약 (v1 → v2)
//  · 부서 PIN / 관리자 비밀번호 로그인 → 서버 서명 토큰(12시간) 발급
//  · 직원목록(emp)·기록 저장/조회·설정 변경은 토큰 필수, 부서 토큰은 자기 부서만
//  · GitHub 저장소 대체: 구글 드라이브 비공개 폴더에 같은 경로 구조로 저장
//    (HTML은 GitHub API 형식 그대로 호출 → 화면 쪽 프록시가 이 서버로 전달)
//  · 버그 수정: 날짜 하루 당김(UTC), 동시 제출 순번 꼬임(LockService), 수식 주입
//  · deleteRecords 구현 (앱에서 삭제한 기록을 시트에서도 삭제)
//
// 최초 설정: 시트 새로고침 → 상단 메뉴 [보건관리 시스템] → [① 초기 설정]
// ════════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();
const PROPS = PropertiesService.getScriptProperties();
const TZ = 'Asia/Seoul';

const TOKEN_TTL_MS = 12 * 60 * 60 * 1000;   // 로그인 유지 12시간
const LOCK_MAX_FAILS = 5;                    // PIN/비밀번호 5회 오류 시
const LOCK_SECONDS = 15 * 60;                // 15분 잠금
const INLINE_LIMIT = 1024 * 1024;            // 1MB 초과 파일은 분할 다운로드 (GitHub API와 동일 기준)
const CHUNK_CHARS = 4 * 1024 * 1024;         // 분할 단위 (base64 문자 수, 4의 배수)
const AUTH_SHEET = '부서인증';

// 부서 토큰이 쓸 수 있는 기록 파일 (key: 경로, value: 배열 위치)
const RECORD_FILES = {
  'records/edu_records.json':  'array',
  'records/ppe_records.json':  'records',
  'records/chem_records.json': 'array',
};
// 부서 토큰이 파일을 올리고 받을 수 있는 폴더
const DEPT_FILE_PREFIXES = ['교육일지/', '보호구/', '화학물질/', 'records/chem_msds/'];

// ════════════════════════════════════════════════════════════════
// 시트 메뉴 (최초 설정 · 관리자 비밀번호)
// ════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu('보건관리 시스템')
    .addItem('① 초기 설정', 'setupSystem')
    .addItem('② 관리자 비밀번호 변경', 'changeAdminPassword')
    .addItem('③ 모든 로그인 강제 해제', 'revokeAllSessions')
    .addToUi();
}

function setupSystem() {
  const ui = SpreadsheetApp.getUi();
  if (!PROPS.getProperty('TOKEN_SECRET')) {
    PROPS.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  if (!PROPS.getProperty('DATA_ROOT_ID')) {
    const folder = DriveApp.createFolder('부민보건관리_데이터 (비공개)');
    PROPS.setProperty('DATA_ROOT_ID', folder.getId());
  }
  let sheet = SS.getSheetByName(AUTH_SHEET);
  if (!sheet) {
    sheet = SS.insertSheet(AUTH_SHEET);
    sheet.appendRow(['부서', 'PIN', '담당자', '사용(Y/N)', '부서장 이메일', '비고']);   // 이메일: 신규입사 알림(5단계)용
    sheet.getRange('B:B').setNumberFormat('@');   // PIN 앞자리 0 유지
    sheet.setFrozenRows(1);
    const depts = getEmployeeDepts_();
    depts.forEach(d => sheet.appendRow([d, '', '', 'N', '', 'PIN 입력 후 사용을 Y로 변경']));
  }
  if (!PROPS.getProperty('ADMIN_PW_HASH')) changeAdminPassword();
  const root = DriveApp.getFolderById(PROPS.getProperty('DATA_ROOT_ID'));
  ui.alert('초기 설정 완료',
    '1) [' + AUTH_SHEET + '] 시트에 부서별 PIN(4~6자리)을 넣고 사용을 Y로 바꾸세요.\n' +
    '   부서 이름은 직원목록 시트의 부서 이름과 같아야 합니다.\n' +
    '2) 기존 GitHub 자료(records, 교육일지, 보호구, 화학물질 폴더)를 아래 드라이브 폴더에 그대로 올리세요.\n' +
    '   ' + root.getUrl() + '\n' +
    '3) 배포 → 배포 관리 → 새 버전으로 웹 앱을 다시 배포하세요.', ui.ButtonSet.OK);
}

function changeAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('관리자 비밀번호 설정', '새 관리자 비밀번호 (8자 이상)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const pw = r1.getResponseText().trim();
  if (pw.length < 8) { ui.alert('8자 이상이어야 합니다.'); return; }
  const r2 = ui.prompt('관리자 비밀번호 확인', '한 번 더 입력하세요', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK || r2.getResponseText().trim() !== pw) {
    ui.alert('두 입력이 일치하지 않습니다.'); return;
  }
  const salt = Utilities.getUuid();
  PROPS.setProperty('ADMIN_PW_SALT', salt);
  PROPS.setProperty('ADMIN_PW_HASH', sha256Hex_(salt + pw));
  ui.alert('관리자 비밀번호가 변경되었습니다. 기존 관리자 로그인은 모두 해제됩니다.');
}

function revokeAllSessions() {
  PROPS.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  SpreadsheetApp.getUi().alert('모든 기기의 로그인이 해제되었습니다.');
}

// ════════════════════════════════════════════════════════════════
// 진입점
// ════════════════════════════════════════════════════════════════

// GET (JSONP) — 기존 HTML 호출 호환. 보호 대상은 t=토큰 필요
function doGet(e) {
  const p = e.parameter || {};
  const cb = /^[A-Za-z0-9_$.]+$/.test(p.callback || '') ? p.callback : 'callback';
  let result;
  try {
    const s = verifyToken_(p.t);
    switch (p.action) {
      // 공개 (개인정보 없음)
      case 'chems':    result = getChemicals();   break;
      case 'edutpl':   result = getEduTemplate(); break;
      case 'ppeItems': result = getPpeItems();    break;
      // 로그인 필요
      case 'emp':        result = requireLogin_(s) || getEmployees(s);  break;
      case 'record':     result = requireLogin_(s) || saveRecord(p, s); break;
      case 'records':    result = requireLogin_(s) || getRecords(p, s); break;
      case 'allRecords': result = requireLogin_(s) || getAllRecords(p, s); break;
      case 'adminCfg':
        result = requireLogin_(s) ||
          (p.method === 'set' ? (requireAdmin_(s) || setAdminCfg(p)) : getAdminCfg());
        break;
      case 'deleteRecords': result = requireAdmin_(s) || deleteSheetRecords(p); break;
      default: result = { success: false, error: 'unknown action: ' + p.action };
    }
  } catch (err) {
    result = { success: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// POST (JSON) — 로그인 · 파일 저장소
function doPost(e) {
  let result;
  try {
    const req = JSON.parse((e.postData && e.postData.contents) || '{}');
    const s = verifyToken_(req.t);
    switch (req.op) {
      case 'depts':     result = { success: true, data: listActiveDepts_() }; break;
      case 'login':     result = login_(req); break;
      case 'whoami':    result = s ? { success: true, role: s.role, dept: s.dept || '', exp: s.exp }
                                   : { success: false, status: 401, error: '로그인 필요' }; break;
      case 'fs.get':    result = requireLogin_(s) || fsGet_(req, s); break;
      case 'fs.put':    result = requireLogin_(s) || fsPut_(req, s); break;
      default:          result = { success: false, error: 'unknown op: ' + req.op };
    }
  } catch (err) {
    result = { success: false, status: 500, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// 인증
// ════════════════════════════════════════════════════════════════
function requireLogin_(s) {
  return s ? null : { success: false, status: 401, error: '로그인 필요' };
}
function requireAdmin_(s) {
  if (!s) return { success: false, status: 401, error: '로그인 필요' };
  return s.role === 'admin' ? null : { success: false, status: 403, error: '관리자 권한 필요' };
}

function login_(req) {
  const kind = req.kind === 'admin' ? 'admin' : 'dept';
  const dept = String(req.dept || '').trim();
  const lockKey = 'fail_' + (kind === 'admin' ? 'admin' : sha256Hex_(dept).slice(0, 16));
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get(lockKey) || '0', 10);
  if (fails >= LOCK_MAX_FAILS) {
    return { success: false, status: 423, error: '오류 ' + LOCK_MAX_FAILS + '회로 15분간 잠겼습니다. 잠시 후 다시 시도하세요.' };
  }

  let ok = false, payload = null;
  if (kind === 'admin') {
    const salt = PROPS.getProperty('ADMIN_PW_SALT') || '';
    const hash = PROPS.getProperty('ADMIN_PW_HASH') || '';
    ok = !!hash && sha256Hex_(salt + String(req.password || '')) === hash;
    if (ok) payload = { role: 'admin', v: hash.slice(0, 12) };
  } else {
    const row = findDeptAuth_(dept);
    ok = !!row && row.pin !== '' && row.pin === String(req.pin || '').trim();
    if (ok) payload = { role: 'dept', dept: row.dept, v: sha256Hex_(row.pin).slice(0, 12) };
  }

  if (!ok) {
    cache.put(lockKey, String(fails + 1), LOCK_SECONDS);
    const remain = LOCK_MAX_FAILS - fails - 1;
    return { success: false, status: 401,
      error: remain > 0 ? '부서 또는 PIN/비밀번호가 올바르지 않습니다 (남은 시도 ' + remain + '회)'
                        : '오류 ' + LOCK_MAX_FAILS + '회로 15분간 잠겼습니다.' };
  }
  cache.remove(lockKey);
  payload.exp = Date.now() + TOKEN_TTL_MS;
  return { success: true, token: signToken_(payload), role: payload.role, dept: payload.dept || '', exp: payload.exp };
}

function signToken_(payload) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify(payload), Utilities.Charset.UTF_8);
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(body, PROPS.getProperty('TOKEN_SECRET')));
  return body + '.' + sig;
}

// 토큰 검증: 서명·만료 + PIN/비밀번호가 바뀌었으면 무효
function verifyToken_(t) {
  if (!t || typeof t !== 'string' || t.indexOf('.') < 0) return null;
  const secret = PROPS.getProperty('TOKEN_SECRET');
  if (!secret) return null;
  const parts = t.split('.');
  const expected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(parts[0], secret));
  if (expected !== parts[1]) return null;
  let s;
  try { s = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString('UTF-8')); }
  catch (_) { return null; }
  if (!s.exp || s.exp < Date.now()) return null;
  if (s.role === 'admin') {
    return (PROPS.getProperty('ADMIN_PW_HASH') || '').slice(0, 12) === s.v ? s : null;
  }
  if (s.role === 'dept') {
    const row = findDeptAuth_(s.dept);
    return row && row.pin !== '' && sha256Hex_(row.pin).slice(0, 12) === s.v ? s : null;
  }
  return null;
}

function readAuthRows_() {
  const sheet = SS.getSheetByName(AUTH_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues()
    .map(r => ({ dept: String(r[0]).trim(), pin: String(r[1]).trim(), active: /^(Y|예|사용|O)$/i.test(String(r[3]).trim()) }))
    .filter(r => r.dept);
}
function findDeptAuth_(dept) {
  const d = String(dept || '').trim();
  return readAuthRows_().find(r => r.active && r.dept === d) || null;
}
function listActiveDepts_() {
  return readAuthRows_().filter(r => r.active && r.pin).map(r => r.dept);
}

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(b => ((b & 0xff) + 0x100).toString(16).slice(1)).join('');
}
function sha1HexBytes_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, bytes)
    .map(b => ((b & 0xff) + 0x100).toString(16).slice(1)).join('');
}

// ════════════════════════════════════════════════════════════════
// 파일 저장소 (GitHub contents API 형식 호환, 구글 드라이브 비공개 폴더)
// ════════════════════════════════════════════════════════════════
function normPath_(p) {
  const path = String(p || '').replace(/^\/+|\/+$/g, '');
  if (!path || path.split('/').some(seg => seg === '' || seg === '.' || seg === '..')) throw new Error('잘못된 경로: ' + p);
  return path;
}
function rootFolder_() {
  const id = PROPS.getProperty('DATA_ROOT_ID');
  if (!id) throw new Error('초기 설정이 필요합니다 (시트 메뉴 → 보건관리 시스템 → 초기 설정)');
  return DriveApp.getFolderById(id);
}
function getFolder_(segments, create) {
  const cache = CacheService.getScriptCache();
  let folder = rootFolder_();
  let key = '';
  for (const seg of segments) {
    key += '/' + seg;
    const cached = cache.get('fid' + key);
    if (cached) {
      try { folder = DriveApp.getFolderById(cached); continue; } catch (_) { cache.remove('fid' + key); }
    }
    const it = folder.getFoldersByName(seg);
    if (it.hasNext()) folder = it.next();
    else if (create) folder = folder.createFolder(seg);
    else return null;
    cache.put('fid' + key, folder.getId(), 21600);
  }
  return folder;
}
function getFile_(path) {
  const segs = path.split('/');
  const folder = getFolder_(segs.slice(0, -1), false);
  if (!folder) return null;
  const it = folder.getFilesByName(segs[segs.length - 1]);
  let newest = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
  }
  return newest;
}
function fileMeta_(file) {
  let m = {};
  try { m = JSON.parse(file.getDescription() || '{}'); } catch (_) {}
  if (!m.sha) {   // 드라이브에 직접 올린 기존 파일: 최초 1회 계산 후 저장
    m.sha = sha1HexBytes_(file.getBlob().getBytes());
    try { file.setDescription(JSON.stringify(m)); } catch (_) {}
  }
  return m;
}
function writeFile_(path, bytes, mime, dept) {
  const segs = path.split('/');
  const folder = getFolder_(segs.slice(0, -1), true);
  const name = segs[segs.length - 1];
  const blob = Utilities.newBlob(bytes, mime, name);
  const sha = sha1HexBytes_(bytes);
  const old = getFile_(path);
  let file;
  if (old) {
    const prevMeta = fileMeta_(old);
    old.setTrashed(true);
    file = folder.createFile(blob);
    dept = prevMeta.dept || dept;
  } else {
    file = folder.createFile(blob);
  }
  file.setDescription(JSON.stringify({ sha: sha, dept: dept || '' }));
  return { sha: sha, size: bytes.length };
}
function mimeOf_(path) {
  const ext = (path.split('.').pop() || '').toLowerCase();
  return { pdf: 'application/pdf', json: 'application/json', png: 'image/png', jpg: 'image/jpeg',
           jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' }[ext] || 'application/octet-stream';
}

// 부서 토큰이 이 파일에 접근할 수 있는지 (파일 소유 부서 또는 파일명에 부서명)
function deptCanAccessFile_(s, path, meta) {
  if (!DEPT_FILE_PREFIXES.some(pre => path.indexOf(pre) === 0)) return false;
  if (meta && meta.dept) return meta.dept === s.dept;
  const name = path.split('/').pop();
  return name.indexOf('-' + s.dept + '-') >= 0 || name.indexOf('_' + s.dept + '_') >= 0;
}

function fsGet_(req, s) {
  const path = normPath_(req.path);
  const isAdmin = s.role === 'admin';

  // 폴더 목록
  const folder = getFolder_(path.split('/'), false);
  if (folder) {
    if (!isAdmin && !DEPT_FILE_PREFIXES.some(pre => (path + '/').indexOf(pre) === 0)) {
      return { success: false, status: 403, error: '권한 없음' };
    }
    const entries = [];
    const fit = folder.getFiles();
    while (fit.hasNext()) {
      const f = fit.next();
      let meta = {};
      try { meta = JSON.parse(f.getDescription() || '{}'); } catch (_) {}
      const p = path + '/' + f.getName();
      if (!isAdmin && !deptCanAccessFile_(s, p, meta)) continue;
      entries.push({ type: 'file', name: f.getName(), path: p, sha: meta.sha || '', size: f.getSize() });
    }
    const dit = folder.getFolders();
    while (dit.hasNext()) {
      const d = dit.next();
      entries.push({ type: 'dir', name: d.getName(), path: path + '/' + d.getName(), sha: '', size: 0 });
    }
    return { success: true, status: 200, type: 'dir', entries: entries };
  }

  // 기록 파일 (부서는 자기 부서 기록만)
  if (RECORD_FILES[path] || path === 'records/admin_cfg.json' || path === 'records/audit_log.json' || path === 'records/settings.json') {
    const file = getFile_(path);
    if (!file) return { success: false, status: 404, error: 'Not Found' };
    const text = file.getBlob().getDataAsString('UTF-8');
    if (isAdmin) {
      return jsonFileResult_(path, text, fileMeta_(file).sha);
    }
    const view = deptView_(path, JSON.parse(text), s);
    if (view === undefined) return { success: false, status: 403, error: '권한 없음' };
    const viewText = JSON.stringify(view);
    return jsonFileResult_(path, viewText, sha256Hex_(viewText).slice(0, 40));
  }

  // 일반 파일 (PDF·이미지)
  const file = getFile_(path);
  if (!file) return { success: false, status: 404, error: 'Not Found' };
  const meta = fileMeta_(file);
  if (!isAdmin && !deptCanAccessFile_(s, path, meta)) return { success: false, status: 403, error: '권한 없음' };
  const size = file.getSize();
  const out = { success: true, status: 200, type: 'file', path: path, name: file.getName(), sha: meta.sha,
                size: size, mime: file.getMimeType() || mimeOf_(path), encoding: 'base64' };
  const wantChunk = req.offset !== undefined && req.offset !== null;
  if (!wantChunk && size > INLINE_LIMIT) { out.content = ''; out.chunked = true; return out; }
  const b64 = Utilities.base64Encode(file.getBlob().getBytes());
  const offset = wantChunk ? Math.max(0, parseInt(req.offset, 10) || 0) : 0;
  out.content = wantChunk ? b64.substr(offset, CHUNK_CHARS) : b64;
  out.offset = offset;
  out.totalChars = b64.length;
  return out;
}

function jsonFileResult_(path, text, sha) {
  return { success: true, status: 200, type: 'file', path: path, name: path.split('/').pop(), sha: sha,
           size: text.length, mime: 'application/json', encoding: 'base64',
           content: Utilities.base64Encode(text, Utilities.Charset.UTF_8), offset: 0 };
}

// 부서 토큰에게 보여줄 기록 파일 내용 (undefined = 접근 불가)
function deptView_(path, data, s) {
  const mine = id => typeof id === 'string' && id.indexOf('|' + s.dept + '|') >= 0;
  if (RECORD_FILES[path]) {
    const list = recordList_(data, RECORD_FILES[path]).filter(r => r && r.dept === s.dept);
    return RECORD_FILES[path] === 'records' ? Object.assign({}, data, { records: list }) : list;
  }
  if (path === 'records/admin_cfg.json') {
    const out = {};
    Object.keys(data || {}).forEach(k => {
      out[k] = Array.isArray(data[k]) ? data[k].filter(mine) : data[k];
    });
    return out;
  }
  if (path === 'records/audit_log.json') return { entries: [] };
  if (path === 'records/settings.json') {
    const pub = {};
    ['apv1Name', 'stampName', 'updated'].forEach(k => { if (data && data[k]) pub[k] = data[k]; });
    return pub;
  }
  return undefined;
}

function recordList_(data, shape) {
  if (shape === 'records') return (data && Array.isArray(data.records)) ? data.records : [];
  return Array.isArray(data) ? data : [];
}
function recordKey_(r) {
  if (r.recordId) return 'id:' + r.recordId;
  return 'nk:' + [r.name, r.empNo, r.dept, r.type, String(r.dateStart || r.date || '').slice(0, 10)]
    .map(v => String(v || '').trim()).join('|');
}

function fsPut_(req, s) {
  const path = normPath_(req.path);
  const isAdmin = s.role === 'admin';

  // 분할 업로드: 조각을 임시 파일로 모은 뒤 마지막 조각에서 합침
  let b64 = String(req.content || '');
  if (req.uploadId) {
    const tmp = getFolder_(['_tmp'], true);
    const base = String(req.uploadId).replace(/[^A-Za-z0-9_-]/g, '');
    tmp.createFile(base + '_' + req.index, b64, MimeType.PLAIN_TEXT);
    if (Number(req.index) < Number(req.count) - 1) return { success: true, status: 202, received: req.index };
    const pieces = [];
    for (let i = 0; i < Number(req.count); i++) {
      const it = tmp.getFilesByName(base + '_' + i);
      if (!it.hasNext()) return { success: false, status: 400, error: '업로드 조각 누락: ' + i };
      const f = it.next();
      pieces.push(f.getBlob().getDataAsString());
      f.setTrashed(true);
    }
    b64 = pieces.join('');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, status: 503, error: '서버가 바쁩니다. 잠시 후 다시 시도하세요.' };
  try {
    const file = getFile_(path);

    // 기록 파일
    if (RECORD_FILES[path] || path === 'records/admin_cfg.json' || path === 'records/audit_log.json' || path === 'records/settings.json') {
      const incoming = JSON.parse(Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('UTF-8'));
      const curText = file ? file.getBlob().getDataAsString('UTF-8') : '';
      const cur = curText ? JSON.parse(curText) : null;

      if (isAdmin) {
        const curSha = file ? fileMeta_(file).sha : null;
        if (file && !req.sha) return { success: false, status: 422, error: 'sha 필요 (파일이 이미 있음)' };
        if (file && req.sha !== curSha) return { success: false, status: 409, error: 'is at ' + curSha + ' but expected ' + req.sha };
        const text = JSON.stringify(incoming);
        const w = writeFile_(path, Utilities.newBlob(text).getBytes(), 'application/json', '');
        return { success: true, status: file ? 200 : 201, sha: w.sha, path: path };
      }

      // 부서: 자기 부서 기록만 추가·수정 (삭제·다른 부서 변경 불가)
      const view = deptView_(path, cur || (RECORD_FILES[path] === 'records' ? { records: [] } : (RECORD_FILES[path] ? [] : {})), s);
      if (view === undefined || path === 'records/settings.json') return { success: false, status: 403, error: '권한 없음' };
      const viewSha = sha256Hex_(JSON.stringify(view)).slice(0, 40);
      if (cur && req.sha && req.sha !== viewSha) return { success: false, status: 409, error: 'is at ' + viewSha + ' but expected ' + req.sha };
      const merged = mergeDeptWrite_(path, cur, incoming, s);
      const text = JSON.stringify(merged);
      writeFile_(path, Utilities.newBlob(text).getBytes(), 'application/json', '');
      const newView = JSON.stringify(deptView_(path, merged, s));
      return { success: true, status: 200, sha: sha256Hex_(newView).slice(0, 40), path: path };
    }

    // 일반 파일
    if (!isAdmin) {
      const meta = file ? fileMeta_(file) : null;
      if (!DEPT_FILE_PREFIXES.some(pre => path.indexOf(pre) === 0)) return { success: false, status: 403, error: '권한 없음' };
      if (file && !deptCanAccessFile_(s, path, meta)) return { success: false, status: 403, error: '다른 부서 파일은 덮어쓸 수 없습니다' };
    }
    if (file) {
      const curSha = fileMeta_(file).sha;
      if (!req.sha) return { success: false, status: 422, error: 'sha 필요 (파일이 이미 있음)' };
      if (req.sha !== curSha) return { success: false, status: 409, error: 'is at ' + curSha + ' but expected ' + req.sha };
    }
    const bytes = Utilities.base64Decode(b64);
    const w = writeFile_(path, bytes, mimeOf_(path), isAdmin ? '' : s.dept);
    return { success: true, status: file ? 200 : 201, sha: w.sha, path: path, size: w.size };
  } finally {
    lock.releaseLock();
  }
}

// 부서 쓰기 병합: 서버의 다른 부서 기록은 그대로 두고, 자기 부서 기록만 추가·갱신
function mergeDeptWrite_(path, cur, incoming, s) {
  if (RECORD_FILES[path]) {
    const shape = RECORD_FILES[path];
    const curList = recordList_(cur, shape).slice();
    const index = {};
    curList.forEach((r, i) => { if (r) index[recordKey_(r)] = i; });
    recordList_(incoming, shape).forEach(r => {
      if (!r || r.dept !== s.dept) return;
      const k = recordKey_(r);
      if (index[k] !== undefined) {
        if (curList[index[k]].dept === s.dept) curList[index[k]] = r;
      } else {
        index[k] = curList.length;
        curList.push(r);
      }
    });
    return shape === 'records' ? Object.assign({}, cur || {}, { records: curList }) : curList;
  }
  if (path === 'records/admin_cfg.json') {
    // 부서는 자기 부서 삭제표시(묘비)를 "지우는 것"만 가능 (같은 기록 재제출 시)
    const out = Object.assign({}, cur || {});
    Object.keys(out).forEach(k => {
      if (!Array.isArray(out[k])) return;
      const keep = new Set(Array.isArray(incoming[k]) ? incoming[k] : []);
      out[k] = out[k].filter(id => !(typeof id === 'string' && id.indexOf('|' + s.dept + '|') >= 0) || keep.has(id));
    });
    return out;
  }
  if (path === 'records/audit_log.json') {
    const curEntries = (cur && Array.isArray(cur.entries)) ? cur.entries.slice() : [];
    const seen = new Set(curEntries.map(x => JSON.stringify(x)));
    ((incoming && incoming.entries) || []).forEach(x => {
      const k = JSON.stringify(x);
      if (!seen.has(k)) { seen.add(k); curEntries.push(Object.assign({}, x, { dept: s.dept })); }
    });
    return Object.assign({}, cur || {}, { entries: curEntries });
  }
  return cur;
}

// ════════════════════════════════════════════════════════════════
// 공통 유틸
// ════════════════════════════════════════════════════════════════
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v === undefined || v === null ? '' : v).trim();
}
// 시트 셀 값 표현: 날짜는 KST 기준 (기존 toISOString은 UTC라 하루 당겨짐)
function cellStr_(v) {
  if (v instanceof Date) {
    const hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, TZ, hasTime ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
  }
  return String(v === undefined || v === null ? '' : v).trim();
}
// 수식 주입 방지: =,+,-,@ 로 시작하는 입력은 문자로 저장
function safeCell_(v) {
  const s = String(v === undefined || v === null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// ════════════════════════════════════════════════════════════════
// 직원목록 조회 (부서 토큰은 자기 부서 직원만)
// 시트명: 직원목록 / 컬럼: 이름(또는 성명), 부서, 사번, 입사일, 성별
// ════════════════════════════════════════════════════════════════
function findEmpSheet_() {
  return SS.getSheetByName('직원목록') || SS.getSheetByName('직원') || SS.getSheetByName('직원정보')
      || SS.getSheetByName('employees') || SS.getSheetByName('Employees');
}

function readEmployees_() {
  const sheet = findEmpSheet_();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  const findCol = keys => { for (const k of keys) if (col[k] !== undefined) return col[k]; return -1; };
  const nameIdx     = findCol(['이름','성명','직원이름','직원명','성함','name','성명(이름)','이름(성명)']);
  const deptIdx     = findCol(['부서','부서명','소속','소속부서','팀','dept','department','부서(팀)','팀(부서)']);
  const empNoIdx    = findCol(['사번','직원번호','사원번호','사원코드','직원코드','empNo','employeeNo','사번(직원번호)']);
  const joinDateIdx = findCol(['입사일','입사일자','입사년월일','joinDate','join','입사일(joinDate)','입사년월']);
  const genderIdx   = findCol(['성별','gender','남녀','성(남/여)']);
  const getCell = (row, idx, fb) => {
    const i = idx >= 0 ? idx : fb;
    return (i < 0 || i >= row.length) ? '' : fmtDate_(row[i]);
  };
  return data.slice(1).map(row => ({
    name:     getCell(row, nameIdx, 1),
    dept:     getCell(row, deptIdx, 2),
    empNo:    getCell(row, empNoIdx, 0),
    joinDate: getCell(row, joinDateIdx, 3),
    gender:   getCell(row, genderIdx, 4),
  })).filter(e => e.name);
}

function getEmployees(s) {
  const list = readEmployees_();
  if (list === null) return { success: false, error: '직원목록 시트가 없습니다. 시트명을 "직원목록"으로 해주세요.' };
  const data = s && s.role === 'dept' ? list.filter(e => e.dept === s.dept) : list;
  return { success: true, data: data };
}

function getEmployeeDepts_() {
  const list = readEmployees_() || [];
  return [...new Set(list.map(e => e.dept).filter(Boolean))].sort();
}

// ════════════════════════════════════════════════════════════════
// MSDS 화학물질 목록 조회 (공개 — 개인정보 없음)
// ════════════════════════════════════════════════════════════════
function getChemicals() {
  const sheet = SS.getSheetByName('chemicals') || SS.getSheetByName('Chemicals') || SS.getSheetByName('화학물질')
             || SS.getSheetByName('MSDS목록') || SS.getSheetByName('MSDS');
  if (!sheet) return { success: false, error: 'chemicals 시트가 없습니다' };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  const headers = data[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  const get = (row, ...keys) => {
    for (const k of keys) {
      if (col[k] !== undefined) {
        const v = row[col[k]];
        if (v === undefined || v === null) continue;
        const sv = fmtDate_(v);
        if (sv !== '') return sv;
      }
    }
    return '';
  };
  const getBool = (row, ...keys) => {
    const v = get(row, ...keys).toUpperCase();
    return ['TRUE','YES','예','Y','1','O','○','V','해당'].indexOf(v) >= 0;
  };
  const fixUrl = u => (u || '').replace(/https\/\//g, 'https://').replace(/http\/\//g, 'http://');

  const chems = [];
  let lastChem = null;
  data.slice(1).forEach(row => {
    const rowName   = get(row, 'name', '물질명', '제품명');
    const compName  = get(row, 'compName', '구성성분명', '성분명', '성분');
    const casNumber = get(row, 'casNumber', 'CAS', 'CAS번호', 'cas');
    const amount    = get(row, 'amount', '함유량', '함유량(%)', '농도');
    if (!rowName) {
      if (lastChem && compName) {
        lastChem.components.push({ compName, casNumber, amount,
          workEnv: get(row, 'workEnv', '작업환경측정', '작업환경'),
          healthCheck: get(row, 'healthCheck', '건강검진', '특수건강검진') });
      }
      return;
    }
    const deptStr = get(row, 'department', 'departments', '부서', '사용부서');
    const firstWorkEnv     = get(row, 'workEnv', '작업환경측정', '작업환경');
    const firstHealthCheck = get(row, 'healthCheck', '건강검진', '특수건강검진');
    const skinEyeVal = get(row, '[피부,눈]', '[피부]', '[눈]', 'skin_eye', 'skin', '피부/눈', '피부,눈', 'emergency');
    const inhalVal   = get(row, '[흡입]', '흡입', 'inhalation');
    const ingVal     = get(row, '[먹었을때]', '[섭취]', '섭취', 'ingestion', '[삼켰을때]');
    const emergencyRaw = get(row, 'emergency', '응급처치', '비상조치');
    const iconUrlStr = fixUrl(get(row, 'iconUrl', 'icon', '아이콘', 'icons', '위험그림문자'));
    const chemObj = {
      name:          get(row, 'name', '물질명', '제품명', '화학물질명'),
      manufacturer:  get(row, 'manufacturer', '제조사', '제조자', '공급자'),
      department:    deptStr,
      departments:   deptStr.split(/[,\/]/).map(d => d.trim()).filter(Boolean),
      usage:         get(row, 'usage', '용도', '사용용도'),
      targetManage:  getBool(row, 'targetManage', '관리대상', '관리대상물질'),
      specialManage: getBool(row, 'specialManage', '특별관리', '특별관리물질'),
      workEnv:       firstWorkEnv,
      healthCheck:   firstHealthCheck,
      emergency:     emergencyRaw || [skinEyeVal && ('[피부,눈] ' + skinEyeVal), inhalVal && ('[흡입] ' + inhalVal),
                                      ingVal && ('[먹었을때] ' + ingVal)].filter(Boolean).join('\n'),
      emergencyParsed: { skin_eye: skinEyeVal || '', inhalation: inhalVal || '', ingestion: ingVal || '' },
      pdfUrl:        fixUrl(get(row, 'pdfUrl', 'PDF', 'PDF링크', 'MSDS링크', 'pdf')),
      iconUrl:       iconUrlStr,
      icons:         iconUrlStr ? iconUrlStr.split(',').map(part => {
                       const pts = part.trim().split('|');
                       return { url: fixUrl(pts[0] ? pts[0].trim() : ''), label: pts[1] ? pts[1].trim() : '' };
                     }).filter(ic => ic.url) : [],
      components:    compName ? [{ compName, casNumber, amount, workEnv: firstWorkEnv, healthCheck: firstHealthCheck }] : [],
    };
    if (chemObj.name) { chems.push(chemObj); lastChem = chemObj; }
  });
  return { success: true, data: chems };
}

// ════════════════════════════════════════════════════════════════
// 교육 템플릿 조회 (공개)
// ════════════════════════════════════════════════════════════════
function getEduTemplate() {
  const result = { success: true, specialEdu: {}, msdsEdu: {} };
  const read = (sheet, target) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    sheet.getDataRange().getValues().slice(1).forEach(row => {
      const dept = String(row[0] || '').trim();
      if (!dept) return;
      target[dept] = { time: String(row[1] || ''), place: String(row[2] || ''), method: String(row[3] || ''),
        instructor: String(row[4] || ''), instrPos: String(row[5] || ''), individual: String(row[6] || ''),
        common: String(row[7] || '') };
    });
  };
  read(SS.getSheetByName('특별교육대장') || SS.getSheetByName('특별안전보건교육') || SS.getSheetByName('특별교육'), result.specialEdu);
  read(SS.getSheetByName('MSDS교육대장') || SS.getSheetByName('msds교육대장') || SS.getSheetByName('MSDS교육'), result.msdsEdu);
  return result;
}

// ════════════════════════════════════════════════════════════════
// 기록 저장 (교육 + 보호구) — 부서 토큰은 자기 부서로만 저장
// ════════════════════════════════════════════════════════════════
function saveRecord(p, s) {
  if (s.role === 'dept' && String(p.dept || '').trim() !== s.dept) {
    return { success: false, status: 403, error: '다른 부서 기록은 저장할 수 없습니다' };
  }
  const type = p.type || '교육';
  const isPpe = type.indexOf('보호구') >= 0;
  const sheetName = isPpe ? '보호구이력대장' : '교육이력대장';
  const EDU_HEADER = ['순번','제출일시','교육구분','교육부서','취급물질','교육시작일','교육종료일','교육일수','총교육시간','사번','성명','입사일','서명여부','제출계정'];
  const PPE_HEADER = ['순번','제출일시','부서','이름','지급물품','지급일','지급자','서명여부','제출계정'];
  const expectedHeader = isPpe ? PPE_HEADER : EDU_HEADER;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '서버가 바쁩니다. 잠시 후 다시 시도하세요.' };
  try {
    let sheet = SS.getSheetByName(sheetName);
    if (!sheet) sheet = SS.insertSheet(sheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(expectedHeader);
    } else {
      const cur = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeader.length)).getValues()[0].map(h => String(h).trim());
      // v1 헤더(제출계정 열 없음)는 마지막 열만 추가, 그 외 불일치만 새 헤더 삽입
      const v1Match = expectedHeader.slice(0, -1).every((h, i) => h === cur[i]);
      if (v1Match && cur[expectedHeader.length - 1] !== '제출계정') {
        sheet.getRange(1, expectedHeader.length).setValue('제출계정');
      } else if (!expectedHeader.every((h, i) => h === cur[i])) {
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, expectedHeader.length).setValues([expectedHeader]);
      }
    }

    const now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
    const account = s.role === 'admin' ? '관리자' : s.dept;
    const names     = p.names     ? p.names.split('|')     : [];
    const empNos    = p.empNos    ? p.empNos.split('|')    : [];
    const joinDates = p.joinDates ? p.joinDates.split('|') : [];
    const sigCount  = parseInt(p.sigCount || '0', 10);

    let totalHours = p.totalHours || '';
    if (!totalHours && p.time) {
      const m1 = p.time.match(/(\d+)\s*시간\s*\*\s*(\d+)/);
      const m2 = p.time.match(/(\d+)\s*시간/);
      totalHours = m1 ? (parseInt(m1[1], 10) * parseInt(m1[2], 10)) + '시간' : (m2 ? parseInt(m2[1], 10) + '시간' : p.time);
    }

    let nextNo = sheet.getLastRow();   // 헤더 1행 → 다음 순번
    const rows = [];
    if (isPpe) {
      rows.push([nextNo, now, p.dept || '', p.names || '', p.chem || '', p.dateStart || '', '안전관리자',
                 sigCount > 0 ? '✓ 서명완료' : '- 미서명', account]);
    } else {
      names.forEach((name, i) => {
        rows.push([nextNo++, now, p.type || '', p.dept || '', (p.chem || '').slice(0, 100), p.dateStart || '',
          p.dateEnd || p.dateStart || '', p.days || '1', totalHours, empNos[i] || '', name.trim(),
          joinDates[i] || '', i < sigCount ? '✓ 서명완료' : '- 미서명', account]);
      });
    }
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
        .setValues(rows.map(r => r.map((v, i) => i === 0 ? v : safeCell_(v))));
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true, message: '저장완료' };
}

// ════════════════════════════════════════════════════════════════
// 기록 조회 (교육 + 보호구) — 부서 토큰은 자기 부서만
// ════════════════════════════════════════════════════════════════
function getRecords(p, s) {
  const results = [];
  const readSheet = (sheet, rowFn) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    const col = {};
    data[0].forEach((h, i) => col[String(h).trim()] = i);
    const get = (row, ...keys) => {
      for (const k of keys) {
        if (col[k] !== undefined) {
          const v = cellStr_(row[col[k]]);
          if (v !== '') return v;
        }
      }
      return '';
    };
    data.slice(1).forEach(row => { if (row[0] !== '' && row[0] !== null) rowFn(row, get); });
  };

  ['교육이력대장','특별교육대장','MSDS교육대장','특별안전보건교육','MSDS교육'].forEach(name => {
    readSheet(SS.getSheetByName(name), (row, get) => {
      const type = get(row, '교육구분') || (name.indexOf('MSDS') >= 0 ? 'MSDS교육' : '특별안전보건교육');
      if (type.indexOf('보호구') >= 0) return;
      results.push({
        submitted: get(row, '제출일시'), type: type, dept: get(row, '교육부서', '부서'), chem: get(row, '취급물질', 'chem'),
        name: get(row, '성명', '이름'), empNo: get(row, '사번'), joinDate: get(row, '입사일'),
        dateStart: get(row, '교육시작일', '교육일'), dateEnd: get(row, '교육종료일', '교육일'),
        days: get(row, '교육일수'), totalHours: get(row, '총교육시간', '교육시간'), signed: get(row, '서명여부', '서명'),
      });
    });
  });

  readSheet(SS.getSheetByName('보호구이력대장') || SS.getSheetByName('보호구지급대장') || SS.getSheetByName('보호구지급'),
    (row, get) => {
      results.push({
        submitted: get(row, '제출일시'), type: '보호구지급', dept: get(row, '부서'), name: get(row, '이름', '성명'),
        items: get(row, '지급물품'), date: get(row, '지급일'), giver: get(row, '지급자'),
        signed: get(row, '서명') || '✓ 서명완료',
      });
    });

  const data = s && s.role === 'dept' ? results.filter(r => r.dept === s.dept) : results;
  return { success: true, data: data };
}

// ════════════════════════════════════════════════════════════════
// 시트 기록 삭제 (관리자) — 앱 영구삭제와 시트를 일치시킴
// ids: recordId 목록 "제출시각|이름|사번|부서|구분|교육일"
// ════════════════════════════════════════════════════════════════
function deleteSheetRecords(p) {
  let ids = [];
  try { ids = JSON.parse(p.ids || '[]'); } catch (_) {}
  const category = p.category || 'edu';
  const targets = ids.map(id => String(id).split('|')).filter(a => a.length >= 6)
    .map(a => ({ name: a[1].replace(/\s+/g, ''), dept: a[3].replace(/\s+/g, ''), type: a[4], date: a[5].slice(0, 10) }));
  if (!targets.length) return { success: true, deleted: 0 };

  const sheetName = category === 'ppe' ? '보호구이력대장' : '교육이력대장';
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return { success: true, deleted: 0 };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '서버가 바쁩니다' };
  let deleted = 0;
  try {
    const data = sheet.getDataRange().getValues();
    const col = {};
    data[0].forEach((h, i) => col[String(h).trim()] = i);
    const val = (row, k) => col[k] === undefined ? '' : cellStr_(row[col[k]]);
    for (let r = data.length - 1; r >= 1; r--) {
      const row = data[r];
      const name = (category === 'ppe' ? val(row, '이름') : val(row, '성명')).replace(/\s+/g, '');
      const dept = (category === 'ppe' ? val(row, '부서') : val(row, '교육부서')).replace(/\s+/g, '');
      const date = (category === 'ppe' ? val(row, '지급일') : val(row, '교육시작일')).slice(0, 10);
      const type = category === 'ppe' ? '보호구지급' : val(row, '교육구분').replace(/\s+/g, '');
      const hit = targets.some(t => t.dept === dept && t.date === date &&
        (category === 'ppe' ? name.split('|').indexOf(t.name) >= 0 || name === t.name : name === t.name) &&
        (category === 'ppe' || type.indexOf(t.type) === 0 || t.type.indexOf(type) === 0));
      if (hit) { sheet.deleteRow(r + 1); deleted++; }
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true, deleted: deleted };
}

// ════════════════════════════════════════════════════════════════
// 보호구 품목 조회 (공개)
// ════════════════════════════════════════════════════════════════
function getPpeItems() {
  const sheet = SS.getSheetByName('보호구품목') || SS.getSheetByName('보호구') || SS.getSheetByName('PPE');
  if (!sheet) return { success: true, data: [] };
  const items = sheet.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), unit: String(r[1] || '개').trim() }));
  return { success: true, data: items };
}

// ════════════════════════════════════════════════════════════════
// 통합 조회: adminCfg + records
// ════════════════════════════════════════════════════════════════
function getAllRecords(p, s) {
  const cfg = getAdminCfg();
  const recs = getRecords(p, s);
  return { success: true, cfgData: (cfg.success && cfg.data) ? cfg.data : {}, data: (recs.success && recs.data) ? recs.data : [] };
}

// ════════════════════════════════════════════════════════════════
// 관리자 설정 (초기화 시각 등) — 읽기: 로그인 / 쓰기: 관리자
// ════════════════════════════════════════════════════════════════
function getAdminCfg() {
  const sheet = SS.getSheetByName('관리자설정');
  if (!sheet) return { success: true, data: {} };
  const cfg = {};
  sheet.getDataRange().getValues().forEach(row => {
    const key = String(row[0] || '').trim();
    if (key && key !== '키') cfg[key] = String(row[1] || '').trim();
  });
  return { success: true, data: cfg };
}

function setAdminCfg(p) {
  const key = String(p.key || '').trim();
  const value = String(p.value || '').trim();
  if (!key) return { success: false, error: 'key 없음' };
  let sheet = SS.getSheetByName('관리자설정');
  if (!sheet) { sheet = SS.insertSheet('관리자설정'); sheet.appendRow(['키', '값', '수정일시']); }
  const now = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[safeCell_(value), now]]);
      return { success: true };
    }
  }
  sheet.appendRow([safeCell_(key), safeCell_(value), now]);
  return { success: true };
}
