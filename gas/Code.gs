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
const SERVER_VERSION = 'v20.1 (2026-10-21)';   // 관리자 화면에 표시 — 배포된 서버가 최신인지 확인용
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
    .addSeparator()
    .addItem('④ 인사 사이트(Firebase) 연결·자동 로그인 설정', 'saveHrCredentials')
    .addItem('④-1 인사 사이트 자동 로그인 정보 삭제', 'clearHrCredentials')
    .addItem('⑤ 매일 아침 자동 확인 켜기', 'enableDailyCheck')
    .addItem('⑤-1 매일 아침 자동 확인 끄기', 'disableDailyCheck')
    .addSeparator()
    .addItem('⑥ AI(Gemini) 키 저장', 'saveGeminiKey')
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
  deptMapSheet_();
  if (!PROPS.getProperty('ADMIN_PW_HASH')) changeAdminPassword();
  // 신규입사 알림 기본값 (관리자설정 시트에서 수정 가능)
  const cfg = getAdminCfg().data || {};
  if (!cfg['앱주소']) setAdminCfg({ key: '앱주소', value: 'https://buminhealth.github.io/health/' + encodeURI('부민병원_MSDS_통합시스템_v5 - 2026-04-17T095042.287.html') });
  if (!cfg['보건관리자이메일']) setAdminCfg({ key: '보건관리자이메일', value: 'clover8477@bumin.co.kr' });
  if (!cfg['신규입사기준일수']) setAdminCfg({ key: '신규입사기준일수', value: '7' });
  if (!cfg['재알림간격일']) setAdminCfg({ key: '재알림간격일', value: '3' });
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

// 권한 다시 승인: Apps Script 편집기 위쪽에서 이 함수(authorizeAll)를 고르고 ▶ 실행
// (드라이브·메일·외부 사이트(Firebase) 접속·자동 실행 권한을 한 번에 요청)
function authorizeAll() {
  SS.getName();
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  ScriptApp.getProjectTriggers();
  Logger.log('권한 승인 완료 — 시트 메뉴 ④를 다시 실행하세요');
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
  DEPT_ALIAS_CACHE = null;
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
  DEPT_ALIAS_CACHE = null;
  let result;
  try {
    const req = JSON.parse((e.postData && e.postData.contents) || '{}');
    const s = verifyToken_(req.t);
    switch (req.op) {
      case 'depts':     result = { success: true, data: listActiveDepts_() }; break;
      case 'login':     result = login_(req); break;
      case 'whoami':    result = s ? { success: true, role: s.role, dept: s.dept || '', exp: s.exp, serverVersion: SERVER_VERSION }
                                   : { success: false, status: 401, error: '로그인 필요' }; break;
      case 'fs.get':    result = requireLogin_(s) || fsGet_(req, s); break;
      case 'fs.put':    result = requireLogin_(s) || fsPut_(req, s); break;
      case 'rec.has':   result = requireLogin_(s) || recHas_(req, s); break;
      case 'emp.status': case 'emp.sync': case 'hr.fetch':
      case 'nh.list': case 'nh.answer': case 'nh.resend': case 'nh.run':
                        result = staffOp_(req, s); break;
      case 'chem.audit': case 'chem.seedLaw': case 'chem.apply': case 'chem.lawImport':
                        result = chemOp_(req, s); break;
      case 'ai.status': case 'ai.setKey': case 'ai.list': case 'ai.analyze': case 'ai.judge': case 'ai.apply': case 'ai.products': case 'ai.compare':
                        result = aiOp_(req, s); break;
      case 'edu.dupes': case 'edu.dedupe': case 'edu.status': case 'edu.exclude': case 'edu.askMail': case 'edu.newChem':
                        result = eduOp_(req, s); break;
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
  return { success: true, token: signToken_(payload), role: payload.role, dept: payload.dept || '', exp: payload.exp,
           head: payload.role === 'dept' ? deptHead_(payload.dept) : '' };
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

// ════════════════════════════════════════════════════════════════
// 부서 이름 맞추기
//  · 띄어쓰기·전각 괄호 차이는 같은 부서로 봄 (예: "61 병동" = "61병동")
//  · 인사 사이트 이름이 앱 이름과 다르면 '부서매핑' 시트에 한 번 적어 두면 동기화 때 앱 이름으로 저장
//  · chemicals 시트의 '병동'은 앱 화면과 같은 세부 병동 목록에 적용
// ════════════════════════════════════════════════════════════════
const DEPT_MAP_SHEET = '부서매핑';
const WARDS = ['51병동','61병동','71병동','6병동','7병동','8병동','9병동','응급실','중환자실','인공신장실'];   // HTML의 WARD_EXPAND와 같게 유지

function normDept_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')');
}
function sameDept_(a, b) { return normDept_(a) !== '' && normDept_(a) === normDept_(b); }

function deptMapSheet_() {
  let sh = SS.getSheetByName(DEPT_MAP_SHEET);
  if (!sh) {
    sh = SS.insertSheet(DEPT_MAP_SHEET);
    sh.appendRow(['인사 사이트 부서명', '앱 부서명', '비고']);
    sh.setFrozenRows(1);
  }
  return sh;
}
// { 정규화된 인사 부서명: 앱 부서명('' = 그대로 사용) }
function readDeptMap_() {
  const sh = SS.getSheetByName(DEPT_MAP_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
    const raw = normDept_(r[0]);
    if (raw) map[raw] = String(r[1] || '').trim();
  });
  return map;
}
// v20.1: 직원목록에 남아 있는 인사 사이트 부서명도 부서매핑대로 앱 부서명으로 봄
//  (동기화 때만 바꾸면 매핑을 나중에 적은 경우 기존 직원은 그대로라 교육 대상 · 부서 로그인 명단에서 빠짐)
let DEPT_ALIAS_CACHE = null;   // 요청마다 doGet/doPost 시작 때 비움
function deptAlias_(d) {
  if (!DEPT_ALIAS_CACHE) DEPT_ALIAS_CACHE = readDeptMap_();
  const v = DEPT_ALIAS_CACHE[normDept_(d)];
  return v ? v : d;
}
// 앱에서 쓰는 부서 이름들 (부서인증 + chemicals + 세부 병동)
function knownAppDepts_() {
  const set = {};
  readAuthRows_().forEach(r => { set[normDept_(r.dept)] = r.dept; });
  (getChemicals().data || []).forEach(c => (c.departments || []).forEach(d => { if (d !== '병동') set[normDept_(d)] = d; }));
  WARDS.forEach(d => { set[normDept_(d)] = d; });
  return set;
}

function readAuthRows_() {
  const sheet = SS.getSheetByName(AUTH_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues()
    .map(r => ({ dept: String(r[0]).trim(), pin: String(r[1]).trim(), active: /^(Y|예|사용|O)$/i.test(String(r[3]).trim()) }))
    .filter(r => r.dept);
}
function findDeptAuth_(dept) {
  return readAuthRows_().find(r => r.active && sameDept_(r.dept, dept)) || null;
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
// 같은 이름 폴더가 여러 개여도(드라이브에 폴더를 두 번 올린 경우) 모두 찾음
function getFolders_(segments) {
  let level = [rootFolder_()];
  for (const seg of segments) {
    const next = [];
    level.forEach(f => { const it = f.getFoldersByName(seg); while (it.hasNext()) next.push(it.next()); });
    if (!next.length) return [];
    level = next;
  }
  return level;
}
function getFile_(path) {
  const segs = path.split('/');
  let newest = null;
  getFolders_(segs.slice(0, -1)).forEach(folder => {
    const it = folder.getFilesByName(segs[segs.length - 1]);
    while (it.hasNext()) {
      const f = it.next();
      if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
    }
  });
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
  const folders = getFolders_(path.split('/'));
  if (folders.length) {
    if (!isAdmin && !DEPT_FILE_PREFIXES.some(pre => (path + '/').indexOf(pre) === 0)) {
      return { success: false, status: 403, error: '권한 없음' };
    }
    const entries = [];
    const seenNames = {};
    folders.forEach(folder => {
      const fit = folder.getFiles();
      while (fit.hasNext()) {
        const f = fit.next();
        if (seenNames['f:' + f.getName()]) continue;
        let meta = {};
        try { meta = JSON.parse(f.getDescription() || '{}'); } catch (_) {}
        const p = path + '/' + f.getName();
        if (!isAdmin && !deptCanAccessFile_(s, p, meta)) continue;
        seenNames['f:' + f.getName()] = true;
        entries.push({ type: 'file', name: f.getName(), path: p, sha: meta.sha || '', size: f.getSize() });
      }
      const dit = folder.getFolders();
      while (dit.hasNext()) {
        const d = dit.next();
        if (seenNames['d:' + d.getName()]) continue;
        seenNames['d:' + d.getName()] = true;
        entries.push({ type: 'dir', name: d.getName(), path: path + '/' + d.getName(), sha: '', size: 0 });
      }
    });
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
    const list = recordList_(data, RECORD_FILES[path]).filter(r => r && sameDept_(r.dept, s.dept));
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
// 서명·사진 필드: 브라우저는 용량 때문에 이 값을 빼고 저장할 수 있음 → 서버는 원본을 절대 잃지 않게 보존
const IMG_FIELDS = ['sigImg', 'sigsByDay', 'photos', 'newSig', 'stopSig'];
function hasImg_(r) {
  return !!r && IMG_FIELDS.some(f => { const v = r[f]; return Array.isArray(v) ? v.some(Boolean) : (v && typeof v === 'object' ? Object.keys(v).length > 0 : !!v); });
}
function isEmptyImg_(v) {
  return v === undefined || v === null || v === '' || (Array.isArray(v) && !v.some(Boolean)) || (typeof v === 'object' && !Array.isArray(v) && !Object.keys(v).length);
}
// 들어온 기록에 서명·사진이 비어 있으면 서버에 있던 값을 그대로 유지
function preserveImages_(curList, incomingList) {
  const byKey = {};
  (curList || []).forEach(r => { if (r && hasImg_(r)) byKey[recordKey_(r)] = r; });
  (incomingList || []).forEach(r => {
    if (!r) return;
    const old = byKey[recordKey_(r)];
    if (old) IMG_FIELDS.forEach(f => { if (isEmptyImg_(r[f]) && !isEmptyImg_(old[f])) r[f] = old[f]; });
    delete r._img;
  });
}
// 브라우저가 보낸 기록 키 중 서버에 서명·사진이 있는 것만 돌려줌 (그 기록만 브라우저에서 이미지를 비워도 안전)
function recHas_(req, s) {
  const path = normPath_(req.path);
  if (!RECORD_FILES[path]) return { success: false, error: '기록 파일이 아닙니다' };
  const file = getFile_(path);
  if (!file) return { success: true, keys: [] };
  const list = recordList_(JSON.parse(file.getBlob().getDataAsString('UTF-8')), RECORD_FILES[path]);
  const have = new Set(list.filter(r => r && hasImg_(r) && (s.role === 'admin' || sameDept_(r.dept, s.dept))).map(recordKey_));
  return { success: true, keys: (req.keys || []).filter(k => have.has(k)) };
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
        if (RECORD_FILES[path] && cur) preserveImages_(recordList_(cur, RECORD_FILES[path]), recordList_(incoming, RECORD_FILES[path]));
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

// 결재 · 반려 칸은 보건관리자만 바꿈 — 부서가 올린 기록에서는 서버 값 유지 (새 기록은 비움)
const ADMIN_ONLY_FIELDS = ['approved', 'approvedAt', 'approverName', 'rejected', 'rejectReason', 'rejectedAt', 'rejectedBy'];
function keepAdminFields_(next, prev) {
  ADMIN_ONLY_FIELDS.forEach(f => {
    if (prev && prev[f] !== undefined) next[f] = prev[f];
    else delete next[f];
  });
  return next;
}

// 부서 쓰기 병합: 서버의 다른 부서 기록은 그대로 두고, 자기 부서 기록만 추가·갱신
function mergeDeptWrite_(path, cur, incoming, s) {
  if (RECORD_FILES[path]) {
    const shape = RECORD_FILES[path];
    const curList = recordList_(cur, shape).slice();
    preserveImages_(curList, recordList_(incoming, shape));
    const index = {};
    curList.forEach((r, i) => { if (r) index[recordKey_(r)] = i; });
    recordList_(incoming, shape).forEach(r => {
      if (!r || !sameDept_(r.dept, s.dept)) return;
      const k = recordKey_(r);
      if (index[k] !== undefined) {
        if (sameDept_(curList[index[k]].dept, s.dept)) curList[index[k]] = keepAdminFields_(r, curList[index[k]]);
      } else {
        index[k] = curList.length;
        curList.push(keepAdminFields_(r, null));
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

// 직원목록 열 이름 (첫 번째가 새로 만들 때 쓰는 이름, 나머지는 기존 시트·엑셀 붙여넣기 호환 이름)
const EMP_COLS = {
  empNo:     ['사번','직원번호','사원번호','사원코드','직원코드','empNo','employeeNo','사번(직원번호)','empId','employeeId','staffNo','sabun','empCode','emp_code','employeeCode','_key'],   // _key: Firestore 문서 ID (휴직자관리 employees_public은 문서 ID가 사번)
  name:      ['이름','성명','직원이름','직원명','성함','name','성명(이름)','이름(성명)','empName','userName','fullName'],
  dept:      ['부서','부서명','소속','소속부서','팀','dept','department','부서(팀)','팀(부서)','deptName','team'],
  joinDate:  ['입사일','입사일자','입사년월일','joinDate','join','입사일(joinDate)','입사년월','hireDate','startDate','entryDate'],
  status:    ['재직상태','상태','재직구분','재직여부','status','state','workStatus'],
  leaveDate: ['퇴사일','퇴직일','퇴사일자','leaveDate','resignDate','retireDate','quitDate'],
  email:     ['이메일','메일','email','E-mail','e-mail','이메일주소','mail'],
  position:  ['직위','직급','직책','position','rank','title','jobRank','직무직급'],
  firstSeen: ['최초확인일'],
  lastSeen:  ['최종확인일'],
};
const EMP_FALLBACK = { empNo: 0, name: 1, dept: 2, joinDate: 3 };   // v1 위치 기반 폴백

function empColIndex_(headers) {
  const col = {};
  headers.forEach((h, i) => { const k = String(h).trim(); if (k && col[k] === undefined) col[k] = i; });
  const idx = {};
  Object.keys(EMP_COLS).forEach(f => {
    idx[f] = -1;
    for (const k of EMP_COLS[f]) if (col[k] !== undefined) { idx[f] = col[k]; break; }
  });
  return idx;
}

function readEmployees_() {
  const sheet = findEmpSheet_();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const idx = empColIndex_(data[0]);
  const getCell = (row, f) => {
    const i = idx[f] >= 0 ? idx[f] : (EMP_FALLBACK[f] !== undefined ? EMP_FALLBACK[f] : -1);
    return (i < 0 || i >= row.length) ? '' : fmtDate_(row[i]);
  };
  return data.slice(1).map(row => {
    const e = {};
    Object.keys(EMP_COLS).forEach(f => { e[f] = getCell(row, f); });
    return e;
  }).filter(e => e.name);
}

function isResigned_(e) { return /퇴사|퇴직/.test(e.status || ''); }

// 부서 토큰: 자기 부서 재직·휴직자만 (퇴사자 제외) / 관리자: 전체 + 재직상태
function getEmployees(s) {
  const list = readEmployees_();
  if (list === null) return { success: false, error: '직원목록 시트가 없습니다. 시트명을 "직원목록"으로 해주세요.' };
  const pick = e => ({ name: e.name, dept: e.dept, empNo: e.empNo, joinDate: e.joinDate, status: e.status });
  const data = s && s.role === 'dept'
    ? list.filter(e => (sameDept_(e.dept, s.dept) || sameDept_(deptAlias_(e.dept), s.dept)) && !isResigned_(e)).map(pick)
    : list.map(pick);
  return { success: true, data: data };
}

function getEmployeeDepts_() {
  const list = readEmployees_() || [];
  return [...new Set(list.filter(e => !isResigned_(e)).map(e => e.dept).filter(Boolean))].sort();
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
          healthCheck: get(row, 'healthCheck', '건강검진', '특수건강검진'),
          memo: get(row, '메모1', '비고', 'memo'), memo2: get(row, '메모2', 'memo2') });
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
      components:    compName ? [{ compName, casNumber, amount, workEnv: firstWorkEnv, healthCheck: firstHealthCheck,
                                       memo: get(row, '메모1', '비고', 'memo'), memo2: get(row, '메모2', 'memo2') }] : [],
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
  if (s.role === 'dept' && !sameDept_(p.dept, s.dept)) {
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
  if (!isPpe) { try { nhMarkEducated_(String(p.dept || '').trim(), p.names ? p.names.split('|') : []); } catch (_) {} }
  eduCacheClear_();   // v20.1 교육 현황 캐시 비움
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

// ════════════════════════════════════════════════════════════════
// 직원정보 동기화 (관리자)
//  · 사번 기준으로 추가·갱신. 행을 지우지 않음 (퇴사자도 남김 → 교육기록 조회 가능)
//  · "전체 명단"으로 올린 경우에만, 명단에 없는 재직자를 '퇴사(명단없음)'으로 표시
//  · 교육·보호구 기록은 별도 파일/시트라 동기화가 건드리지 않음
// ════════════════════════════════════════════════════════════════
const EMP_NEW_HEADER = ['사번','이름','부서','입사일','재직상태','퇴사일','이메일','직위','최초확인일','최종확인일'];

function todayKst_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowKst_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

// 엑셀 붙여넣기·인사 사이트 결과(헤더 + 행 배열)를 직원 객체로 변환
function rowsToEmployees_(headers, rows) {
  // 항목마다 해당하는 열을 모두 찾고, 값이 있는 첫 열을 사용 (예: 사번이 empNo·empCode로 섞여 있어도 인식)
  const cols = {};
  Object.keys(EMP_COLS).forEach(f => {
    cols[f] = EMP_COLS[f].map(k => headers.map(h => String(h).trim()).indexOf(k)).filter(i => i >= 0);
  });
  if (!cols.name.length) throw new Error('이름(성명) 열을 찾지 못했습니다. 첫 줄에 열 제목을 포함해 주세요.');
  return rows.map(r => {
    const e = {};
    Object.keys(EMP_COLS).forEach(f => {
      const i = cols[f].find(c => String(r[c] == null ? '' : r[c]).trim() !== '');
      e[f] = i === undefined ? '' : String(r[i]).trim();
    });
    e.joinDate = normDate_(e.joinDate);
    e.leaveDate = normDate_(e.leaveDate);
    return e;
  }).filter(e => e.name);
}
function normDate_(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})[.\-\/년\s]*(\d{1,2})[.\-\/월\s]*(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s;
}

function syncEmployees_(incoming, fullList) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('서버가 바쁩니다. 잠시 후 다시 시도하세요.');
  try {
    let sheet = findEmpSheet_();
    if (!sheet) { sheet = SS.insertSheet('직원목록'); sheet.appendRow(EMP_NEW_HEADER); }
    if (sheet.getLastRow() === 0) sheet.appendRow(EMP_NEW_HEADER);
    // 없는 열은 끝에 추가 (기존 열 순서는 유지)
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    let idx = empColIndex_(headers);
    Object.keys(EMP_COLS).forEach(f => {
      if (idx[f] < 0) { headers.push(EMP_COLS[f][0]); sheet.getRange(1, headers.length).setValue(EMP_COLS[f][0]); }
    });
    idx = empColIndex_(headers);
    const width = headers.length;
    const last = sheet.getLastRow();
    const data = last > 1 ? sheet.getRange(2, 1, last - 1, width).getValues() : [];

    const today = todayKst_();
    // 사번으로 찾고, 사번이 없으면 이름+입사일로 찾음
    const noKey = v => String(v || '').trim() ? 'no:' + String(v).trim() : '';
    const nmKey = (name, joinDate) => 'nm:' + String(name).trim() + '|' + normDate_(fmtDate_(joinDate));
    const pos = {};
    data.forEach((r, i) => {
      const kn = noKey(r[idx.empNo]);
      if (kn) pos[kn] = i;
      const km = nmKey(r[idx.name], r[idx.joinDate]);
      if (pos[km] === undefined) pos[km] = i;
    });
    const rowKey = r => noKey(r[idx.empNo]) || nmKey(r[idx.name], r[idx.joinDate]);

    // 부서 이름: 매핑 시트 적용 → 앱에 없는 이름은 매핑 시트에 '확인 필요'로 추가
    const dmap = readDeptMap_();
    const known = knownAppDepts_();
    const unknown = {};
    incoming.forEach(e => {
      const raw = String(e.dept || '').trim();
      const n = normDept_(raw);
      if (!n) return;
      if (dmap[n]) { e.dept = dmap[n]; return; }
      if (known[n]) { e.dept = known[n]; return; }   // 띄어쓰기만 다른 경우 앱 이름으로 통일
      if (!(n in dmap)) unknown[n] = raw;
    });
    const unknownNames = Object.keys(unknown).map(k => unknown[k]).sort();
    if (unknownNames.length) {
      const msh = deptMapSheet_();
      unknownNames.forEach(raw => msh.appendRow([safeCell_(raw), '', '확인 필요: 부서인증·chemicals에 없는 이름 (같은 부서면 앱 부서명 입력, 새 부서면 비워 두기)']));
    }

    const seen = new Set();
    const summary = { added: 0, updated: 0, resigned: 0, unchanged: 0, newEmployees: [], unknownDepts: unknownNames };
    incoming.forEach(e => {
      const kn = noKey(e.empNo), km = nmKey(e.name, e.joinDate);
      let at = kn ? pos[kn] : undefined;
      if (at === undefined && pos[km] !== undefined) {
        // 이름+입사일이 같아도 양쪽 사번이 다르면 다른 사람
        const other = noKey(data[pos[km]][idx.empNo]);
        if (!kn || !other) at = pos[km];
      }
      if (at !== undefined) seen.add(rowKey(data[at]));
      if (at === undefined) {
        const row = new Array(width).fill('');
        ['empNo','name','dept','joinDate','leaveDate','email','position'].forEach(f => { row[idx[f]] = e[f] || ''; });
        row[idx.status] = e.status || '재직';
        row[idx.firstSeen] = today;
        row[idx.lastSeen] = today;
        if (kn) pos[kn] = data.length;
        pos[km] = data.length;
        seen.add(rowKey(row));
        data.push(row);
        summary.added++;
        summary.newEmployees.push({ name: e.name, dept: e.dept, joinDate: e.joinDate });
        return;
      }
      const row = data[at];
      let changed = false;
      ['name','dept','joinDate','status','leaveDate','email','position'].forEach(f => {
        if (!e[f]) return;
        if (fmtDate_(row[idx[f]]) !== e[f]) { row[idx[f]] = e[f]; changed = true; }
      });
      if (!e.status && isResigned_({ status: fmtDate_(row[idx.status]) })) { row[idx.status] = '재직'; changed = true; }   // 재입사
      row[idx.lastSeen] = today;
      if (changed) summary.updated++; else summary.unchanged++;
    });

    if (fullList) {
      data.forEach(row => {
        if (seen.has(rowKey(row)) || isResigned_({ status: fmtDate_(row[idx.status]) })) return;
        row[idx.status] = '퇴사(명단없음)';
        if (!fmtDate_(row[idx.leaveDate])) row[idx.leaveDate] = today;
        summary.resigned++;
      });
    }
    if (data.length) {
      sheet.getRange(2, 1, data.length, width).setValues(data.map(r => r.map(v => v instanceof Date ? v : safeCell_(v))));
    }
    setAdminCfg({ key: 'emp_last_sync', value: nowKst_() + ' (추가 ' + summary.added + ' · 변경 ' + summary.updated + ' · 퇴사표시 ' + summary.resigned + ')' });
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function empStatus_() {
  const list = readEmployees_() || [];
  const cfg = getAdminCfg().data || {};
  return {
    success: true,
    total: list.length,
    active: list.filter(e => !isResigned_(e) && !/휴직/.test(e.status || '')).length,
    leave: list.filter(e => /휴직/.test(e.status || '')).length,
    resigned: list.filter(isResigned_).length,
    lastSync: cfg.emp_last_sync || '',
    hrAuto: hrAutoReady_(),
    hrConnector: hrConnectorReady_(),
    hrSite: PROPS.getProperty('FB_SITE_URL') || '',
    dailyOn: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'dailyCheck'),
    hrLastError: cfg.hr_last_error || '',
    deptCheck: deptCheck_(list),
    serverVersion: SERVER_VERSION,
  };
}

// 재직 중인 직원이 있는 부서별: 로그인 가능 · 부서장 메일 · 사용 물질 수
function deptCheck_(list) {
  const auth = readAuthRows_();
  const emails = deptEmails_();
  const chemMap = handlingDepts_();
  const counts = {};
  list.filter(e => !isResigned_(e) && e.dept).forEach(e => {
    const k = normDept_(e.dept);
    (counts[k] = counts[k] || { dept: e.dept, staff: 0 }).staff++;
  });
  return Object.keys(counts).sort().map(k => {
    const c = counts[k];
    const a = auth.find(r => sameDept_(r.dept, c.dept));
    const chems = chemsForDept_(chemMap, c.dept);
    const em = Object.keys(emails).find(k => sameDept_(k, c.dept));
    return { dept: c.dept, staff: c.staff, login: !!(a && a.active && a.pin), email: !!(em && emails[em]),
             chems: chems ? chems.length : 0 };
  });
}

// ════════════════════════════════════════════════════════════════
// 인사 사이트 연결 (Firebase: 이메일·비밀번호 로그인 → Firestore 또는 Realtime Database 조회)
//  · 연결 정보(apiKey·projectId·데이터 위치)는 시트 메뉴 ④에서 입력 → 스크립트 속성(비공개)에 저장
//  · 수동: 관리자가 화면에서 비밀번호 입력 → 이 요청에만 사용, 저장하지 않음. 메일 발송 없음
//  · 자동(매일): 메뉴 ④에서 저장한 아이디·비밀번호 사용 → 이때만 신규입사 알림 메일 발송
// ════════════════════════════════════════════════════════════════
function fbCfg_() {
  return {
    apiKey: PROPS.getProperty('FB_API_KEY') || '',
    projectId: PROPS.getProperty('FB_PROJECT_ID') || '',
    dbUrl: PROPS.getProperty('FB_DB_URL') || '',
    mode: PROPS.getProperty('FB_MODE') || 'firestore',      // firestore | rtdb
    path: PROPS.getProperty('FB_PATH') || '',               // 예: employees
    fullList: PROPS.getProperty('FB_FULL_LIST') === 'Y',    // Y면 사이트에 없는 재직자를 퇴사로 표시
    // 휴직 기록 위치 (휴직자관리: leaves). 직원 위치가 employees_public이면 기본으로 leaves 사용
    leavePath: PROPS.getProperty('FB_LEAVE_PATH') || (PROPS.getProperty('FB_PATH') === 'employees_public' ? 'leaves' : ''),
    fieldMap: (function () { try { return JSON.parse(PROPS.getProperty('FB_FIELD_MAP') || '{}'); } catch (_) { return {}; } })(),
  };
}
function hrConnectorReady_() {
  const c = fbCfg_();
  return !!(c.apiKey && c.path && (c.mode === 'rtdb' ? c.dbUrl : c.projectId));
}
function hrAutoReady_() {
  return hrConnectorReady_() && !!PROPS.getProperty('HR_ID') && !!PROPS.getProperty('HR_PW');
}

function fbSignIn_(c, email, password) {
  const res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=' + encodeURIComponent(c.apiKey), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ email: email, password: password, returnSecureToken: true }),
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !j.idToken) {
    const code = (j.error && j.error.message) || ('HTTP ' + res.getResponseCode());
    const msg = /INVALID_PASSWORD|INVALID_LOGIN_CREDENTIALS|EMAIL_NOT_FOUND/.test(code) ? '인사 사이트 아이디 또는 비밀번호가 올바르지 않습니다'
              : /TOO_MANY_ATTEMPTS/.test(code) ? '인사 사이트 로그인 시도가 너무 많아 잠시 막혔습니다'
              : /USER_DISABLED/.test(code) ? '인사 사이트 계정이 비활성화되어 있습니다'
              : '인사 사이트 로그인 실패: ' + code;
    throw new Error(msg);
  }
  return j.idToken;
}

// Firestore 값 → 일반 값 (날짜는 KST yyyy-MM-dd)
function fsValue_(v) {
  if (!v || typeof v !== 'object') return '';
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return String(v.integerValue);
  if ('doubleValue' in v) return String(v.doubleValue);
  if ('booleanValue' in v) return v.booleanValue ? 'Y' : 'N';
  if ('timestampValue' in v) return Utilities.formatDate(new Date(v.timestampValue), TZ, 'yyyy-MM-dd');
  if ('nullValue' in v) return '';
  if ('mapValue' in v) { const o = {}; Object.keys((v.mapValue && v.mapValue.fields) || {}).forEach(k => { o[k] = fsValue_(v.mapValue.fields[k]); }); return o; }
  if ('arrayValue' in v) return ((v.arrayValue && v.arrayValue.values) || []).map(fsValue_).join(', ');
  if ('referenceValue' in v) return v.referenceValue.split('/').pop();
  return '';
}

// 조회 실패 사유를 알기 쉽게 (Firebase가 보낸 메시지 포함)
function fbErr_(res, path) {
  const code = res.getResponseCode();
  let msg = '';
  try { const j = JSON.parse(res.getContentText() || '{}'); msg = (j.error && (j.error.message || j.error)) || ''; } catch (_) {}
  const hint = code === 403 ? ' → 이 계정에 "' + path + '" 읽기 권한이 없습니다 (데이터 위치 이름 또는 Firebase 보안 규칙 확인)'
             : code === 404 ? ' → 프로젝트 또는 데이터 위치 "' + path + '"를 찾을 수 없습니다'
             : '';
  return '인사 데이터 조회 실패 (HTTP ' + code + (msg ? ', ' + String(msg).slice(0, 150) : '') + ')' + hint;
}

function fbFetchObjects_(c, idToken, path) {
  path = path || c.path;
  const out = [];
  if (c.mode === 'rtdb') {
    const url = c.dbUrl.replace(/\/+$/, '') + '/' + path.split('/').map(encodeURIComponent).join('/') + '.json?auth=' + encodeURIComponent(idToken);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error(fbErr_(res, path));
    const data = JSON.parse(res.getContentText() || 'null') || {};
    (Array.isArray(data) ? data : Object.keys(data).map(k => Object.assign({ _key: k }, data[k])))
      .forEach(o => { if (o && typeof o === 'object') out.push(o); });
    return out;
  }
  let pageToken = '';
  for (let page = 0; page < 50; page++) {
    const url = 'https://firestore.googleapis.com/v1/projects/' + encodeURIComponent(c.projectId) + '/databases/(default)/documents/' +
      path.split('/').map(encodeURIComponent).join('/') + '?pageSize=300' + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + idToken }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error(fbErr_(res, path));
    const j = JSON.parse(res.getContentText() || '{}');
    (j.documents || []).forEach(d => {
      const o = { _key: String(d.name || '').split('/').pop() };
      Object.keys(d.fields || {}).forEach(k => { o[k] = fsValue_(d.fields[k]); });
      out.push(o);
    });
    if (!j.nextPageToken) break;
    pageToken = j.nextPageToken;
  }
  return out;
}

// 반환: { headers, rows } (rowsToEmployees_가 열 이름으로 사번·이름·부서·입사일 등을 찾음)
function hrFetchEmployees_(id, password) {
  const c = fbCfg_();
  if (!hrConnectorReady_()) throw new Error('인사 사이트 연결 정보가 없습니다 (시트 메뉴 ④에서 설정)');
  if (!id || !password) throw new Error('인사 사이트 아이디와 비밀번호를 입력하세요');
  const idToken = fbSignIn_(c, id, password);
  const objs = fbFetchObjects_(c, idToken);
  // 재직상태가 없으면 퇴사일·휴직 기록으로 계산 (퇴사일이 오늘 이전 → 퇴사, 오늘이 휴직 기간 → 휴직)
  const today = todayKst_();
  let onLeave = {};
  if (c.leavePath) {
    try {
      fbFetchObjects_(c, idToken, c.leavePath).forEach(l => {
        const code = String(l.empCode || '').trim();
        const st = normDate_(l.start || ''), en = normDate_(l.end || '');
        if (code && (!st || st <= today) && (!en || en >= today)) onLeave[code] = l.type || '휴직';
      });
    } catch (e) { onLeave = {}; }   // 휴직 기록을 못 읽어도 직원 명단은 계속 사용
  }
  objs.forEach(o => {
    if (o.status || o['재직상태']) return;
    const code = String(o.empCode || o.empNo || o['사번'] || o._key || '').trim();
    const quit = normDate_(o.quitDate || o.leaveDate || o['퇴사일'] || '');
    o.status = quit && quit <= today ? '퇴사' : (onLeave[code] ? '휴직(' + onLeave[code] + ')' : '재직');
  });
  const flat = objs.map(o => {
    const f = {};
    Object.keys(o).forEach(k => {
      const v = o[k];
      if (v && typeof v === 'object') Object.keys(v).forEach(k2 => { f[k + '.' + k2] = typeof v[k2] === 'object' ? JSON.stringify(v[k2]) : v[k2]; });
      else f[k] = v;
    });
    return f;
  });
  // 주민등록번호 등 민감정보 열은 가져오지 않음 (화면·시트·서버 응답 어디에도 남기지 않음)
  const SENSITIVE = /rrn|jumin|ssn|resident|주민|생년월일|birth|phone|mobile|tel|연락처|전화|주소|address|account|계좌/i;
  const headers = [];
  flat.forEach(f => Object.keys(f).forEach(k => { if (headers.indexOf(k) < 0 && !SENSITIVE.test(k)) headers.push(k); }));
  // 사이트 필드명 → 표준 이름 (메뉴 ④의 필드 매핑, 예: {"성명":"nm"} → nm 열을 이름으로)
  const outHeaders = headers.map(h => {
    const std = Object.keys(c.fieldMap).find(k => c.fieldMap[k] === h);
    return std || h;
  });
  return { headers: outHeaders, rows: flat.map(f => headers.map(h => f[h] === undefined ? '' : String(f[h]))) };
}

// 메뉴 ④ 1/5 입력 → 연결 정보. 사이트 주소(https://○○.web.app)면 Firebase 호스팅의
// /__/firebase/init.json 에서 자동으로 읽고, 아니면 붙여넣은 firebaseConfig에서 찾음
function fbConfigFromInput_(text) {
  const t = String(text || '').trim();
  if (/^https?:\/\//i.test(t)) {
    const base = t.replace(/[?#].*$/, '').replace(/\/+$/, '');
    const res = UrlFetchApp.fetch(base + '/__/firebase/init.json', { muteHttpExceptions: true });
    let j = null;
    try { j = JSON.parse(res.getContentText() || 'null'); } catch (_) {}
    if (res.getResponseCode() !== 200 || !j || !j.apiKey) {
      throw new Error('사이트 주소에서 연결 정보를 읽지 못했습니다 (HTTP ' + res.getResponseCode() + '). 페이지 소스의 firebaseConfig를 붙여넣어 주세요.');
    }
    return { apiKey: j.apiKey, projectId: j.projectId || '', databaseURL: j.databaseURL || '', site: base };
  }
  const pick = k => { const m = t.match(new RegExp(k + '\\s*["\']?\\s*:\\s*["\']([^"\']+)["\']')); return m ? m[1] : ''; };
  return { apiKey: pick('apiKey'), projectId: pick('projectId'), databaseURL: pick('databaseURL'), site: '' };
}

function saveHrCredentials() {
  const ui = SpreadsheetApp.getUi();
  const ask = (title, msg, def) => {
    const r = ui.prompt(title, msg + (def ? '\n(현재: ' + def + ' — 비우면 유지)' : ''), ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) throw new Error('취소');
    return r.getResponseText().trim() || def || '';
  };
  try {
    const c = fbCfg_();
    const cfgText = ask('인사 사이트(Firebase) 연결 1/5', '휴직자관리 사이트 주소를 넣으세요 (예: https://bumin-leave.web.app)\n또는 페이지 소스의 firebaseConfig 내용을 붙여넣어도 됩니다', c.apiKey ? '저장됨' : '');
    if (cfgText && cfgText !== '저장됨') {
      const fc = fbConfigFromInput_(cfgText);
      if (!fc.apiKey) throw new Error('apiKey를 찾지 못했습니다');
      PROPS.setProperty('FB_API_KEY', fc.apiKey);
      if (fc.projectId) PROPS.setProperty('FB_PROJECT_ID', fc.projectId);
      if (fc.databaseURL) PROPS.setProperty('FB_DB_URL', fc.databaseURL);
      if (fc.site) PROPS.setProperty('FB_SITE_URL', fc.site);
    }
    PROPS.setProperty('FB_MODE', /^r/i.test(ask('인사 사이트(Firebase) 연결 2/5', '데이터 종류: firestore 또는 rtdb', c.mode)) ? 'rtdb' : 'firestore');
    PROPS.setProperty('FB_PATH', ask('인사 사이트(Firebase) 연결 3/5', '직원 데이터 위치 (휴직자관리 사이트는 employees_public)', c.path || 'employees_public'));
    PROPS.setProperty('HR_ID', ask('인사 사이트(Firebase) 연결 4/5', '자동 로그인 아이디(이메일)', PROPS.getProperty('HR_ID') || ''));
    const pw = ui.prompt('인사 사이트(Firebase) 연결 5/5', '자동 로그인 비밀번호 (스크립트 속성에만 저장, 코드·시트에 남지 않음)\n비우면 기존 값 유지', ui.ButtonSet.OK_CANCEL);
    if (pw.getSelectedButton() !== ui.Button.OK) throw new Error('취소');
    if (pw.getResponseText()) PROPS.setProperty('HR_PW', pw.getResponseText());
  } catch (e) {
    if (String(e.message) !== '취소') ui.alert('저장 실패: ' + e.message);
    return;
  }
  try {
    const r = hrFetchEmployees_(PROPS.getProperty('HR_ID'), PROPS.getProperty('HR_PW'));
    const list = rowsToEmployees_(r.headers, r.rows);
    ui.alert('연결 성공: 인사 사이트에서 ' + r.rows.length + '건을 읽었고, 그중 ' + list.length + '명을 직원으로 인식했습니다.\n' +
             '읽은 열: ' + r.headers.slice(0, 15).join(', ') + (r.headers.length > 15 ? ' …' : '') +
             (list.length ? '' : '\n\n이름 열을 찾지 못했습니다. Claude에게 위 열 목록을 알려주면 필드 매핑을 맞춰 드립니다.'));
  } catch (e) {
    ui.alert('연결 정보는 저장했지만 시험 조회에 실패했습니다: ' + e.message);
  }
}
function clearHrCredentials() {
  ['HR_ID', 'HR_PW'].forEach(k => PROPS.deleteProperty(k));
  SpreadsheetApp.getUi().alert('인사 사이트 자동 로그인 아이디·비밀번호를 삭제했습니다. (연결 정보는 유지)');
}

// ════════════════════════════════════════════════════════════════
// 신규입사 알림
//  흐름: 새 입사자 중 화학물질 사용부서 소속 → 부서장 메일([확인하기] 링크)
//        → 부서장 PIN 로그인 후 취급 여부 회신 → '예'면 교육일지 작성
//        → 보건관리자: 관리자 모드에서 회신·미회신·미교육 확인 + 회신 알림 메일
//        → 회신이 없으면 3일마다 재알림 (최대 3회)
// ════════════════════════════════════════════════════════════════
const NH_SHEET = '신규입사알림';
const NH_HEADER = ['알림ID','생성일시','사번','이름','부서','입사일','취급물질','부서장이메일','최초발송','최근재알림',
                   '재알림횟수','회신일시','취급여부','회신메모','교육일지제출일','상태','교육기한'];
const EDU_PERIOD_DAYS = 3;   // 입사일로부터 3일 안에 교육
const NH_MAX_REMIND = 3;

function sysCfg_() {
  const c = getAdminCfg().data || {};
  return {
    appUrl: c['앱주소'] || '',
    adminEmail: c['보건관리자이메일'] || '',
    newHireDays: parseInt(c['신규입사기준일수'] || '7', 10),
    remindDays: parseInt(c['재알림간격일'] || '3', 10),
  };
}

function nhSheet_() {
  let sh = SS.getSheetByName(NH_SHEET);
  if (!sh) { sh = SS.insertSheet(NH_SHEET); sh.appendRow(NH_HEADER); sh.setFrozenRows(1); }
  return sh;
}
function nhRows_() {
  const sh = nhSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, NH_HEADER.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    NH_HEADER.forEach((h, j) => { o[h] = cellStr_(r[j]); });
    return o;
  });
}
function nhWrite_(o) {
  nhSheet_().getRange(o._row, 1, 1, NH_HEADER.length).setValues([NH_HEADER.map(h => safeCell_(o[h] === undefined ? '' : o[h]))]);
}

// 부서 → 그 부서가 사용하는 화학물질 이름 목록 (chemicals 시트의 부서 열 기준)
function handlingDepts_() {
  const map = {};
  (getChemicals().data || []).forEach(c => {
    (c.departments || []).forEach(d => { (map[d] = map[d] || []).push(c.name); });
  });
  return map;
}
// 그 부서 전용 물질 + (병동이면) chemicals 시트 '병동' 공통 물질. 없으면 null
function chemsForDept_(map, dept) {
  dept = deptAlias_(dept);
  const key = Object.keys(map).find(k => k !== '병동' && sameDept_(k, dept));
  // 시트의 '병동'은 앱 화면과 같은 세부 병동 목록(+ 이름이 '병동'으로 끝나는 부서)에 적용
  const isWard = WARDS.some(w => sameDept_(w, dept)) || /병동(\(.*\))?$/.test(normDept_(dept));
  const list = (key ? map[key] : []).concat(isWard && map['병동'] ? map['병동'] : []);
  return list.length ? list.filter((n, i) => list.indexOf(n) === i) : null;
}

// { 부서: 부서장 이메일 } — 부서인증 시트 1회 읽기
function deptEmails_() {
  const sheet = SS.getSheetByName(AUTH_SHEET);
  const out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const ci = headers.indexOf('부서장 이메일');
  if (ci < 0) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    .forEach(r => { if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[ci]).trim(); });
  return out;
}

// 부서장 이름 — 부서인증 시트의 '부서장'(또는 '부서장 이름' · '부서장명') 열. 교육일지 교육 실시자 자동 입력용
function deptHead_(dept) {
  const hs = SS.getSheetByName('부서장');   // 별도 '부서장' 시트(부서 · 부서장명)가 있으면 우선
  if (hs && hs.getLastRow() >= 2) {
    const h = hs.getRange(1, 1, 1, hs.getLastColumn()).getValues()[0].map(x => String(x).trim());
    const ni = ['부서장명', '부서장', '이름'].map(x => h.indexOf(x)).find(i => i >= 0);
    if (ni !== undefined) {
      const r = hs.getRange(2, 1, hs.getLastRow() - 1, hs.getLastColumn()).getDisplayValues().find(x => sameDept_(x[0], dept));
      if (r && String(r[ni]).trim()) return String(r[ni]).trim();
    }
  }
  const sheet = SS.getSheetByName(AUTH_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const ci = ['부서장', '부서장 이름', '부서장명'].map(h => headers.indexOf(h)).find(i => i >= 0);
  if (ci === undefined) return '';
  const row = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    .find(r => sameDept_(r[0], dept));
  return row ? String(row[ci]).trim() : '';
}

function deptEmail_(dept) {
  const sheet = SS.getSheetByName(AUTH_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const ci = headers.indexOf('부서장 이메일');
  if (ci < 0) return '';
  const row = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues()
    .find(r => sameDept_(r[0], dept));
  return row ? String(row[ci]).trim() : '';
}

function detectNewHires_() {
  const cfg = sysCfg_();
  const map = handlingDepts_();
  const existing = new Set(nhRows_().map(o => (o['사번'] || '') + '|' + o['이름'] + '|' + o['부서']));
  const today = new Date(todayKst_() + 'T00:00:00');
  const from = new Date(today.getTime() - cfg.newHireDays * 86400000);
  const created = [];
  (readEmployees_() || []).forEach(e => {
    if (isResigned_(e) || !e.joinDate) return;
    const jd = new Date(normDate_(e.joinDate) + 'T00:00:00');
    if (isNaN(jd.getTime()) || jd < from) return;
    const chems = chemsForDept_(map, e.dept);
    if (!chems) return;
    const key = (e.empNo || '') + '|' + e.name + '|' + e.dept;
    if (existing.has(key)) return;
    existing.add(key);
    const o = { '알림ID': 'NH' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
                '생성일시': nowKst_(), '사번': e.empNo, '이름': e.name, '부서': e.dept, '입사일': normDate_(e.joinDate),
                '취급물질': chems.slice(0, 30).join(', '), '부서장이메일': deptEmail_(e.dept), '재알림횟수': '0',
                '교육기한': addDays_(normDate_(e.joinDate), EDU_PERIOD_DAYS) };
    nhSheet_().appendRow(NH_HEADER.map(h => safeCell_(o[h] || '')));
    o._row = nhSheet_().getLastRow();
    o['상태'] = nhSend_(o, false) ? '회신대기' : '이메일없음';
    if (o['상태'] === '회신대기') o['최초발송'] = nowKst_();
    nhWrite_(o);
    created.push({ name: o['이름'], dept: o['부서'], status: o['상태'] });
  });
  return created;
}

function addDays_(ymd, n) {
  const d = new Date(ymd + 'T12:00:00Z');
  if (isNaN(d.getTime())) return '';
  return new Date(d.getTime() + n * 86400000).toISOString().slice(0, 10);
}

function nhSend_(o, isReminder) {
  if (!o['부서장이메일']) return false;
  const cfg = sysCfg_();
  const link = cfg.appUrl ? cfg.appUrl + (cfg.appUrl.indexOf('?') >= 0 ? '&' : '?') + 'nh=' + encodeURIComponent(o['알림ID']) : '';
  const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const html =
    '<div style="font-family:sans-serif;font-size:14px;line-height:1.6">' +
    (isReminder ? '<p style="color:#b91c1c"><b>[재알림] 아직 회신이 없어 다시 보내드립니다.</b></p>' : '') +
    '<p>' + esc(o['부서']) + ' 부서장님, 유해화학물질 사용 부서에 신규 입사자가 있습니다.</p>' +
    '<table style="border-collapse:collapse" cellpadding="6" border="1">' +
    '<tr><th align="left">이름</th><td>' + esc(o['이름']) + '</td></tr>' +
    '<tr><th align="left">입사일</th><td>' + esc(o['입사일']) + '</td></tr>' +
    '<tr><th align="left">교육 기한</th><td><b>' + esc(o['교육기한'] || addDays_(o['입사일'], EDU_PERIOD_DAYS)) + '</b> (입사일로부터 ' + EDU_PERIOD_DAYS + '일)</td></tr>' +
    '<tr><th align="left">부서 사용 물질</th><td>' + esc(o['취급물질']) + '</td></tr></table>' +
    '<p>이 직원이 위 물질을 <b>취급하는지</b> 확인해 주세요. 취급하면 기한 안에 MSDS 교육을 하고 교육일지를 작성해 주세요.</p>' +
    (link ? '<p><a href="' + esc(link) + '" style="display:inline-block;background:#1c3557;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none">확인하기</a></p>'
          : '<p>앱에 부서 PIN으로 로그인하면 확인 요청이 표시됩니다.</p>') +
    '<p style="color:#64748b;font-size:12px">부민병원 보건관리실 · 이 메일에 답장하지 말고 위 버튼으로 회신해 주세요.</p></div>';
  const opts = { to: o['부서장이메일'], subject: (isReminder ? '[재알림] ' : '') + '[보건관리] 신규 입사자 유해화학물질 취급 확인 요청 - ' + o['이름'],
                 htmlBody: html, name: '부민병원 보건관리' };
  if (cfg.adminEmail) opts.replyTo = cfg.adminEmail;
  MailApp.sendEmail(opts);
  return true;
}

function sendReminders_() {
  const cfg = sysCfg_();
  const now = Date.now();
  let sent = 0;
  nhRows_().forEach(o => {
    if (o['상태'] !== '회신대기') return;
    const n = parseInt(o['재알림횟수'] || '0', 10);
    if (n >= NH_MAX_REMIND) return;
    const lastSent = new Date((o['최근재알림'] || o['최초발송']).replace(' ', 'T') + '+09:00').getTime();
    if (isNaN(lastSent) || now - lastSent < cfg.remindDays * 86400000) return;
    if (nhSend_(o, true)) { o['최근재알림'] = nowKst_(); o['재알림횟수'] = String(n + 1); nhWrite_(o); sent++; }
  });
  return sent;
}

function nhList_(s) {
  const rows = nhRows_().map(o => { const c = Object.assign({}, o); delete c._row; return c; });
  if (s.role === 'admin') return { success: true, data: rows };
  return { success: true, data: rows.filter(o => sameDept_(o['부서'], s.dept) && (o['상태'] === '회신대기' || o['상태'] === '이메일없음' || o['상태'] === '취급-교육대기')) };
}

function nhAnswer_(req, s) {
  const o = nhRows_().find(x => x['알림ID'] === req.id);
  if (!o) return { success: false, status: 404, error: '알림을 찾을 수 없습니다' };
  if (s.role !== 'admin' && !sameDept_(o['부서'], s.dept)) return { success: false, status: 403, error: '다른 부서 알림입니다' };
  const handles = req.handles === true;
  o['회신일시'] = nowKst_();
  o['취급여부'] = handles ? '예' : '아니오';
  o['회신메모'] = String(req.memo || '').slice(0, 300);
  o['상태'] = handles ? (o['교육일지제출일'] ? '취급-교육완료' : '취급-교육대기') : '미취급';
  nhWrite_(o);
  const cfg = sysCfg_();
  if (cfg.adminEmail) {
    try {
      MailApp.sendEmail({ to: cfg.adminEmail, name: '부민병원 보건관리',
        subject: '[보건관리] 신규입사 회신 - ' + o['부서'] + ' ' + o['이름'] + ' (취급: ' + o['취급여부'] + ')',
        htmlBody: '<p>' + o['부서'] + ' ' + o['이름'] + ' (입사 ' + o['입사일'] + ')</p><p>유해화학물질 취급: <b>' + o['취급여부'] + '</b>' +
                  (o['회신메모'] ? '<br>메모: ' + o['회신메모'].replace(/[<>&]/g, '') : '') + '</p><p>관리자 모드 → 직원정보관리에서 확인하세요.</p>' });
    } catch (e) { /* 메일 실패는 회신 기록에 영향 없음 */ }
  }
  return { success: true, status: o['상태'] };
}

// 교육일지 저장 시: 같은 부서·이름의 '취급-교육대기' 알림을 교육완료로 표시
function nhMarkEducated_(dept, names) {
  if (!SS.getSheetByName(NH_SHEET)) return;
  const set = new Set(names.map(n => String(n).trim()));
  nhRows_().forEach(o => {
    if (!sameDept_(o['부서'], dept) || !set.has(o['이름']) || o['교육일지제출일']) return;
    o['교육일지제출일'] = todayKst_();
    if (o['상태'] === '취급-교육대기') o['상태'] = '취급-교육완료';   // 회신 전 제출이면 회신 '예' 때 완료 처리
    nhWrite_(o);
  });
}

function nhResend_(req) {
  const o = nhRows_().find(x => x['알림ID'] === req.id);
  if (!o) return { success: false, status: 404, error: '알림을 찾을 수 없습니다' };
  if (!o['부서장이메일']) o['부서장이메일'] = deptEmail_(o['부서']);
  if (!nhSend_(o, true)) return { success: false, error: '부서장 이메일이 없습니다 (부서인증 시트에 입력)' };
  o['최근재알림'] = nowKst_();
  if (o['상태'] === '이메일없음') { o['상태'] = '회신대기'; o['최초발송'] = o['최초발송'] || nowKst_(); }
  nhWrite_(o);
  return { success: true };
}

// ════════════════════════════════════════════════════════════════
// 매일 아침 자동 확인 (시트 메뉴 ⑤에서 켜기)
//  · 인사 사이트 자동 로그인이 가능할 때만: 직원정보 갱신 → 신규입사자 확인·메일
//  · 재알림(3일 무회신)은 매일 확인
// ════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════
// 인사 사이트 데이터 이상 감지 (사이트 구조가 바뀌어도 조용히 틀린 데이터가 들어오지 않게)
//  · 인원이 지난번 정상 조회보다 30% 넘게 줄면 중단
//  · 재직자 중 입사일·부서·사번이 빈 비율이 지난번 정상 때보다 30%p 넘게 늘면 중단 (필드 이름이 바뀐 경우)
//    (원래부터 비어 있는 데이터는 오탐하지 않도록 '지난번 정상값' 기준. 첫 실행은 기준값만 기록)
// ════════════════════════════════════════════════════════════════
const HR_DROP_LIMIT = 0.3;
const HR_FIELD_JUMP = 0.3;
function hrSanity_(list) {
  const cfg = getAdminCfg().data || {};
  const prev = parseInt(cfg.hr_last_count || '0', 10);
  let base = null;
  try { base = JSON.parse(cfg.hr_last_ratios || 'null'); } catch (_) {}
  const active = list.filter(e => !isResigned_(e));
  const ratio = f => active.length ? active.filter(e => !String(e[f] || '').trim()).length / active.length : 1;
  const ratios = { joinDate: ratio('joinDate'), dept: ratio('dept'), empNo: ratio('empNo') };
  const LABEL = { joinDate: '입사일', dept: '부서', empNo: '사번' };
  const SUBJ = { joinDate: '입사일이', dept: '부서가', empNo: '사번이' };
  const problems = [];
  if (!list.length) problems.push('직원을 한 명도 인식하지 못했습니다 (이름 필드가 바뀌었을 수 있음)');
  if (prev && list.length < prev * (1 - HR_DROP_LIMIT)) problems.push('인원이 지난번 ' + prev + '명 → ' + list.length + '명으로 크게 줄었습니다');
  if (base && active.length) Object.keys(LABEL).forEach(f => {
    if (ratios[f] - (base[f] || 0) > HR_FIELD_JUMP) {
      problems.push('재직자 중 ' + SUBJ[f] + ' 빈 비율이 ' + Math.round((base[f] || 0) * 100) + '% → ' + Math.round(ratios[f] * 100) +
                    '%로 늘었습니다 (' + LABEL[f] + ' 필드 이름이 바뀌었을 수 있음)');
    }
  });
  return { ok: !problems.length, problems: problems, count: list.length, prev: prev, ratios: ratios };
}
function saveHrBaseline_(chk) {
  setAdminCfg({ key: 'hr_last_count', value: String(chk.count) });
  setAdminCfg({ key: 'hr_last_ratios', value: JSON.stringify(chk.ratios) });
}

// 자동 확인 실패 시 보건관리자에게 메일 (같은 날 같은 사유는 1번만)
function notifyHrFailure_(msg) {
  const cfg = sysCfg_();
  if (!cfg.adminEmail) return false;
  const key = todayKst_() + '|' + sha256Hex_(msg).slice(0, 12);
  if ((getAdminCfg().data || {}).hr_fail_mail_key === key) return false;
  try {
    MailApp.sendEmail({ to: cfg.adminEmail, name: '부민병원 보건관리',
      subject: '[보건관리] 인사 사이트 자동 확인 실패 - 신규입사 알림이 멈춰 있습니다',
      htmlBody: '<div style="font-family:sans-serif;font-size:14px;line-height:1.6">' +
        '<p>매일 아침 인사 사이트(휴직자관리) 자동 확인이 실패해 <b>직원정보를 반영하지 않았습니다.</b> ' +
        '이 문제가 해결될 때까지 신규입사 알림 메일이 나가지 않습니다.</p>' +
        '<p style="background:#fef2f2;padding:10px;border-radius:6px;color:#b91c1c">' + String(msg).replace(/[<>&]/g, '') + '</p>' +
        '<p>확인할 것: ① 인사 사이트 계정 비밀번호가 바뀌었으면 시트 메뉴 ④에서 다시 저장 ' +
        '② 인사 사이트의 필드 이름·저장 위치가 바뀌었는지 담당자에게 확인 ③ 관리자 모드 → 직원정보관리에서 오류 내용 확인</p>' +
        '<p style="color:#64748b;font-size:12px">같은 사유의 메일은 하루 1번만 보냅니다.</p></div>' });
    setAdminCfg({ key: 'hr_fail_mail_key', value: key });
    return true;
  } catch (e) { return false; }
}

function runAutoCheck_() {
  const out = { created: [], reminded: 0 };
  try {
    const r = hrFetchEmployees_(PROPS.getProperty('HR_ID'), PROPS.getProperty('HR_PW'));
    const list = rowsToEmployees_(r.headers, r.rows);
    const chk = hrSanity_(list);
    if (!chk.ok) throw new Error('인사 사이트 데이터가 평소와 달라 반영하지 않았습니다: ' + chk.problems.join(' / '));
    out.sync = syncEmployees_(list, fbCfg_().fullList);
    saveHrBaseline_(chk);
    out.created = detectNewHires_();
    setAdminCfg({ key: 'hr_last_error', value: '' });
  } catch (e) {
    const msg = String(e.message || e);
    setAdminCfg({ key: 'hr_last_error', value: nowKst_() + ' ' + msg.slice(0, 300) });
    out.error = msg;
    out.mailed = notifyHrFailure_(msg);
  }
  out.reminded = sendReminders_();   // 자동 로그인으로 보낸 알림의 재알림
  return out;
}

// 자동 로그인이 될 때만 진행 (안 되면 아무 메일도 보내지 않음)
function dailyCheck() {
  if (hrAutoReady_()) runAutoCheck_();
}

function enableDailyCheck() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyCheck').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('매일 아침 7시경 자동 확인을 켰습니다.' +
    (hrAutoReady_() ? '' : '\n(인사 사이트 자동 로그인(메뉴 ④)이 설정되기 전에는 아무 것도 하지 않습니다)'));
}
function disableDailyCheck() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyCheck').forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('매일 자동 확인을 껐습니다.');
}

// 관리자 화면 요청 처리
function staffOp_(req, s) {
  const deny = requireAdmin_(s);
  switch (req.op) {
    case 'emp.status': return deny || empStatus_();
    case 'emp.sync': {
      if (deny) return deny;
      const list = rowsToEmployees_(req.headers || [], req.rows || []);
      if (!list.length) return { success: false, error: '올릴 직원이 없습니다' };
      const summary = syncEmployees_(list, req.fullList === true);   // 수동 업데이트: 알림 메일 없음
      return Object.assign({ success: true }, summary);
    }
    case 'hr.fetch': {
      if (deny) return deny;
      // 비밀번호를 비우면 메뉴 ④에서 저장한 자동 로그인 계정 사용 (수동 가져오기는 메일 없음)
      const useSaved = !req.password;
      const r = hrFetchEmployees_(useSaved ? (PROPS.getProperty('HR_ID') || '') : String(req.id || PROPS.getProperty('HR_ID') || ''),
                                  useSaved ? (PROPS.getProperty('HR_PW') || '') : String(req.password));
      // 수동 가져오기도 이상 감지 결과를 함께 돌려줌 (화면에서 경고 후 관리자가 판단)
      const chk = hrSanity_(rowsToEmployees_(r.headers, r.rows));
      if (chk.ok && useSaved) saveHrBaseline_(chk);
      return { success: true, headers: r.headers, rows: r.rows, warnings: chk.problems };
    }
    case 'nh.list':   return requireLogin_(s) || nhList_(s);
    case 'nh.answer': return requireLogin_(s) || nhAnswer_(req, s);
    case 'nh.resend': return deny || nhResend_(req);
    case 'nh.run':    // '지금 확인' = 매일 자동 확인과 같은 동작 (자동 로그인 가능할 때만 메일)
      if (deny) return deny;
      if (!hrAutoReady_()) return { success: false, error: '인사 사이트 자동 로그인이 설정되어야 합니다 (시트 메뉴 ④). 알림 메일은 자동 로그인일 때만 발송됩니다.' };
      return Object.assign({ success: true }, runAutoCheck_());
  }
  return { success: false, error: 'unknown op: ' + req.op };
}

// ════════════════════════════════════════════════════════════════
// 2단계: 화학물질 마스터 정비 (AI 판정의 기준 데이터)
//  · 점검: CAS 형식·검증숫자 오류, 날짜로 바뀐 CAS(구글시트 자동 변환) 복원 제안, 같은 CAS 판정 불일치, 빈칸
//  · 법령대조표 시트: CAS 기준 특수건강진단·작업환경측정 주기 (기존 시트 판정으로 시작, 고시 목록 붙여넣기)
//  · 빈칸 채우기는 '제안'만 — 관리자가 승인한 것만, 빈 칸에만 기록 (기존 값 절대 덮어쓰지 않음)
// ════════════════════════════════════════════════════════════════
const LAW_SHEET = '법령대조표';
const LAW_HEADER = ['CAS', '물질명', '특수건강진단', '작업환경측정', '관리대상', '특별관리', '출처', '확인일', '비고'];

function findChemSheet_() {
  return SS.getSheetByName('chemicals') || SS.getSheetByName('Chemicals') || SS.getSheetByName('화학물질')
      || SS.getSheetByName('MSDS목록') || SS.getSheetByName('MSDS');
}

// CAS 검증숫자: 끝자리 = (앞 숫자들을 오른쪽부터 1,2,3…배 한 합) mod 10
function casCheck_(cas) {
  const m = String(cas || '').trim().match(/^(\d{2,7})-(\d{2})-(\d)$/);
  if (!m) return null;
  const digits = (m[1] + m[2]).split('').reverse();
  return digits.reduce((sum, d, i) => sum + (i + 1) * Number(d), 0) % 10 === Number(m[3]);
}
// 시트 셀 값 → CAS 문자열 + 복원 제안 (날짜로 바뀐 값, 끝의 점 등)
function casFromCell_(v) {
  if (v instanceof Date) {
    const y = v.getFullYear(), mo = v.getMonth() + 1, d = v.getDate();
    const guess = y + '-' + ('0' + mo).slice(-2) + '-' + d;
    return { raw: Utilities.formatDate(v, TZ, 'yyyy-MM-dd'), fixed: casCheck_(guess) ? guess : '', why: '날짜로 바뀐 CAS' };
  }
  const raw = String(v == null ? '' : v).trim();
  let m = raw.match(/^(\d{2,7})-(\d{1,2})-(\d{1,2})(?:\s+00:00:00)?$/);   // 날짜 문자열로 바뀐 경우 (9014-01-01)
  if (m && raw.length >= 10 && !casCheck_(raw)) {
    const guess = m[1].replace(/^0+(?=\d{2})/, '') + '-' + ('0' + Number(m[2])).slice(-2) + '-' + Number(m[3]);
    if (Number(m[3]) < 10 && casCheck_(guess)) return { raw: raw, fixed: guess, why: '날짜로 바뀐 CAS' };
  }
  const trimmed = raw.replace(/[.\s]+$/, '');
  if (trimmed !== raw && casCheck_(trimmed)) return { raw: raw, fixed: trimmed, why: '끝에 불필요한 문자' };
  // v19.6: 앞의 점 · 중간 띄어쓰기 · 앞자리 0 (예: .64-02-8, 64-17- 5, 0584-08-07)
  const squeezed = raw.replace(/\s+/g, '').replace(/^[^0-9]+/, '').replace(/[^0-9]+$/, '');
  if (squeezed !== raw && casCheck_(squeezed)) return { raw: raw, fixed: squeezed, why: '불필요한 문자 · 띄어쓰기' };
  m = squeezed.match(/^0*(\d{2,7})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const guess = m[1] + '-' + ('0' + Number(m[2])).slice(-2) + '-' + Number(m[3]);
    if (guess !== raw && Number(m[3]) < 10 && casCheck_(guess)) return { raw: raw, fixed: guess, why: '날짜로 바뀐 CAS' };
  }
  if (/^\d{3}-\d{3}-\d$/.test(raw)) return { raw: raw, fixed: '', why: 'EC 번호 (CAS 아님) — MSDS에서 CAS 확인' };
  return { raw: raw, fixed: '', why: '' };
}
// 함유량 칸이 날짜로 바뀐 경우 (예: 9-10 → 2026-09-10) → 원래 '월-일' 범위로 복원
function amountFromCell_(v) {
  if (v instanceof Date) return { raw: Utilities.formatDate(v, TZ, 'yyyy-MM-dd'), fixed: (v.getMonth() + 1) + '-' + v.getDate() };
  const raw = String(v == null ? '' : v).trim();
  const m = raw.match(/^\d{4}-(\d{1,2})-(\d{1,2})(?:[\sT]00:00(?::00)?(?:\.000Z?)?)?$/);
  return m ? { raw: raw, fixed: Number(m[1]) + '-' + Number(m[2]) } : { raw: raw, fixed: '' };
}

// chemicals 시트를 성분 행 단위로 (시트 행 번호 포함)
function chemRows_() {
  const sheet = findChemSheet_();
  if (!sheet) return { sheet: null, rows: [] };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { if (h && col[h] === undefined) col[h] = i; });
  const find = (...keys) => { for (const k of keys) if (col[k] !== undefined) return col[k]; return -1; };
  const c = {
    name: find('name', '물질명', '제품명'), comp: find('compName', '구성성분명', '성분명', '성분'),
    cas: find('casNumber', 'CAS', 'CAS번호', 'cas'), amount: find('amount', '함유량', '함유량(%)', '농도'),
    hc: find('healthCheck', '건강검진', '특수건강검진', '특수건강진단'), we: find('workEnv', '작업환경측정', '작업환경'),
    dept: find('department', 'departments', '부서', '사용부서'),
    memo: find('메모1', '비고', 'memo'),
  };
  const out = [];
  let product = '', dept = '';
  values.slice(1).forEach((r, i) => {
    const nm = c.name >= 0 ? String(r[c.name] || '').trim() : '';
    if (nm) { product = nm; dept = c.dept >= 0 ? String(r[c.dept] || '').trim() : ''; }
    const comp = c.comp >= 0 ? String(r[c.comp] || '').trim() : '';
    const casCell = c.cas >= 0 ? r[c.cas] : '';
    if (!product || (!comp && (casCell === '' || casCell === null))) return;
    const ci = casFromCell_(casCell);
    const ai = c.amount >= 0 ? amountFromCell_(r[c.amount]) : { raw: '', fixed: '' };
    out.push({ row: i + 2, product: product, dept: dept, comp: comp, cas: ci.fixed || ci.raw, casRaw: ci.raw, casFix: ci.fixed, casWhy: ci.why,
               amount: ai.fixed || ai.raw, amountRaw: ai.raw, amountFix: ai.fixed,
               hc: c.hc >= 0 ? String(r[c.hc] || '').trim() : '', we: c.we >= 0 ? String(r[c.we] || '').trim() : '',
               memo: c.memo >= 0 ? String(r[c.memo] || '').trim() : '' });
  });
  return { sheet: sheet, rows: out, col: c };
}

function lawSheet_() {
  let sh = SS.getSheetByName(LAW_SHEET);
  if (!sh) {
    sh = SS.insertSheet(LAW_SHEET);
    sh.appendRow(LAW_HEADER);
    sh.getRange('A:A').setNumberFormat('@');   // CAS가 날짜로 바뀌지 않게
    sh.setFrozenRows(1);
  }
  return sh;
}
function readLaw_() {
  const sh = SS.getSheetByName(LAW_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, LAW_HEADER.length).getDisplayValues().forEach(r => {
    const cas = String(r[0]).trim();
    if (cas) map[cas] = { name: r[1], hc: String(r[2]).trim(), we: String(r[3]).trim(), target: r[4], special: r[5], src: r[6], date: r[7] };
  });
  return map;
}
const isCycle_ = v => /개월|년/.test(String(v || '')) && !/제외/.test(String(v || ''));

// 점검 + 제안
function chemAudit_() {
  const { rows } = chemRows_();
  const law = readLaw_();
  const byCas = {};
  rows.forEach(r => { if (casCheck_(r.cas)) (byCas[r.cas] = byCas[r.cas] || []).push(r); });
  const casIssues = rows.filter(r => r.casRaw && !/영업비밀|trade secret|비공개|자료없음|^-$/i.test(r.casRaw) && (r.casFix || !casCheck_(r.casRaw)))
    .map(r => {
      const o = { row: r.row, product: r.product, comp: r.comp, raw: r.casRaw, fix: r.casFix, why: r.casFix ? r.casWhy : (r.casWhy || (casCheck_(r.casRaw) === false ? '검증숫자 틀림' : '형식 아님')) };
      if (!o.fix) {   // v19.6: 오타 추정 — 같은 시트에 있는 CAS 중 앞자리 하나 더하기/빼기로 맞는 것 (관리자가 확인 후 체크)
        const t = String(r.casRaw).replace(/\s+/g, '');
        const cands = [];
        if (/^\d{2,7}-\d{2}-\d$/.test(t)) {
          for (let d = 1; d <= 9; d++) cands.push(d + t);
          if (/^\d{3,}/.test(t)) cands.push(t.slice(1));
        }
        const hit = cands.filter(x => byCas[x]);
        if (hit.length === 1) { o.fix = hit[0]; o.guess = true; o.why = o.why + ' → 같은 시트의 ' + hit[0] + '(' + byCas[hit[0]][0].comp + ')로 추정'; }
      }
      return o;
    });
  const conflicts = [];
  Object.keys(byCas).forEach(cas => {
    ['hc', 'we'].forEach(f => {
      const vals = [...new Set(byCas[cas].map(r => r[f]).filter(isCycle_))];
      if (vals.length > 1) {
        const memos = byCas[cas].filter(r => r.memo).map(r => r.product + ': ' + r.memo);
        conflicts.push({ cas: cas, comp: byCas[cas][0].comp, field: f, values: vals, explained: memos.length > 0, memos: memos,
                         rows: byCas[cas].map(r => r.product + ' (' + (r[f] || '미해당') + ')') });
      }
    });
  });
  const amountIssues = rows.filter(r => r.amountFix).map(r => ({ row: r.row, product: r.product, comp: r.comp, raw: r.amountRaw, fix: r.amountFix }));
  // 빈칸 제안: 법령대조표 → 없으면 같은 CAS의 다른 행 판정(값이 하나로 일치할 때만)
  const suggestions = [];
  rows.forEach(r => {
    if (!casCheck_(r.cas)) return;
    ['hc', 'we'].forEach(f => {
      if (r[f]) return;
      let val = '', src = '';
      if (law[r.cas] && isCycle_(law[r.cas][f])) { val = law[r.cas][f]; src = '법령대조표(' + (law[r.cas].src || '') + ')'; }
      else {
        const vals = [...new Set((byCas[r.cas] || []).map(x => x[f]).filter(isCycle_))];
        if (vals.length === 1) { val = vals[0]; src = '같은 CAS의 다른 제품 판정'; }
      }
      if (val) suggestions.push({ row: r.row, field: f, product: r.product, comp: r.comp, cas: r.cas, amount: r.amount, value: val, src: src });
    });
  });
  const blankHc = rows.filter(r => !r.hc).length, blankWe = rows.filter(r => !r.we).length;
  return { success: true, total: rows.length, products: new Set(rows.map(r => r.product)).size, uniqueCas: Object.keys(byCas).length,
           blankHc: blankHc, blankWe: blankWe, lawCount: Object.keys(law).length,
           casIssues: casIssues, conflicts: conflicts, suggestions: suggestions, amountIssues: amountIssues,
           unknownCas: Object.keys(byCas).filter(cas => !law[cas]).length };
}

// 법령대조표 시작값: 기존 시트에서 보건관리자가 이미 판정한 값 (CAS 정상 + 판정이 하나로 일치하는 것만)
function chemSeedLaw_() {
  const { rows } = chemRows_();
  const law = readLaw_();
  const byCas = {};
  rows.forEach(r => { if (casCheck_(r.cas)) (byCas[r.cas] = byCas[r.cas] || []).push(r); });
  const sh = lawSheet_();
  const today = todayKst_();
  const add = [];
  Object.keys(byCas).forEach(cas => {
    if (law[cas]) return;
    const hcs = [...new Set(byCas[cas].map(r => r.hc).filter(isCycle_))];
    const wes = [...new Set(byCas[cas].map(r => r.we).filter(isCycle_))];
    if (hcs.length > 1 || wes.length > 1 || (!hcs.length && !wes.length)) return;
    add.push([cas, byCas[cas][0].comp, hcs[0] || '', wes[0] || '', '', '', '기존 시트(보건관리자 판정)', today, '고시 확인 필요']);
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, LAW_HEADER.length).setValues(add.map(r => r.map(safeCell_)));
  return { success: true, added: add.length };
}

// 관리자가 고른 제안·CAS 수정을 시트에 반영 (빈 칸만 / CAS는 텍스트 형식으로)
function chemApply_(req) {
  const { sheet, rows, col } = chemRows_();
  if (!sheet) return { success: false, error: 'chemicals 시트가 없습니다' };
  const byRow = {};
  rows.forEach(r => { byRow[r.row] = r; });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '서버가 바쁩니다' };
  let filled = 0, fixed = 0, amounts = 0, skipped = 0;
  try {
    (req.fills || []).forEach(f => {
      const r = byRow[f.row];
      const ci = f.field === 'hc' ? col.hc : col.we;
      if (!r || ci < 0 || r[f.field] || !isCycle_(f.value)) { skipped++; return; }   // 그 사이 값이 생겼으면 건너뜀
      sheet.getRange(f.row, ci + 1).setValue(safeCell_(f.value));
      filled++;
    });
    if (col.cas >= 0) sheet.getRange(2, col.cas + 1, Math.max(sheet.getLastRow() - 1, 1), 1).setNumberFormat('@');   // 재발 방지
    (req.casFixes || []).forEach(f => {
      const r = byRow[f.row];
      const fix = String(f.fix || '').trim();
      // 자동 복원값이거나, 문제 있는 칸을 관리자가 검증숫자가 맞는 CAS로 고친 경우만 (정상 CAS는 바꾸지 않음)
      const isIssue = r && (r.casFix || !casCheck_(r.casRaw));
      if (!r || col.cas < 0 || !isIssue || !casCheck_(fix) || (f.raw !== undefined && String(f.raw) !== r.casRaw)) { skipped++; return; }
      sheet.getRange(f.row, col.cas + 1).setValue(fix);
      fixed++;
    });
    if ((req.amountFixes || []).length && col.amount >= 0) {
      sheet.getRange(2, col.amount + 1, Math.max(sheet.getLastRow() - 1, 1), 1).setNumberFormat('@');   // 다시 날짜로 바뀌지 않게
      req.amountFixes.forEach(f => {
        const r = byRow[f.row];
        if (!r || !r.amountFix || r.amountFix !== f.fix) { skipped++; return; }
        sheet.getRange(f.row, col.amount + 1).setValue(f.fix);
        amounts++;
      });
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true, filled: filled, fixed: fixed, amounts: amounts, skipped: skipped };
}

// 법령대조표: 고시 목록 붙여넣기 (별표 표를 복사해 붙이면 줄마다 CAS를 찾아 해당 항목에 기록)
//  kind: hc 특수건강진단 · we 작업환경측정 · target 관리대상 · special 특별관리 / value: 주기 또는 '해당'
//  이미 값이 있는 칸은 덮어쓰지 않음 (다르면 '다름'으로 알려줌)
function chemLawImport_(req) {
  const KIND = { hc: 2, we: 3, target: 4, special: 5 };
  const ci = KIND[req.kind];
  if (ci === undefined) return { success: false, error: '항목을 고르세요' };
  const value = String(req.value || '').trim() || (req.kind === 'hc' || req.kind === 'we' ? '' : '해당');
  if (!value) return { success: false, error: '주기(예: 12개월)를 입력하세요' };
  const src = String(req.src || '').trim().slice(0, 80) || '고시 목록 붙여넣기';
  const lines = String(req.text || '').split(/\r?\n/).slice(0, 3000);
  const sh = lawSheet_();
  const today = todayKst_();
  const last = sh.getLastRow();
  const cur = last >= 2 ? sh.getRange(2, 1, last - 1, LAW_HEADER.length).getDisplayValues() : [];
  const at = {};
  cur.forEach((r, i) => { if (String(r[0]).trim()) at[String(r[0]).trim()] = i; });
  let added = 0, updated = 0, same = 0, invalid = 0;
  const differ = [], seen = {};
  lines.forEach(line => {
    const found = String(line).match(/\b\d{2,7}-\d{2}-\d\b/g) || [];
    found.forEach(cas => {
      if (seen[cas]) return;
      seen[cas] = 1;
      if (!casCheck_(cas)) { invalid++; return; }
      const name = String(line).replace(/\b\d{2,7}-\d{2}-\d\b/g, '').replace(/[\t;|]+/g, ' ').replace(/^[\s\d.)]+/, '').replace(/\s{2,}/g, ' ').trim().slice(0, 80);
      if (at[cas] === undefined) {
        const row = [cas, name, '', '', '', '', src, today, '고시 붙여넣기'];
        row[ci] = value;
        cur.push(row); at[cas] = cur.length - 1; added++;
      } else {
        const row = cur[at[cas]];
        const old = String(row[ci] || '').trim();
        if (!old) { row[ci] = value; row[6] = String(row[6] || '').indexOf(src) >= 0 ? row[6] : (row[6] ? row[6] + ' · ' : '') + src; row[7] = today; if (!row[1]) row[1] = name; updated++; }
        else if (old === value) same++;
        else differ.push(cas + ' (지금 ' + old + ' / 붙여넣기 ' + value + ')');
      }
    });
  });
  if (cur.length) {
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(2, 1, cur.length, LAW_HEADER.length).setValues(cur.map(r => r.map(safeCell_)));
  }
  return { success: true, added: added, updated: updated, same: same, invalid: invalid, differ: differ.slice(0, 50), differCount: differ.length };
}

// ════════════════════════════════════════════════════════════════
// 3단계: AI MSDS 자동 해석 (Gemini) — 관리자 업로드 또는 부서 신규사용 보고서의 MSDS
//  · 키: 스크립트 속성 GEMINI_API_KEY (시트 메뉴 ⑥ 또는 관리자 화면) — 화면에는 저장하지 않음
//  · 판정 근거 3가지를 나란히: 법령대조표(CAS) · 기존 시트의 같은 CAS 판정 · MSDS 15번(법적규제) 항목
//  · 결과는 '제안'만 — 관리자가 고쳐서 승인한 것만 chemicals 시트에 새 제품으로 추가 (기존 제품은 바꾸지 않음)
//  · 개인정보는 보내지 않음 (MSDS 문서만)
// ════════════════════════════════════════════════════════════════
const AI_MAX_BYTES = 15 * 1024 * 1024;
function saveGeminiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('AI(Gemini) 키 저장', 'Google AI Studio(aistudio.google.com)에서 만든 API 키를 붙여넣으세요.\n비워 두고 확인하면 저장된 키를 지웁니다.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const k = r.getResponseText().trim();
  if (k) PROPS.setProperty('GEMINI_API_KEY', k); else PROPS.deleteProperty('GEMINI_API_KEY');
  ui.alert(k ? 'AI 키를 저장했습니다.' : 'AI 키를 지웠습니다.');
}
function aiModel_() { return (getAdminCfg().data || {})['AI모델'] || 'gemini-2.5-flash'; }

const AI_PROMPT = [
  '너는 산업보건 전문가다. 첨부한 한국어 물질안전보건자료(MSDS) PDF에서 아래 항목만 JSON으로 추출해라.',
  '문서에 없는 값은 빈 문자열로 두고 추측하지 마라. CAS 번호는 문서에 적힌 그대로 옮겨라(영업비밀이면 "영업비밀").',
  '함유량은 문서 표기 그대로(예: "1~5", "<1", "30-40").',
  '{"product":{"name":"제품명","manufacturer":"제조사 또는 공급자","usage":"권고 용도","signalWord":"신호어(위험/경고)",',
  '"pictograms":["그림문자 이름(한국어: 폭발성, 인화성, 산화성, 고압가스, 부식성, 급성독성, 경고, 호흡기과민성, 수생환경유해성, 발암성 등)"],',
  '"emergency":{"eye":"4번 눈에 들어갔을 때","skin":"피부에 접촉했을 때","inhalation":"흡입했을 때","ingestion":"먹었을 때"}},',
  '"components":[{"name":"구성성분명","cas":"CAS 번호","amount":"함유량(%)",',
  '"reg15":{"workEnv":"15번 법적규제에서 작업환경측정 대상 여부와 주기(예: 6개월, 해당없음)","healthCheck":"특수건강진단 대상 여부와 주기(예: 12개월, 해당없음)",',
  '"managed":"관리대상유해물질 해당 여부(해당/해당없음)","special":"특별관리물질 해당 여부(해당/해당없음)"}}],',
  '"warnings":["추출이 불확실한 부분"]}'
].join('\n');

function aiCallGemini_(bytes) {
  const key = PROPS.getProperty('GEMINI_API_KEY');
  if (!key) return { success: false, noKey: true, error: 'AI 키가 없습니다 — 시트 메뉴 ⑥ 또는 관리자 화면에서 Gemini 키를 저장하세요. (키 없이도 아래에 직접 입력해 등록할 수 있습니다)' };
  if (bytes.length > AI_MAX_BYTES) return { success: false, error: 'PDF가 너무 큽니다 (15MB 이하)' };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(aiModel_()) + ':generateContent';
  const body = { contents: [{ parts: [{ inline_data: { mime_type: 'application/pdf', data: Utilities.base64Encode(bytes) } }, { text: AI_PROMPT }] }],
                 generationConfig: { temperature: 0, responseMimeType: 'application/json' } };
  let res;
  try {
    res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(body),
                                   headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  } catch (e) { return { success: false, error: 'AI 연결 실패: ' + e.message }; }
  const code = res.getResponseCode();
  let j = null;
  try { j = JSON.parse(res.getContentText()); } catch (_) {}
  if (code === 429) return { success: false, error: 'AI 무료 사용 한도를 넘었습니다. 잠시(1분~내일) 후 다시 시도하거나 직접 입력하세요.' };
  if (code !== 200) return { success: false, error: 'AI 오류 (' + code + '): ' + ((j && j.error && j.error.message) || '').slice(0, 200) };
  const text = (((j.candidates || [])[0] || {}).content || { parts: [] }).parts.map(p => p.text || '').join('');
  try { return { success: true, data: JSON.parse(text.replace(/^```(?:json)?\s*|\s*```$/g, '')) }; }
  catch (_) { return { success: false, error: 'AI 응답을 읽지 못했습니다 — 다시 시도하거나 직접 입력하세요' }; }
}

// AI 추출값(또는 직접 입력값)에 판정 근거를 붙임
function aiJudge_(data) {
  const { rows } = chemRows_();
  const law = readLaw_();
  const byCas = {};
  rows.forEach(r => { if (casCheck_(r.cas)) (byCas[r.cas] = byCas[r.cas] || []).push(r); });
  const yes = v => /해당|대상|예|yes|true|o/i.test(String(v || '')) && !/해당\s*없음|비해당|미해당|아님|없음|no|false/i.test(String(v || ''));
  const cycleOf = v => { const m = String(v || '').match(/(\d+)\s*개월/); return m ? m[1] + '개월' : ''; };
  const p = data.product || {};
  const comps = (data.components || []).map(c => {
    const ci = casFromCell_(String(c.cas || '').trim());
    const cas = ci.fixed || ci.raw;
    const ok = casCheck_(cas) === true;
    const same = ok ? (byCas[cas] || []) : [];
    const sheetHc = [...new Set(same.map(r => r.hc).filter(Boolean))], sheetWe = [...new Set(same.map(r => r.we).filter(Boolean))];
    const L = ok ? law[cas] : null;
    const r15 = c.reg15 || {};
    // 제안 우선순위: 법령대조표 → 기존 시트(같은 CAS, 값이 하나) → MSDS 15번
    const pick = (lawV, sheetVals, msdsV) => {
      if (lawV && isCycle_(lawV)) return { v: lawV, src: '법령대조표' };
      if (sheetVals.length === 1) return { v: sheetVals[0], src: '기존 시트(같은 CAS)' };
      if (cycleOf(msdsV)) return { v: cycleOf(msdsV), src: 'MSDS 15번' };
      return { v: '', src: '' };
    };
    const hc = pick(L && L.hc, sheetHc, r15.healthCheck), we = pick(L && L.we, sheetWe, r15.workEnv);
    return {
      name: String(c.name || '').trim(), cas: cas, casRaw: String(c.cas || ''), casOk: ok, casNote: ci.fixed ? ci.why : (ok ? '' : (ci.why || '검증숫자 틀림 · MSDS 확인')),
      amount: String(c.amount || '').trim(),
      hc: hc.v, hcSrc: hc.src, we: we.v, weSrc: we.src,
      managed: (L && yes(L.target)) || yes(r15.managed), special: (L && yes(L.special)) || yes(r15.special),
      evidence: { law: L ? { hc: L.hc, we: L.we, target: L.target, special: L.special } : null,
                  sheet: { hc: sheetHc, we: sheetWe, products: same.slice(0, 3).map(r => r.product) },
                  msds15: { hc: r15.healthCheck || '', we: r15.workEnv || '', managed: r15.managed || '', special: r15.special || '' } }
    };
  });
  const iconMap = {};
  try {
    const sh = findChemSheet_();
    const v = sh ? sh.getDataRange().getValues() : [];
    const ic = v.length ? v[0].map(h => String(h).trim()).indexOf('iconUrl') : -1;
    if (ic >= 0) v.slice(1).forEach(r => String(r[ic] || '').split(',').forEach(part => {
      const pts = part.trim().split('|'); if (pts[1] && !iconMap[pts[1].trim()]) iconMap[pts[1].trim()] = pts[0].trim();
    }));
  } catch (_) {}
  const names = rows.map(r => r.product);
  const pname = String(p.name || '').trim();
  return {
    product: { name: pname, manufacturer: String(p.manufacturer || '').trim(), usage: String(p.usage || '').trim(), signalWord: String(p.signalWord || '').trim(),
               pictograms: (p.pictograms || []).map(x => String(x).trim()).filter(Boolean),
               emergency: p.emergency || {} },
    components: comps,
    targetManage: comps.some(c => c.managed), specialManage: comps.some(c => c.special),
    iconKnown: Object.keys(iconMap), duplicate: pname && names.indexOf(pname) >= 0,
    warnings: (data.warnings || []).map(String).slice(0, 10)
  };
}

function aiBytes_(req) {
  if (req.path) {
    const path = normPath_(req.path);
    if (path.indexOf('records/chem_msds/') !== 0 && path.indexOf('MSDS원본/') !== 0) return { error: '읽을 수 없는 경로' };
    const f = getFile_(path);
    if (!f) return { error: '파일을 찾지 못했습니다' };
    return { bytes: f.getBlob().getBytes(), path: path };
  }
  if (!req.content) return { error: 'PDF가 없습니다' };
  const bytes = Utilities.base64Decode(String(req.content));
  const name = String(req.name || 'msds.pdf').replace(/[\\/:*?"<>|]/g, '_').slice(-80);
  const path = 'MSDS원본/' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + '_' + Date.now() + '_' + name;
  writeFile_(path, bytes, 'application/pdf', '');   // 관리자 업로드 원본 보관
  return { bytes: bytes, path: path };
}

function aiOp_(req, s) {
  const deny = requireAdmin_(s);
  if (deny) return deny;
  switch (req.op) {
    case 'ai.status':
      return { success: true, hasKey: !!PROPS.getProperty('GEMINI_API_KEY'), model: aiModel_() };
    case 'ai.setKey': {
      const k = String(req.key || '').trim();
      if (k) PROPS.setProperty('GEMINI_API_KEY', k); else PROPS.deleteProperty('GEMINI_API_KEY');
      return { success: true, hasKey: !!k };
    }
    case 'ai.list': {   // 부서가 신규사용 보고서에 첨부한 MSDS
      const out = [];
      getFolders_(['records', 'chem_msds']).forEach(folder => {
        const it = folder.getFiles();
        while (it.hasNext()) {
          const f = it.next();
          let m = {}; try { m = JSON.parse(f.getDescription() || '{}'); } catch (_) {}
          out.push({ path: 'records/chem_msds/' + f.getName(), name: f.getName(), dept: m.dept || '', size: f.getSize(),
                     date: Utilities.formatDate(f.getLastUpdated(), TZ, 'yyyy-MM-dd'), aiDone: !!m.aiDone });
        }
      });
      out.sort((a, b) => (b.date + b.name).localeCompare(a.date + a.name));
      return { success: true, data: out.slice(0, 100) };
    }
    case 'ai.analyze': {
      const b = aiBytes_(req);
      if (b.error) return { success: false, error: b.error };
      const r = aiCallGemini_(b.bytes);
      if (!r.success) return Object.assign(r, { path: b.path, draft: aiJudge_({ product: {}, components: [] }) });
      return { success: true, path: b.path, model: aiModel_(), draft: aiJudge_(r.data) };
    }
    case 'ai.products': {   // 비교할 기존 제품 목록
      const names = [];
      chemRows_().rows.forEach(r => { if (names.indexOf(r.product) < 0) names.push(r.product); });
      return { success: true, data: names };
    }
    case 'ai.compare': {    // 기존 제품 묶음 (개정 비교용)
      const sh = findChemSheet_();
      const b = sh ? chemBlock_(sh, String(req.name || '').trim()) : null;
      if (!b) return { success: false, error: '기존 제품을 찾지 못했습니다' };
      return { success: true, row: b.row, count: b.count, product: b.product, components: b.components };
    }
    case 'ai.judge':   // 직접 입력 · 고친 값으로 다시 판정 (AI 없이)
      return { success: true, draft: aiJudge_(req.data || {}) };
    case 'ai.apply':
      return aiApply_(req);
  }
  return { success: false, error: 'unknown op: ' + req.op };
}

// chemicals 시트 열 위치
function chemCols_(sheet) {
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const ix = (...ks) => { for (const k of ks) { const i = head.indexOf(k); if (i >= 0) return i; } return -1; };
  return { head: head, C: { name: ix('name', '물질명', '제품명'), mf: ix('manufacturer', '제조사'), dept: ix('department', '부서', '사용부서'), usage: ix('usage', '용도'),
    comp: ix('compName', '구성성분명', '성분명'), cas: ix('casNumber', 'CAS', 'CAS번호'), amt: ix('amount', '함유량'),
    tg: ix('targetManage', '관리대상'), sp: ix('specialManage', '특별관리'), we: ix('workEnv', '작업환경측정'), hc: ix('healthCheck', '특수건강검진', '건강검진'),
    em: ix('emergency', '응급처치'), pdf: ix('pdfUrl', 'PDF'), icon: ix('iconUrl', '위험그림문자'), memo: ix('메모1', '비고'), memo2: ix('메모2') } };
}
// 제품 한 묶음(제품 줄 + 아래 성분 줄)의 위치와 값
function chemBlock_(sheet, name) {
  const { head, C } = chemCols_(sheet);
  const v = sheet.getDataRange().getValues();
  const disp = sheet.getDataRange().getDisplayValues();
  const nm = r => C.name >= 0 ? String(v[r][C.name] || '').trim() : '';
  let start = -1;
  for (let r = 1; r < v.length; r++) if (nm(r) === name) { start = r; break; }
  if (start < 0) return null;
  let end = start + 1;
  while (end < v.length && !nm(end) && disp[end].some(x => String(x).trim())) end++;
  const g = (r, k) => C[k] >= 0 ? String(disp[r][C[k]] || '').trim() : '';
  return {
    row: start + 1, count: end - start, raw: v.slice(start, end), width: head.length,
    product: { name: name, manufacturer: g(start, 'mf'), department: g(start, 'dept'), usage: g(start, 'usage'), target: g(start, 'tg'), special: g(start, 'sp'),
               emergency: g(start, 'em'), pdfUrl: g(start, 'pdf'), iconUrl: g(start, 'icon') },
    components: v.slice(start, end).map((_, i) => ({ name: g(start + i, 'comp'), cas: g(start + i, 'cas'), amount: g(start + i, 'amt'),
      hc: g(start + i, 'hc'), we: g(start + i, 'we'), memo: g(start + i, 'memo'), memo2: g(start + i, 'memo2') })).filter(c => c.name || c.cas)
  };
}

// 승인: chemicals 시트에 새 제품 추가 (첫 줄 = 제품 정보 + 첫 성분, 다음 줄 = 성분만)
//  mode 'revise' + target: 기존 제품 묶음을 새 내용으로 바꿈 (이전 줄은 '개정이력' 시트에 보관)
function aiApply_(req) {
  const d = req.data || {};
  const p = d.product || {};
  const comps = (d.components || []).filter(c => String(c.name || '').trim());
  const name = String(p.name || '').trim();
  if (!name) return { success: false, error: '제품명을 입력하세요' };
  if (!comps.length) return { success: false, error: '구성성분을 1개 이상 입력하세요' };
  const bad = comps.filter(c => /^\d/.test(String(c.cas || '').trim()) && !casCheck_(String(c.cas).trim()));
  if (bad.length) return { success: false, error: 'CAS 검증숫자가 맞지 않습니다: ' + bad.map(c => c.cas).join(', ') };
  const sheet = findChemSheet_();
  if (!sheet) return { success: false, error: 'chemicals 시트가 없습니다' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: '서버가 바쁩니다' };
  try {
    const { rows } = chemRows_();
    const revise = req.mode === 'revise';
    const target = String(req.target || '').trim();
    let old = null;
    if (revise) {
      old = chemBlock_(sheet, target);
      if (!old) return { success: false, error: '개정할 기존 제품을 찾지 못했습니다: ' + target };
      if (name !== target && rows.some(r => r.product === name)) return { success: false, error: '바꿀 제품명과 같은 이름의 다른 제품이 있습니다' };
    } else if (rows.some(r => r.product === name)) {
      return { success: false, error: '같은 이름의 제품이 이미 있습니다 — [기존 제품과 비교]로 개정 반영하세요' };
    }
    const { head, C } = chemCols_(sheet);
    // 원본 PDF 링크 (보기 전용 공유 — 실패해도 진행). 직접 입력 + PDF면 여기서 보관
    if (!req.path && req.content) { const b = aiBytes_(req); if (!b.error) req.path = b.path; }
    let pdfUrl = '';
    if (req.path) {
      try {
        const f = getFile_(normPath_(req.path));
        if (f) { try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {} pdfUrl = f.getUrl();
                 let m = {}; try { m = JSON.parse(f.getDescription() || '{}'); } catch (_) {} m.aiDone = 1; f.setDescription(JSON.stringify(m)); }
      } catch (_) {}
    }
    // 그림문자: 시트에 이미 있는 같은 이름 그림의 주소 재사용
    const iconMap = {};
    if (C.icon >= 0) sheet.getRange(2, C.icon + 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().forEach(r => String(r[0] || '').split(',').forEach(part => {
      const pts = part.trim().split('|'); if (pts[1] && !iconMap[pts[1].trim()]) iconMap[pts[1].trim()] = pts[0].trim();
    }));
    let icons = (p.pictograms || []).map(x => String(x).trim()).filter(x => iconMap[x]).map(x => iconMap[x] + '|' + x).join(',');
    if (revise) {   // 개정: 비어 있는 값은 기존 값 유지
      if (!icons) icons = old.product.iconUrl;
      if (!pdfUrl) pdfUrl = old.product.pdfUrl;
      if (!(d.departments || []).length) d.departments = old.product.department.split(/[,\/]/).map(x => x.trim()).filter(Boolean);
    }
    const em = p.emergency || {};
    let emText = [em.eye || em.skin ? '[피부,눈] ' + [em.eye, em.skin].filter(Boolean).join(' ') : '', em.inhalation ? '[흡입] ' + em.inhalation : '',
                    em.ingestion ? '[먹었을때] ' + em.ingestion : ''].filter(Boolean).join('\n');
    if (revise && !emText) emText = old.product.emergency;
    const width = head.length;
    const out = comps.map((c, i) => {
      const r = new Array(width).fill('');
      const set = (k, v) => { if (C[k] >= 0) r[C[k]] = safeCell_(String(v == null ? '' : v)); };
      if (i === 0) {
        set('name', name); set('mf', p.manufacturer); set('dept', (d.departments || []).join(',')); set('usage', p.usage);
        set('tg', d.targetManage ? 'o' : ''); set('sp', d.specialManage ? 'o' : ''); set('em', emText); set('pdf', pdfUrl); set('icon', icons);
      }
      set('comp', c.name); set('cas', String(c.cas || '').trim()); set('amt', c.amount); set('we', c.we); set('hc', c.hc); set('memo', c.memo); set('memo2', c.memo2);
      return r;
    });
    let start = sheet.getLastRow() + 1;
    if (revise) {
      // 이전 줄 보관 → 묶음 크기 맞춘 뒤 덮어쓰기
      let hist = SS.getSheetByName('개정이력');
      if (!hist) { hist = SS.insertSheet('개정이력'); hist.appendRow(['개정일시', '제품명', '새 제품명', '원본 PDF'].concat(head)); hist.setFrozenRows(1); }
      const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
      hist.getRange(hist.getLastRow() + 1, 1, old.raw.length, 4 + width).setValues(
        old.raw.map((r, i) => [i ? '' : now, i ? '' : target, i ? '' : name, i ? '' : pdfUrl].concat(r.map(x => x instanceof Date ? fmtDate_(x) : safeCell_(x)))));
      start = old.row;
      if (out.length > old.count) sheet.insertRowsAfter(old.row + old.count - 1, out.length - old.count);
      if (out.length < old.count) sheet.deleteRows(old.row + out.length, old.count - out.length);
    }
    [C.cas, C.amt].forEach(ci => { if (ci >= 0) sheet.getRange(start, ci + 1, out.length, 1).setNumberFormat('@'); });
    sheet.getRange(start, 1, out.length, width).setValues(out);
    // v20.0: 새 물질 → 사용 부서 교육 대상 (개정은 성분이 새로 들어온 경우만)
    let eduAdded = 0;
    try {
      const depts = d.departments || [];
      if (!revise) eduAdded = newChemAdd_(name, depts, 'AI MSDS 검토 승인', '');
      else {
        const oldKeys = old.components.map(c => casCheck_(c.cas) ? c.cas : normChem_(c.name));
        const added = comps.filter(c => oldKeys.indexOf(casCheck_(String(c.cas || '').trim()) ? String(c.cas).trim() : normChem_(c.name)) < 0).map(c => c.name);
        if (added.length) eduAdded = newChemAdd_(name, depts, 'MSDS 개정 승인', '성분 추가: ' + added.join(', '));
      }
    } catch (_) {}
    return { success: true, eduAdded: eduAdded, mode: revise ? 'revise' : 'new', row: start, rows: out.length, oldRows: old ? old.count : 0, pdfUrl: pdfUrl, icons: icons ? icons.split(',').length : 0 };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════
// 교육이력대장 중복 정리: 같은 사람(사번, 없으면 성명) · 교육구분 · 부서 · 시작일 · 종료일
//  · 남길 줄: 서명완료 우선 → 먼저 제출한 줄 / 나머지는 관리자가 확인한 것만 삭제
// ════════════════════════════════════════════════════════════════
function eduDupes_() {
  const sh = SS.getSheetByName('교육이력대장');
  if (!sh || sh.getLastRow() < 3) return { groups: [], total: 0 };
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0].map(x => String(x).trim());
  const c = k => h.indexOf(k);
  const C = { sub: c('제출일시'), type: c('교육구분'), dept: c('교육부서'), ds: c('교육시작일'), de: c('교육종료일'), no: c('사번'), name: c('성명'), sig: c('서명여부'), hrs: c('총교육시간') };
  if (C.type < 0 || C.name < 0) return { groups: [], total: v.length - 1, error: '교육이력대장 제목 줄을 찾지 못했습니다' };
  const g = {};
  v.slice(1).forEach((r, i) => {
    const name = String(r[C.name] || '').trim();
    if (!name) return;
    const no = C.no >= 0 ? String(r[C.no] || '').trim() : '';
    const who = /^\d{6,}$/.test(no) ? 'no:' + no : 'nm:' + name;
    const k = [who, r[C.type], r[C.dept], r[C.ds], r[C.de]].map(x => String(x).trim()).join('|');
    (g[k] = g[k] || []).push({ row: i + 2, sub: String(r[C.sub] || ''), signed: /서명완료/.test(String(r[C.sig] || '')),
      name: name, no: no, type: r[C.type], dept: r[C.dept], ds: r[C.ds], de: r[C.de], hrs: C.hrs >= 0 ? r[C.hrs] : '' });
  });
  const groups = Object.keys(g).filter(k => g[k].length > 1).map(k => {
    const rows = g[k].slice().sort((a, b) => (b.signed - a.signed) || a.sub.localeCompare(b.sub));
    return { key: k, keep: rows[0], remove: rows.slice(1) };
  }).sort((a, b) => a.keep.row - b.keep.row);
  return { groups: groups, total: v.length - 1, removeCount: groups.reduce((n, x) => n + x.remove.length, 0) };
}
// ════════════════════════════════════════════════════════════════
// 교육 보완 (v20.0): 이수 현황 · 교육 제외 · 새 물질 교육
//  · 대상: 재직자 중 화학물질 사용 부서 직원 — 기본은 모두 대상
//  · 제외: 관리자 또는 부서장이 '취급 안 함'으로 확인한 사람만 (교육제외 시트, 신규입사 '아니오' 회신 포함)
//          단, 교육일지를 작성한 사람은 제외 확인이 있어도 대상으로 봄
//  · 특별교육: 현장 교육 시간 + 공통 온라인 8시간 ≥ 16시간이면 이수, 남은 시간은 첫 교육일부터 3개월 안에
//  · 새 물질: AI MSDS 검토에서 새 제품 · 성분을 승인하면 사용 부서에 교육 대상 생성 (기한 14일)
// ════════════════════════════════════════════════════════════════
const EXCL_SHEET = '교육제외', EXCL_HEADER = ['사번', '이름', '부서', '사유', '확인자', '확인일', '경로', '해제일'];
const NEWCHEM_SHEET = '새물질교육', NEWCHEM_HEADER = ['등록일', '물질', '부서', '기한', '출처', '비고'];
const SPECIAL_HOURS = 16, SPECIAL_ONLINE = 8, SPECIAL_MONTHS = 3, NEWCHEM_DAYS = 14, NEWHIRE_WINDOW = 90;

function eduSheet_(name, header) {
  let sh = SS.getSheetByName(name);
  if (!sh) { sh = SS.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function addMonths_(ymd, n) {
  const d = new Date(ymd + 'T00:00:00');
  d.setMonth(d.getMonth() + n);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
const eduKey_ = (no, name, dept) => /^\d{6,}$/.test(String(no || '').trim()) ? 'no:' + String(no).trim() : 'nm:' + String(name || '').trim() + '|' + normDept_(dept);
const normChem_ = x => String(x || '').toLowerCase().replace(/[\s()（）\[\]_\-.,·%]/g, '');

function exclMap_() {
  const out = {};
  const sh = SS.getSheetByName(EXCL_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, EXCL_HEADER.length).getDisplayValues().forEach((r, i) => {
      if (!String(r[1]).trim() || String(r[7]).trim()) return;   // 해제된 줄은 무시
      const o = { row: i + 2, reason: r[3], by: r[4], date: r[5], route: r[6] };
      out[eduKey_(r[0], r[1], r[2])] = o;
      out['nm:' + String(r[1]).trim() + '|' + normDept_(r[2])] = o;
    });
  }
  try {   // 신규입사 알림에서 '취급 안 함'으로 회신한 직원
    nhRows_().forEach(o => {
      if (o['상태'] !== '미취급') return;
      const x = { reason: '신규입사 회신: 취급 안 함', by: o['부서'] + ' 부서장', date: String(o['회신일시'] || '').slice(0, 10), route: '신규입사 회신' };
      const k1 = eduKey_(o['사번'], o['이름'], o['부서']), k2 = 'nm:' + o['이름'] + '|' + normDept_(o['부서']);
      if (!out[k1]) out[k1] = x;
      if (!out[k2]) out[k2] = x;
    });
  } catch (_) {}
  return out;
}
function ledgerRows_() {
  const sh = SS.getSheetByName('교육이력대장');
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0].map(x => String(x).trim());
  const c = k => h.indexOf(k);
  const C = { type: c('교육구분'), dept: c('교육부서'), chem: c('취급물질'), ds: c('교육시작일'), de: c('교육종료일'), hrs: c('총교육시간'), no: c('사번'), name: c('성명') };
  if (C.name < 0 || C.type < 0) return [];
  return v.slice(1).filter(r => String(r[C.name] || '').trim()).map(r => ({
    no: C.no >= 0 ? String(r[C.no] || '').trim() : '', name: String(r[C.name]).trim(), dept: C.dept >= 0 ? String(r[C.dept] || '').trim() : '',
    type: String(r[C.type] || ''), special: /특별/.test(String(r[C.type] || '')), chem: C.chem >= 0 ? String(r[C.chem] || '') : '',
    ds: C.ds >= 0 ? String(r[C.ds] || '').trim() : '', de: C.de >= 0 ? String(r[C.de] || '').trim() : '',
    hrs: (() => { const m = String(C.hrs >= 0 ? r[C.hrs] : '').match(/(\d+(?:\.\d+)?)/); return m ? Number(m[1]) : 0; })()
  }));
}
function newChemRows_() {
  const sh = SS.getSheetByName(NEWCHEM_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, NEWCHEM_HEADER.length).getDisplayValues()
    .filter(r => String(r[1]).trim() && String(r[2]).trim())
    .map(r => ({ date: String(r[0]).trim(), chem: String(r[1]).trim(), dept: String(r[2]).trim(), due: String(r[3]).trim() || addDays_(String(r[0]).trim(), NEWCHEM_DAYS), src: r[4] }));
}
function newChemAdd_(chem, depts, src, note) {
  const sh = eduSheet_(NEWCHEM_SHEET, NEWCHEM_HEADER);
  const today = todayKst_();
  const have = newChemRows_().map(x => normChem_(x.chem) + '|' + normDept_(x.dept) + '|' + x.date);
  const rows = (depts || []).map(d => String(d).trim()).filter(Boolean)
    .filter(d => have.indexOf(normChem_(chem) + '|' + normDept_(d) + '|' + today) < 0)
    .map(d => [today, chem, d, addDays_(today, NEWCHEM_DAYS), src || '', note || ''].map(safeCell_));
  if (rows.length) { sh.getRange(sh.getLastRow() + 1, 1, rows.length, NEWCHEM_HEADER.length).setValues(rows); eduCacheClear_(); }
  return rows.length;
}

function eduStatus_(s) {
  const today = todayKst_();
  const map = handlingDepts_();
  const chems = getChemicals().data || [];
  const yes = c => c.targetManage || c.specialManage;
  const deptInfo = {};
  const info = dept => {
    const k = normDept_(dept);
    if (deptInfo[k] !== undefined) return deptInfo[k];
    const names = chemsForDept_(map, dept);
    if (!names) return (deptInfo[k] = null);
    const special = chems.some(c => names.indexOf(c.name) >= 0 && yes(c));
    return (deptInfo[k] = { need: special ? '특별' : 'MSDS', chems: names });
  };
  const excl = exclMap_();
  const recs = ledgerRows_();
  const byNo = {}, byName = {};
  recs.forEach(r => {
    if (/^\d{6,}$/.test(r.no)) (byNo[r.no] = byNo[r.no] || []).push(r);
    (byName[r.name] = byName[r.name] || []).push(r);
  });
  const newChems = newChemRows_();
  const emps = (readEmployees_() || []).filter(e => !isResigned_(e) && !(e.leaveDate && e.leaveDate < today) && e.dept)
    .map(e => Object.assign({}, e, { dept: deptAlias_(e.dept) }));
  const people = [];
  emps.forEach(e => {
    if (s.role === 'dept' && !sameDept_(e.dept, s.dept)) return;
    const di = info(e.dept);
    if (!di) return;
    let mine = /^\d{6,}$/.test(e.empNo) ? (byNo[e.empNo] || []) : [];
    if (!mine.length) mine = (byName[e.name] || []).filter(r => !r.no || !/^\d{6,}$/.test(r.no) || r.no === e.empNo).filter(r => sameDept_(r.dept, e.dept) || /^\d{6,}$/.test(e.empNo) === false);
    const ex = excl[eduKey_(e.empNo, e.name, e.dept)] || excl['nm:' + e.name + '|' + normDept_(e.dept)];
    const p = { no: e.empNo, name: e.name, dept: e.dept, joinDate: e.joinDate, need: di.need, records: mine.length,
                lastEdu: mine.map(r => r.ds).sort().pop() || '' };
    if (ex && !mine.length) { p.status = '제외'; p.excluded = ex; people.push(p); return; }
    const isNew = e.joinDate && e.joinDate >= addDays_(today, -NEWHIRE_WINDOW);
    if (di.need === '특별') {
      const sp = mine.filter(r => r.special);
      const seen = {};
      let h = 0;
      sp.forEach(r => { const k = r.ds + '|' + r.de; if (!seen[k]) { seen[k] = 1; h += r.hrs; } });
      const first = sp.map(r => r.ds).filter(Boolean).sort()[0] || '';
      p.hours = sp.length ? h + SPECIAL_ONLINE : 0;
      if (!sp.length) { p.status = '미이수'; p.deadline = isNew ? e.joinDate : ''; }
      else if (p.hours >= SPECIAL_HOURS) p.status = '이수';
      else { p.status = '진행중'; p.remain = SPECIAL_HOURS - p.hours; p.deadline = addMonths_(first, SPECIAL_MONTHS); }
    } else {
      if (mine.length) p.status = '이수';
      else { p.status = '미이수'; p.deadline = isNew ? addDays_(e.joinDate, EDU_PERIOD_DAYS) : ''; }
    }
    if (p.deadline && p.status !== '이수' && today > p.deadline) p.overdue = true;
    // 새 물질 교육
    p.newChem = newChems.filter(n => (sameDept_(n.dept, e.dept) || (n.dept === '병동' && (WARDS.some(w => sameDept_(w, e.dept)) || /병동/.test(e.dept))))
                                   && !(e.joinDate && e.joinDate >= n.date))
      .map(n => {
        const done = mine.some(r => r.ds >= n.date && normChem_(r.chem).indexOf(normChem_(n.chem)) >= 0);
        return { chem: n.chem, date: n.date, due: n.due, done: done, overdue: !done && today > n.due };
      });
    if (p.newChem.some(n => !n.done)) p.newChemPending = true;
    people.push(p);
  });
  const depts = {};
  people.forEach(p => {
    const d = depts[p.dept] = depts[p.dept] || { dept: p.dept, need: p.need, total: 0, done: 0, progress: 0, missing: 0, overdue: 0, excluded: 0, newChem: 0 };
    if (p.status === '제외') { d.excluded++; return; }
    d.total++;
    if (p.status === '이수') d.done++; else if (p.status === '진행중') d.progress++; else d.missing++;
    if (p.overdue) d.overdue++;
    if (p.newChemPending) d.newChem++;
  });
  return { success: true, today: today, people: people, depts: Object.keys(depts).sort().map(k => depts[k]),
           rules: { specialHours: SPECIAL_HOURS, online: SPECIAL_ONLINE, months: SPECIAL_MONTHS, newChemDays: NEWCHEM_DAYS } };
}

function eduExclude_(req, s) {
  const name = String(req.name || '').trim(), dept = String(req.dept || '').trim(), no = String(req.no || '').trim();
  if (!name || !dept) return { success: false, error: '직원을 고르세요' };
  if (s.role === 'dept' && !sameDept_(dept, s.dept)) return { success: false, status: 403, error: '다른 부서 직원입니다' };
  const emp = (readEmployees_() || []).map(e => Object.assign({}, e, { dept: deptAlias_(e.dept) }))
    .find(e => e.name === name && sameDept_(e.dept, dept) && (!no || !e.empNo || e.empNo === no));
  if (!emp) return { success: false, error: '직원목록에서 찾지 못했습니다' };
  const sh = eduSheet_(EXCL_SHEET, EXCL_HEADER);
  const today = todayKst_();
  const by = s.role === 'admin' ? '보건관리자' : (deptHead_(s.dept) || s.dept + ' 부서장');
  const route = s.role === 'admin' ? '관리자' : (req.fromMail ? '부서장(메일)' : '부서장(프로그램)');
  if (req.undo) {
    const ex = exclMap_()[eduKey_(emp.empNo, emp.name, emp.dept)];
    if (!ex || !ex.row) return { success: false, error: '제외 기록이 없습니다 (신규입사 회신 제외는 직원정보관리에서 바꾸세요)' };
    sh.getRange(ex.row, 8).setValue(today + ' ' + by);
    return { success: true, undone: true };
  }
  const reason = String(req.reason || '').trim();
  if (!reason) return { success: false, error: '제외 사유를 적어 주세요 (예: 행정 업무만 · 물질 취급 안 함)' };
  sh.appendRow([emp.empNo || '', emp.name, emp.dept, reason, by, today, route, ''].map(safeCell_));
  return { success: true };
}

function eduAskMail_(req) {
  const st = eduStatus_({ role: 'admin' });
  const cfg = sysCfg_();
  const emails = deptEmails_();
  try {   // '부서장' 시트의 이메일도 사용
    const hs = SS.getSheetByName('부서장');
    if (hs && hs.getLastRow() >= 2) {
      const h = hs.getRange(1, 1, 1, hs.getLastColumn()).getValues()[0].map(x => String(x).trim());
      const ei = h.indexOf('이메일');
      if (ei >= 0) hs.getRange(2, 1, hs.getLastRow() - 1, hs.getLastColumn()).getDisplayValues().forEach(r => { if (String(r[ei]).trim()) emails[String(r[0]).trim()] = String(r[ei]).trim(); });
    }
  } catch (_) {}
  const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const want = (req.depts || []).map(String);
  let sent = 0;
  const noMail = [];
  st.depts.forEach(d => {
    if (want.length && want.indexOf(d.dept) < 0) return;
    const todo = st.people.filter(p => p.dept === d.dept && p.status !== '제외' && (p.status !== '이수' || p.newChemPending));
    if (!todo.length) return;
    const to = Object.keys(emails).filter(k => sameDept_(k, d.dept)).map(k => emails[k]).find(Boolean);
    if (!to) { noMail.push(d.dept); return; }
    const link = cfg.appUrl ? cfg.appUrl + (cfg.appUrl.indexOf('?') >= 0 ? '&' : '?') + 'dv=edu' : '';
    const rows = todo.map(p => '<tr><td>' + esc(p.name) + '</td><td>' + esc(p.need === '특별' ? '특별안전보건교육' : 'MSDS 교육') + '</td><td>' +
      esc(p.status === '진행중' ? '진행중 (남은 ' + p.remain + '시간, 기한 ' + p.deadline + ')' : (p.status === '이수' ? '새 물질 교육 필요' : '미이수' + (p.deadline ? ' (기한 ' + p.deadline + ')' : ''))) + '</td></tr>').join('');
    const html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.6"><p>' + esc(d.dept) + ' 부서장님, 유해화학물질 교육 대상자 확인을 부탁드립니다.</p>' +
      '<table style="border-collapse:collapse" cellpadding="6" border="1"><tr><th>이름</th><th>교육</th><th>상태</th></tr>' + rows + '</table>' +
      '<p>위 직원 중 <b>화학물질을 취급하지 않는 직원</b>은 앱에서 [제외]로 확인해 주세요. 확인하지 않으면 교육 대상으로 봅니다.</p>' +
      (link ? '<p><a href="' + esc(link) + '" style="display:inline-block;background:#2D581B;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none">확인하기</a></p>' : '') +
      '<p style="color:#64748b;font-size:12px">부민병원 보건관리실</p></div>';
    const opts = { to: to, subject: '[보건관리] 화학물질 교육 대상자 확인 요청 - ' + d.dept, htmlBody: html, name: '부민병원 보건관리' };
    if (cfg.adminEmail) opts.replyTo = cfg.adminEmail;
    MailApp.sendEmail(opts);
    sent++;
  });
  return { success: true, sent: sent, noMail: noMail };
}

function eduCacheKey_(s) { return 'edu:' + (s.role === 'admin' ? 'admin' : sha256Hex_(normDept_(s.dept)).slice(0, 16)); }
function eduCacheClear_() { try { CacheService.getScriptCache().put('edu:ver', String(Date.now()), 21600); } catch (_) {} }
function eduStatusCached_(s) {
  const c = CacheService.getScriptCache();
  const ver = c.get('edu:ver') || '0';
  const k = eduCacheKey_(s) + ':' + ver;
  const hit = c.get(k);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  const r = eduStatus_(s);
  const txt = JSON.stringify(r);
  if (txt.length < 95000) c.put(k, txt, 60);
  return r;
}
function eduOp_(req, s) {
  if (req.op === 'edu.status') return requireLogin_(s) || (req.fresh ? eduStatus_(s) : eduStatusCached_(s));
  if (req.op === 'edu.exclude') { const d = requireLogin_(s); if (d) return d; const r = eduExclude_(req, s); if (r.success) eduCacheClear_(); return r; }
  const deny = requireAdmin_(s);
  if (deny) return deny;
  if (req.op === 'edu.askMail') return eduAskMail_(req);
  if (req.op === 'edu.newChem') {
    const n = newChemAdd_(String(req.chem || '').trim(), req.depts || [], '관리자 등록', String(req.note || ''));
    if (n) eduCacheClear_();
    return n ? { success: true, added: n } : { success: false, error: '물질명과 부서를 확인하세요 (같은 날 같은 등록은 한 번만)' };
  }
  if (req.op === 'edu.dupes') return Object.assign({ success: true }, eduDupes_());
  if (req.op === 'edu.dedupe') {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return { success: false, error: '서버가 바쁩니다' };
    try {
      // 지금 다시 계산한 '지울 줄' 중에서 관리자가 고른 것만 (그 사이 시트가 바뀌어도 안전)
      const allowed = {};
      eduDupes_().groups.forEach(gp => gp.remove.forEach(r => { allowed[r.row] = r.sub + '|' + r.name; }));
      const want = (req.rows || []).filter(x => allowed[x.row] && allowed[x.row] === x.sub + '|' + x.name).map(x => x.row);
      const sh = SS.getSheetByName('교육이력대장');
      const uniq = [...new Set(want)].sort((a, b) => b - a);   // 아래 줄부터 지워야 행 번호가 안 밀림
      uniq.forEach(r => sh.deleteRow(r));
      if (uniq.length) eduCacheClear_();
      return { success: true, deleted: uniq.length, skipped: (req.rows || []).length - uniq.length };
    } finally {
      lock.releaseLock();
    }
  }
  return { success: false, error: 'unknown op: ' + req.op };
}

function chemOp_(req, s) {
  const deny = requireAdmin_(s);
  if (deny) return deny;
  switch (req.op) {
    case 'chem.audit': return chemAudit_();
    case 'chem.seedLaw': return chemSeedLaw_();
    case 'chem.apply': return chemApply_(req);
    case 'chem.lawImport': return chemLawImport_(req);
  }
  return { success: false, error: 'unknown op: ' + req.op };
}
