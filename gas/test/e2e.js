// End-to-end check of v18 BMAUTH proxy + Code.gs v2 against the mock GAS server.
const { chromium } = require('playwright');
const { start } = require('./server');

const path = require('path');
const vm = require('vm');
const CODE = path.join(__dirname, '..', 'Code.gs');
const HTML = path.join(__dirname, '..', '..', '부민병원_MSDS_통합시스템_v5 - 2026-04-17T095042.287.html');
const PORT = 8787;
const MOCK = `http://localhost:${PORT}/script.google.com/macros/s/x/exec`;

let failures = 0;
function g_lawHas(g, cas) {
  const sh = g.sheets['법령대조표'];
  return !!sh && sh.rows.some(r => String(r[0]) === cas);
}
function check(name, cond, extra) {
  console.log((cond ? 'PASS ' : 'FAIL ') + name + (extra !== undefined ? '  → ' + JSON.stringify(extra) : ''));
  if (!cond) failures++;
}

(async () => {
  const { server, g } = await start(CODE, HTML, PORT);
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  await ctx.route('**/*', async route => {
    const u = route.request().url();
    if (u.startsWith('https://script.google.com/macros/')) {
      const q = u.indexOf('?') >= 0 ? u.slice(u.indexOf('?')) : '';
      const resp = await route.fetch({ url: MOCK + q });
      return route.fulfill({ response: resp, headers: { ...resp.headers(), 'access-control-allow-origin': '*' } });
    }
    if (u.startsWith(`http://localhost:${PORT}/`)) return route.continue();
    if (u.startsWith('https://api.github.com') || u.startsWith('https://raw.githubusercontent.com')) {
      check('GitHub로 직접 나가는 요청 없음', false, u);
      return route.abort();
    }
    return route.abort();   // CDN 등 외부 리소스 (테스트 환경에서 차단됨)
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('dialog', d => d.accept());

  await page.goto(`http://localhost:${PORT}/index.html`);
  await page.waitForTimeout(2500);

  // 1. 비로그인 상태
  check('비로그인: 배지 표시', (await page.textContent('#bm-auth-badge')).includes('부서 로그인'));
  check('비로그인: 실제 GitHub 토큰 없음', await page.evaluate(() => CFG.ghToken === 'gas-proxy' && localStorage.getItem('bm_ghToken') === 'gas-proxy'));
  check('비로그인: 직원목록 미로드', await page.evaluate(() => empList.length === 0));
  const anonEmp = await page.evaluate(async () => (await fetch('https://script.google.com/macros/s/X/exec?action=emp&callback=cb')).text());
  check('비로그인: 서버가 직원목록 거부', anonEmp.includes('로그인 필요'), anonEmp.slice(0, 80));
  const anonFs = await page.evaluate(async () => (await fetch('https://api.github.com/repos/buminhealth/health/contents/records/edu_records.json')).status);
  check('비로그인: 기록 파일 401', anonFs === 401, anonFs);

  // 1-1. v19.1 MSDS 관리 화면: 목록 · 검색 · 부서 선택 · 상세 · 교육일지 종류 판정 · 휴대폰 전환
  const ms = await page.evaluate(() => {
    // 시험용 물질 2개를 잠시 추가: 관리대상 부서 / 일반 부서
    const fx = [
      { name: '시험물질A', manufacturer: '가제조', usage: '소독', targetManage: '해당', specialManage: '', departments: ['시험특별부'], components: [{ compName: '성분A', casNumber: '7722-84-1', amount: '50' }], icons: [], emergency: '[눈] 물로 씻으시오' },
      { name: '시험물질B', manufacturer: '나제조', usage: '세척', targetManage: false, specialManage: false, departments: ['시험일반부'], components: [], icons: [], emergency: '' },
    ];
    fx.forEach(c => { CHEMS.push(c); productMap[c.name] = c; });
    renderList('all');
    const n = document.querySelectorAll('#gridContainer .ms-li').length;
    const cnt = document.getElementById('ms-cnt').textContent;
    const inList = !!document.querySelector('#productList #deptSelect');
    showDetail('시험물질A');
    const k1 = document.querySelector('.ms-kind').textContent, b1 = document.querySelector('.ms-btn.p').getAttribute('onclick');
    const h2 = document.querySelector('.ms-dh h2').textContent === '시험물질A';
    const em = !!document.querySelector('.ms-et');
    showDetail('시험물질B');
    const k2 = document.querySelector('.ms-kind').textContent, b2 = document.querySelector('.ms-btn.p').getAttribute('onclick');
    searchChem('7722-84'); const nS = document.querySelectorAll('#gridContainer .ms-li').length;
    searchChem(''); filterDept('시험일반부'); const nD = document.querySelectorAll('#gridContainer .ms-li').length;
    filterDept('all');
    return { n, total: CHEMS.length, cnt, inList, k1, b1, h2, em, k2, b2, nS, nD };
  });
  check('MSDS: 전체 물질 목록과 개수', ms.n === ms.total && ms.cnt.includes(String(ms.total)), ms);
  check('MSDS: 부서 선택은 목록 안', ms.inList);
  check('MSDS: CAS 검색 · 부서 선택으로 줄어듦', ms.nS === 1 && ms.nD === 1 && ms.n >= 3, ms);
  check('MSDS: 응급처치 표', ms.em);
  check('MSDS: 상세 제목', ms.h2);
  check('MSDS: 관리대상 물질 부서 → 특별안전보건교육일지', ms.k1.includes('특별안전보건교육') && ms.b1.startsWith('startEdu1'), ms.k1);
  check('MSDS: 관리대상 없는 부서 → MSDS 교육일지', ms.k2.includes('MSDS 교육') && ms.b2.startsWith('startEdu2'), ms.k2);
  await page.setViewportSize({ width: 390, height: 844 });
  const mob = await page.evaluate(() => {
    const vis = id => getComputedStyle(document.getElementById(id)).display !== 'none';
    msdsBack(); const a = [vis('productList'), vis('detailView')];
    document.querySelector('#gridContainer .ms-li').click(); const b = [vis('productList'), vis('detailView')];
    msdsBack(); const c = [vis('productList'), vis('detailView')];
    const sw = document.documentElement.scrollWidth;
    ['시험물질A', '시험물질B'].forEach(nm => { CHEMS.splice(CHEMS.findIndex(x => x.name === nm), 1); delete productMap[nm]; });
    curChem = null; renderList('all');
    return { a, b, c, sw };
  });
  check('휴대폰: 목록 → 누르면 상세 → 목록으로', JSON.stringify(mob.a) === '[true,false]' && JSON.stringify(mob.b) === '[false,true]' && JSON.stringify(mob.c) === '[true,false]', mob);
  check('휴대폰: 가로 넘침 없음', mob.sw <= 390, mob.sw);
  await page.setViewportSize({ width: 1366, height: 768 });

  // 2. 작성 탭 → 로그인 창 → 틀린 PIN → 맞는 PIN
  await page.evaluate(() => goTab('edu2'));
  await page.waitForSelector('#bm-login', { state: 'visible' });
  await page.waitForFunction(() => document.querySelectorAll('#bm-login-dept option').length > 1);
  const deptOptions = await page.$$eval('#bm-login-dept option', os => os.map(o => o.value));
  check('로그인 창: PIN 등록된 부서만 표시', JSON.stringify(deptOptions) === JSON.stringify(['', '61병동', '71병동']), deptOptions);
  await page.selectOption('#bm-login-dept', '61병동');
  await page.fill('#bm-login-pin', '0000');
  await page.click('#bm-login-go');
  await page.waitForFunction(() => document.getElementById('bm-login-err').textContent.length > 0);
  const errText = await page.textContent('#bm-login-err');
  check('틀린 PIN: 남은 횟수 안내', errText.includes('남은 시도 4회'), errText);
  await page.fill('#bm-login-pin', '1234');
  await Promise.all([page.waitForNavigation(), page.click('#bm-login-go')]);
  await page.waitForTimeout(2500);

  // 3. 부서 로그인 후
  check('부서 로그인: 배지', (await page.textContent('#bm-auth-badge')).includes('61병동'));
  const emps = await page.evaluate(() => empList.map(e => e.name + '/' + e.dept + '/' + e.joinDate));
  check('부서 로그인: 자기 부서 직원만', emps.length === 2 && emps.every(x => x.includes('/61병동/')), emps);
  check('직원 입사일 KST 기준 (하루 당김 없음)', emps.includes('김가나/61병동/2026-07-01'), emps);
  await page.evaluate(() => goTab('edu2'));
  const e2dept = await page.evaluate(() => [...document.getElementById('e2-dept').options].map(o => o.value));
  check('부서 로그인: 교육일지 부서 고정', JSON.stringify(e2dept) === '["61병동"]', e2dept);

  // 4. 기록 저장 (자기 부서 + 다른 부서 섞어서 보냄)
  const w1 = await page.evaluate(async () => {
    const recs = [
      { recordId: 'r1|김가나|1|61병동|MSDS교육|2026-07-01', name: '김가나', dept: '61병동', type: 'MSDS교육', dateStart: '2026-07-01' },
      { recordId: 'r2|이다라|2|71병동|MSDS교육|2026-07-01', name: '이다라', dept: '71병동', type: 'MSDS교육', dateStart: '2026-07-01' },
    ];
    const cur = await readGitHubJSON('records/edu_records.json');
    await writeGitHubJSON('records/edu_records.json', recs, cur.sha, 'test');
    return (await readGitHubJSON('records/edu_records.json')).data;
  });
  check('부서 저장: 자기 부서 기록만 저장됨', w1.length === 1 && w1[0].dept === '61병동', w1);

  // 5. 큰 PDF 분할 업로드 (≈6MB base64 → 2조각)
  const big = await page.evaluate(async () => {
    const n = 4.5 * 1024 * 1024;
    const bytes = new Uint8Array(n);
    for (let i = 0; i < n; i++) bytes[i] = (i * 31 + 7) & 0xff;
    let bin = ''; for (let i = 0; i < n; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    const b64 = btoa(bin);
    const url = await uploadPdfToGitHub(b64, 'MSDS교육일지-61병동-김가나-2026-07-01.pdf');
    window.__bigB64 = b64;
    return url;
  });
  check('큰 PDF 분할 업로드 성공', typeof big === 'string' && big.includes('raw.githubusercontent.com'), big);

  // 6. 관리자 로그인 → 전체 조회 + 같은 PDF 분할 다운로드
  const adminBad = await page.evaluate(async () => (await BMAUTH.adminLogin('wrong')).error);
  check('관리자 틀린 비밀번호 거부', /올바르지 않습니다/.test(adminBad || ''), adminBad);
  const adm = await page.evaluate(async () => {
    const r = await BMAUTH.adminLogin('admin-pass-1');
    if (!r.success) return { err: r.error };
    const all = (await readGitHubJSON('records/edu_records.json')).data;
    const res = await fetch(window.__lastPdfUrl || '');
    return { role: BMAUTH.session().role, count: all.length, adminOk: localStorage.getItem('bm_admin_ok') };
  });
  check('관리자 로그인 성공', adm.role === 'admin', adm);
  const same = await page.evaluate(async () => {
    const ym = todayStr().slice(0, 7);
    const res = await fetch(`https://raw.githubusercontent.com/buminhealth/health/main/교육일지/${ym}/MSDS교육일지-61병동-김가나-2026-07-01.pdf?_t=1`);
    const buf = new Uint8Array(await res.arrayBuffer());
    let bin = ''; for (let i = 0; i < buf.length; i += 0x8000) bin += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
    return { status: res.status, equal: btoa(bin) === window.__bigB64, size: buf.length };
  });
  check('큰 PDF 분할 다운로드 = 원본과 동일', same.status === 200 && same.equal, same);
  const sameGit = await page.evaluate(async () => {
    const ym = todayStr().slice(0, 7);
    const res = await fetch(`https://api.github.com/repos/buminhealth/health/contents/교육일지/${ym}`);
    return (await res.json()).map(f => f.name);
  });
  check('폴더 목록 (GitHub 형식)', sameGit.includes('MSDS교육일지-61병동-김가나-2026-07-01.pdf'), sameGit);

  // 7. 다른 부서는 61병동 PDF를 못 받음
  const other = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '71병동', pin: '5678' });
    const ym = todayStr().slice(0, 7);
    const g = await BMAUTH.post('fs.get', { path: `교육일지/${ym}/MSDS교육일지-61병동-김가나-2026-07-01.pdf`, t: r.token });
    const recs = await BMAUTH.post('fs.get', { path: 'records/edu_records.json', t: r.token });
    return { pdf: g.status, recs: JSON.parse(decodeURIComponent(escape(atob(recs.content)))) };
  });
  check('다른 부서: 61병동 PDF 403', other.pdf === 403, other.pdf);
  check('다른 부서: 61병동 기록 안 보임', Array.isArray(other.recs) && other.recs.length === 0, other.recs);

  // 8. JSONP 기록 저장 (토큰 자동 첨부 + 수식 주입 차단) — 관리자 세션
  await page.evaluate(() => new Promise(res => {
    window.__cb = r => { window.__rec = r; res(); };
    const sc = document.createElement('script');
    sc.src = CFG.url + '?action=record&type=MSDS교육&dept=61병동&names=' + encodeURIComponent('=HYPERLINK("x")') + '&callback=__cb';
    document.head.appendChild(sc);
  }));
  const rec = await page.evaluate(() => window.__rec);
  check('JSONP 기록 저장 (토큰 자동 첨부)', rec && rec.success, rec);
  const cellVal = g.sheets['교육이력대장'] && g.sheets['교육이력대장'].rows[1][10];
  check('시트: 수식 주입 차단 (작은따옴표 처리)', cellVal === "'=HYPERLINK(\"x\")", cellVal);

  // 9. PIN 5회 오류 잠금
  const lock = await page.evaluate(async () => {
    let last;
    for (let i = 0; i < 6; i++) last = await BMAUTH.post('login', { kind: 'dept', dept: '71병동', pin: '9999' });
    return last;
  });
  check('PIN 5회 오류 → 잠금', lock.status === 423, lock);

  // 10. PIN 변경 시 기존 토큰 무효
  const revoked = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '1234' });
    window.__t61 = r.token;
    return r.success;
  });
  g.sheets['부서인증'].rows[1][1] = '4321';
  const after = await page.evaluate(async () => (await BMAUTH.post('fs.get', { path: 'records/edu_records.json', t: window.__t61 })).status);
  check('PIN 변경 → 기존 부서 토큰 무효', revoked && after === 401, after);

  // 11. 직원정보관리: 엑셀 붙여넣기 → 업데이트 (관리자) — 수동이므로 메일 없음
  const mail = g.ctx.MailApp;
  const fb = g.ctx.FIREBASE;
  const kstDay = n => new Date(Date.now() + 9 * 3600e3 + n * 86400e3).toISOString().slice(0, 10);
  await page.evaluate(() => goTab('admin'));
  await page.fill('#admin-pw', 'admin-pass-1');
  await page.evaluate(() => adminLogin());   // 실제 관리자 로그인 폼 경로
  await page.waitForSelector('#admin-main', { state: 'visible' });
  await page.click('.adm-nav-item[data-pane="staff"]');
  await page.waitForSelector('#staff-paste');
  check('직원정보관리: 연결 전에는 [지금 자동 확인] 비활성', await page.$eval('#staff-nh-run', b => b.disabled));
  await page.fill('#staff-paste', [
    '사번\t이름\t부서\t입사일\t재직상태',
    '126051302\t김가나\t61병동\t2026-07-01\t재직',
    `126101801\t최신입\t61병동\t${kstDay(-1)}\t재직`,
    '126051303\t이다라\t71병동\t2026.07.02\t휴직',
  ].join('\n'));
  await page.click('#staff-preview');
  await page.check('#staff-full');
  await page.click('#staff-sync');
  await page.waitForFunction(() => document.getElementById('staff-stat').textContent.indexOf('퇴사 1') >= 0, null, { timeout: 15000 });
  const empRows = g.sheets['직원목록'].rows;
  const byName = n => empRows.find(r => r[1] === n);
  const stCol = empRows[0].indexOf('재직상태');
  check('직원 동기화: 행 삭제 없음 (4명 유지)', empRows.length === 5, empRows.length);
  check('직원 동기화: 명단에 없는 박마바 → 퇴사 표시', byName('박마바') && byName('박마바')[stCol] === '퇴사(명단없음)', byName('박마바'));
  check('직원 동기화: 이다라 휴직 반영', byName('이다라')[stCol] === '휴직', byName('이다라')[stCol]);
  check('직원 동기화: 신규 최신입 추가', !!byName('최신입'));
  check('수동 업데이트: 알림 메일 없음', mail.sent.length === 0, mail.sent.length);
  const nhRun = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  check('자동 로그인 미설정: 지금 확인 거부 + 메일 없음', !nhRun.success && mail.sent.length === 0, nhRun.error);
  g.api.dailyCheck();
  check('자동 로그인 미설정: 매일 확인은 아무 메일도 안 보냄', mail.sent.length === 0, mail.sent.length);

  // 12. Firebase 인사 사이트 연결 (메뉴 ④ 설정을 직접 입력)
  Object.assign(g.props, { FB_API_KEY: 'AIzaTest', FB_PROJECT_ID: 'bumin-hr', FB_MODE: 'firestore', FB_PATH: 'employees' });
  fb.users['clover8477@bumin.co.kr'] = 'hr-pass';
  const S = v => ({ stringValue: v });
  fb.docs = [
    { name: 'projects/bumin-hr/databases/(default)/documents/employees/a1', fields: { empNo: S('126051302'), name: S('김가나'), dept: S('61병동'), hireDate: { timestampValue: '2026-06-30T15:00:00Z' }, status: S('재직') } },
    { name: 'projects/bumin-hr/databases/(default)/documents/employees/a2', fields: { empCode: S('126101801'), name: S('최신입'), dept: S('61병동'), hireDate: S(kstDay(-1)), status: S('재직'), rrn: S('900101-1234567'), phone: S('010-0000-0000') } },
    { name: 'projects/bumin-hr/databases/(default)/documents/employees/a3', fields: { empNo: S('126101802'), name: S('한새봄'), dept: S('71병동'), hireDate: S(kstDay(-1)), status: S('재직') } },
  ];
  await page.evaluate(() => BMSTAFF.open());
  await page.waitForSelector('#staff-hr-pw');
  await page.fill('#staff-hr-id', 'clover8477@bumin.co.kr');
  await page.fill('#staff-hr-pw', 'wrong');
  await page.click('#staff-hr-go');
  await page.waitForTimeout(800);
  check('인사 사이트 가져오기: 틀린 비밀번호는 거부', !(await page.$eval('#staff-sync', b => !b.disabled && document.getElementById('staff-preview-out').textContent.includes('총'))));
  await page.fill('#staff-hr-pw', 'hr-pass');
  await page.click('#staff-hr-go');
  await page.waitForFunction(() => document.getElementById('staff-preview-out').textContent.includes('총 3명'), null, { timeout: 15000 });
  check('인사 사이트 가져오기: Firestore 3명 미리보기', true);
  const prevText = await page.textContent('#staff-preview-out');
  check('인사 사이트 가져오기: 주민번호·전화번호 열 제외', !prevText.includes('900101') && !prevText.includes('010-0000') && !/rrn|phone/.test(prevText), prevText.slice(0, 120));
  check('인사 사이트 가져오기: 비밀번호 칸 비움 (저장 안 함)', (await page.$eval('#staff-hr-pw', i => i.value)) === '' && !g.props.HR_PW);
  check('수동 가져오기: 메일 없음', mail.sent.length === 0, mail.sent.length);

  // 13. 자동 로그인 저장 후 [지금 자동 확인] → 이때만 신규입사 메일
  Object.assign(g.props, { HR_ID: 'clover8477@bumin.co.kr', HR_PW: 'hr-pass' });
  const savedFetch = await page.evaluate(() => BMAUTH.post('hr.fetch', {}));
  check('저장된 계정으로 가져오기 (비밀번호 입력 없이)', savedFetch.success && savedFetch.rows.length === 3, savedFetch.error);
  await page.evaluate(() => BMSTAFF.open());
  await page.waitForSelector('#staff-nh-run:not([disabled])');
  await page.click('#staff-nh-run');
  await page.waitForFunction(() => document.querySelectorAll('#staff-nh-out tr').length >= 2, null, { timeout: 15000 });
  const nhMail = mail.sent.find(m => m.to === 'head61@bumin.co.kr');
  check('자동 로그인: 61병동 부서장에게 신규입사 메일', !!nhMail && /최신입/.test(nhMail.subject), mail.sent.map(m => m.to + ' ' + m.subject));
  check('메일: [확인하기] 링크 포함', !!nhMail && /app\.html\?nh=NH/.test(nhMail.htmlBody));
  check('메일: 교육 기한 = 입사일 + 3일', !!nhMail && nhMail.htmlBody.includes(kstDay(2)), kstDay(2));
  check('메일: 오래된 입사자·비취급 부서(71병동)는 제외', mail.sent.length === 1, mail.sent.length);
  const nh0 = g.api.nhRows_()[0];
  check('알림 기록: 교육기한 저장', nh0['교육기한'] === kstDay(2), nh0['교육기한']);
  check('자동 동기화: 한새봄 추가', !!g.sheets['직원목록'].rows.find(r => r[1] === '한새봄'));
  const tableText = await page.textContent('#staff-nh-out');
  check('관리자 화면: 회신대기 + 교육기한 표시', tableText.includes('회신대기') && tableText.includes(kstDay(2)));

  // 13-0. 메뉴 ④: 사이트 주소만 넣으면 연결 정보 자동
  const fc1 = g.api.fbConfigFromInput_('https://bumin-leave.web.app/');
  check('메뉴 ④: 사이트 주소 → 연결 정보 자동 (init.json)', fc1.apiKey === 'AIzaFromSite' && fc1.projectId === 'bumin-leave' && fc1.site === 'https://bumin-leave.web.app', fc1);
  const fc2 = g.api.fbConfigFromInput_('const firebaseConfig = {\n apiKey: "AIzaPaste",\n projectId: "bumin-leave",\n};');
  check('메뉴 ④: firebaseConfig 붙여넣기도 그대로 지원', fc2.apiKey === 'AIzaPaste' && fc2.projectId === 'bumin-leave', fc2);
  let fc3err = '';
  try { g.api.fbConfigFromInput_('https://unknown.example.com'); } catch (e) { fc3err = e.message; }
  check('메뉴 ④: 연결 정보가 없는 주소는 안내', /firebaseConfig를 붙여넣어/.test(fc3err), fc3err);

  // 13-1. 휴직자관리 실제 구조: employees_public(문서 ID=사번, quitDate) + leaves(현재 휴직)
  g.props.FB_PATH = 'employees_public';
  const D2 = (id, f) => ({ name: 'projects/bumin-leave/databases/(default)/documents/employees_public/' + id, fields: f });
  fb.collections = {
    employees_public: [
      D2('126051302', { name: S('김가나'), dept: S('61병동'), jobRank: S('간호사'), joinDate: S('2026-07-01'), quitDate: S('') }),
      D2('126051303', { name: S('이다라'), dept: S('71병동'), joinDate: S('2026-07-02'), quitDate: S('') }),
      D2('120000001', { name: S('정퇴사'), dept: S('61병동'), joinDate: S('2020-01-01'), quitDate: S(kstDay(-5)) }),
    ],
    leaves: [
      { name: 'x/leaves/126051303_2026-09-01', fields: { empCode: S('126051303'), type: S('육아휴직'), start: S('2026-09-01'), end: S(kstDay(30)) } },
    ],
  };
  const r131 = await page.evaluate(() => BMAUTH.post('hr.fetch', { id: 'clover8477@bumin.co.kr', password: 'hr-pass' }));
  const hdr131 = r131.headers || [];
  const rowOf = n => (r131.rows || []).find(r => r.indexOf(n) >= 0) || [];
  const stI = hdr131.indexOf('status');
  check('휴직자관리 구조: 문서 ID를 사번으로 (_key)', hdr131.indexOf('_key') >= 0 && rowOf('김가나')[hdr131.indexOf('_key')] === '126051302', hdr131);
  check('휴직자관리 구조: 퇴사일 지난 직원 → 퇴사', rowOf('정퇴사')[stI] === '퇴사', rowOf('정퇴사'));
  check('휴직자관리 구조: 현재 휴직 → 휴직(육아휴직)', rowOf('이다라')[stI] === '휴직(육아휴직)', rowOf('이다라'));
  check('휴직자관리 구조: 그 외 → 재직', rowOf('김가나')[stI] === '재직');
  fb.denied = ['wrong_name'];
  g.props.FB_PATH = 'wrong_name';
  const r403 = await page.evaluate(() => BMAUTH.post('hr.fetch', { id: 'clover8477@bumin.co.kr', password: 'hr-pass' }));
  check('잘못된 데이터 위치 → 403 사유 안내', !r403.success && /wrong_name/.test(r403.error) && /권한/.test(r403.error), r403.error);
  g.props.FB_PATH = 'employees'; fb.collections = null; fb.denied = null;

  // 13-2. 부서 이름 맞추기
  g.sheets['chemicals'].appendRow(['병동공용소독제', 'B사', '병동', '이소프로판올', '67-63-0', '70']);
  const syncDept = rows => page.evaluate(rs => BMAUTH.post('emp.sync', {
    headers: ['사번', '이름', '부서', '입사일', '재직상태'], rows: rs, fullList: false }), rows);
  const ds1 = await syncDept([
    ['126200001', '띄어쓰기', '61 병동', '2026-01-01', '재직'],
    ['126200002', '매핑대상', '간호부 71병동', '2026-01-01', '재직'],
    ['126200003', '신장실', '인공신장실', '2026-01-01', '재직'],
  ]);
  const empDept = n => { const rr = g.sheets['직원목록'].rows; const h = rr[0]; const r = rr.find(x => x[h.indexOf('이름')] === n); return r && r[h.indexOf('부서')]; };
  check('부서: 띄어쓰기 차이는 앱 이름으로 통일 ("61 병동" → 61병동)', empDept('띄어쓰기') === '61병동', empDept('띄어쓰기'));
  check('부서: 처음 보는 이름은 부서매핑 시트에 추가', ds1.unknownDepts && ds1.unknownDepts.indexOf('간호부 71병동') >= 0 &&
        g.sheets['부서매핑'].rows.some(r => r[0] === '간호부 71병동'), ds1.unknownDepts);
  g.sheets['부서매핑'].rows.find(r => r[0] === '간호부 71병동')[1] = '71병동';
  const ds2 = await syncDept([['126200002', '매핑대상', '간호부 71병동', '2026-01-01', '재직']]);
  check('부서: 매핑 적용 ("간호부 71병동" → 71병동)', empDept('매핑대상') === '71병동' && ds2.unknownDepts.length === 0, empDept('매핑대상'));
  const stx = await page.evaluate(() => BMAUTH.post('emp.status', {}));
  const dc = n => (stx.deptCheck || []).find(d => d.dept === n) || {};
  check("부서 점검: 인공신장실도 chemicals '병동' 물질 적용", dc('인공신장실').chems >= 1, dc('인공신장실'));
  check('부서 점검: 61병동 로그인 ✓ · 메일 ✓', dc('61병동').login && dc('61병동').email, dc('61병동'));
  check("부서 점검: 61병동 물질 = 전용 1 + '병동' 공통 1", dc('61병동').chems === 2, dc('61병동'));
  check('부서 점검: 인공신장실 로그인 ✕ (PIN 없음)', dc('인공신장실').login === false);
  await page.evaluate(() => BMSTAFF.open());
  await page.waitForFunction(() => document.getElementById('staff-body').textContent.includes('④ 부서 이름 점검'), null, { timeout: 15000 });
  check('관리자 화면: 부서 이름 점검표 표시', (await page.textContent('#staff-body')).includes('PIN 없음'));
  check('관리자 화면: 서버 버전 표시', (await page.textContent('#staff-body')).includes('서버 버전: v18.7'));

  // 13-3. 인사 사이트 이상 감지 + 실패 메일 (v18.6)
  g.props.FB_PATH = 'employees_public';
  const mk = (id, f) => ({ name: 'x/employees_public/' + id, fields: f });
  const many = [];
  for (let i = 0; i < 20; i++) many.push(mk('1300000' + String(i).padStart(2, '0'), { name: S('직원' + i), dept: S('61병동'), joinDate: S('2020-01-01'), quitDate: S('') }));
  fb.collections = { employees_public: many, leaves: [] };
  const okRun = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  check('정상 데이터: 자동 확인 성공·기준 인원 기록', okRun.success && !okRun.error, okRun.error);
  const cfgVal = k => { const r = g.sheets['관리자설정'].rows.find(x => x[0] === k); return r ? r[1] : ''; };
  check('정상 데이터: 기준 인원 20명 기록', cfgVal('hr_last_count') === '20', cfgVal('hr_last_count'));
  const empBefore = g.sheets['직원목록'].rows.length;
  const mailBefore = mail.sent.length;
  fb.collections.employees_public = many.slice(0, 5);   // 인원 급감
  const dropRun = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  check('이상 감지: 인원 급감 → 반영 중단', !!dropRun.error && /크게 줄었습니다/.test(dropRun.error) && g.sheets['직원목록'].rows.length === empBefore, dropRun.error);
  const failMail = mail.sent.slice(mailBefore).filter(m => m.to === 'admin@bumin.co.kr' && /자동 확인 실패/.test(m.subject));
  check('실패 시 보건관리자 메일 1통', failMail.length === 1, mail.sent.slice(mailBefore).map(m => m.subject));
  await page.evaluate(() => BMAUTH.post('nh.run', {}));
  check('같은 날 같은 실패는 메일 다시 안 보냄', mail.sent.slice(mailBefore).filter(m => /자동 확인 실패/.test(m.subject)).length === 1);
  fb.collections.employees_public = many.map(d => ({ name: d.name, fields: { name: d.fields.name, dept: d.fields.dept, hireDt: d.fields.joinDate } }));   // 필드 이름 변경
  const renRun = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  check('이상 감지: 입사일 필드 이름 변경(빈 비율 0%→100%) → 반영 중단', !!renRun.error && /입사일/.test(renRun.error), renRun.error);
  check('다른 사유의 실패는 새로 메일', mail.sent.slice(mailBefore).filter(m => /자동 확인 실패/.test(m.subject)).length === 2);
  const manual = await page.evaluate(() => BMAUTH.post('hr.fetch', {}));
  check('수동 가져오기: 이상 경고를 함께 반환', manual.success && manual.warnings.length > 0, manual.warnings);
  const st186 = await page.evaluate(() => BMAUTH.post('emp.status', {}));
  check('관리자 화면용 최근 오류에 사유 표시', /입사일/.test(st186.hrLastError), st186.hrLastError);
  // 원래부터 입사일이 대부분 비어 있는 데이터는 오탐하지 않음 (기준값을 지우고 첫 실행처럼)
  g.sheets['관리자설정'].rows = g.sheets['관리자설정'].rows.filter(x => !/^hr_last_(count|ratios)$/.test(x[0]));
  fb.collections.employees_public = many.map((d, i) => ({ name: d.name, fields: i < 15 ? { name: d.fields.name, dept: d.fields.dept }
                                                                                    : { name: d.fields.name, dept: d.fields.dept, joinDate: d.fields.joinDate } }));
  const noFalse1 = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  const noFalse2 = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  check('오탐 방지: 원래 입사일 75%가 빈 데이터도 정상 반영 (첫 실행·다음 실행 모두)', !noFalse1.error && !noFalse2.error, [noFalse1.error, noFalse2.error]);
  fb.collections = null; g.props.FB_PATH = 'employees';

  // 13-4. 물질 마스터 점검 (v18.7): CAS 복원·검증, 빈칸 제안, 법령대조표, 빈 칸에만 반영
  const ch = g.sheets['chemicals'];
  ch.rows[0] = ['제품명', '제조사', '부서', '성분명', 'CAS', '함유량', '특수건강검진', '작업환경측정'];
  ch.appendRow(['효소세정제', 'C사', '중앙공급실', 'Protease', vm.runInContext('new Date(9014, 0, 1)', g.ctx), '1', '', '']);
  ch.appendRow(['메탄올A', 'D사', '병리과', '메탄올', '67-56-1', '99', '12개월', '6개월']);
  ch.appendRow(['메탄올B', 'D사', '병리과', '메탄올', '67-56-1', '50', '', '']);
  ch.appendRow(['정제수제품', 'E사', '약제과', '정제수', '732-18-5', '90', '', '']);
  ch.appendRow(['EDTA용액', 'F사', '진단검사의학과', 'EDTA', '64-02-8.', '5', '', '']);
  const au = await page.evaluate(() => BMAUTH.post('chem.audit', {}));
  const iss = (au.casIssues || []).reduce((m, x) => (m[x.comp] = x, m), {});
  check('물질 점검: 날짜로 바뀐 CAS → 9014-01-1 복원 제안', iss.Protease && iss.Protease.fix === '9014-01-1', iss.Protease);
  check('물질 점검: 끝의 점 → 64-02-8', iss.EDTA && iss.EDTA.fix === '64-02-8', iss.EDTA);
  check('물질 점검: 검증숫자 오류(732-18-5) 표시', iss['정제수'] && !iss['정제수'].fix && /검증숫자/.test(iss['정제수'].why), iss['정제수']);
  const sugB = (au.suggestions || []).filter(x => x.product === '메탄올B');
  check('물질 점검: 같은 CAS 다른 제품 판정으로 빈칸 제안 (특수검진·측정)', sugB.length === 2 && sugB.some(x => x.field === 'hc' && x.value === '12개월'), sugB);
  const seed = await page.evaluate(() => BMAUTH.post('chem.seedLaw', {}));
  check('법령대조표: 기존 판정으로 시작값 생성', seed.success && seed.added >= 1 && g_lawHas(g, '67-56-1'), seed);
  const seed2 = await page.evaluate(() => BMAUTH.post('chem.seedLaw', {}));
  check('법령대조표: 다시 눌러도 중복 추가 없음', seed2.success && seed2.added === 0, seed2);
  const au2 = await page.evaluate(() => BMAUTH.post('chem.audit', {}));
  check('물질 점검: 법령대조표 근거로 제안', au2.lawCount >= 1 && au2.suggestions.some(x => x.product === '메탄올B' && /법령대조표/.test(x.src)), au2.suggestions);
  const rowB = ch.rows.findIndex(r => r[0] === '메탄올B') + 1, rowA = ch.rows.findIndex(r => r[0] === '메탄올A') + 1;
  const rowP = ch.rows.findIndex(r => r[0] === '효소세정제') + 1;
  const ap = await page.evaluate(a => BMAUTH.post('chem.apply', a), {
    fills: [{ row: rowB, field: 'hc', value: '12개월' }, { row: rowA, field: 'hc', value: '6개월' }],
    casFixes: [{ row: rowP, fix: '9014-01-1' }] });
  check('반영: 빈 칸만 채움 · 기존 값은 그대로', ap.success && ap.filled === 1 && ap.skipped === 1 && ch.rows[rowB - 1][6] === '12개월' && ch.rows[rowA - 1][6] === '12개월', ap);
  check('반영: CAS 원래 값으로 복원', ap.fixed === 1 && ch.rows[rowP - 1][4] === '9014-01-1', ch.rows[rowP - 1][4]);
  const denyChem = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' });
    return BMAUTH.post('chem.audit', { t: r.token });
  });
  check('물질 점검: 부서 계정은 거부', !denyChem.success && denyChem.status === 403, denyChem);
  await page.evaluate(() => { admSwitchPane('chemm'); return BMCHEM.open(); });
  check('관리자 화면: 물질 마스터 점검 표시', /CAS 번호 문제/.test(await page.textContent('#chemm-body')));

  // 14. 부서장: 메일 링크(?nh=) → 로그인 창 → PIN 로그인 → 확인 창 → '예' → 교육일지 대상자 자동 추가
  await page.evaluate(() => { localStorage.removeItem('bm_session'); localStorage.removeItem('bm_admin_ok'); });
  await page.goto(`http://localhost:${PORT}/index.html?nh=${nh0['알림ID']}`);
  await page.waitForSelector('#bm-login', { state: 'visible', timeout: 15000 });
  check('링크 열기(비로그인) → 부서 로그인 창', true);
  await page.waitForFunction(() => document.querySelectorAll('#bm-login-dept option').length > 1);
  await page.selectOption('#bm-login-dept', '61병동');
  await page.fill('#bm-login-pin', '4321');
  await Promise.all([page.waitForNavigation(), page.click('#bm-login-go')]);
  await page.waitForSelector('#bm-nh [data-yes]', { timeout: 15000 });
  check('로그인 후 같은 링크 → 확인 요청 창', true);
  check('확인 창: 교육 기한 표시', (await page.textContent('#bm-nh')).includes('교육 기한: ' + kstDay(2)));
  const deptEmps = await page.evaluate(() => empList.map(e => e.name));
  check('부서 직원 검색: 퇴사자 제외', !deptEmps.includes('박마바') && deptEmps.includes('최신입'), deptEmps);
  await page.click('#bm-nh [data-yes]');
  await page.waitForFunction(() => typeof e2 !== 'undefined' && e2.parts.some(p => p.name === '최신입'), null, { timeout: 15000 });
  const part = await page.evaluate(() => e2.parts.find(p => p.name === '최신입'));
  check("'예' → MSDS 교육일지에 대상자 자동 추가", part && part.empNo === '126101801' && part.join === kstDay(-1), part);
  check("'예' → MSDS 교육일지 화면 표시 + 부서 고정", await page.evaluate(() => document.getElementById('pg-edu2').style.display === 'block' && document.getElementById('e2-dept').value === '61병동'));
  check("'예' → 주소의 ?nh 제거", !(await page.evaluate(() => location.search)).includes('nh='));
  check("'예' → 상태 취급-교육대기", g.api.nhRows_()[0]['상태'] === '취급-교육대기', g.api.nhRows_()[0]['상태']);
  check('보건관리자에게 회신 알림 메일', mail.sent.some(m => m.to === 'admin@bumin.co.kr' && /최신입/.test(m.subject)));

  // 14-1. 부서 제출 현황 탭
  check('부서 제출 현황 탭 표시', (await page.textContent('#tab-deptview')).includes('부서현황') && (await page.getAttribute('#tab-deptview','title')).includes('61병동'));
  await page.evaluate(() => goTab('deptview'));
  await page.waitForFunction(() => document.getElementById('dv-body').textContent.includes('교육일지'), null, { timeout: 15000 });
  await page.selectOption('#dv-range', '0');
  await page.waitForFunction(() => document.getElementById('dv-body').textContent.includes('김가나'), null, { timeout: 15000 });
  const dvText = await page.textContent('#dv-body');
  const dvHead = await page.$$eval('#dv-body table', ts => [...ts[0].querySelectorAll('th')].map(t => t.textContent));
  check('교육일지 열 순서: 구분-대상자-입사일-교육일-취급물질-결재-미리보기', dvHead.join(',') === '구분,대상자,입사일,교육일,취급물질,결재,미리보기', dvHead);
  check('부서 제출 현황: 자기 부서 교육일지 표시', dvText.includes('김가나') && dvText.includes('결재대기'), dvText.slice(0, 150));
  check('부서 제출 현황: 다른 부서(71병동) 기록 없음', !dvText.includes('이다라'));
  check('부서 제출 현황: 다른 화면은 숨김', await page.evaluate(() => document.getElementById('pg-edu2').style.display === 'none' && document.getElementById('pg-deptview').style.display === 'block'));
  // 서명이 있는 기록 → 관리자 모드와 같은 데이터 미리보기
  await page.evaluate(async () => {
    const cur = await readGitHubJSON('records/edu_records.json');
    await writeGitHubJSON('records/edu_records.json', cur.data.concat([{ recordId: 'r9|최신입|126101801|61병동|MSDS교육|2026-10-20', submitted: '2026-10-20T01:00:00Z',
      name: '최신입', dept: '61병동', type: 'MSDS교육', dateStart: '2026-10-20', joinDate: '2026-10-18', chem: '테스트소독제',
      sigImg: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==' }]), cur.sha, 'test');
  });
  await page.click('#dv-refresh');
  await page.waitForSelector('#dv-body [data-edu]');
  const btnCount = await page.$$eval('#dv-body [data-edu]', bs => bs.length);
  await page.evaluate(() => [...document.querySelectorAll('#dv-body tr')].find(tr => tr.textContent.includes('최신입')).querySelector('[data-edu]').click());
  await page.waitForSelector('#data-view-modal', { timeout: 15000 });
  check('미리보기: 관리자 모드와 같은 일지 화면(서명 포함 기록)', (await page.textContent('#data-view-modal')).includes('최신입'), btnCount);
  await page.evaluate(() => document.getElementById('data-view-modal').remove());
  const topKeys = await page.evaluate(() => _lsTopKeys());
  check('저장 용량: 항목별 용량 계산', Array.isArray(topKeys) && topKeys.length > 0 && 'imgKB' in topKeys[0], topKeys.slice(0, 2));
  await page.evaluate(() => goTab('edu2'));
  check('다른 탭으로 가면 제출 현황 숨김', await page.evaluate(() => document.getElementById('pg-deptview').style.display === 'none'));

  // 14-2. 서명·사진 분리 저장 (v18.4)
  const SIG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
  // (가) 서버 보존: 관리자가 서명 빠진 기록으로 덮어써도 서버 서명 유지
  const keep = await page.evaluate(async () => {
    const a = await BMAUTH.post('login', { kind: 'admin', password: 'admin-pass-1' });
    const g1 = await BMAUTH.post('fs.get', { path: 'records/edu_records.json', t: a.token });
    const dec = b => JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(b), c => c.charCodeAt(0))));
    const enc = o => btoa(String.fromCharCode(...new TextEncoder().encode(JSON.stringify(o))));
    const list = dec(g1.content);
    list.forEach(r => { if (r.sigImg) { delete r.sigImg; r._img = 1; } });
    const put = await BMAUTH.post('fs.put', { path: 'records/edu_records.json', sha: g1.sha, content: enc(list), t: a.token });
    const g2 = await BMAUTH.post('fs.get', { path: 'records/edu_records.json', t: a.token });
    const after = dec(g2.content).find(r => r.name === '최신입');
    return { put: put.success, sig: !!(after && after.sigImg), marker: !!(after && after._img) };
  });
  check('서버: 서명 빠진 기록으로 저장해도 기존 서명 보존', keep.put && keep.sig && !keep.marker, keep);

  // (나) PC 저장: 서버에 있는 기록만 서명을 비우고, 서버에 없는 기록은 서명 유지
  await page.waitForFunction(() => BMIMG.enabled(), null, { timeout: 15000 });
  check('서버 v18.4 확인 후에만 서명 비우기 활성화', true);
  const local = await page.evaluate(async (sig) => {
    const srv = (await readGitHubJSON('records/edu_records.json')).data.find(r => r.name === '최신입');
    const onlyLocal = { recordId: 'L1|미업로드|0|61병동|MSDS교육|2026-10-21', name: '미업로드', dept: '61병동', type: 'MSDS교육', dateStart: '2026-10-21', sigImg: sig };
    localStorage.setItem('bm_edu_records', JSON.stringify([srv, onlyLocal]));
    await BMIMG.compact();
    const raw = JSON.parse(BMIMG.rawGet('bm_edu_records'));
    const viaGet = JSON.parse(localStorage.getItem('bm_edu_records'));
    return {
      srvRawHasSig: !!raw.find(r => r.name === '최신입').sigImg, srvRawMarker: raw.find(r => r.name === '최신입')._img === 1,
      localRawHasSig: !!raw.find(r => r.name === '미업로드').sigImg,
      srvGetHasSig: !!viaGet.find(r => r.name === '최신입').sigImg,
      rawLen: BMIMG.rawGet('bm_edu_records').length, getLen: localStorage.getItem('bm_edu_records').length,
    };
  }, SIG);
  check('PC: 서버에 있는 기록은 서명을 비워 저장 (_img 표시)', !local.srvRawHasSig && local.srvRawMarker, local);
  check('PC: 서버에 없는 기록(업로드 전·실패)은 서명 유지', local.localRawHasSig, local);
  check('PC: 읽을 때는 서명이 다시 붙어 기존 기능 그대로', local.srvGetHasSig, local);
  // (다) 새로 열면(메모리 비움) 서버에서 서명을 다시 받아 둠
  await page.reload();
  await page.waitForFunction(() => window.BMIMG && !!JSON.parse(localStorage.getItem('bm_edu_records') || '[]').find(r => r.name === '최신입' && r.sigImg), null, { timeout: 20000 }).catch(() => {});
  const reloaded = await page.evaluate(() => ({
    get: !!JSON.parse(localStorage.getItem('bm_edu_records')).find(r => r.name === '최신입').sigImg,
    raw: BMIMG.rawGet('bm_edu_records').indexOf('iVBOR') }));
  check('새로 열어도 서버에서 서명을 받아 복원 (PC에는 미업로드 건만 이미지)', reloaded.get, reloaded);
  // (라) 용량 표시는 실제 저장값 기준
  const storage = await page.evaluate(() => ({ edu: apvGetStorageInfo().eduKB, raw: Math.round(BMIMG.rawGet('bm_edu_records').length * 2 / 1024) }));
  check('저장 용량: 실제 저장된 크기로 표시', storage.edu === storage.raw, storage);

  // 15. 성별 삭제
  const g2 = await page.evaluate(() => ({
    input: !!document.getElementById('e2-gender') || !!document.getElementById('e1-gender'),
    head: [...document.querySelectorAll('.ptbl th')].some(th => th.textContent.includes('성별')),
    partKey: e2.parts.some(p => 'gender' in p),
    rowSelect: document.getElementById('e2-ptbody').innerHTML.includes('>남<'),
  }));
  check('성별: 입력란·표 머리·대상자 데이터·표 행 모두 없음', !g2.input && !g2.head && !g2.partKey && !g2.rowSelect, g2);

  // 16. 교육일지 저장 → 교육완료
  await page.evaluate(() => new Promise(res => {
    window.__cb2 = r => { window.__rec2 = r; res(); };
    const sc = document.createElement('script');
    sc.src = CFG.url + '?action=record&type=MSDS교육&dept=' + encodeURIComponent('61병동') + '&names=' + encodeURIComponent('최신입') + '&callback=__cb2';
    document.head.appendChild(sc);
  }));
  check('교육일지 저장 → 취급-교육완료', g.api.nhRows_()[0]['상태'] === '취급-교육완료', g.api.nhRows_()[0]['상태']);

  // 17. 3일 무회신 재알림 (자동 확인 시)
  const o = g.api.nhRows_()[0];
  o['상태'] = '회신대기';
  o['최초발송'] = new Date(Date.now() + 9 * 3600e3 - 4 * 86400e3).toISOString().slice(0, 19);
  g.api.nhWrite_(o);
  const before = mail.sent.length;
  g.api.dailyCheck();
  const reminders = mail.sent.slice(before).filter(m => /재알림/.test(m.subject));
  check('3일 무회신 → 매일 자동 확인에서 재알림 1건', reminders.length === 1, mail.sent.slice(before).map(m => m.subject));
  check('같은 날 다시 실행해도 중복 없음', (g.api.dailyCheck(), mail.sent.filter(m => /재알림/.test(m.subject)).length === 1));
  check('자동 확인을 반복해도 신규입사 메일 중복 없음', mail.sent.filter(m => m.to === 'head61@bumin.co.kr' && !/재알림/.test(m.subject)).length === 1);

  check('페이지 스크립트 오류 없음', errors.length === 0, errors.slice(0, 5));
  await ctx.unrouteAll({ behavior: 'ignoreErrors' });
  await browser.close();
  server.close();
  console.log(failures ? `\n${failures} FAILED` : '\nALL PASSED');
  process.exit(failures ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
