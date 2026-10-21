// End-to-end check of v18 BMAUTH proxy + Code.gs v2 against the mock GAS server.
const { chromium } = require('playwright');
const { start } = require('./server');

const path = require('path');
const CODE = path.join(__dirname, '..', 'Code.gs');
const HTML = path.join(__dirname, '..', '..', '부민병원_MSDS_통합시스템_v5 - 2026-04-17T095042.287.html');
const PORT = 8787;
const MOCK = `http://localhost:${PORT}/script.google.com/macros/s/x/exec`;

let failures = 0;
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

  check('페이지 스크립트 오류 없음', errors.length === 0, errors.slice(0, 5));
  await ctx.unrouteAll({ behavior: 'ignoreErrors' });
  await browser.close();
  server.close();
  console.log(failures ? `\n${failures} FAILED` : '\nALL PASSED');
  process.exit(failures ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
