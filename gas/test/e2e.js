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
  let ghOldHits = 0;
  const ctx = await browser.newContext();
  await ctx.route('**/*', async route => {
    const u = route.request().url();
    if (u.startsWith('https://script.google.com/macros/')) {
      const q = u.indexOf('?') >= 0 ? u.slice(u.indexOf('?')) : '';
      try {
        const resp = await route.fetch({ url: MOCK + q });
        return route.fulfill({ response: resp, headers: { ...resp.headers(), 'access-control-allow-origin': '*' } });
      } catch (e) { return route.abort().catch(() => {}); }   // 시험 끝날 때 남은 요청 (연결 끊김) 무시
    }
    if (u.startsWith(`http://localhost:${PORT}/`)) return route.continue();
    if (u.startsWith('https://raw.githubusercontent.com/buminhealth/health/main/' + encodeURIComponent('교육일지') + '/2026-03/' + encodeURIComponent('옛일지-61병동.pdf'))) {
      ghOldHits++;   // v20.3: 드라이브에 없는 옛 일지 PDF만 공개 저장소에서 읽기 (관리자)
      return route.fulfill({ status: 200, contentType: 'application/pdf', body: '%PDF-1.4 old' });
    }
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
  check('비로그인: 배지 표시', (await page.textContent('#bm-auth-badge')).includes('부서 / 관리자 로그인'));
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
      { name: '시험물질B', manufacturer: '나제조', usage: '세척', targetManage: false, specialManage: false, departments: ['시험일반부'], components: [{ compName: '정제수', casNumber: '7732-18-5', amount: '90', workEnv: '', healthCheck: '', memo: '26년 상반기부터 작측 제외', memo2: '혈액학' }], icons: [], emergency: '' },
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
    const naText = document.querySelector('.ms-ct').textContent;
    searchChem('7722-84'); const nS = document.querySelectorAll('#gridContainer .ms-li').length;
    searchChem(''); filterDept('시험일반부'); const nD = document.querySelectorAll('#gridContainer .ms-li').length;
    filterDept('all');
    return { n, total: CHEMS.length, cnt, inList, k1, b1, h2, em, k2, b2, nS, nD, naText };
  });
  check('MSDS: 전체 물질 목록과 개수', ms.n === ms.total && ms.cnt.includes(String(ms.total)), ms);
  check('MSDS: 부서 선택은 목록 안', ms.inList);
  check('MSDS: CAS 검색 · 부서 선택으로 줄어듦', ms.nS === 1 && ms.nD === 1 && ms.n >= 3, ms);
  check('MSDS: 응급처치 표', ms.em);
  check('MSDS: 정상 CAS는 ⚠ 없음', /7732-18-5/.test(ms.naText) && !/⚠/.test(ms.naText), ms.naText);
  check('MSDS: 특검 · 작측 빈칸은 미해당, 비고(메모1 · 메모2) 표시', /미해당/.test(ms.naText) && /비고/.test(ms.naText) && /26년 상반기부터 작측 제외/.test(ms.naText) && /혈액학/.test(ms.naText), ms.naText);
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

  // 1-2. v19.2 홈 (로그인 전)
  const home0 = await page.evaluate(() => ({
    vis: getComputedStyle(document.getElementById('pg-home')).display !== 'none',
    msdsHidden: document.getElementById('pg-msds').style.display === 'none',
    tabOn: document.getElementById('tab-home').classList.contains('on'),
    tiles: [...document.querySelectorAll('#hm-tiles .hm-tile strong')].map(x => x.textContent),
    ems: [...document.querySelectorAll('#hm-tiles .hm-tile em')].map(x => x.textContent),
    todo: document.getElementById('hm-todo-t').textContent,
    login: !!document.getElementById('hm-login'),
    foot: (document.querySelector('#pg-home > .bm-foot') || {}).textContent || '',
  }));
  check('홈: 처음 화면은 홈', home0.vis && home0.msdsHidden && home0.tabOn, home0);
  check('홈: 타일 5개 (v21.1 관리자 모드 타일 삭제)', JSON.stringify(home0.tiles) === JSON.stringify(['MSDS 관리', '교육일지 작성', '보호구 지급', '화학물질 보고서', 'AI 위험성평가']), home0.tiles);
  check('홈(로그인 전): 작성 타일 버튼은 부서 로그인', home0.ems.slice(1, 4).every(x => x === '부서 로그인'), home0.ems);
  check('홈(로그인 전): 로그인하면 할 수 있는 일 (v21.1 로그인 줄 삭제 · 오른쪽 위 버튼 하나)', home0.todo.includes('할 수 있는 일') && !home0.login, home0.todo);
  check('홈: 아래 띠 문구 (v21.1)', home0.foot.includes('안전보건 통합관리 시스템') && home0.foot.includes('3293') && !home0.foot.includes('['), home0.foot);
  await page.click('#hm-tiles .hm-tile[data-go="msds"]');
  check('홈: MSDS 타일 → MSDS 화면', await page.evaluate(() => document.getElementById('pg-msds').style.display === 'flex'));
  await page.click('.bm-logo');
  check('로고 → 홈', await page.evaluate(() => document.getElementById('pg-home').style.display === 'block'));

  // 2. 작성 탭(홈 교육일지 타일) → 로그인 창 → 틀린 PIN → 맞는 PIN → 가려던 화면으로
  await page.click('#hm-tiles .hm-tile[data-go="edu"]');
  await page.waitForSelector('#bm-login', { state: 'visible' });
  check('로그인 창: 작성할 것 표시', (await page.textContent('#bm-login-what')).includes('MSDS 교육일지'));
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
  check('로그인 후 가려던 교육일지 화면으로', await page.evaluate(() => document.getElementById('pg-edu2').style.display === 'block' && document.getElementById('pg-home').style.display === 'none'));
  await page.evaluate(() => goTab('home'));
  await page.waitForFunction(() => !document.getElementById('hm-todo').textContent.includes('불러오는 중'));
  const home1 = await page.evaluate(() => ({
    ems: [...document.querySelectorAll('#hm-tiles .hm-tile em')].map(x => x.textContent),
    edu: document.querySelectorAll('#hm-tiles .hm-tile span')[1].textContent,
    todo: document.getElementById('hm-todo-t').textContent,
    rows: document.getElementById('hm-todo').textContent,
  }));
  check('홈(로그인 후): 작성하기 · 보고하기', JSON.stringify(home1.ems.slice(1, 4)) === '["작성하기","작성하기","보고하기"]', home1.ems);
  check('홈(로그인 후): 교육일지 타일에 부서 · 일지 종류', home1.edu.includes('61병동') && /교육일지/.test(home1.edu), home1.edu);
  check('홈(로그인 후): 부서 할 일', home1.todo === '61병동 할 일' && home1.rows.length > 0, home1);
  const emps = await page.evaluate(() => empList.map(e => e.name + '/' + e.dept + '/' + e.joinDate));
  check('부서 로그인: 자기 부서 직원만', emps.length === 2 && emps.every(x => x.includes('/61병동/')), emps);
  check('직원 입사일 KST 기준 (하루 당김 없음)', emps.includes('김가나/61병동/2026-07-01'), emps);
  await page.evaluate(() => goTab('edu2'));
  const e2dept = await page.evaluate(() => [...document.getElementById('e2-dept').options].map(o => o.value));
  check('부서 로그인: 교육일지 부서 고정', JSON.stringify(e2dept) === '["61병동"]', e2dept);
  const edu4 = await page.evaluate(() => {
    e2DeptCh();
    const vis = id => { const e = document.getElementById(id); return !!e && e.style.display !== 'none'; };
    return {
      head: BMAUTH.session().head,
      instr: document.getElementById('e2-instr').value,
      note: document.getElementById('e2-kind').textContent,
      noteVis: vis('e2-kind'),
      tabs: [vis('tab-edu1'), vis('tab-edu2')],
      label: document.querySelector('#tab-edu2 .tbl').textContent,
      redirect: (goTab('edu1'), [document.getElementById('pg-edu1').style.display, document.getElementById('pg-edu2').style.display]),
    };
  });
  check('서버: 로그인 응답에 부서장 이름', edu4.head === '김부장', edu4.head);
  check('교육 실시자 = 부서장 자동 입력', edu4.instr === '김부장 / 61병동 부서장', edu4.instr);
  check('교육일지 위 종류 안내 (관리대상 없음 → MSDS 교육일지)', edu4.noteVis && edu4.note.includes('MSDS 교육일지로 작성') && edu4.note.includes('김부장'), edu4.note);
  check('부서 로그인: 위쪽 메뉴 교육일지 하나', JSON.stringify(edu4.tabs) === '[false,true]' && edu4.label === '교육일지', edu4);
  check('부서 로그인: 다른 일지로 가도 부서에 맞는 일지로', JSON.stringify(edu4.redirect) === '["none","block"]', edu4.redirect);

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
  const selfApv = await page.evaluate(async () => {
    const cur = await readGitHubJSON('records/edu_records.json');
    const list = cur.data.map(r => Object.assign({}, r, { approved: true, approverName: '부서가씀', rejected: true }));
    await writeGitHubJSON('records/edu_records.json', list, cur.sha, 'test');
    return (await readGitHubJSON('records/edu_records.json')).data;
  });
  check('서버: 부서는 결재 · 반려 칸을 바꿀 수 없음', selfApv.length === 1 && !selfApv[0].approved && !selfApv[0].approverName && !selfApv[0].rejected, selfApv);

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

  // 6-1. v19.5 관리자: 결재 요약 · 반려 · PDF 원본 보기
  const rj = await page.evaluate(async () => {
    const cur = await readGitHubJSON('records/edu_records.json');
    localStorage.setItem('bm_edu_records', JSON.stringify(cur.data.map(r => Object.assign({}, r, { pdfFileName: 'MSDS교육일지-61병동-김가나-2026-07-01.pdf', pdfYm: todayStr().slice(0, 7) }))));
    apvLoadAndRender();
    const key = apvState.groups[0].groupKey;
    const kpi0 = document.getElementById('adm-kpi').textContent;
    window.prompt = () => '교육 시간 누락';
    await apvBatchReject([key]);
    const srv = (await readGitHubJSON('records/edu_records.json')).data;
    const g = apvState.groups.find(x => x.groupKey === key);
    const chip = [...document.querySelectorAll('#apv-list .apv-row')].find(r => r.dataset.key === key).textContent;
    const kpi1 = document.getElementById('adm-kpi').textContent;
    document.getElementById('apv-status').value = 'pending'; apvRender();
    const pendingRows = document.querySelectorAll('#apv-list .apv-row[data-key]').length;
    document.getElementById('apv-status').value = 'rejected'; apvRender();
    const rjRows = document.querySelectorAll('#apv-list .apv-row[data-key]').length;
    document.getElementById('apv-status').value = ''; apvRender();
    // 줄 누르면 PDF 원본 보기
    document.querySelector('#apv-list .apv-row[data-key]').click();
    const t0 = Date.now();
    while (Date.now() - t0 < 8000 && !document.querySelector('#apv-v-frame, #apv-v-data')) await new Promise(r => setTimeout(r, 100));
    const viewer = { open: !!document.getElementById('apv-viewer'), frame: !!document.getElementById('apv-v-frame'),
      btns: [...document.querySelectorAll('#apv-viewer .apv-v-bar button, #apv-viewer .apv-v-ac button')].map(b => b.textContent),
      st: (document.querySelector('#apv-viewer .apv-v-st') || {}).textContent || '' };
    apvCloseViewer();
    // 시험 뒤 원래대로 (반려 지움)
    const back = (await readGitHubJSON('records/edu_records.json'));
    const clean = back.data.map(r => { const c = Object.assign({}, r); ['rejected', 'rejectReason', 'rejectedAt', 'rejectedBy'].forEach(f => delete c[f]); return c; });
    await writeGitHubJSON('records/edu_records.json', clean, back.sha, 'test');
    localStorage.setItem('bm_edu_records', JSON.stringify(clean));
    return { kpi0, kpi1, srv: srv.map(r => [r.rejected, r.rejectReason]), g: [g.rejected, g.rejectReason], chip, pendingRows, rjRows, viewer, closed: !document.getElementById('apv-viewer') };
  });
  check('관리자: 결재 요약 4칸', /결재 대기/.test(rj.kpi0) && /반려/.test(rj.kpi0) && /이번 달 결재/.test(rj.kpi0), rj.kpi0);
  check('반려: 사유와 함께 서버 기록에 저장', rj.srv.length === 1 && rj.srv[0][0] === true && rj.srv[0][1] === '교육 시간 누락', rj.srv);
  check('반려: 목록 상태 · 필터', rj.g[0] === true && rj.chip.includes('반려') && rj.pendingRows === 0 && rj.rjRows === 1, rj);
  check('반려: 요약에 반려 1건', /반려1 건/.test(rj.kpi1.replace(/\s+/g, ' ').replace('반려 1', '반려1')), rj.kpi1);
  check('PDF 원본 보기: 줄 누르면 오른쪽에 열림 (새 창 · 내려받기 · 인쇄 · 반려 · 결재)', rj.viewer.open && rj.viewer.frame &&
    ['새 창으로 열기', '내려받기', '인쇄'].every(t => rj.viewer.btns.includes(t)) && rj.viewer.btns.some(t => t.includes('반려')) && rj.viewer.btns.includes('이 문서 결재'), rj.viewer);
  check('PDF 원본 보기: 닫기', rj.closed);

  // 6-2. v20.3 옛 일지 PDF (드라이브에 없음) → 공개 저장소에서 읽기 · 대시보드 · 설정 · 메뉴
  const v203 = await page.evaluate(async () => {
    const miss = await apvResolvePdfBytes({ pdfFileName: '옛일지-61병동.pdf', pdfYm: '2026-03', type: 'MSDS교육', dateStart: '2026-03-02', submitted: '2026-03-02' });
    admSwitchPane('dash');
    const t0 = Date.now();
    while (Date.now() - t0 < 8000 && !/%/.test((document.getElementById('dash-edu') || {}).textContent || '')) await new Promise(r => setTimeout(r, 100));
    const dash = document.getElementById('dash-body').textContent;
    const kpis = document.querySelectorAll('#dash-body .adm-kpi-c').length;
    const row = document.querySelector('#dash-body .dash-row[data-go^="apv:"]');
    if (row) row.click();
    const t2 = Date.now();
    while (Date.now() - t2 < 6000 && !document.getElementById('apv-viewer')) await new Promise(r => setTimeout(r, 100));
    const jump = !!row && document.getElementById('adm-pane-apv').classList.contains('active') && !!document.getElementById('apv-viewer');
    apvCloseViewer();
    admSwitchPane('cfg'); BMDASH.settings();
    const t1 = Date.now();
    while (Date.now() - t1 < 8000 && /확인 중/.test((document.getElementById('cfg-ver') || {}).textContent || '')) await new Promise(r => setTimeout(r, 100));
    const cfg = document.getElementById('cfg-body').textContent;
    const nav = [...document.querySelectorAll('.adm-nav-item .adm-nav-label')].map(e => e.textContent);
    const zero = [...document.querySelectorAll('.adm-nav-badge')].filter(b => b.textContent === '0').length;
    const h = document.querySelector('.adm-content').clientHeight, ph = document.getElementById('pg-admin').clientHeight;
    admSwitchPane('apv');
    return { miss: miss ? new TextDecoder().decode(miss) : null, dash, kpis, jump, cfg, nav, zero, fit: h > 0 && h < ph,
             disp: getComputedStyle(document.getElementById('admin-main')).display };
  });
  check('옛 일지 PDF: 드라이브에 없으면 공개 저장소에서 읽음 (관리자)', v203.miss === '%PDF-1.4 old' && ghOldHits >= 1, [v203.miss, ghOldHits]);
  check('관리자 메뉴: A-8 순서 (대시보드 … 신규입사 알림 · 설정)', ['대시보드', '결재 관리', '교육 현황', '보호구 지급 현황', '화학물질 보고서 검토', '위험성평가', '물질 마스터 점검', 'AI MSDS 검토', '직원정보관리', '신규입사 알림', '설정', '홈으로 나가기'].join() === v203.nav.join(), v203.nav);
  check('대시보드: 결재 대기 줄 → 결재 관리로 이동 + PDF 보기 열림', v203.jump);
  check('관리자 메뉴: 0인 배지는 숨김', v203.zero === 0, v203.zero);
  check('대시보드: 요약 6칸 + 결재 대기 · 이수율 · 신규입사 · 최근 제출', v203.kpis === 6 && ['결재 대기 일지', '부서별 교육 이수율', '신규입사 알림', '최근 제출'].every(t => v203.dash.includes(t)) && /\d+%/.test(v203.dash), v203.dash.slice(0, 300));
  check('설정: 결재자 · 서버 버전 · AI 키', /결재자/.test(v203.cfg) && /v21\.6/.test(v203.cfg) && /AI \(Gemini\)/.test(v203.cfg), v203.cfg.slice(0, 300));

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
  const lay = await page.evaluate(() => ({ disp: getComputedStyle(document.getElementById('admin-main')).display,
    c: document.querySelector('.adm-content').getBoundingClientRect().bottom, p: document.getElementById('pg-admin').getBoundingClientRect().bottom }));
  check('관리자 화면: 세로 배치 · 내용 영역이 화면 안에서 끝남(잘림 없이 스크롤)', lay.disp === 'flex' && lay.c <= lay.p + 1, lay);
  await page.click('.adm-nav-item[data-pane="staff"]');
  await page.waitForSelector('#staff-paste');
  await page.click('.adm-nav-item[data-pane="nh"]');
  await page.waitForSelector('#staff-nh-run');
  check('신규입사 알림 메뉴: 연결 전에는 [지금 자동 확인] 비활성', await page.$eval('#staff-nh-run', b => b.disabled));
  await page.click('.adm-nav-item[data-pane="staff"]');
  await page.waitForSelector('#staff-paste');
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
  await page.evaluate(() => { admSwitchPane('nh'); BMSTAFF.openNh(); });
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
  await page.waitForFunction(() => document.getElementById('staff-body').textContent.includes('③ 부서 이름 점검'), null, { timeout: 15000 });
  check('관리자 화면: 부서 이름 점검표 표시', (await page.textContent('#staff-body')).includes('PIN 없음'));
  check('관리자 화면: 서버 버전 표시', (await page.textContent('#staff-body')).includes('서버 버전: v21.6'));

  // 13-3. 인사 사이트 이상 감지 + 실패 메일 (v18.6)
  g.props.FB_PATH = 'employees_public';
  const mk = (id, f) => ({ name: 'x/employees_public/' + id, fields: f });
  const many = [];
  for (let i = 0; i < 20; i++) many.push(mk('1300000' + String(i).padStart(2, '0'), { name: S('직원' + i), dept: S('61병동'), joinDate: S('2020-01-01'), quitDate: S('') }));
  fb.collections = { employees_public: many, leaves: [] };
  const okRun = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  check('정상 데이터: 자동 확인 성공·기준 인원 기록', okRun.success && !okRun.error, okRun.error);
  const cfgVal = k => { const r = g.sheets['관리자설정'].rows.find(x => x[0] === k); return r ? r[1] : ''; };
  check('정상 데이터: 기준 인원 20명 기록', cfgVal('hr_last_count') === '20', cfgVal('hr_last_count'));
  const empBefore = g.sheets['직원목록'].rows.length;
  const mailBefore = mail.sent.length;
  fb.collections.employees_public = many.slice(0, 5);   // 인원 급감
  const dropRun = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  check('이상 감지: 인원 급감 → 반영 중단', !!dropRun.error && /크게 줄었습니다/.test(dropRun.error) && g.sheets['직원목록'].rows.length === empBefore, dropRun.error);
  const failMail = mail.sent.slice(mailBefore).filter(m => m.to === 'admin@bumin.co.kr' && /자동 확인 실패/.test(m.subject));
  check('실패 시 보건관리자 메일 1통', failMail.length === 1, mail.sent.slice(mailBefore).map(m => m.subject));
  await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  check('같은 날 같은 실패는 메일 다시 안 보냄', mail.sent.slice(mailBefore).filter(m => /자동 확인 실패/.test(m.subject)).length === 1);
  fb.collections.employees_public = many.map(d => ({ name: d.name, fields: { name: d.fields.name, dept: d.fields.dept, hireDt: d.fields.joinDate } }));   // 필드 이름 변경
  const renRun = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  check('이상 감지: 입사일 필드 이름 변경(빈 비율 0%→100%) → 반영 중단', !!renRun.error && /입사일/.test(renRun.error), renRun.error);
  check('다른 사유의 실패는 새로 메일', mail.sent.slice(mailBefore).filter(m => /자동 확인 실패/.test(m.subject)).length === 2);
  const manual = await page.evaluate(() => BMAUTH.post('hr.fetch', { fresh: true }));
  check('수동 가져오기: 이상 경고를 함께 반환', manual.success && manual.warnings.length > 0, manual.warnings);
  const st186 = await page.evaluate(() => BMAUTH.post('emp.status', {}));
  check('관리자 화면용 최근 오류에 사유 표시', /입사일/.test(st186.hrLastError), st186.hrLastError);
  // 원래부터 입사일이 대부분 비어 있는 데이터는 오탐하지 않음 (기준값을 지우고 첫 실행처럼)
  g.sheets['관리자설정'].rows = g.sheets['관리자설정'].rows.filter(x => !/^hr_last_(count|ratios)$/.test(x[0]));
  fb.collections.employees_public = many.map((d, i) => ({ name: d.name, fields: i < 15 ? { name: d.fields.name, dept: d.fields.dept }
                                                                                    : { name: d.fields.name, dept: d.fields.dept, joinDate: d.fields.joinDate } }));
  const noFalse1 = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  const noFalse2 = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  check('오탐 방지: 원래 입사일 75%가 빈 데이터도 정상 반영 (첫 실행·다음 실행 모두)', !noFalse1.error && !noFalse2.error, [noFalse1.error, noFalse2.error]);
  // v20.4 인사 사이트 읽기 줄이기: 매일 확인은 조건 조회(부분), 전체는 주 1회 · 가져오기 6시간 재사용 · 한도 초과 안내
  fb.collections = { employees_public: many.map(d => ({ name: d.name, fields: Object.assign({}, d.fields) })), leaves: [] };
  fb.collections.employees_public.push(mk('126102101', { name: S('부분신입'), dept: S('61병동'), joinDate: S(kstDay(-1)), quitDate: S('') }));
  fb.collections.employees_public[3].fields.quitDate = S(kstDay(-1));   // 직원3 어제 퇴사
  fb.collections.leaves = [{ name: 'x/leaves/a', fields: { empCode: S('130000005'), type: S('육아휴직'), start: S(kstDay(-2)), end: S(kstDay(60)) } }];
  const setCfg = (k, v) => { const rr = g.sheets['관리자설정'].rows; const r = rr.find(x => x[0] === k); if (r) r[1] = v; else rr.push([k, v, '']); };
  setCfg('hr_last_full', kstDay(-1)); setCfg('hr_last_check', kstDay(-1));
  const sunday = new Date(kstDay(0) + 'T12:00:00Z').getUTCDay() === 0;
  fb.queries = [];
  const callsBefore = fb.calls.length;
  const pr = await page.evaluate(() => BMAUTH.post('nh.run', {}));
  const eRows = g.sheets['직원목록'].rows, EH2 = eRows[0];
  const empOf = nm => eRows.find(r => r[EH2.indexOf('이름')] === nm) || [];
  if (!sunday) {
    check('매일 확인 = 부분 조회 (입사일 · 퇴사일 · 휴직 종료일 조건) · 읽은 문서 수가 전체보다 훨씬 적음', pr.success && pr.mode === 'partial' && pr.reads < 15 &&
      ['joinDateGREATER_THAN_OR_EQUAL', 'quitDateGREATER_THAN_OR_EQUAL', 'endGREATER_THAN_OR_EQUAL', 'endEQUAL'].every(q => fb.queries.some(x => x.filters.indexOf(q) >= 0)), [pr.mode, pr.reads, fb.queries]);
    check('부분 조회 반영: 신규 입사 추가 · 어제 퇴사 → 퇴사 · 현재 휴직 → 휴직 (명단 없는 직원은 그대로)', !!empOf('부분신입').length && /퇴사/.test(empOf('직원3')[EH2.indexOf('재직상태')]) &&
      /휴직\(육아휴직\)/.test(empOf('직원5')[EH2.indexOf('재직상태')]) && !/퇴사/.test(empOf('직원7')[EH2.indexOf('재직상태')]), [empOf('직원3'), empOf('직원5')]);
  }
  const full = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  check('전체 동기화(요청 · 주 1회)는 전체 조회 + 이상 감지', full.success && full.mode === 'full' && full.reads >= 20, [full.mode, full.reads, full.error]);
  const c0 = fb.calls.length;
  const f1 = await page.evaluate(() => BMAUTH.post('hr.fetch', {}));
  const c1 = fb.calls.length;
  const f2 = await page.evaluate(() => BMAUTH.post('hr.fetch', { fresh: true }));
  check('[저장된 계정으로 가져오기]: 6시간 안이면 다시 읽지 않음 · 새로 읽기는 fresh', f1.success && f1.cached && c1 === c0 && f2.success && !f2.cached && fb.calls.length > c1, [f1.cached, c1 - c0, f2.cached]);
  fb.quota = true;
  const q429 = await page.evaluate(() => BMAUTH.post('nh.run', { full: true }));
  fb.quota = false;
  check('읽기 한도 초과(429) → "오후 4시 이후 다시" 안내', /한도/.test(q429.error || '') && /오후 4시/.test(q429.error || ''), q429.error);
  const stR = await page.evaluate(() => BMAUTH.post('emp.status', {}));
  check('직원정보관리 상태: 오늘 인사 사이트 읽기 건수 · 마지막 전체 동기화', stR.hrReads > 0 && stR.hrLastFull === kstDay(0), [stR.hrReads, stR.hrLastFull]);
  fb.collections = null; g.props.FB_PATH = 'employees';

  // 13-4. 물질 마스터 점검 (v18.7): CAS 복원·검증, 빈칸 제안, 법령대조표, 빈 칸에만 반영
  const ch = g.sheets['chemicals'];
  ch.rows[0] = ['제품명', '제조사', '부서', '성분명', 'CAS', '함유량', '특수건강검진', '작업환경측정'];
  ch.appendRow(['효소세정제', 'C사', '중앙공급실', 'Protease', vm.runInContext('new Date(9014, 0, 1)', g.ctx), '1', '', '']);
  ch.appendRow(['메탄올A', 'D사', '병리과', '메탄올', '67-56-1', '99', '12개월', '6개월']);
  ch.appendRow(['메탄올B', 'D사', '병리과', '메탄올', '67-56-1', '50', '', '']);
  ch.appendRow(['정제수제품', 'E사', '약제과', '정제수', '732-18-5', '90', '', '']);
  ch.appendRow(['EDTA용액', 'F사', '진단검사의학과', 'EDTA', '64-02-8.', '5', '', '']);
  ch.appendRow(['정제수B', 'G사', '약제과', '물', '7732-18-5', '95', '', '']);
  ch.appendRow(['그람탈색액', 'H사', '진단검사의학과', '에틸알코올', '64-17- 5', '2026-09-10 00:00:00', '', '']);
  ch.appendRow(['세정제K', 'I사', '영양팀', '탄산 칼륨', '0584-08-07', vm.runInContext('new Date(2026, 0, 5)', g.ctx), '', '']);
  ch.appendRow(['광택제L', 'J사', '미화부', 'Lemon oil', '226-394-6', '1', '', '']);
  const au = await page.evaluate(() => BMAUTH.post('chem.audit', {}));
  const iss = (au.casIssues || []).reduce((m, x) => (m[x.comp] = x, m), {});
  check('물질 점검: 날짜로 바뀐 CAS → 9014-01-1 복원 제안', iss.Protease && iss.Protease.fix === '9014-01-1', iss.Protease);
  check('물질 점검: 끝의 점 → 64-02-8', iss.EDTA && iss.EDTA.fix === '64-02-8', iss.EDTA);
  check('물질 점검: 검증숫자 오류(732-18-5) → 같은 시트의 7732-18-5로 추정', iss['정제수'] && iss['정제수'].fix === '7732-18-5' && iss['정제수'].guess && /검증숫자/.test(iss['정제수'].why), iss['정제수']);
  check('물질 점검: 띄어쓰기 · 앞자리 0 복원 (64-17-5, 584-08-7)', iss['에틸알코올'] && iss['에틸알코올'].fix === '64-17-5' && iss['탄산 칼륨'] && iss['탄산 칼륨'].fix === '584-08-7', [iss['에틸알코올'], iss['탄산 칼륨']]);
  check('물질 점검: EC 번호 구분 (CAS 아님)', iss['Lemon oil'] && !iss['Lemon oil'].fix && /EC 번호/.test(iss['Lemon oil'].why), iss['Lemon oil']);
  const amt = (au.amountIssues || []).reduce((m, x) => (m[x.comp] = x.fix, m), {});
  check('물질 점검: 함유량이 날짜로 바뀜 → 9-10, 1-5 복원 제안', amt['에틸알코올'] === '9-10' && amt['탄산 칼륨'] === '1-5', au.amountIssues);
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
  const chRow = n => ch.rows.findIndex(r => r[0] === n) + 1;
  const ap2 = await page.evaluate(a => BMAUTH.post('chem.apply', a), {
    casFixes: [{ row: chRow('광택제L'), fix: '8008-56-8', raw: '226-394-6' }, { row: chRow('정제수B'), fix: '7732-18-5' }, { row: chRow('메탄올B'), fix: '64-17-5' }, { row: chRow('정제수제품'), fix: '7732-18-4' }],
    amountFixes: [{ row: chRow('그람탈색액'), fix: '9-10' }, { row: chRow('세정제K'), fix: '1-5' }] });
  check('반영: 관리자가 입력한 CAS (검증숫자 맞을 때만, 정상 CAS는 안 바꿈)', ap2.fixed === 1 && ch.rows[chRow('광택제L') - 1][4] === '8008-56-8' && ch.rows[chRow('메탄올B') - 1][4] === '67-56-1' && ap2.skipped === 3, ap2);
  check('반영: 함유량 원래 값으로 복원', ap2.amounts === 2 && ch.rows[chRow('그람탈색액') - 1][5] === '9-10' && ch.rows[chRow('세정제K') - 1][5] === '1-5', ap2);
  const imp = await page.evaluate(() => BMAUTH.post('chem.lawImport', { kind: 'hc', value: '12개월', src: '시행규칙 별표22',
    text: '1\t메틸알코올(Methanol; 67-56-1)\n2\t에틸렌글리콜(Ethylene glycol; 107-21-1)\n3\t잘못된 번호 50-00-1\n제목 줄' }));
  check('법령대조표 붙여넣기: 새 CAS 추가 · 같은 값은 그대로 · 검증숫자 오류 건너뜀', imp.success && imp.added === 1 && imp.same === 1 && imp.invalid === 1 && g_lawHas(g, '107-21-1'), imp);
  const imp2 = await page.evaluate(() => BMAUTH.post('chem.lawImport', { kind: 'hc', value: '6개월', text: '메틸알코올 67-56-1' }));
  check('법령대조표 붙여넣기: 기존 값과 다르면 바꾸지 않고 알림', imp2.success && imp2.differCount === 1 && imp2.updated === 0, imp2);
  const imp3 = await page.evaluate(() => BMAUTH.post('chem.lawImport', { kind: 'target', text: '메틸알코올 67-56-1' }));
  check('법령대조표 붙여넣기: 관리대상 칸 채우기', imp3.success && imp3.updated === 1, imp3);
  const guard = await page.evaluate(() => BMAUTH.post('chem.lawImport', { kind: 'hc', value: '12개월', text: 'Sodium Hydroxide\tSodium Hydroxide\t1310-73-2\t100\t특수건강진단\t12개월\t같은 CAS의 다른 제품 판정' }));
  check('법령대조표 붙여넣기: 앱 ④ 목록을 붙여넣으면 거부 (고시 원문만)', !guard.success && /원문/.test(guard.error), guard);
  const lawSh = g.sheets['법령대조표'];
  const lawRow = lawSh.rows.find(r => r[0] === '67-56-1');
  const keepHc = lawRow[2]; lawRow[2] = '미대상';
  const au4 = await page.evaluate(() => BMAUTH.post('chem.audit', {}));
  check("법령대조표 '미대상' → ④ 제안 안 함 · 시트에 주기가 적힌 성분은 따로 표시", !au4.suggestions.some(x => x.cas === '67-56-1' && x.field === 'hc') && au4.lawMismatch.some(x => x.cas === '67-56-1' && x.field === 'hc'), [au4.lawMismatch]);
  let lsh0;
  lsh0 = g.sheets['법령대조표']; lsh0.appendRow(['7722-84-1', '과산화수소', '', '6개월(임/단시)', '', '', '', '', '']);
  const imp6 = await page.evaluate(() => BMAUTH.post('chem.lawImport', { kind: 'we', value: '6개월', text: '과산화수소 7722-84-1' }));
  check("법령대조표 넣기: '6개월(임/단시)'와 '6개월'은 같은 주기로 봄 (다름으로 안 알림)", imp6.success && imp6.same === 1 && imp6.differCount === 0, imp6);
  // v20.7 ③ 화면에서 직접 고침: 미해당(비움) + 메모1 사유 · 그 사이 바뀐 값(old 불일치)은 건너뜀
  const mm = au4.lawMismatch.find(x => x.cas === '67-56-1' && x.field === 'hc');
  const hcCol = ch.rows[0].findIndex(h => /healthCheck|특수건강/.test(h)), memoCol = ch.rows[0].findIndex(h => /메모1|비고|memo/.test(h));
  const ed0 = await page.evaluate(m => BMAUTH.post('chem.apply', { edits: [{ row: m.row, field: 'hc', old: '99개월', value: '' }] }), mm);
  check('③ 직접 고침: 시트 값이 그 사이 바뀌었으면 건너뜀', ed0.success && ed0.edited === 0 && ed0.skipped === 1 && ch.rows[mm.row - 1][hcCol] === mm.value, ed0);
  const nRows = ch.rows.length;
  ch.appendRow(['석회A', 'Z사', '병리과', 'Limestone', '1317-65-3', '50', '24개월', '']);
  ch.appendRow(['석회B', 'Z사', '병리과', 'Limestone', '1317-65-3', '50', '12개월', '']);
  const uiEd = await page.evaluate(async m => {
    admSwitchPane('chemm'); await BMCHEM.open();
    const sel = document.querySelector('[data-medit][data-row="' + m.row + '"]'); sel.value = '';
    document.querySelector('[data-cmemo="' + m.row + '"]').value = '고시 미대상 확인';
    document.getElementById('chemm-edit').click();
    await new Promise(r => setTimeout(r, 1500));
    return !!document.querySelector('[data-cedit]');
  }, mm);
  const limeB = ch.rows[nRows + 1];
  ch.rows.length = nRows;
  check('③ 직접 고침(화면): 미해당(비움) 선택 + 메모1 사유 → 시트 반영 · 같은 CAS 주기 다름 표에도 고르기 칸', ch.rows[mm.row - 1][hcCol] === '' && (memoCol < 0 || /고시 미대상 확인/.test(ch.rows[mm.row - 1][memoCol])), [ch.rows[mm.row - 1], uiEd]);
  lawRow[2] = keepHc;
  // v20.6 국가법령정보 OPEN API: OC 없으면 안내 · 여러 OC 중 되는 것만 남김 · 별표22 원문 → 칸에 채움 (반영은 관리자가)
  const lf0 = await page.evaluate(() => BMAUTH.post('chem.lawFetch', { kind: 'hc' }));
  check('법령 불러오기: OC 없으면 설정 안내', !lf0.success && lf0.noOc, lf0);
  const so = await page.evaluate(() => BMAUTH.post('law.setOc', { oc: 'clover8477, bumin3131' }));
  const lf1 = await page.evaluate(() => BMAUTH.post('chem.lawFetch', { kind: 'hc' }));
  check('OC 저장: 바로 시험해 ✓/✕ 사유 표시 · 되는 OC를 앞으로 (안 되는 것도 지우지 않음) · 가려서 표시', so.success && so.tests.length === 2 && !so.tests[0].ok && /검증에 실패/.test(so.tests[0].msg) && so.tests[1].ok && /^bum\*+, clo\*+$/.test(so.oc) && g.props.LAW_OC === 'bumin3131,clover8477', [so, g.props.LAW_OC]);
  check('법령 불러오기: 별표22 원문(CAS 3개) · 서식 22는 무시', lf1.success && lf1.no === 22 && lf1.casCount === 3 && /7647-01-0/.test(lf1.text) && lf1.value === '12개월' && /별표22/.test(lf1.src) && g.props.LAW_OC === 'bumin3131,clover8477', [lf1, g.props.LAW_OC]);
  const lf2 = await page.evaluate(() => BMAUTH.post('chem.lawFetch', { kind: 'cycle' }));
  check('법령 불러오기: 별표23(주기) 보기', lf2.success && /24개월/.test(lf2.text), lf2.title);
  // v20.7 별표23 주기 자동 고침: 12개월 → 6개월 (원문에서 이름 확인된 것만) · 다른 값은 건너뜀 · 미리보기는 시트 안 바꿈
  const lsh = g.sheets['법령대조표'];
  lsh.appendRow(['71-43-2', '벤젠', '12개월', '', '', '', '별표22', '', '']);
  lsh.appendRow(['56-23-5', '사염화탄소', '미대상', '', '', '', '', '', '']);
  lsh.appendRow(['', '광물성 분진(운모)', '12개월', '', '', '', '', '', '']);
  const cy0 = await page.evaluate(() => BMAUTH.post('chem.lawCycle', { dry: true }));
  const benz = () => lsh.rows.find(r => r[0] === '71-43-2');
  check('별표23 주기 고치기 미리보기: 벤젠 12→6 · 사염화탄소(미대상) 건너뜀 · 염화비닐(원문에 없음) 건너뜀 · 시트 그대로', cy0.success && cy0.checked && cy0.changes.some(x => /벤젠/.test(x)) && cy0.skip.some(x => /사염화탄소/.test(x)) && cy0.skip.some(x => /염화비닐/.test(x)) && benz()[2] === '12개월' && cy0.dust.length === 1, cy0);
  const cy1 = await page.evaluate(() => BMAUTH.post('chem.lawCycle', {}));
  check('별표23 주기 고치기 반영: 벤젠 6개월 · 출처에 별표23 · 없던 물질은 6개월로 추가', cy1.success && benz()[2] === '6개월' && /별표23/.test(benz()[6]) && lsh.rows.some(r => r[0] === '68-12-2' && r[2] === '6개월'), [benz(), cy1.added]);
  const re22 = await page.evaluate(() => BMAUTH.post('chem.lawImport', { kind: 'hc', value: '12개월', text: '벤젠 71-43-2\n디메틸포름아미드 68-12-2' }));
  check('별표22를 다시 넣어도 별표23 6개월 물질은 다름으로 안 알리고 6개월 유지', re22.success && re22.keep6 === 2 && re22.differCount === 0 && benz()[2] === '6개월', re22);
  g.props.LAW_OC = 'wrongoc';
  const lf3 = await page.evaluate(() => BMAUTH.post('chem.lawFetch', { kind: 'hc' }));
  check('법령 불러오기: 인증 실패면 사유 안내', !lf3.success && /검증에 실패/.test(lf3.error), lf3.error);
  g.props.LAW_OC = 'bumin3131';
  const lawUi = await page.evaluate(async () => {
    admSwitchPane('chemm'); await BMCHEM.open();
    document.getElementById('chemm-lawfetch').click();
    const t0 = Date.now();
    while (Date.now() - t0 < 6000 && !document.getElementById('chemm-paste').value) await new Promise(r => setTimeout(r, 100));
    return { text: document.getElementById('chemm-paste').value, src: document.getElementById('chemm-src').value, msg: document.getElementById('chemm-law-msg').textContent };
  });
  check('물질 점검 ⑤ [법령에서 불러오기]: 원문을 칸에 채우고 출처 · 안내 표시 (자동 반영 안 함)', /111-30-8/.test(lawUi.text) && /별표22/.test(lawUi.src) && /법령대조표에 넣기/.test(lawUi.msg), lawUi);
  const denyChem = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' });
    return BMAUTH.post('chem.audit', { t: r.token });
  });
  check('물질 점검: 부서 계정은 거부', !denyChem.success && denyChem.status === 403, denyChem);
  ch.appendRow(['혼합물X', 'Z사', '병리과', 'Ester Lubricant', 'N/A', '5', '', '']);
  await page.evaluate(() => { admSwitchPane('chemm'); return BMCHEM.open(); });
  check('관리자 화면: 물질 마스터 점검 표시', /CAS 번호 문제/.test(await page.textContent('#chemm-body')));
  const chUi = await page.evaluate(async () => {
    const inp = document.querySelector('#chemm-body [data-casin]');
    if (!inp) return { none: true };
    inp.value = '1334-28-1'; inp.dispatchEvent(new Event('input'));
    const bad = document.querySelector('#chemm-body [data-casok="' + inp.getAttribute('data-casin') + '"]').textContent;
    inp.value = '해당없음'; inp.dispatchEvent(new Event('input'));
    const mark = document.querySelector('#chemm-body [data-casok="' + inp.getAttribute('data-casin') + '"]').textContent;
    return { bad, mark, btn: !!document.querySelector('#chemm-body .chemm-apply2') };
  });
  check('물질 점검 화면: 직접 입력 즉시 검증(✕/✓) · ① 표 아래 [CAS 수정 반영] 버튼', /틀림/.test(chUi.bad) && /CAS 없음/.test(chUi.mark) && chUi.btn, chUi);
  const apN = await page.evaluate(r => BMAUTH.post('chem.apply', { casFixes: [{ row: r, fix: '해당없음' }] }), chRow('혼합물X'));
  const au3 = await page.evaluate(() => BMAUTH.post('chem.audit', {}));
  check('CAS가 원래 없는 성분: "해당없음"으로 표시 → 다음 점검에서 문제 목록에서 빠짐', apN.fixed === 1 && ch.rows[chRow('혼합물X') - 1][4] === '해당없음' && !au3.casIssues.some(x => x.product === '혼합물X'), [apN, au3.casIssues.map(x => x.product)]);

  // 13-5. v19.7 AI MSDS 검토 (Gemini) — 키 없음 → 직접 입력 / 키 저장 → AI 추출 · 판정 근거 · 승인 → 시트 추가
  const gm = g.env.GEMINI;
  const ai0 = await page.evaluate(() => BMAUTH.post('ai.status', {}));
  check('AI: 처음엔 키 없음', ai0.success && ai0.hasKey === false && /gemini/.test(ai0.model), ai0);
  const fakePdf = Buffer.from('%PDF-1.4 fake msds').toString('base64');
  const noKey = await page.evaluate(c => BMAUTH.post('ai.analyze', { content: c, name: '시험.pdf' }), fakePdf);
  check('AI: 키 없으면 직접 입력 안내 + 빈 초안', !noKey.success && noKey.noKey && noKey.draft && /MSDS원본\//.test(noKey.path), noKey);
  await page.evaluate(() => BMAUTH.post('ai.setKey', { key: 'test-gemini-key' }));
  gm.reply = { product: { name: 'AI시험세정제', manufacturer: 'K사', usage: '세척', signalWord: '위험', pictograms: ['부식성', '없는그림'],
      emergency: { eye: '물로 씻으시오', skin: '비누로 씻으시오', inhalation: '신선한 공기', ingestion: '구토 금지' } },
    components: [
      { name: '메탄올', cas: '67-56-1', amount: '1~5', reg15: { workEnv: '작업환경측정 대상(6개월)', healthCheck: '특수건강진단 대상(12개월)', managed: '해당', special: '해당없음' } },
      { name: '물', cas: '7732-18-5', amount: '90-95', reg15: { workEnv: '해당없음', healthCheck: '해당없음', managed: '해당없음' } },
      { name: '오타성분', cas: '64-17-6', amount: '<1', reg15: {} },
      { name: '날짜성분', cas: '9014-01-01', amount: '<1', reg15: {} }],
    warnings: ['3번 성분 CAS 흐림'] };
  const an = await page.evaluate(c => BMAUTH.post('ai.analyze', { content: c, name: '시험.pdf' }), fakePdf);
  const d0 = an.draft || {}, cm = (d0.components || []).reduce((m, c) => (m[c.name] = c, m), {});
  check('AI: Gemini에 PDF · 키를 서버에서 보냄', gm.calls.length === 1 && gm.calls[0].key === 'test-gemini-key' && gm.calls[0].body.contents[0].parts[0].inline_data.mime_type === 'application/pdf', gm.calls.length);
  check('AI: 제품 · 성분 추출', an.success && d0.product.name === 'AI시험세정제' && d0.components.length === 4, an);
  check('AI: CAS 검증 (오타 표시 · 날짜 복원)', cm['오타성분'] && !cm['오타성분'].casOk && cm['날짜성분'].cas === '9014-01-1' && cm['날짜성분'].casOk, [cm['오타성분'], cm['날짜성분']]);
  check('AI: 판정 근거 — 법령대조표 우선, MSDS 15번 함께 표시', cm['메탄올'].hcSrc === '법령대조표' && cm['메탄올'].hc === '12개월' && /12개월/.test(cm['메탄올'].evidence.msds15.hc), cm['메탄올']);
  check('AI: 15번 관리대상 → 제품 관리대상 체크', d0.targetManage === true && d0.specialManage === false, d0);
  check('AI: 근거 없으면 미해당(빈칸)', cm['물'].hc === '' && cm['물'].we === '', cm['물']);
  gm.status = 429;
  const lim = await page.evaluate(c => BMAUTH.post('ai.analyze', { content: c, name: '시험.pdf' }), fakePdf);
  check('AI: 무료 한도 초과 안내 + 직접 입력 초안', !lim.success && /한도/.test(lim.error) && lim.draft, lim.error);
  gm.status = 200;
  const badApply = await page.evaluate(d => BMAUTH.post('ai.apply', { data: d }), Object.assign({}, d0, { components: d0.components }));
  check('승인: CAS 검증숫자 틀리면 거부', !badApply.success && /64-17-6/.test(badApply.error), badApply);
  const okDraft = Object.assign({}, d0, { departments: ['61병동'], components: d0.components.filter(c => c.name !== '오타성분').map(c => Object.assign({}, c, c.name === '물' ? { memo: '희석용' } : {})) });
  const chBefore = ch.rows.length;
  const ap3 = await page.evaluate(([d, path]) => BMAUTH.post('ai.apply', { data: d, path: path }), [okDraft, an.path]);
  const added = ch.rows.slice(chBefore);
  const H = ch.rows[0];
  check('승인: chemicals 시트에 새 제품 추가 (첫 줄 제품 정보 + 성분 줄)', ap3.success && added.length === 3 && added[0][H.indexOf('제품명')] === 'AI시험세정제' && added[1][H.indexOf('제품명')] === '' && added[1][H.indexOf('성분명')] === '물', { ap3, added });
  check('승인: 특검 · 측정 · CAS 값', added[0][H.indexOf('CAS')] === '67-56-1' && added[0][H.indexOf('특수건강검진')] === '12개월' && added[2][H.indexOf('CAS')] === '9014-01-1', added[0]);
  check('승인: 원본 PDF 링크 연결', /drive\.google\.com\/file/.test(ap3.pdfUrl || ''), ap3);
  const dup = await page.evaluate(d => BMAUTH.post('ai.apply', { data: d }), okDraft);
  check('승인: 같은 이름 제품은 거부 (기존 제품 안 바꿈)', !dup.success && /이미/.test(dup.error), dup);
  // 13-6. v19.8 개정 MSDS: 기존 제품과 비교 → 개정 반영 (이전 줄은 개정이력에 보관)
  const follow = await page.evaluate(() => BMAUTH.post('ai.apply', { data: { product: { name: 'AI후속제품' }, components: [{ name: '에탄올', cas: '64-17-5', amount: '70' }] } }));
  check('개정 준비: 뒤에 다른 제품 추가', follow.success, follow);
  const cmpRes = await page.evaluate(() => BMAUTH.post('ai.compare', { name: 'AI시험세정제' }));
  check('개정: 기존 제품 묶음 읽기', cmpRes.success && cmpRes.count === 3 && cmpRes.components.map(c => c.name).join(',') === '메탄올,물,날짜성분' && cmpRes.components[0].hc === '12개월', cmpRes);
  const plist = await page.evaluate(() => BMAUTH.post('ai.products', {}));
  check('개정: 기존 제품 목록', plist.success && plist.data.includes('AI시험세정제') && plist.data.includes('AI후속제품'), plist.data);
  const rowBefore = ch.rows.findIndex(r => r[0] === 'AI시험세정제') + 1;
  const rev = await page.evaluate(() => BMAUTH.post('ai.apply', { mode: 'revise', target: 'AI시험세정제', data: {
    product: { name: 'AI시험세정제', manufacturer: 'K사(개정)', usage: '세척' },
    components: [{ name: '메탄올', cas: '67-56-1', amount: '3~7', hc: '12개월', we: '6개월' }, { name: '새성분', cas: '64-17-5', amount: '<1' }] } }));
  const H2 = ch.rows[0];
  const blk = ch.rows.slice(rowBefore - 1, rowBefore + 2);
  check('개정 반영: 같은 자리에서 3줄 → 2줄로 바꿈', rev.success && rev.mode === 'revise' && rev.row === rowBefore && rev.oldRows === 3 && rev.rows === 2 &&
    blk[0][H2.indexOf('제조사')] === 'K사(개정)' && blk[0][H2.indexOf('함유량')] === '3~7' && blk[1][H2.indexOf('성분명')] === '새성분' && blk[2][0] === 'AI후속제품', { rev, blk });
  check('개정 반영: 기존 사용부서 · PDF 링크 유지', blk[0][H2.indexOf('부서')] === '61병동', blk[0]);
  const hist = g.sheets['개정이력'];
  check('개정 반영: 이전 3줄을 개정이력 시트에 보관', !!hist && hist.rows.length === 4 && hist.rows[1][1] === 'AI시험세정제' && hist.rows[3].includes('날짜성분'), hist && hist.rows);
  const revBad = await page.evaluate(() => BMAUTH.post('ai.apply', { mode: 'revise', target: '없는제품', data: { product: { name: '없는제품' }, components: [{ name: 'a', cas: '64-17-5' }] } }));
  check('개정 반영: 없는 제품은 거부', !revBad.success, revBad);

  // 부서가 신규사용 보고서에 첨부한 MSDS 목록 → 경로로 분석
  await page.evaluate(() => uploadFileToGitHub('data:application/pdf;base64,' + btoa('%PDF dept msds'), 'records/chem_msds/신규_1_부서MSDS.pdf'));
  const lst = await page.evaluate(() => BMAUTH.post('ai.list', {}));
  check('AI: 부서 신규사용 MSDS 목록', lst.success && lst.data.some(f => f.name === '신규_1_부서MSDS.pdf'), lst.data);
  const an2 = await page.evaluate(() => BMAUTH.post('ai.analyze', { path: 'records/chem_msds/신규_1_부서MSDS.pdf' }));
  check('AI: 부서 MSDS 경로로 분석', an2.success && an2.path === 'records/chem_msds/신규_1_부서MSDS.pdf', an2.error);
  const deptAi = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' });
    return BMAUTH.post('ai.analyze', { content: 'AAAA', t: r.token });
  });
  check('AI: 부서 계정은 거부', !deptAi.success && deptAi.status === 403, deptAi);
  await page.evaluate(() => { admSwitchPane('ai'); return BMAI.open(); });
  check('관리자 화면: AI MSDS 검토 표시', /관리자가 MSDS 올리기/.test(await page.textContent('#ai-body')) && /신규_1_부서MSDS/.test(await page.textContent('#ai-body')));
  // 화면: 같은 이름 제품 → 비교 → 추가 · 삭제 · 변경 · 기존 판정 유지
  gm.reply = { product: { name: 'AI시험세정제', manufacturer: 'K사(2차개정)' }, components: [
    { name: '메탄올', cas: '67-56-1', amount: '3~7', reg15: {} }, { name: '새성분', cas: '64-17-5', amount: '2' }, { name: '추가성분', cas: '7732-18-5', amount: '90' }] };
  await page.click('[data-ai-path]');
  await page.waitForSelector('#ai-cmp', { timeout: 15000 });
  const guessed = await page.$eval('#ai-target', s => s.value);
  await page.click('#ai-cmp');
  await page.waitForFunction(() => /기존 제품과 비교 —/.test(document.getElementById('ai-review').textContent), null, { timeout: 15000 });
  const cmpUi = await page.evaluate(() => ({ t: document.getElementById('ai-review').textContent, btn: document.getElementById('ai-apply').textContent,
    hc: document.querySelector('#ai-review [data-k="hc"]').value }));
  check('개정 비교 화면: 같은 이름 제품 자동 선택', guessed === 'AI시험세정제', guessed);
  check('개정 비교 화면: 추가 · 변경 · 같음 표시와 개정 반영 버튼', /추가 1/.test(cmpUi.t.replace(/\s+/g, ' ')) && /변경/.test(cmpUi.t) && /개정 반영/.test(cmpUi.btn), cmpUi.t.slice(0, 300));
  check('개정 비교 화면: 특검 판정 유지 (12개월)', cmpUi.hc === '12개월', cmpUi.hc);
  const keepUi = await page.evaluate(async () => {
    // 법령대조표 · 시트 근거가 없는 성분의 기존 판정 · 메모는 새 MSDS에서 비어 있으면 유지
    const r = await BMAUTH.post('ai.apply', { data: { product: { name: 'AI유지시험' }, components: [{ name: '특이성분', cas: '7440-02-0', amount: '5', hc: '6개월', memo: '임/단시' }] } });
    document.getElementById('ai-name').value = 'AI유지시험';
    document.querySelectorAll('#ai-review [data-k="name"]')[0].value = '특이성분';
    document.querySelectorAll('#ai-review [data-k="cas"]')[0].value = '7440-02-0';
    document.querySelectorAll('#ai-review [data-k="hc"]')[0].value = '';
    window.BMAI && 0;
    const opt = document.createElement('option'); opt.textContent = 'AI유지시험'; document.getElementById('ai-target').appendChild(opt);
    document.getElementById('ai-target').value = 'AI유지시험';
    document.getElementById('ai-cmp').click();
    const t0 = Date.now();
    while (Date.now() - t0 < 8000 && !/AI유지시험 \(시트/.test(document.getElementById('ai-review').textContent)) await new Promise(x => setTimeout(x, 100));
    return { ok: r.success, hc: document.querySelectorAll('#ai-review [data-k="hc"]')[0].value, memo: document.querySelectorAll('#ai-review [data-k="memo"]')[0].value, t: document.getElementById('ai-review').textContent };
  });
  check("개정 비교 화면: 비어 있는 특검은 기존 값 유지 (기존 판정 유지 표시)", keepUi.ok && keepUi.hc === "6개월" && /기존 판정 유지/.test(keepUi.t), [keepUi.hc, keepUi.memo]);
  // 13-7. v19.9 교육이력대장 중복 정리
  const led = g.sheets['교육이력대장'] || g.ctx.SpreadsheetApp.getActiveSpreadsheet().insertSheet('교육이력대장');
  const ledBak = led.rows.slice();
  const LH = ['순번', '제출일시', '교육구분', '교육부서', '취급물질', '교육시작일', '교육종료일', '교육일수', '총교육시간', '사번', '성명', '입사일', '서명여부'];
  const L = (n, sub, type, dept, ds, no, nm, sig) => [n, sub, type, dept, 'x', ds, ds, '1', '1시간', no, nm, '', sig];
  led.rows = [LH,
    L(1, '2026-05-13T07:00Z', 'MSDS교육', '중환자실', '2026-05-11', '126030313', '배건우', '✓ 서명완료'),
    L(2, '2026-05-14T07:00Z', 'MSDS교육', '중환자실', '2026-05-11', '126030313', '배건우', '✓ 서명완료'),
    L(3, '2026-05-15T07:00Z', 'MSDS교육', '총무팀', '2026-05-14', '121042601', '장은선', '- 미서명'),
    L(4, '2026-05-15T08:00Z', 'MSDS교육', '총무팀', '2026-05-14', '121042601', '장은선', '✓ 서명완료'),
    L(5, '2026-05-16T07:00Z', 'MSDS교육', '응급실', '2026-06-02', '', '정예나', '✓ 서명완료'),
    L(6, '2026-05-17T07:00Z', 'MSDS교육', '응급실', '2026-06-02', '', '정예나', '✓ 서명완료'),
    L(7, '2026-05-18T07:00Z', '특별안전보건교육', '중환자실', '2026-05-11', '126030313', '배건우', '✓ 서명완료')];
  const du = await page.evaluate(() => BMAUTH.post('edu.dupes', {}));
  const keepRows = (du.groups || []).map(x => x.keep.row).join(','), rmRows = (du.groups || []).flatMap(x => x.remove.map(r => r.row)).join(',');
  check('중복 정리: 같은 사람 · 구분 · 부서 · 기간 묶음 (사번 없으면 이름)', du.success && du.groups.length === 3 && du.removeCount === 3, du);
  check('중복 정리: 서명완료 · 먼저 제출한 줄을 남김', keepRows === '2,5,6' && rmRows === '3,4,7', { keepRows, rmRows });
  const dd = await page.evaluate(rows => BMAUTH.post('edu.dedupe', { rows }), [{ row: 3, sub: '2026-05-14T07:00Z', name: '배건우' }, { row: 4, sub: '2026-05-15T07:00Z', name: '장은선' }, { row: 2, sub: '2026-05-13T07:00Z', name: '배건우' }]);
  check('중복 정리: 체크한 줄만 삭제 (남길 줄 · 틀린 값은 거부)', dd.success && dd.deleted === 2 && dd.skipped === 1 && led.rows.length === 6 && led.rows.filter(r => r[10] === '배건우').length === 2, { dd, n: led.rows.length });
  await page.evaluate(() => { admSwitchPane('edu'); return BMEDU.dupes(); });
  check('관리자 화면: 교육이력대장 중복 정리 표시', /교육이력대장 중복 1묶음/.test(await page.textContent('#edu-dup-box')));
  // 13-8. v20.0 교육 이수 현황 · 교육 제외 · 새 물질 교육
  const empSh = g.sheets['직원목록'];
  const chHead = ch.rows[0].slice();
  ch.rows[0][10] = '관리대상';
  const spRow = new Array(11).fill(''); spRow[0] = '특별시험물질'; spRow[2] = '61병동'; spRow[3] = '포름알데히드'; spRow[4] = '50-00-0'; spRow[10] = 'o';
  ch.rows.push(spRow);
  empSh.appendRow(['126000001', '교육가', '61병동', '2020-01-01', '여']);
  empSh.appendRow(['126000002', '교육나', '61병동', '2020-01-01', '여']);
  empSh.appendRow(['126000003', '교육다', '61병동', '2020-01-01', '남']);
  led.rows = [LH,
    ['1', '2026-08-01T01:00Z', '특별안전보건교육', '61병동', '포름알데히드', '2026-08-01', '2026-08-02', '2', '8시간', '126000001', '교육가', '', '✓ 서명완료'],
    ['2', '2026-08-01T01:00Z', '특별안전보건교육', '61병동', '포름알데히드', '2026-08-01', '2026-08-01', '1', '4시간', '126000002', '교육나', '', '✓ 서명완료']];
  const est = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  const P = n => (est.people || []).find(p => p.name === n) || {};
  check('교육 현황: 61병동은 관리대상 물질 → 특별교육 대상', P('교육가').need === '특별', P('교육가'));
  check('교육 현황: 현장 8 + 온라인 8 = 16시간 → 이수', P('교육가').status === '이수' && P('교육가').hours === 16, P('교육가'));
  check('교육 현황: 현장 4시간 → 진행중 (남은 4시간, 기한 첫 교육일+3개월)', P('교육나').status === '진행중' && P('교육나').remain === 4 && P('교육나').deadline === '2026-11-01', P('교육나'));
  check('교육 현황: 일지 없음 → 미이수 (확인 안 해도 대상)', P('교육다').status === '미이수', P('교육다'));
  const d61 = (est.depts || []).find(d => d.dept === '61병동') || {};
  check('교육 현황: 부서 요약', d61.total >= 3 && d61.done >= 1 && d61.progress >= 1 && d61.missing >= 1, d61);
  const ex1 = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000003', name: '교육다', dept: '61병동', reason: '행정 업무만' }));
  const ex0 = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000003', name: '교육다', dept: '61병동', reason: '' }));
  const est2 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  const Q = (st, n) => (st.people || []).find(p => p.name === n) || {};
  check('교육 제외: 사유 없으면 거부', !ex0.success, ex0);
  check('교육 제외: 관리자 확인 → 제외 (사유 · 확인자 기록)', ex1.success && Q(est2, '교육다').status === '제외' && Q(est2, '교육다').excluded.by === '보건관리자' && g.sheets['교육제외'].rows[1][3] === '행정 업무만', Q(est2, '교육다'));
  const exA = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000001', name: '교육가', dept: '61병동', reason: '시험' }));
  const est3 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('교육 제외: 확인일 전 교육만 있으면 제외 (예전 기록은 보존 표시)', exA.success && Q(est3, '교육가').status === '제외' && Q(est3, '교육가').kept.n === 1, Q(est3, '교육가'));
  led.rows.push(['8', '2026-10-22T01:00Z', '특별안전보건교육', '61병동', '포름알데히드', est3.today, est3.today, '1', '8시간', '126000001', '교육가', '', '✓ 서명완료']);
  const est3b = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('교육 제외: 확인일 이후 지금 부서에서 일지를 쓰면 다시 대상', Q(est3b, '교육가').status === '이수', Q(est3b, '교육가').status);
  led.rows.pop();
  await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000001', name: '교육가', dept: '61병동', undo: true }));
  const un = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000003', name: '교육다', dept: '61병동', undo: true }));
  const est4 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('교육 제외: 해제 → 다시 대상', un.success && Q(est4, '교육다').status === '미이수', Q(est4, '교육다'));
  const nc = await page.evaluate(() => BMAUTH.post('edu.newChem', { chem: '신물질X', depts: ['61병동'] }));
  const est5 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  const ncA = (Q(est5, '교육가').newChem || []).find(n => n.chem === '신물질X') || {};
  check('새 물질 교육: 등록 → 부서 직원 대상 (기한 = 취급 전 · 등록일)', nc.success && ncA.done === false && ncA.due === ncA.date, ncA);
  led.rows.push(['3', '2026-10-21T01:00Z', 'MSDS교육', '61병동', '신물질X, 기타', est5.today, est5.today, '1', '1시간', '126000001', '교육가', '', '✓ 서명완료']);
  const est6 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('새 물질 교육: 그 물질이 든 일지를 쓰면 완료', ((Q(est6, '교육가').newChem || []).find(n => n.chem === '신물질X') || {}).done === true, Q(est6, '교육가').newChem);
  check('새 물질 교육: AI 승인 때 사용 부서에 자동 등록', ap3.eduAdded === 1 && g.sheets['새물질교육'].rows.some(r => r[1] === 'AI시험세정제' && r[2] === '61병동'), ap3.eduAdded);
  const deptEx = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' });
    const own = await BMAUTH.post('edu.status', { t: r.token });
    const other = await BMAUTH.post('edu.exclude', { name: '이다라', dept: '71병동', reason: 'x', t: r.token });
    const mail = await BMAUTH.post('edu.askMail', { t: r.token });
    return { own: own.success && own.people.length > 0 && own.people.every(p => p.dept === '61병동'), other, mail: mail.status };
  });
  check('교육 현황: 부서는 자기 부서만 · 다른 부서 제외 불가 · 메일은 관리자만', deptEx.own && !deptEx.other.success && deptEx.other.status === 403 && deptEx.mail === 403, deptEx);
  g.env.MailApp.sent.length = 0;
  const am = await page.evaluate(() => BMAUTH.post('edu.askMail', { depts: ['61병동'] }));
  check('확인 요청 메일: 61병동 부서장에게 미이수 명단 + 링크', am.success && am.sent === 1 && /head61@/.test(g.env.MailApp.sent[0].to) && /교육나/.test(g.env.MailApp.sent[0].htmlBody) && /dv=edu/.test(g.env.MailApp.sent[0].htmlBody), am);
  // v20.1: 직원목록의 인사 사이트 부서명도 부서매핑대로 (동기화 전 직원 포함) · 교육 현황 캐시
  const dm = g.sheets['부서매핑'] || g.ctx.SpreadsheetApp.getActiveSpreadsheet().insertSheet('부서매핑');
  if (!dm.rows.length) dm.appendRow(['인사 사이트 부서명', '앱 부서명', '비고']);
  dm.appendRow(['진검실', '61병동', '']);
  empSh.appendRow(['126000009', '별칭가', '진검실', '2020-01-01', '여']);
  const alias = await page.evaluate(async () => {
    const st = await BMAUTH.post('edu.status', { fresh: true });
    const p = st.people.find(x => x.name === '별칭가');
    return { dept: p && p.dept };
  });
  alias.inEmp = (vm.runInContext("getEmployees({ role: 'dept', dept: '61병동' })", g.ctx).data || []).some(e => e.name === '별칭가');
  check('부서매핑: 직원목록의 옛 부서명도 앱 부서로 (교육 대상 · 부서 직원 명단)', alias.dept === '61병동' && alias.inEmp, alias);
  const cache = await page.evaluate(async () => {
    const a = await BMAUTH.post('edu.status', {});
    await BMAUTH.post('edu.exclude', { no: '126000009', name: '별칭가', dept: '61병동', reason: '시험' });
    const b = await BMAUTH.post('edu.status', {});
    return [a.people.find(x => x.name === '별칭가').status, b.people.find(x => x.name === '별칭가').status];
  });
  check('교육 현황 캐시: 제외하면 바로 반영', cache[0] === '미이수' && cache[1] === '제외', cache);
  await page.evaluate(() => { admSwitchPane('edu'); return BMEDU.board(); });
  check('관리자 화면: 교육 이수 현황판', /교육 이수 현황/.test(await page.textContent('#edu-board')) && /61병동/.test(await page.textContent('#edu-board')));
  // 13-9. v20.3 교육 대상 규칙 (부서 · 직종) · 부서 이동 · 부서 매칭 저장
  const seeded = g.sheets['교육대상규칙'];
  check('대상 규칙: 시트 자동 생성 + 기본 예시 (응급센터 · 응급실 · 8·9병동)', !!seeded && seeded.rows.length === 5 && seeded.rows[3][1] === '8병동,9병동' && seeded.rows[3][2] === '간호조무사', seeded && seeded.rows);
  const EH = empSh.rows[0];
  const col = h => { let i = EH.indexOf(h); if (i < 0) { EH.push(h); i = EH.length - 1; } return i; };
  const empRow = (no, nm, dept, st, pos) => { const r = new Array(EH.length + 2).fill(''); r[col('사번')] = no; r[col('이름')] = nm; r[col('부서')] = dept; r[col('입사일')] = '2020-01-01'; r[col('재직상태')] = st; r[col('직위')] = pos; return r; };
  col('재직상태'); col('직위');
  empSh.appendRow(empRow('126000031', '육간호', '61병동', '재직', '간호사5급'));
  empSh.appendRow(empRow('126000032', '육조무', '61병동', '휴직(육아휴직)', '간호조무사6급'));
  empSh.appendRow(empRow('126000033', '수술이동', '수술실', '재직', '간호사4급'));
  empSh.appendRow(empRow('126000034', '칠이동', '71병동', '재직', '간호사4급'));
  empSh.appendRow(empRow('126000041', '동명가', '61병동', '재직', '간호사5급'));
  empSh.appendRow(empRow('126000042', '동명가', '수술실', '재직', '간호사5급'));
  empSh.appendRow(empRow('126000043', '중복가', '61병동', '', ''));
  empSh.appendRow(empRow('126000043', '중복가', '61병동', '재직', '간호사5급'));
  const mkC = (nm, dept) => { const r = new Array(11).fill(''); r[0] = nm; r[2] = dept; r[3] = '성분'; r[4] = '7732-18-5'; return r; };
  const cOp = mkC('수술물질', '수술실'), c71 = mkC('테스트소독제', '71병동');
  ch.rows.push(cOp, c71);
  led.rows.push(['9', '2026-01-01T01:00Z', 'MSDS교육', '61병동', '테스트소독제', '2026-01-02', '2026-01-02', '1', '1시간', '126000033', '수술이동', '', '✓ 서명완료']);
  led.rows.push(['10', '2026-01-01T01:00Z', 'MSDS교육', '61병동', 'x', '2026-01-02', '2026-01-02', '1', '1시간', '126000034', '칠이동', '', '✓ 서명완료']);
  led.rows.push(['11', '2026-01-01T01:00Z', '특별안전보건교육', '61병동(신관)', '테스트소독제', '2026-01-02', '2026-01-03', '2', '8시간', '126000042', '동명가', '', '✓ 서명완료']);
  const est0 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  const DM = d => (est0.people || []).find(p => p.name === '동명가' && p.dept === d) || {};
  check('교육 기록 주인: 일지에 동명이인 사번 · 부서 표기 61병동(신관) → 61병동 동명가 이수', DM('61병동').status === '이수' && DM('수술실').status === '미이수' && !DM('수술실').moved, [DM('61병동').status, DM('수술실')]);
  check('직원목록에 같은 사번 두 줄 → 한 명만', (est0.people || []).filter(p => p.name === '중복가').length === 1 && Q(est0, '중복가').job === '간호사', (est0.people || []).filter(p => p.name === '중복가'));
  const sv = await page.evaluate(() => BMAUTH.post('edu.saveRules', { rules: [
    { use: true, dept: '61병동', job: '간호조무사', decision: '대상', note: '조무사만 사용' },
    { use: true, dept: '61병동', job: '', decision: '제외', note: '그 외 사용 안 함' },
    { use: true, dept: '', job: '', decision: '제외' } ] }));
  const est7 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('대상 규칙 저장: 빈 줄은 빼고 저장', sv.success && sv.saved === 2 && g.sheets['교육대상규칙'].rows.length === 3, sv);
  check('대상 규칙: 61병동 간호조무사 → 대상 (규칙 1) · 휴직자도 교육 1회 대상', Q(est7, '육조무').status === '미이수' && Q(est7, '육조무').rule === 1 && Q(est7, '육조무').onLeave === true && Q(est7, '육조무').job === '간호조무사', Q(est7, '육조무'));
  check('대상 규칙: 61병동 그 외 직종 → 제외 (사유 = 규칙 2)', Q(est7, '육간호').status === '제외' && /대상 규칙 2/.test(Q(est7, '육간호').excluded.reason), Q(est7, '육간호'));
  check('대상 규칙: 일지를 쓴 사람은 규칙 제외여도 이수', Q(est7, '교육가').status === '이수', Q(est7, '교육가').status);
  check('대상 규칙: 적용 인원 집계', est7.ruleHits[0] === 1 && est7.ruleHits[1] >= 3, est7.ruleHits);
  check('부서 이동: 병동 교육만 있는 수술실 직원 → 수술실 물질 교육 1회 필요', Q(est7, '수술이동').status === '미이수' && Q(est7, '수술이동').moved === '61병동', Q(est7, '수술이동'));
  check('부서 이동: 같은 물질 부서(61→71병동)로 옮기면 다시 안 받아도 이수', Q(est7, '칠이동').status === '이수', Q(est7, '칠이동'));
  // v20.3 개별 포함 (규칙 제외보다 우선) · 신규입사 알림 연동 · 수기(종이) 교육 등록
  const nhSh = g.sheets['신규입사알림'];
  const NHH = nhSh.rows[0];
  const nhRow = new Array(NHH.length).fill(''); nhRow[NHH.indexOf('알림ID')] = 'NH-T1'; nhRow[NHH.indexOf('사번')] = '126000031'; nhRow[NHH.indexOf('이름')] = '육간호'; nhRow[NHH.indexOf('부서')] = '61병동'; nhRow[NHH.indexOf('부서장이메일')] = 'head61@bumin.co.kr'; nhRow[NHH.indexOf('상태')] = '회신대기';
  nhSh.rows.push(nhRow);
  const nhSt = () => nhSh.rows.find(r => r[0] === 'NH-T1')[NHH.indexOf('상태')];
  const incR = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000031', name: '육간호', dept: '61병동', include: true, reason: '인터벤션 근무' }));
  const est9 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('개별 포함: 규칙으로 제외된 사람 → 대상 (규칙보다 우선)', incR.success && Q(est9, '육간호').status === '미이수' && Q(est9, '육간호').included.reason === '인터벤션 근무', Q(est9, '육간호'));
  check('개별 포함 → 신규입사 알림도 취급-교육대기 (부서 확인 창 안 뜸)', nhSt() === '취급-교육대기', nhSt());
  const undoInc = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000031', name: '육간호', dept: '61병동', undo: true }));
  const est10 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('포함 해제 → 다시 규칙대로 제외 · 알림은 회신대기로', undoInc.success && Q(est10, '육간호').status === '제외' && nhSt() === '회신대기', [Q(est10, '육간호').status, nhSt()]);
  const undoRule = await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000031', name: '육간호', dept: '61병동', undo: true }));
  check('규칙 제외를 [제외 해제]하면 안내 (대상으로 포함 쓰기)', !undoRule.success && /대상으로 포함/.test(undoRule.error), undoRule);
  await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000031', name: '육간호', dept: '61병동', reason: '행정 업무만' }));
  check('개별 제외 → 신규입사 알림 미취급', nhSt() === '미취급', nhSt());
  await page.evaluate(() => BMAUTH.post('edu.exclude', { no: '126000031', name: '육간호', dept: '61병동', undo: true }));
  const ledN = led.rows.length;
  const md = await page.evaluate(() => BMAUTH.post('edu.manual', { dry: true, rows: [
    { name: '육조무', dept: '61병동', type: 'MSDS', ds: '2024.3.4', hours: '1', note: '2024 종이일지' },
    { name: '동명가', ds: '2024-03-04' },
    { name: '없는사람', dept: '61병동', ds: '2024-03-04' },
    { name: '육간호', dept: '61병동', type: '특별', ds: '2024-03-04', hours: '' },
    { no: '126000033', name: '수술이동', dept: '수술실', ds: '2099-01-01' } ] }));
  const R = i => md.rows[i];
  check('수기 등록 미리보기: 날짜 2024.3.4 → 2024-03-04 · 동명이인 · 못 찾음 · 특별은 시간 필수 · 미래 날짜', md.success && R(0).ok && R(0).ds === '2024-03-04' && /동명이인/.test(R(1).reason) && /못 찾음/.test(R(2).reason) && /시간/.test(R(3).reason) && /이후/.test(R(4).reason) && led.rows.length === ledN, md.rows.map(r => r.reason || r.ds));
  const mg = await page.evaluate(() => BMAUTH.post('edu.manual', { rows: [{ name: '육조무', dept: '61병동', type: '특별', ds: '2024.3.4', de: '2024.3.5', hours: '8', note: '2024 종이일지' }] }));
  const last = led.rows[led.rows.length - 1];
  const est11 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('수기 등록: 교육이력대장에 한 줄 (서명여부 = 수기(종이) 일지) → 이수', mg.added === 1 && last[10] === '육조무' && last[9] === '126000032' && /수기\(종이\) 일지 · 2024 종이일지/.test(last[12]) && last[2] === '특별안전보건교육' && last[7] === '2' && Q(est11, '육조무').status === '이수' && Q(est11, '육조무').hours === 16, [last, Q(est11, '육조무')]);
  const mg2 = await page.evaluate(() => BMAUTH.post('edu.manual', { rows: [{ name: '육조무', dept: '61병동', type: '특별', ds: '2024-03-04', hours: '8' }] }));
  check('수기 등록: 같은 날 같은 교육은 중복으로 건너뜀', mg2.added === 0 && /이미/.test(mg2.rows[0].reason), mg2.rows[0]);
  const mgDeny = await page.evaluate(async () => { const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' }); return (await BMAUTH.post('edu.manual', { rows: [], t: r.token })).status; });
  check('수기 등록은 관리자만', mgDeny === 403, mgDeny);
  await page.evaluate(async () => { admSwitchPane('edu'); await new Promise(r => setTimeout(r, 50)); await BMEDU.board(); document.querySelector('#edu-board [data-dept="61병동"]').click(); });
  await page.waitForFunction(() => !/계산 중/.test(document.getElementById('edu-board').textContent) && document.querySelector('#eb-people [data-pick]'), null, { timeout: 15000 });
  const ui2 = await page.evaluate(async () => {
    const f = document.getElementById('eb-f'); f.value = ''; f.dispatchEvent(new Event('change'));
    document.getElementById('eb-manual-btn').click();
    const box = [...document.querySelectorAll('#eb-people [data-pick]')].find(c => /미이수|진행중/.test(c.closest('tr').textContent));
    const inc = !!document.querySelector('#eb-people [data-inc]');
    box.checked = true;
    document.getElementById('mn-ds').value = '2023-05-05';
    document.getElementById('mn-dry').click();
    const t0 = Date.now();
    while (Date.now() - t0 < 6000 && !/등록 가능/.test(document.getElementById('mn-sum').textContent)) await new Promise(r => setTimeout(r, 100));
    return { sum: document.getElementById('mn-sum').textContent, go: !document.getElementById('mn-go').disabled, inc };
  });
  check('관리자 화면: 현황판에서 체크 → 수기 등록 미리보기 · 규칙 제외자는 [대상으로 포함] 버튼', /등록 가능 1명/.test(ui2.sum) && ui2.go && ui2.inc, ui2);
  const rowMan = await page.evaluate(async () => {
    const b = document.querySelector('#eb-people [data-man]');
    if (!b) return { none: true };
    const who = b.closest('tr').querySelector('b').textContent;
    const besideX = !!b.parentNode.querySelector('[data-x]');
    b.click();
    const box = b.parentNode.querySelector('.mn-inline');
    box.querySelector('[data-f="ds"]').value = '2023-04-05';
    box.querySelector('[data-f="h"]').value = '8';
    box.querySelector('[data-f="ok"]').click();
    const t0 = Date.now();
    let st = null;
    while (Date.now() - t0 < 8000) {
      await new Promise(r => setTimeout(r, 200));
      const r = await BMAUTH.post('edu.status', { fresh: true });
      st = (r.people.find(p => p.name === who) || {}).status;
      if (st === '이수') break;
    }
    return { besideX, st, who };
  });
  const lastMan = led.rows[led.rows.length - 1];
  check('현황판 줄마다 [수기 이수] (취급 안 함 옆) → 교육일 · 시간 넣고 이수 처리', rowMan.st === '이수' && lastMan[10] === rowMan.who && lastMan[5] === '2023-04-05' && /수기\(종이\) 일지 · 예전 종이 일지/.test(lastMan[12]), [rowMan, lastMan]);
  // v20.3 근무지 지정: 규칙으로 제외된 61병동 직원을 근무지 '수술실'로 → 수술실 물질 · 교육으로 판정 (부서 규칙보다 우선)
  const un1 = await page.evaluate(() => BMAUTH.post('edu.unit', { rows: [{ no: '126000031', name: '육간호', dept: '61병동', unit: '수술실' }] }));
  const est12 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('근무지 지정: 61병동 직원 → 근무지 수술실 (규칙 제외 무시, 수술실 물질로 대상)', un1.success && un1.set === 1 && Q(est12, '육간호').unit === '수술실' && Q(est12, '육간호').status === '미이수' && !!g.sheets['근무지지정'], Q(est12, '육간호'));
  const un0 = await page.evaluate(() => BMAUTH.post('edu.unit', { rows: [{ no: '126000031', name: '육간호', dept: '61병동', unit: '' }] }));
  const est13 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('근무지 해제 → 다시 부서 규칙대로', un0.cleared === 1 && Q(est13, '육간호').status === '제외' && !Q(est13, '육간호').unit, Q(est13, '육간호').status);
  // 전 부서 일괄 수기 이수 (기준일 · 입사일 맞춤)
  const bulk = await page.evaluate(async () => {
    await BMEDU.board();
    document.getElementById('eb-bulk-btn').click();
    document.getElementById('bk-ds').value = '2025-01-01';
    document.getElementById('bk-dry').click();
    let t0 = Date.now();
    while (Date.now() - t0 < 8000 && document.getElementById('bk-go').disabled) await new Promise(r => setTimeout(r, 100));
    const sum = document.getElementById('bk-sum').textContent;
    document.getElementById('bk-go').click();
    t0 = Date.now();
    let r;
    while (Date.now() - t0 < 10000) { await new Promise(x => setTimeout(x, 300)); r = await BMAUTH.post('edu.status', { fresh: true }); if (!r.people.some(p => p.status === '미이수' || p.status === '진행중')) break; }
    return { sum, left: r.people.filter(p => p.status === '미이수' || p.status === '진행중').map(p => p.dept + ' ' + p.name) };
  });
  const bulkRows = led.rows.filter(r => /일괄/.test(String(r[12])));
  check('전 부서 일괄 수기 이수: 미리보기 → 처리 → 할 일 0명 · 입사일이 늦은 사람은 입사일로', /등록 \d+명/.test(bulk.sum) && bulk.left.length === 0 && bulkRows.length > 0 && bulkRows.every(r => r[5] >= '2025-01-01') && bulkRows.some(r => r[5] === '2025-01-01'), [bulk, bulkRows.length]);
  // v20.3 규칙 제외는 일괄 수기 기록보다 우선 (치과 의사 사례) · 수기 기록 삭제(백업) · 앱 일지는 삭제 불가
  const onlyBulk = bulkRows.map(r => r[10]).find(nm => led.rows.filter(r => r[10] === nm).every(r => /일괄/.test(String(r[12]))));
  const bulkDept = bulkRows.find(r => r[10] === onlyBulk)[3];
  const rulesBak = (await page.evaluate(() => BMAUTH.post('edu.rules', {}))).rules;
  await page.evaluate(d => BMAUTH.post('edu.saveRules', { rules: [{ use: true, dept: d, job: '', decision: '제외', note: '의사 — 사용 안 함' }] }), bulkDept);
  const est14 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  const BP = Q(est14, onlyBulk);
  check('규칙 제외 > 일괄 수기 기록: 일괄로 이수 처리된 사람도 규칙으로 제외 (기록 보존 표시)', BP.status === '제외' && BP.kept && BP.kept.n >= 1 && BP.manualRows.length >= 1 && BP.manualRows[0].bulk, BP);
  const appRow = led.rows.findIndex(r => !/^수기/.test(String(r[12])) && r[10]) + 1;
  const ledLen = led.rows.length;
  const del = await page.evaluate(([items, app]) => BMAUTH.post('edu.manualDelete', { reason: '시험', items: items.concat([app]) }),
    [BP.manualRows.map(r => ({ row: r.row, name: onlyBulk, ds: r.ds })), { row: appRow, name: led.rows[appRow - 1][10], ds: led.rows[appRow - 1][5] }]);
  const bak = g.sheets['교육이력_삭제'];
  check('수기 기록 삭제: 수기 줄만 삭제 · 앱 일지는 건너뜀 · 교육이력_삭제 시트에 백업(사유)', del.success && del.deleted === BP.manualRows.length && del.skipped === 1 && led.rows.length === ledLen - del.deleted && !led.rows.some(r => r[10] === onlyBulk) && !!bak && bak.rows.some(r => r.includes(onlyBulk) && r.includes('시험')), [del, ledLen, led.rows.length]);
  await page.evaluate(r => BMAUTH.post('edu.saveRules', { rules: r }), rulesBak);
  const oldSrv = await page.evaluate(async () => {
    const orig = BMAUTH.post;
    BMAUTH.post = async (op, b) => { const r = await orig(op, b); if (op === 'edu.status') { delete r.serverVersion; r.people.forEach(p => delete p.manualRows); } return r; };
    try {
      await BMEDU.board();
      const banner = /이전 버전/.test(document.getElementById('edu-board').textContent);
      document.querySelector('#edu-board [data-dept]').click();
      const c = document.querySelector('#eb-people [data-pick]'); c.checked = true;
      const msgs = [];
      const t0 = window.toast; window.toast = (m) => msgs.push(m);
      document.getElementById('eb-mdel-go').click();
      window.toast = t0;
      return { banner, msg: msgs.join(' ') };
    } finally { BMAUTH.post = orig; await BMEDU.board(); }
  });
  check('서버가 이전 버전이면: 현황판 맨 위 배포 안내 · [수기 기록 삭제]는 배포 필요 메시지', oldSrv.banner && /이전 버전/.test(oldSrv.msg) && /배포/.test(oldSrv.msg), oldSrv);
  led.rows.splice(ledN);
  // v20.3 직원목록 ↔ 인사 사이트 합치기: #N/A · 빈 사번 채움, 입사일 표기 달라도(9/1) 같은 부서 같은 이름이면 같은 사람, 같은 사번 두 줄 합침
  const r1 = empRow('#N/A', '엔에이', '71병동', '재직', ''); empSh.appendRow(r1);
  const r2 = empRow('', '빈사번', '71병동', '재직', ''); r2[EH.indexOf('입사일')] = '9/1'; empSh.appendRow(r2);
  const empN0 = empSh.rows.length;
  const mrgRows = [['126000051', '엔에이', '71병동', '2020-01-01'], ['126000052', '빈사번', '71병동', '2026-09-01']];
  const dryM = await page.evaluate(rows => BMAUTH.post('emp.sync', { headers: ['사번', '이름', '부서', '입사일'], rows, dry: true }), mrgRows);
  check('합치기 미리보기: 사번 채움 2 · 같은 사번 두 줄 합침(중복가) · 시트는 그대로', dryM.success && dryM.dry && dryM.filledNo.length === 2 && dryM.added === 0 && dryM.merged.some(m => m.name === '중복가') && dryM.deleted >= 1 && empSh.rows.length === empN0 && empSh.rows.some(r => r[0] === '#N/A'), dryM);
  const realM = await page.evaluate(rows => BMAUTH.post('emp.sync', { headers: ['사번', '이름', '부서', '입사일'], rows }), mrgRows);
  const rowOfName = nm => empSh.rows.filter(r => r[EH.indexOf('이름')] === nm);
  check('합치기 적용: #N/A → 인사 사번, 빈 사번 + 입사일 9/1 → 사번 · 입사일 채움 (새 줄 없음)', realM.success && realM.added === 0 && rowOfName('엔에이').length === 1 && rowOfName('엔에이')[0][0] === '126000051' && rowOfName('빈사번').length === 1 && rowOfName('빈사번')[0][0] === '126000052' && rowOfName('빈사번')[0][EH.indexOf('입사일')] === '2026-09-01', [rowOfName('엔에이'), rowOfName('빈사번')]);
  check('합치기 적용: 같은 사번 두 줄 → 한 줄 (빈 칸은 다른 줄 값)', rowOfName('중복가').length === 1 && rowOfName('중복가')[0][EH.indexOf('직위')] === '간호사5급' && rowOfName('중복가')[0][EH.indexOf('재직상태')] === '재직', rowOfName('중복가'));
  const mergeWarn = await page.evaluate(() => BMAUTH.post('emp.merge', { dry: true, fresh: true }));
  const mergeAuto = await page.evaluate(() => BMAUTH.post('emp.merge', { dry: true, force: true, fresh: true }));
  check('인사 사이트와 합치기: 인원이 크게 줄면 경고 후 멈춤 → 확인하면 미리보기', !mergeWarn.success && mergeWarn.warnings.length > 0 && mergeAuto.success && mergeAuto.dry && mergeAuto.hrCount > 0 && Array.isArray(mergeAuto.missingNo), [mergeWarn.error, mergeAuto.error || mergeAuto.hrCount]);
  const info = await page.evaluate(() => BMAUTH.post('edu.rules', {}));
  check('규칙 화면 정보: 규칙 · 적용 인원 · 부서별 직종 · 부서 매칭', info.success && info.rules.length === 2 && info.hits[0] === 1 && info.jobs['61병동']['간호조무사'] === 1 && info.deptMap.some(m => m.raw === '진검실' && m.app === '61병동' && m.staff === 1), info.jobs);
  const msv = await page.evaluate(() => BMAUTH.post('dept.saveMap', { rows: [{ raw: '진검실', app: '71병동' }, { raw: '새부서명', app: '71병동' }] }));
  const est8 = await page.evaluate(() => BMAUTH.post('edu.status', { fresh: true }));
  check('부서 매칭 저장: 바꾸면 바로 반영 · 없는 이름은 새 줄', msv.success && msv.changed === 2 && Q(est8, '별칭가').dept === '71병동' && dm.rows.some(r => r[0] === '새부서명'), [msv, Q(est8, '별칭가').dept]);
  const deptDeny = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' });
    const a = await BMAUTH.post('edu.saveRules', { rules: [], t: r.token });
    const b = await BMAUTH.post('dept.saveMap', { rows: [], t: r.token });
    return [a.status, b.status];
  });
  check('규칙 · 부서 매칭 저장은 관리자만', deptDeny[0] === 403 && deptDeny[1] === 403, deptDeny);
  await page.evaluate(() => { admSwitchPane('staff'); return BMSTAFF.open(); });
  await page.waitForFunction(() => document.querySelectorAll('#rl-tbl tr').length === 3, null, { timeout: 15000 });
  const ui = await page.evaluate(() => ({ t: document.getElementById('staff-rules').textContent, maps: document.querySelectorAll('#staff-rules [data-map]').length }));
  check('관리자 화면: 직원정보관리 ④ 규칙 표 + 부서 매칭 표', /ⓐ 교육 대상 규칙/.test(ui.t) && /ⓑ 부서 매칭/.test(ui.t) && ui.maps >= 2 && /1명/.test(ui.t), ui.t.slice(0, 200));
  await page.evaluate(() => BMAUTH.post('edu.saveRules', { rules: [] }));
  ch.rows.splice(ch.rows.indexOf(cOp), 1); ch.rows.splice(ch.rows.indexOf(c71), 1);
  led.rows = ledBak;
  ch.rows.splice(ch.rows.indexOf(spRow), 1); ch.rows[0] = chHead;

  // 메모1 · 메모2 → MSDS 목록 데이터, 부서장 시트
  ch.rows[0][8] = '메모1'; ch.rows[0][9] = '메모2';
  const mi = ch.rows.findIndex(r => r[0] === '메탄올A');
  ch.rows[mi][8] = '26년 상반기부터 작측 제외'; ch.rows[mi][9] = '혈액학';
  const gc = vm.runInContext('getChemicals()', g.ctx);
  const ma = gc.data.find(c => c.name === '메탄올A');
  check('MSDS 목록: 메모1 · 메모2 읽기', ma && ma.components[0].memo === '26년 상반기부터 작측 제외' && ma.components[0].memo2 === '혈액학', ma && ma.components[0]);
  const hs = g.ctx.SpreadsheetApp.getActiveSpreadsheet().insertSheet('부서장');
  hs.appendRow(['부서', '부서장명', '이메일']); hs.appendRow(['61병동', '박부서장', '']);
  check('부서장: 부서장 시트 우선', vm.runInContext("deptHead_('61병동')", g.ctx) === '박부서장');

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
  const dv5 = await page.evaluate(() => {
    const sum = [...document.querySelectorAll('#dv-body .dv-card b')].map(b => b.textContent);
    const vis = () => [...document.querySelectorAll('#dv-body .dv-sec')].filter(x => x.style.display !== 'none').map(x => x.getAttribute('data-sec'));
    const all = vis();
    document.querySelector('#dv-body .dv-tab[data-tab="ppe"]').click(); const ppe = vis();
    document.querySelector('#dv-body .dv-tab[data-tab="all"]').click();
    const q = document.getElementById('dv-q'); q.value = '없는이름'; q.oninput();
    const hidden = [...document.querySelectorAll('#dv-body .dv-tbl tr')].filter(tr => tr.querySelector('td') && tr.style.display !== 'none').length;
    q.value = '김가나'; q.oninput();
    const shown = [...document.querySelectorAll('#dv-body .dv-tbl tr')].filter(tr => tr.querySelector('td') && tr.style.display !== 'none').length;
    q.value = ''; q.oninput();
    return { sum, all, ppe, hidden, shown };
  });
  check('부서현황: 요약 (처리 필요 0 · 결재 대기 1명 · 완료 0명)', dv5.sum[0] === '0명' && dv5.sum[1] === '1명' && dv5.sum[2] === '0명', dv5.sum);
  check('부서현황: 종류 탭', dv5.all.join(',') === 'edu,ppe,chem' && dv5.ppe.join(',') === 'ppe', dv5);
  check('부서현황: 이름 검색', dv5.hidden === 0 && dv5.shown === 1, dv5);
  const rjDept = await page.evaluate(async () => {
    const orig = window.readGitHubJSON;
    window.readGitHubJSON = async (path) => {
      const r = await orig(path);
      if (path === 'records/edu_records.json') r.data = (r.data || []).map(x => Object.assign({}, x, { rejected: true, rejectReason: '교육 시간 누락' }));
      return r;
    };
    try{
      await BMDEPT.render();
      const dv = document.getElementById('dv-body').textContent;
      const sum0 = document.querySelector('#dv-body .dv-card b').textContent;
      const redo = !!document.querySelector('#dv-body [data-redo]');
      goTab('home');
      const t0 = Date.now();
      while (Date.now() - t0 < 8000 && document.getElementById('hm-todo').textContent.includes('불러오는 중')) await new Promise(r => setTimeout(r, 100));
      const home = document.getElementById('hm-todo').textContent;
      return { dv: dv.includes('반려') && dv.includes('사유: 교육 시간 누락'), sum0, redo, home };
    } finally {
      window.readGitHubJSON = orig;
      goTab('deptview');
      await BMDEPT.render();
    }
  });
  check('반려 → 부서현황: 처리 필요 · 사유 · 다시 작성', rjDept.dv && rjDept.sum0 === '1명' && rjDept.redo, rjDept);
  check('반려 → 홈 할 일: 반려 · 다시 작성', rjDept.home.includes('반려') && rjDept.home.includes('다시 작성') && rjDept.home.includes('교육 시간 누락'), rjDept.home);
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

  // 18. v21.0 AI 위험성평가: 부서 로그인 → 사진 AI 분석(점수 · 등급은 서버 계산, 법령은 기준규칙 조문 목록과 대조) → 제출 → 메일 · 시트 → 관리자 처리 → 부서 화면에 결과
  g.props.LAW_OC = 'bumin3131';
  g.sheets['관리자설정'].appendRow(['위험성평가알림이메일', 'safety@bumin.co.kr', '']);
  const tk = await page.evaluate(async () => ({
    d: (await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' })).token,
    a: (await BMAUTH.post('login', { kind: 'admin', password: 'admin-pass-1' })).token }));
  const jpg = 'data:image/jpeg;base64,' + Buffer.from('fakejpeg').toString('base64');
  const raNo = await page.evaluate(() => BMAUTH.post('ra.analyze', { image: 'x', t: '' }));
  check('위험성평가: 로그인 없으면 거부', !raNo.success && raNo.status === 401, raNo);
  const gmCalls = gm.calls.length;
  gm.reply = { scene: '처치실 바닥', privacy: '', quality: '', goodPractice: '',
    hazards: [{ category: '작업환경', hazard: '바닥 물기', situation: '이동 중 미끄러져 넘어짐', evidence: '바닥에 물 고임', freq: 4, sev: 3, score: 99, grade: '낮음',
                controls: [{ type: '관리적', text: '물기 제거 · 미끄럼 주의 표지', who: '근로자 즉시' }, { type: '공학적', text: '미끄럼 방지 매트', who: '관리자' }],
                afterFreq: 2, afterSev: 3, laws: ['제3조', '제9999조'] },
              { category: '생물', hazard: '뚜껑 없는 주사침 용기', situation: '자상', evidence: '용기 넘침', freq: 9, sev: 0, controls: [], laws: ['제597조'] }] };
  const ra1 = await page.evaluate(a => BMAUTH.post('ra.analyze', { image: a.jpg, place: '61병동 처치실', task: '청소 후', t: a.t }), { jpg, t: tk.d });
  const sentPrompt = gm.calls.length > gmCalls ? gm.calls[gm.calls.length - 1].body.contents[0].parts : [];
  check('위험성평가 AI: 사진(image/jpeg) · 장소만 보냄 (부서 · 이름 없음) · 기준규칙 조문 목록 포함', sentPrompt[0] && sentPrompt[0].inline_data.mime_type === 'image/jpeg' && /61병동 처치실/.test(sentPrompt[1].text) && /제3조 전도의 방지/.test(sentPrompt[1].text) && !/작성자/.test(sentPrompt[1].text), (sentPrompt[1] || {}).text && sentPrompt[1].text.slice(0, 80));
  const h0 = (ra1.hazards || [])[0] || {}, h1 = (ra1.hazards || [])[1] || {};
  check('위험성평가 AI: 점수 · 등급 서버 재계산(4×3=12 높음) · 범위 밖 값 맞춤 · 위험한 순 · 분류 보정', ra1.success && h0.score === 12 && h0.grade === '높음' && h0.afterScore === 6 && h1.freq === 5 && h1.sev === 1 && h1.category === '생물학적', ra1);
  check('위험성평가 AI: 법령은 실제 조문만 (제9999조 버림 · 제목은 법령에서)', h0.laws.length === 1 && h0.laws[0].no === '제3조' && h0.laws[0].title === '전도의 방지' && h1.laws[0].title === '혈액노출 예방 조치', [h0.laws, h1.laws]);
  const mailN = mail.sent.length;
  const raHz = ra1.hazards.map(h => Object.assign({}, h, { freq: h === ra1.hazards[0] ? 5 : h.freq, picked: [0], extra: '매트 설치 요청' }));
  const raSub0 = await page.evaluate(a => BMAUTH.post('ra.submit', { image: a.jpg, place: '61병동 처치실', hazards: a.hz, ai: true, author: '', t: a.t }), { jpg, hz: raHz, t: tk.d });
  check('위험성평가 제출: 작성자 없으면 거부', !raSub0.success && /작성자/.test(raSub0.error), raSub0);
  const sub = await page.evaluate(a => BMAUTH.post('ra.submit', { image: a.jpg, place: '61병동 처치실', task: '청소 후', hazards: a.hz, ai: true, immediate: '물기 닦음', author: '박간호', t: a.t }), { jpg, hz: raHz, t: tk.d });
  const raSh = g.sheets['위험성평가'], raRow = raSh.rows[raSh.rows.length - 1], raH = raSh.rows[0];
  const rv = k => raRow[raH.indexOf(k)];
  check('위험성평가 제출: 근로자가 고친 값으로 재계산(5×3=15 매우높음) · 시트 · 사진 저장 · 메일(사진 첨부)', sub.success && sub.maxScore === 15 && sub.maxGrade === '매우높음' && rv('부서') === '61병동' && rv('작성자') === '박간호' && /records\/ra\//.test(rv('사진')) && rv('상태') === '접수' &&
    mail.sent.length === mailN + 1 && mail.sent[mailN].to === 'safety@bumin.co.kr' && /매우높음/.test(mail.sent[mailN].subject) && (mail.sent[mailN].attachments || []).length === 1, [sub, raRow, mail.sent.slice(mailN).map(m => m.subject)]);
  const raDet = JSON.parse(rv('상세(JSON)'));
  check('위험성평가 제출: AI 원래 값 · 선택한 대책 · 근로자 제안 보관', raDet.hazards[0].aiFreq === 4 && raDet.hazards[0].freq === 5 && raDet.hazards[0].picked[0] === 0 && raDet.hazards[0].extra === '매트 설치 요청', raDet.hazards[0]);
  raRow[raH.indexOf('부서')] = '71병동';   // 다른 부서 제출로 바꿔 권한 확인
  const deny2 = await page.evaluate(a => Promise.all([BMAUTH.post('ra.list', { t: a.d }), BMAUTH.post('ra.photo', { id: a.id, t: a.d }), BMAUTH.post('ra.mine', { t: a.d })]), { d: tk.d, id: sub.id });
  raRow[raH.indexOf('부서')] = '61병동';
  check('위험성평가 권한: 부서는 전체 목록 불가 · 다른 부서 사진 · 내역 안 보임', deny2[0].status === 403 && deny2[1].status === 403 && !deny2[2].data.some(x => x.id === sub.id), deny2.map(x => x.status || x.data && x.data.length));
  const raUp = await page.evaluate(a => BMAUTH.post('ra.update', { id: a.id, status: '개선중', owner: '시설팀', due: '2026-11-05', comment: '매트 구매 요청함', t: a.a }), { id: sub.id, a: tk.a });
  const raMine = await page.evaluate(a => BMAUTH.post('ra.mine', { t: a }), tk.d);
  const raPh = await page.evaluate(a => BMAUTH.post('ra.photo', { id: a.id, t: a.d }), { id: sub.id, d: tk.d });
  check('위험성평가 처리: 관리자 상태 · 담당 · 예정일 · 의견 → 부서 내역에 표시 · 자기 부서 사진 보기', raUp.success && raMine.data[0].status === '개선중' && raMine.data[0].comment === '매트 구매 요청함' && raMine.data[0].due === '2026-11-05' && raPh.success && /^data:image\/jpeg;base64,/.test(raPh.image), [raUp, raMine.data[0]]);
  const raUi = await page.evaluate(async () => {
    const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' });
    goTab('home'); await new Promise(z => setTimeout(z, 200));
    const tile = !!document.querySelector('#hm-tiles [data-go="risk"]');
    goTab('risk'); await new Promise(z => setTimeout(z, 1500));
    const v = { tile, page: getComputedStyle(document.getElementById('pg-risk')).display, drop: !!document.getElementById('ra-drop'), mine: document.getElementById('ra-mine') ? document.getElementById('ra-mine').textContent : '' };
    document.getElementById('ra-manual').click(); await new Promise(z => setTimeout(z, 100));
    v.manual = document.querySelectorAll('[data-hz]').length;
    const box = document.querySelector('[data-hz="0"]'); box.querySelector('[data-k="freq"]').value = '5'; box.querySelector('[data-k="sev"]').value = '4'; box.querySelector('[data-k="freq"]').onchange();
    v.score = box.querySelector('[data-score]').textContent;
    admSwitchPane && goTab('admin');
    return v;
  });
  check('위험성평가 화면: 홈 타일 · 탭 · 사진 칸 · 우리 부서 내역(처리 의견) · 직접 입력 시 점수 즉시 계산', raUi.tile && raUi.page === 'block' && raUi.drop && /매트 구매 요청함/.test(raUi.mine) && raUi.manual === 1 && /20점 매우높음/.test(raUi.score), raUi);
  const raAdm = await page.evaluate(async () => {
    await BMAUTH.adminLogin('admin-pass-1');
    admSwitchPane('risk'); await BMRISK.admin(); await new Promise(z => setTimeout(z, 300));
    const tr = document.querySelector('#ra-adm-body tr.hit'); if(tr) tr.click();
    await new Promise(z => setTimeout(z, 800));
    return { rows: document.querySelectorAll('#ra-adm-body tr.hit').length, kpi: document.querySelector('#ra-adm-body .ra-kpi').textContent, img: !!document.querySelector('#ra-adm-body img'), save: !!document.querySelector('#ra-adm-body [data-save]') };
  });
  check('위험성평가 관리자 화면: 목록 · 지표 · 사진 · 처리 저장 칸', raAdm.rows >= 1 && /높음 이상 미완료/.test(raAdm.kpi) && raAdm.img && raAdm.save, raAdm);

  // 19. v21.1 AI 도우미 · 안전보건공단 OPEN API · 법령 본문
  const ks = g.env.KOSHA;
  const tka = (await page.evaluate(() => BMAUTH.post('login', { kind: 'admin', password: 'admin-pass-1' }))).token;
  const tkd = (await page.evaluate(() => BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' }))).token;
  const ksDeny = await page.evaluate(a => BMAUTH.post('kosha.setKey', { key: 'koshaKey1234567890abcdefXYZ', t: a }), tkd);
  check('공단 API: 부서는 키 저장 불가', !ksDeny.success && ksDeny.status === 403, ksDeny);
  const ksBad = await page.evaluate(a => BMAUTH.post('kosha.saveApis', { apis: [{ name: '나쁜 주소', url: 'http://evil.example.com/x' }], t: a }), tka);
  check('공단 API: 요청주소는 apis.data.go.kr 만', !ksBad.success && /apis\.data\.go\.kr/.test(ksBad.error), ksBad);
  g.props.KOSHA_KEY = 'wrong-key';
  const ksSt0 = await page.evaluate(a => BMAUTH.post('kosha.status', { t: a }), tka);
  check('API 기록: 처음엔 기본 2줄 (안전보건자료 링크 · 활용기간 2026-10-23~2028-10-23 / 재해사례 꺼짐)', ksSt0.success && ksSt0.apis.length === 2 && ksSt0.apis[0].to === '2028-10-23' && ksSt0.apis[0].use && !ksSt0.apis[1].use, ksSt0);
  const ksWrong = await page.evaluate(a => BMAUTH.post('kosha.test', { id: 'media', keyword: '포르말린', t: a }), tka);
  check('공단 API 시험: 틀린 키 → 공단 오류 문구', !ksWrong.success && /SERVICE_KEY/.test(ksWrong.error), ksWrong);
  const ksSet = await page.evaluate(a => Promise.all([
    BMAUTH.post('kosha.setKey', { key: 'koshaKey1234567890abcdefXYZ', t: a }),
    BMAUTH.post('kosha.saveApis', { apis: [
      { id: 'media', name: '안전보건자료 링크서비스', url: 'https://apis.data.go.kr/B552468/selectMediaList01', kw: 'keyword', kind: '자료', use: true, from: '2026-10-23', to: '2028-10-23', memo: '참고문서 13' },
      { id: 'disaster', name: '국내재해사례', url: 'https://apis.data.go.kr/B552468/disaster_api02/getdisaster_api02', kw: 'keyword', extra: 'callApiId=1060', kind: '재해사례', use: true },
      { name: '새 API', url: '', kind: '없는종류', use: true, to: '2027-13-99' }], t: a })]), tka);
  ks.okKey = 'koshaKey1234567890abcdefXYZ';
  const apRec = ksSet[1].apis || [];
  check('공단 API: 키 · API 기록 저장 (키 가림 · 주소 없는 줄은 사용 꺼짐 · 종류 · 날짜 바로잡음 · id 부여)', ksSet[0].success && ksSet[0].key === 'kosh…XYZ' && ksSet[1].success && apRec.length === 3 && apRec[1].extra === 'callApiId=1060' && !apRec[2].use && apRec[2].kind === '기타' && apRec[2].to === '' && !!apRec[2].id && g.props.KOSHA_KEY === 'koshaKey1234567890abcdefXYZ', ksSet);
  const ksMedia = await page.evaluate(a => BMAUTH.post('kosha.test', { id: 'media', keyword: '화학물질', t: a }), tka);
  check('공단 API 시험: End Point가 서비스 이름까지만이어도 get+서비스이름으로 다시 시도', ksMedia.success || /selectMediaList01/.test(ks.calls[ks.calls.length - 1]), [ksMedia, ks.calls.slice(-2)]);
  const ksOk = await page.evaluate(a => BMAUTH.post('kosha.test', { id: 'disaster', keyword: '포르말린', t: a }), tka);
  const ksUrl = ks.calls[ks.calls.length - 1] || '';
  check('공단 API 시험: 목록 · 제목 · 링크 읽기 + 고정 변수 · 검색어 붙임', ksOk.success && ksOk.items.length === 2 && /포르말린 취급/.test(ksOk.items[0].title) && /kosha/.test(ksOk.items[0].url) && /callApiId=1060/.test(ksUrl) && /keyword=%ED%8F%AC/.test(ksUrl), [ksOk, ksUrl]);
  const meta = await page.evaluate(() => BMAUTH.post('help.meta', {}));
  check('AI 도우미: 로그인 없이 상태 확인', meta.success && meta.hasKey && meta.kosha, meta);
  const gmN = gm.calls.length;
  gm.status = 200;
  gm.reply = { answer: '포르말린이 눈에 들어가면 즉시 15분 이상 흐르는 물로 씻고 진료를 받으세요.', used: [1, 99], followups: ['포르말린 보호구는?'] };
  const help1 = await page.evaluate(() => BMAUTH.post('help.ask', { q: '포르말린 넘어짐 바닥 미끄러지는 위험 010-1234-5678', history: [{ role: 'user', text: '안녕' }],
    msds: [{ name: '10% 포르말린', text: '응급조치: [눈] 15분 이상 씻기' }] }));
  const hp = gm.calls.length > gmN ? gm.calls[gm.calls.length - 1].body.contents[0].parts[0].text : '';
  check('AI 도우미: 답 · 출처(번호 범위 확인) · 이어서 물을 질문', help1.success && /15분/.test(help1.answer) && help1.sources.length === 1 && help1.followups.length === 1, help1);
  check('AI 도우미: AI에 보낸 자료 = 사용 안내 · MSDS · 법령 본문(제3조) · 공단 자료, 전화번호는 가림', /\(사용 안내\)/.test(hp) && /\(MSDS\) 10% 포르말린/.test(hp) && /전도의 방지/.test(hp) && /공단 (자료|재해사례)/.test(hp) && !/010-1234-5678/.test(hp) && /전화번호 가림/.test(hp), hp.slice(0, 600));
  g.props.GEMINI_API_KEY = '';
  const helpNo = await page.evaluate(() => BMAUTH.post('help.ask', { q: '교육일지 작성법' }));
  g.props.GEMINI_API_KEY = 'gem-key';
  check('AI 도우미: AI 키 없으면 안내 문구', !helpNo.success && helpNo.noKey, helpNo);
  const helpUi = await page.evaluate(async () => {
    const b = document.getElementById('bm-help-btn'); if(!b) return { btn: false };
    b.click(); await new Promise(z => setTimeout(z, 100));
    const box = document.getElementById('bm-help');
    const v = { btn: true, open: box.classList.contains('on'), chips: box.querySelectorAll('.bh-chip').length };
    box.querySelector('textarea').value = '포르말린 눈에 들어가면?';
    box.querySelector('.bh-in button').click();
    await new Promise(z => setTimeout(z, 1500));
    v.answer = (box.querySelectorAll('.bh-m.a')[box.querySelectorAll('.bh-m.a').length - 1] || {}).textContent || '';
    BMHELP.close();
    return v;
  });
  check('AI 도우미 화면: 버튼 · 창 · 빠른 질문 · 답과 출처 표시', helpUi.btn && helpUi.open && helpUi.chips >= 4 && /15분/.test(helpUi.answer) && /출처/.test(helpUi.answer), helpUi);
  const lawPrompt = gm.calls[gm.calls.length - 1].body.contents[0].parts[0].text;
  check('AI 도우미 화면: 질문한 물질의 MSDS 요약을 화면 목록에서 찾아 보냄', /\(MSDS\)/.test(lawPrompt), lawPrompt.slice(0, 300));
  gm.reply = { scene: '처치실', hazards: [{ category: '작업환경', hazard: '바닥 물기', situation: '넘어짐', freq: 3, sev: 2, controls: [], laws: [] }] };
  await page.evaluate(a => BMAUTH.post('ra.analyze', { image: a.jpg, place: '61병동 처치실', task: '청소 후', t: a.t }), { jpg, t: tkd });
  const raP = gm.calls[gm.calls.length - 1].body.contents[0].parts[1].text;
  check('위험성평가 AI: 우리 병원 처리된 평가(관리자 의견) · 공단 재해사례를 참고자료로 보냄', /우리 병원 평가/.test(raP) && /매트 구매 요청함/.test(raP) && /공단 재해사례/.test(raP), raP.slice(raP.indexOf('[참고'), raP.indexOf('[참고') + 400));

  // 20. v21.2 AI 도우미 질문 · 답 관리 · MSDS 한글/영어/CAS 찾기
  const fq0 = await page.evaluate(a => BMAUTH.post('help.faqList', { t: a }), tka);
  const d3 = (fq0.rows || []).find(r => /신규입사자/.test(r.q)) || {};
  check('질문 · 답: 처음엔 기본값 12개 · 신규입사자 교육 기한 = 입사 후 3일 · 빠른 질문 표시', fq0.success && fq0.rows.length === 12 && /3일/.test(d3.a) && d3.quick && g.sheets['AI도우미_FAQ'].rows.length === 13, fq0.rows && fq0.rows.map(r => r.q));
  const fqDeny = await page.evaluate(a => BMAUTH.post('help.faqSave', { rows: [], t: a }), tkd);
  check('질문 · 답: 부서는 저장 불가', !fqDeny.success && fqDeny.status === 403, fqDeny);
  const newRows = fq0.rows.map(r => /신규입사자/.test(r.q) ? Object.assign({}, r, { a: '사내 규정: 입사 후 3일 안에 MSDS 교육 (부서장 작성)' }) : r)
    .concat([{ q: '야간 근무자 교육은?', a: '야간 근무자는 다음 주간 근무일에 교육합니다.', quick: true, memo: '2026-10 보건관리실' }, { q: '', a: '빈 질문은 버림' }]);
  const fqSave = await page.evaluate(a => BMAUTH.post('help.faqSave', { rows: a.rows, t: a.t }), { rows: newRows, t: tka });
  const meta2 = await page.evaluate(() => BMAUTH.post('help.meta', {}));
  check('질문 · 답 저장: 시트에 반영 · 빈 줄 버림 · 빠른 질문 목록 = 체크한 것', fqSave.success && fqSave.count === 13 && meta2.quick.indexOf('야간 근무자 교육은?') >= 0 && meta2.quick.length <= 8, [fqSave, meta2.quick]);
  gm.reply = { answer: '야간 근무자는 다음 주간 근무일에 교육합니다.', used: [1] };
  await page.evaluate(() => BMAUTH.post('help.ask', { q: '야간 근무자 교육은 언제 하나요?' }));
  const fp = gm.calls[gm.calls.length - 1].body.contents[0].parts[0].text;
  check('AI 도우미: 관리자가 고친 질문 · 답을 근거로 보냄 (메모는 안 보냄)', /다음 주간 근무일/.test(fp) && !/2026-10 보건관리실/.test(fp), fp.slice(fp.indexOf('[자료]'), fp.indexOf('[자료]') + 300));
  const msF = await page.evaluate(() => (CHEMS.some(p => p.name === 'STERI-GAS BRAND CARTRIDGES') || CHEMS.push({ name: 'STERI-GAS BRAND CARTRIDGES', manufacturer: '한국쓰리엠', departments: ['공급실'], icons: [],
      components: [{ compName: 'ETHYLENE OXIDE', casNumber: '75-21-8' }], emergency: '[피부,눈] 다량의 미온수로 15분 이상 세척하시오.' }), { eo: BMHELP_MSDS('산화에틸렌에 노출이 됐어'), cas: BMHELP_MSDS('75-21-8 응급조치'), none: BMHELP_MSDS('교육일지 작성법') }));
  check('MSDS 찾기: 한글 "산화에틸렌" → 영어 성분 ETHYLENE OXIDE(75-21-8) 제품 · CAS로도 찾음 · 물질 없는 질문은 안 보냄',
    msF.eo.length >= 1 && /ETHYLENE OXIDE/.test(msF.eo[0].text) && /응급조치: \[피부,눈\]/.test(msF.eo[0].text) && msF.cas.length >= 1 && msF.none.length === 0, msF);
  const faqUi = await page.evaluate(async () => {
    await BMAUTH.adminLogin('admin-pass-1');
    goTab('admin'); admSwitchPane('cfg'); BMDASH.settings(); await new Promise(z => setTimeout(z, 800));
    const v = { rows: document.querySelectorAll('#cfg-faq-list [data-faq]').length };
    document.querySelector('[data-c="faqadd"]').click();
    const first = document.querySelector('#cfg-faq-list [data-faq="0"]');
    first.querySelector('[data-f="q"]').value = '화면에서 추가한 질문'; first.querySelector('[data-f="a"]').value = '화면에서 추가한 답';
    document.querySelector('[data-c="faqsave"]').click(); await new Promise(z => setTimeout(z, 800));
    v.after = document.querySelectorAll('#cfg-faq-list [data-faq]').length;
    v.msg = document.getElementById('cfg-faq-msg').textContent;
    return v;
  });
  check('설정 화면: 질문 · 답 목록 · [+ 질문 추가] · 저장', faqUi.rows === 13 && faqUi.after === 14 && /14개를 저장/.test(faqUi.msg), faqUi);

  // 18-1. v21.5 (가져온 수정) 부서가 { records } 모양으로 올려도 서버의 [ … ] 교육 기록 파일에 들어감 · Gemini 503 → 다시 시도 → 예비 모델
  {
    const post = (b) => JSON.parse(g.api.doPost({ postData: { contents: JSON.stringify(b) } }).text);
    const cur = post({ op: 'fs.get', path: 'records/edu_records.json', t: tk.d });
    const rec = { recordId: 'shape61', name: '모양시험', dept: '61병동', type: 'MSDS교육', dateStart: '2026-10-27', submitted: '2026-10-27T01:00:00Z' };
    const curList = JSON.parse(Buffer.from(cur.content || '', 'base64').toString('utf8') || '[]');
    const mine = (Array.isArray(curList) ? curList : curList.records).concat([rec]);
    const put = post({ op: 'fs.put', path: 'records/edu_records.json', sha: cur.sha, content: Buffer.from(JSON.stringify({ records: mine })).toString('base64'), t: tk.d });
    const adm = post({ op: 'fs.get', path: 'records/edu_records.json', t: tk.a });
    const all = JSON.parse(Buffer.from(adm.content, 'base64').toString('utf8'));
    check('결재 관리 누락 수정: 부서가 { records } 로 올려도 서버 교육 기록([ … ])에 들어감', put.success && Array.isArray(all) && all.some(r => r.recordId === 'shape61'), [put, Array.isArray(all)]);
    const calls0 = gm.calls.length;
    gm.failModel = 'gemini-2.5-flash'; gm.reply = { scene: 'x', hazards: [{ category: '작업환경', hazard: '물기', situation: '넘어짐', freq: 2, sev: 2, controls: [] }] };
    const b1 = post({ op: 'ra.analyze', image: 'data:image/jpeg;base64,' + Buffer.from('x').toString('base64'), place: '복도', t: tk.a });
    gm.failModel = '';
    check('AI 503: 같은 모델 3번 시도 → 예비 모델(gemini-2.5-flash-lite)로 답', b1.success && b1.model === 'gemini-2.5-flash-lite' && gm.calls.length - calls0 === 4, [b1.success, b1.model, b1.error, gm.calls.length - calls0]);
    // v21.6 신규입사: 인사 사이트에 늦게 올라온 입사자(10일 전 입사)도 [누락 찾기]로 알림 · [직접 보내기] · 중복 막기
    const kday = n => new Date(Date.now() + 9 * 3600e3 - n * 86400e3).toISOString().slice(0, 10);
    const er = empRow('126999001', '늦은등록', '61병동', '재직', '간호사6급'); er[col('입사일')] = kday(10); empSh.appendRow(er);
    const er2 = empRow('126999002', '직접대상', '61병동', '재직', '간호사6급'); er2[col('입사일')] = kday(90); empSh.appendRow(er2);
    const m0 = mail.sent.length;
    const det = post({ op: 'nh.detect', t: tk.a });
    const nhOf = nm => g.api.nhRows_().find(o => o['이름'] === nm);
    check('신규입사 누락 찾기: 10일 전 입사(7일 기준 밖)도 교육 대상 · 미이수면 알림 + 부서장 메일', det.success && det.created.some(c => c.name === '늦은등록') && nhOf('늦은등록') && mail.sent.slice(m0).some(m => /늦은등록/.test(m.subject)) && !det.created.some(c => c.name === '직접대상'), det);
    const add1 = post({ op: 'nh.add', name: '직접대상', dept: '61병동', t: tk.a });
    const add2 = post({ op: 'nh.add', name: '직접대상', dept: '61병동', t: tk.a });
    const addD = post({ op: 'nh.add', name: '직접대상', dept: '61병동', t: tk.d });
    check('신규입사 직접 보내기: 고른 직원 알림 · 메일 · 같은 사람 두 번은 막음 · 부서 계정 불가', add1.success && add1.status === '회신대기' && nhOf('직접대상') && !add2.success && add2.exists && addD.status === 403, [add1, add2.error, addD.status]);
    const ui = await page.evaluate(() => typeof apvDeleteOne === 'function' && typeof _getStampedPdfBlob === 'function');
    check('결재 관리: 한 건 삭제 · 결재본 도장 미리보기 함수', ui);
  }

  // 18-2. v21.4 PDF 용량: 예전 방식(PNG → 압축 없는 RGB) PDF를 드라이브에 두고 → 큰 PDF 찾기 → 그림만 JPEG로 바꿔 다시 저장
  {
    const LIB = process.env.PDF_LIBS || '/tmp/claude-0/-home-user-health/380d0980-4304-5ba9-83eb-14958bc872ed/scratchpad/pdfb/node_modules';
    const fs = require('fs');
    if (fs.existsSync(LIB + '/jspdf/dist/jspdf.umd.min.js')) {
      await page.addScriptTag({ path: LIB + '/jspdf/dist/jspdf.umd.min.js' });
      await page.addScriptTag({ path: LIB + '/pdf-lib/dist/pdf-lib.min.js' });
      const pr = await page.evaluate(async () => {
        await BMAUTH.adminLogin('admin-pass-1');
        // 예전 방식 PDF 만들기: 900×1270 그림(글자 · 선 · 투명 영역)을 PNG로 넣음
        const cv = document.createElement('canvas'); cv.width = 900; cv.height = 1270;
        const x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, 900, 1200);
        x.fillStyle = '#123'; x.font = '28px sans-serif'; for (let i = 0; i < 30; i++) x.fillText('MSDS 교육일지 줄 ' + i + ' — 서명 확인', 40, 60 + i * 38);
        x.strokeStyle = '#d63838'; x.lineWidth = 3; x.beginPath(); x.arc(780, 1100, 60, 0, 7); x.stroke();
        const old = new jspdf.jsPDF('p', 'mm', 'a4');
        old.addImage(cv.toDataURL('image/png'), 'PNG', 0, 0, 210, 296);
        old.addPage(); old.addImage(cv.toDataURL('image/png'), 'PNG', 0, 0, 210, 296);
        const bytes = new Uint8Array(old.output('arraybuffer'));
        let bin = ''; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        const path = '교육일지/2026-05/예전큰일지-61병동.pdf';
        const api = 'https://api.github.com/repos/buminhealth/health/contents/' + path.split('/').map(encodeURIComponent).join('/');
        const put = await fetch(api, { method: 'PUT', body: JSON.stringify({ message: 't', content: btoa(bin) }) });
        const big = await BMAUTH.post('pdf.big', { min: 1 });
        const f = (big.data || []).find(z => z.path === path);
        const r = f ? await bmShrinkStoredPdf(f) : null;
        const after = new Uint8Array(await (await fetch('https://raw.githubusercontent.com/buminhealth/health/main/' + path.split('/').map(encodeURIComponent).join('/'))).arrayBuffer());
        const doc = await PDFLib.PDFDocument.load(after);
        const imgs = [];
        doc.getPages().forEach(pg => { const xo = pg.node.Resources().lookup(PDFLib.PDFName.of('XObject')); xo.keys().forEach(k => { const st = doc.context.lookup(xo.get(k)); imgs.push(String(st.dict.get(PDFLib.PDFName.of('Filter')))); }); });
        // 새로 만드는 PDF는 JPEG
        const nw = new jspdf.jsPDF('p', 'mm', 'a4'); nw.addImage(pdfPageImg(cv), 'JPEG', 0, 0, 210, 296, undefined, 'FAST');
        const again = { skipped: (await bmShrinkPdf(after)) === null, listed: (await BMAUTH.post('pdf.big', { min: 1 })).data.some(z => z.path === path) };
        return { putOk: put.ok, listed: !!f, before: bytes.length, after: after.length, r, pages: doc.getPageCount(), imgs, newKB: Math.round(nw.output('arraybuffer').byteLength / 1024), again };
      });
      check('PDF 용량: 예전 PDF 찾기 → 그림만 JPEG로 바꿔 같은 경로에 저장 (페이지 수 그대로 · 1/5 이하) · 다시 하면 건너뜀 · 새 PDF는 JPEG', pr.putOk && pr.listed && pr.pages === 2 && pr.imgs.every(f => f === '/DCTDecode') && pr.after < pr.before / 5 && pr.again.skipped && !pr.again.listed && pr.newKB < 400,
        { before: pr.before, after: pr.after, imgs: pr.imgs, newKB: pr.newKB, again: pr.again, r: pr.r });
    } else check('PDF 용량 시험: jspdf · pdf-lib 없음 (건너뜀)', true);
  }

  // 19. v21.3 저장 공간 가득 참 (같은 주소 다른 앱이 5MB를 거의 다 씀) → 멈추지 않고 메모리 보관 · 앱별 사용량 표시
  const st = await page.evaluate(() => {
    const keep = localStorage.getItem('bm_edu_records');
    localStorage.setItem('bm_chems_cache', 'x'.repeat(1000));
    let chunk = 'h'.repeat(512 * 1024), n = 0;
    try{ for(; n < 40; n++) localStorage.setItem('hk_v2_data_' + n, chunk); }catch(e){}
    try{ for(let m = 0; m < 200; m++) localStorage.setItem('hk_v2_fill_' + m, 'h'.repeat(4096)); }catch(e){}
    const big = JSON.stringify([{ recordId: 'big', name: '용량시험', note: 'z'.repeat(300000) }]);
    let threw = null;
    try{ localStorage.setItem('bm_edu_records', big); }catch(e){ threw = e.name; }
    const back = localStorage.getItem('bm_edu_records') === big;
    const u = BMSTORE.usage();
    const out = { threw, back, cacheGone: localStorage.getItem('bm_chems_cache') === null, mem: u.memoryOnly, groups: u.groups.map(g => g.name), hk: (u.groups.find(g => /건강검진/.test(g.name)) || {}).bytes };
    Object.keys(localStorage).filter(k => /^hk_v2_/.test(k)).forEach(k => localStorage.removeItem(k));
    localStorage.setItem('bm_edu_records', keep || '[]');
    out.restored = BMSTORE.memoryOnly().length === 0;
    return out;
  });
  check('저장 공간 가득 참: 오류 없이 메모리 보관 · 캐시 먼저 비움 · 앱별 사용량(건강검진) · 공간 생기면 다시 PC 저장', !st.threw && st.back && st.cacheGone && st.mem.indexOf('bm_edu_records') >= 0 && st.groups.some(g => /건강검진/.test(g)) && st.hk > 4e6 && st.restored, st);

  check('페이지 스크립트 오류 없음', errors.length === 0, errors.slice(0, 5));
  // 20. v21.3 전용 주소 이전: 예전 주소로 열면 이 PC 기록을 서버에 합치고(로그인 필요) 합친 부서 기록만 지운 뒤 새 주소로 (메일 링크 값 유지)
  {
    const NEW = 'https://new-msds.example/';
    const moved = require('fs').readFileSync(HTML, 'utf8').replace("var BM_HOME = '';", "var BM_HOME = '" + NEW + "';");
    const ctx2 = await browser.newContext();
    let landed = '';
    await ctx2.route('**/*', async route => {
      const u = route.request().url();
      if (u.startsWith('https://script.google.com/macros/')) {
        const q = u.indexOf('?') >= 0 ? u.slice(u.indexOf('?')) : '';
        try { const resp = await route.fetch({ url: MOCK + q }); return route.fulfill({ response: resp, headers: { ...resp.headers(), 'access-control-allow-origin': '*' } }); }
        catch (e) { return route.abort().catch(() => {}); }
      }
      if (u.startsWith(`http://localhost:${PORT}/index.html`)) return route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: moved });
      if (u.startsWith(NEW)) { landed = u; return route.fulfill({ status: 200, contentType: 'text/html', body: '<p>new</p>' }); }
      return route.abort();
    });
    await ctx2.addInitScript(() => {
      if (sessionStorage.getItem('seeded')) return; sessionStorage.setItem('seeded', '1');
      localStorage.setItem('bm_url', 'https://script.google.com/macros/s/x/exec');
      localStorage.setItem('bm_edu_records', JSON.stringify([
        { recordId: 'mv61', name: '이전시험', dept: '61병동', type: 'MSDS교육', dateStart: '2026-10-01' },
        { recordId: 'mv71', name: '다른부서', dept: '71병동', type: 'MSDS교육', dateStart: '2026-10-01' }]));
    });
    const p2 = await ctx2.newPage();
    await p2.goto(`http://localhost:${PORT}/index.html?nh=abc`);
    await p2.waitForTimeout(2500);
    const before = await p2.evaluate(() => ({ box: (document.getElementById('bm-move') || {}).textContent || '', url: location.href }));
    check('주소 이전: 로그인 전엔 이동하지 않고 로그인해서 옮기라고 안내 (PC 기록 보호)', /로그인하고 옮기기/.test(before.box) && /localhost/.test(before.url) && !landed, before);
    const tok = await p2.evaluate(async () => { const r = await BMAUTH.post('login', { kind: 'dept', dept: '61병동', pin: '4321' }); return r; });
    await p2.evaluate(r => localStorage.setItem('bm_session', JSON.stringify({ token: r.token, role: 'dept', dept: '61병동', head: r.head || '', exp: r.exp })), tok);
    await p2.reload();
    for (let i = 0; i < 40 && !landed; i++) await p2.waitForTimeout(250);
    const fg = JSON.parse(g.api.doPost({ postData: { contents: JSON.stringify({ op: 'fs.get', path: 'records/edu_records.json', t: tk.a }) } }).text);
    const srv = JSON.parse(Buffer.from(fg.content || fg.data && fg.data.content || '', 'base64').toString('utf8') || '[]');
    const srvList = Array.isArray(srv) ? srv : srv.records;
    check('주소 이전: 로그인하면 이 PC 기록을 서버에 합친 뒤 새 주소로 (?nh 유지)', srvList.some(r => r.recordId === 'mv61') && landed === NEW + '?nh=abc', [landed, srvList.length]);
    await ctx2.close();
  }

  await ctx.unrouteAll({ behavior: 'ignoreErrors' });
  await browser.close();
  server.close();
  console.log(failures ? `\n${failures} FAILED` : '\nALL PASSED');
  process.exit(failures ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });
