// QI 경진대회 발표 — 병원 그린 디자인(docs/디자인.md) · B 대본 10장 + 목차 · 팀 = 12장
const pptxgen = require('pptxgenjs');
const { applyTheme } = require('/root/.claude/skills/synced/1ea9381d-2a82-4a7a-9b65-5d5f63853442_6dadc0fa-bda1-4702-b74a-0df3a28208f5/pptx/scripts/apply_theme.js');
const OUT = process.argv[2];
const LOGO = __dirname + '/assets/bumin_logo.jpg';

const KO = 'Pretendard';
const THEME = { name: 'Bumin Green', headFontFace: KO, bodyFontFace: KO,
  colors: { dk1: '1E1E1C', lt1: 'FFFFFF', dk2: '45423D', lt2: 'F5F6F2', accent1: '67AF28', accent2: '3F7518', accent3: '345C12',
    accent4: 'A3302A', accent5: '4C8A1D', accent6: 'EEF6E6', hlink: '3F7518', folHlink: '345C12' } };
// 디자인.md 색
const K = { g5: '67AF28', g7: '4C8A1D', tg: '3F7518', g8: '345C12', g50: 'EEF6E6', g40: 'F3F9ED', gb: 'CFE5BA', gb2: 'DCEBCD', gb3: 'A9D67F',
  t1: '5A9E22', t2: '4C8A1D', t3: '3F7518', t4: '325F13', bg2: 'F5F6F2', ink: '1E1E1C', ink2: '45423D', sub: '5A5650',
  line: 'DDDBD5', card: 'E4E2DC', w: 'FFFFFF', red: 'A3302A', redBg: 'FBEDEA', before: 'C9C6BF' };

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: KO, bodyFontFace: KO };
pres.title = 'AI 통합 안전보건 플랫폼 구축을 통한 업무 효율성 증가';
pres.author = '총무-안전보건팀 (Safety First)';
pres.defineSlideMaster({ title: 'BUMIN', background: { color: K.w }, objects: [] });

const p = v => v / 96;
const B = (x, y, w, h) => ({ x: p(x), y: p(y), w: p(w), h: p(h) });
let n = 0; const nm = () => 'o' + (++n);
function T(s, t, x, y, w, h, o = {}) {
  s.addText(t, { ...B(x, y, w, h), isTextBox: true, margin: 0, fontFace: KO, fontSize: 14, color: K.ink, valign: 'top', objectName: nm(), ...o });
}
function card(s, x, y, w, h, fill = K.w, border = K.card) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, y, w, h), fill: { color: fill }, line: border ? { color: border, width: 1 } : { type: 'none' }, rectRadius: p(8), objectName: nm() });
}
function rect(s, x, y, w, h, fill) { s.addShape(pres.shapes.RECTANGLE, { ...B(x, y, w, h), fill: { color: fill }, line: { type: 'none' }, objectName: nm() }); }
function hl(s, x, y, w, c = K.line) { s.addShape(pres.shapes.LINE, { ...B(x, y, w, 0), line: { color: c, width: 0.75 }, objectName: nm() }); }
function vl(s, x, y, h, c = K.line) { s.addShape(pres.shapes.LINE, { ...B(x, y, 0, h), line: { color: c, width: 0.5 }, objectName: nm() }); }
function circ(s, x, y, d, fill, label, color = K.w, fs = 13) {
  s.addShape(pres.shapes.OVAL, { ...B(x, y, d, d), fill: { color: fill }, line: { type: 'none' }, objectName: nm() });
  if (label) T(s, label, x, y, d, d, { fontSize: fs, bold: true, color, align: 'center', valign: 'middle' });
}
function chip(s, x, y, w, text, fill = K.g50, color = K.tg) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, y, w, 22), fill: { color: fill }, line: { type: 'none' }, rectRadius: p(11), objectName: nm() });
  T(s, text, x, y, w, 22, { fontSize: 9, bold: true, color, align: 'center', valign: 'middle' });
}
const PH = ['Plan', 'Do', 'Check', 'Act'];
function band(s, page) {
  rect(s, 0, 680, 1280, 40, K.g8);
  T(s, '부산부민병원 안전보건 통합관리 시스템  ·  Safety First', 64, 680, 700, 40, { fontSize: 9, color: K.w, valign: 'middle' });
  if (page) T(s, String(page).padStart(2, '0'), 1116, 680, 100, 40, { fontSize: 10, bold: true, color: K.w, align: 'right', valign: 'middle' });
}
// 단계 표시(디자인.md 5장): 위 4px 막대 — 현재 · 지난 단계 그린, 남은 단계 회색
function chrome(s, label, phase, title, sub, page) {
  const sw = (1152 - 3 * 8) / 4;
  PH.forEach((ph, i) => {
    rect(s, 64 + i * (sw + 8), 24, sw, 4, i <= phase ? K.g5 : K.line);
    T(s, ph, 64 + i * (sw + 8), 32, sw, 14, { fontSize: 8, bold: i === phase, color: i === phase ? K.tg : K.sub });
  });
  T(s, label, 64, 60, 900, 18, { fontSize: 11, bold: true, color: K.tg, charSpacing: 1 });
  T(s, title, 64, 82, 1152, 44, { fontSize: 26, bold: true });
  if (sub) T(s, sub, 64, 128, 1152, 22, { fontSize: 13, color: K.sub });
  band(s, page);
}
const slide = () => pres.addSlide({ masterName: 'BUMIN' });
const fmt = v => Math.floor(v / 60) + '분 ' + String(v % 60).padStart(2, '0') + '초';

// ───── 01 표지 [0:00–0:20]
{
  const s = slide();
  s.addImage({ path: LOGO, ...B(64, 44, 156, 40), objectName: nm() });
  T(s, '2026 가을 학술대회  ·  QI 경진대회 구연', 64, 150, 640, 20, { fontSize: 12, bold: true, color: K.tg, charSpacing: 1 });
  T(s, 'AI 통합 안전보건\n플랫폼 구축을 통한\n업무 효율성 증가', 64, 184, 640, 200, { fontSize: 36, bold: true, lineSpacingMultiple: 1.12 });
  T(s, '종이 · 수기 · 대면 결재에서\n모바일 · 전자서명 · 비대면 승인으로', 64, 410, 640, 56, { fontSize: 15, color: K.ink2, lineSpacingMultiple: 1.35 });
  hl(s, 64, 500, 80, K.g5);
  T(s, '장은선  ·  김광현  ·  장종문', 64, 520, 640, 24, { fontSize: 16, bold: true });
  T(s, '인당의료재단 부민병원 총무-안전보건팀 (Safety First)', 64, 550, 640, 20, { fontSize: 12, color: K.sub });
  card(s, 760, 72, 456, 576, K.g40, K.gb);
  T(s, '특별안전보건교육 1건당 행정시간', 800, 108, 380, 20, { fontSize: 13, bold: true, color: K.tg });
  T(s, '−93%', 800, 138, 380, 100, { fontSize: 72, bold: true, color: K.tg });
  T(s, '33분 48초  →  2분 22초', 800, 248, 380, 30, { fontSize: 20, bold: true });
  T(s, '2026.05–07 세부 행동 단위 실측 · 목표 70% 초과 달성', 800, 284, 380, 18, { fontSize: 11, color: K.sub });
  [['법정교육 누락', '18건 → 0건'], ['종이 출력', '연 94건 → 0건']].forEach((r, i) => {
    const y = 360 + i * 120;
    hl(s, 800, y, 376, K.gb);
    T(s, r[0], 800, y + 20, 376, 18, { fontSize: 12, color: K.sub });
    T(s, r[1], 800, y + 44, 376, 44, { fontSize: 30, bold: true });
  });
  band(s, null);
  s.addNotes('[0:00–0:20] 표지\n\n안녕하십니까. 부민병원 총무-안전보건팀 보건관리자 장은선입니다.\n\n(도입) "지금 이 순간, 우리 병원 신규 입사자 중 법정 안전보건교육을 받지 않은 사람이 몇 명인지 바로 답하실 수 있습니까?"\n\n오늘은 이 질문에 바로 답할 수 있게 된 과정, AI 통합 안전보건 플랫폼 활동을 PDCA 순서로 말씀드리겠습니다.');
}

// ───── 02 목차 [0:20–0:25]
{
  const s = slide();
  chrome(s, 'CONTENTS', -1, '목차', 'PDCA 4단계 순서로 말씀드립니다', 2);
  const cw = (1152 - 3 * 20) / 4;
  [['P', 'Plan', '01  WHY · 활동 배경\n팀 구성\n02  PROBLEM · 현황\n03  원인 · 목표', 'Slide 03 – 06', K.t1],
   ['D', 'Do', '04  SOLUTION · 6대 모듈\n05  DEMO · 시연\nAI 적용', 'Slide 07 – 09', K.t2],
   ['C', 'Check', '06  RESULT · 전후 비교\n07  IMPACT · 성과 · ESG', 'Slide 10 – 11', K.t3],
   ['A', 'Act', '08  NEXT · 표준화', 'Slide 12', K.t4]].forEach((c, i) => {
    const x = 64 + i * (cw + 20);
    card(s, x, 168, cw, 488, K.bg2, null);
    rect(s, x, 168, cw, 6, c[4]);
    T(s, c[0], x + 28, 196, 120, 90, { fontSize: 64, bold: true, color: K.tg });
    T(s, c[1], x + 30, 288, cw - 60, 22, { fontSize: 15, bold: true, color: K.sub });
    hl(s, x + 28, 322, cw - 56, K.card);
    T(s, c[2], x + 28, 340, cw - 48, 220, { fontSize: 14, color: K.ink2, lineSpacingMultiple: 1.75 });
    T(s, c[3], x + 28, 612, cw - 56, 16, { fontSize: 10, color: K.sub });
  });
  s.addNotes('[0:20–0:25] 목차 — 말없이 넘겨도 됨\n\n"무엇이 달라졌는지 PDCA 순서로 말씀드리겠습니다."');
}

// ───── 03 01 WHY [0:25–1:05]
{
  const s = slide();
  chrome(s, '01   WHY  ·  활동 배경', 0, '법은 실시간 증빙을 요구하는데, 우리는 종이로 일했습니다', '주제 선정: 실현가능성 × 파급력 평가(22.5점) 동점 → 병원 중점과제 정합성 검토 후 최종 선정', 3);
  const cw = (1152 - 2 * 20) / 3;
  [['1', '법 · 제도', '법적 의무의 강화', '중대재해처벌법 · 산업안전보건법은 MSDS 관리와 신규 · 특별안전보건교육의 실시간 증빙을 요구합니다.', '인증기준 10.7 · 11.3'],
   ['2', '현장', '수기 행정의 한계', '출력 → 대면 서명 → 방문 제출 → 아날로그 결재. 교육 이수 누락을 즉시 확인할 수 없는 구조였습니다.', '종이 4단계 흐름'],
   ['3', '경영', '병원 전략과의 정렬', '병원 중점과제 「업무 프로세스 개선」, 그리고 ESG 환경(E) — 종이 행정의 데이터 자산화.', 'ESG · 프로세스 개선']].forEach((c, i) => {
    const x = 64 + i * (cw + 20);
    card(s, x, 168, cw, 352);
    circ(s, x + 28, 192, 36, K.g5, c[0], K.w, 15);
    T(s, c[1], x + 76, 192, cw - 100, 36, { fontSize: 14, bold: true, color: K.tg, valign: 'middle' });
    T(s, c[2], x + 28, 248, cw - 56, 32, { fontSize: 21, bold: true });
    T(s, c[3], x + 28, 296, cw - 56, 130, { fontSize: 14, color: K.ink2, lineSpacingMultiple: 1.5 });
    hl(s, x + 28, 452, cw - 56, K.card);
    chip(s, x + 28, 470, 170, c[4]);
  });
  card(s, 64, 540, 1152, 112, K.g8, null);
  T(s, '모바일 기반 통합 디지털 안전보건 체계가 필요했습니다', 104, 540, 1072, 112, { fontSize: 22, bold: true, color: K.w, valign: 'middle' });
  s.addNotes('[0:25–1:05] 활동 배경\n\n법은 "했다"가 아니라 "했다는 증거"를 요구합니다. MSDS와 신규·특별안전보건교육, 인증기준 10.7·11.3 모두 실시간 증빙이 필요합니다.\n\n그런데 저희는 출력하고, 서명받고, 들고 오고, 결재를 기다렸습니다.\n\n이 활동은 병원 중점과제 \'업무 프로세스 개선\', 그리고 ESG 환경 부문과 정확히 맞닿아 있습니다.\n\n전환: "이 활동은 저희 세 명이 맡았습니다."');
}

// ───── 04 팀 구성 [1:05–1:15]
{
  const s = slide();
  chrome(s, 'PLAN  ·  팀 구성', 0, '3명이 외부 개발 없이 기획부터 표준화까지 직접 수행했습니다', '팀 구성 · 2026년 1–10월 활동 일정', 4);
  const cw = (1152 - 2 * 20) / 3;
  [['김', '김광현', '팀장', '활동 총괄'], ['장', '장은선', '팀원 · 보건관리자', '플랫폼 기획 · 직접 개발'], ['장', '장종문', '팀원 · 안전관리자', 'MSDS · AI 위험성평가']].forEach((m, i) => {
    const x = 64 + i * (cw + 20);
    card(s, x, 168, cw, 104);
    circ(s, x + 24, 192, 56, K.g5, m[0], K.w, 18);
    T(s, m[1], x + 100, 190, cw - 120, 26, { fontSize: 18, bold: true });
    T(s, m[2], x + 100, 218, cw - 120, 18, { fontSize: 11, bold: true, color: K.tg });
    T(s, m[3], x + 100, 238, cw - 120, 18, { fontSize: 11, color: K.sub });
  });
  card(s, 64, 292, 1152, 364, K.bg2, null);
  const lx = 96, gx = 400, gw = 1184 - gx, mw = gw / 10;
  for (let m = 0; m < 10; m++) { T(s, (m + 1) + '월', gx + m * mw, 312, mw, 16, { fontSize: 10, bold: true, color: K.sub, align: 'center' }); vl(s, gx + m * mw, 336, 300); }
  vl(s, gx + 10 * mw, 336, 300);
  [['Plan', '현황 · 원인 분석 · 목표', 1, 3, K.gb3], ['Do', '개발 · 4월 프로토타입', 3, 4, K.t1], ['Do', '61병동 → 전 부서 적용', 5, 7, K.t2],
   ['Check', '결과 확인', 7, 7, K.t3], ['Act', '표준화 · AI 고도화', 8, 10, K.t4], ['', 'QI 경진대회', 9, 9, K.g8]].forEach((r, i) => {
    const y = 352 + i * 46;
    T(s, r[0], lx, y, 60, 26, { fontSize: 10, bold: true, color: K.tg, valign: 'middle' });
    T(s, r[1], lx + 60, y, 244, 26, { fontSize: 12, valign: 'middle' });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(gx + (r[2] - 1) * mw + 5, y + 4, (r[3] - r[2] + 1) * mw - 10, 18), fill: { color: r[4] }, line: { type: 'none' }, rectRadius: p(4), objectName: nm() });
  });
  s.addNotes('[1:05–1:15] 팀 구성 — 10초\n\n"팀장 김광현, 보건관리자인 제가 플랫폼을 직접 개발했고, 안전관리자 장종문이 MSDS와 AI 위험성평가를 맡았습니다."\n\n전환: "먼저 1년치 데이터를 전수조사했습니다."');
}

// ───── 05 02 PROBLEM [1:15–2:00]
{
  const s = slide();
  chrome(s, '02   PROBLEM  ·  현황 파악', 0, '교육 1건에 33분 48초, 그중 78%는 기다림과 이동이었습니다', '2025년 1–12월 신규 입사자 106명 교육일지 전수조사  ·  세부 행동 단위별 소요시간 실측', 5);
  const cw = (1152 - 2 * 20) / 3;
  [['법정교육 누락', '18건', '106명 중 17% · 12개월 중 8개월 발생', 'red'], ['특별안전보건교육 행정시간', '33분 48초', '교육 1건당 · 교육 진행시간 제외', 'hi'], ['종이 기반 처리', '100%', '교육일지 · MSDS 출력 후 바인더 보관 (연 94건)', '']].forEach((k, i) => {
    const x = 64 + i * (cw + 20), hi = k[3] === 'hi';
    card(s, x, 168, cw, 208, hi ? K.g5 : K.w, hi ? null : K.card);
    T(s, k[0], x + 28, 192, cw - 56, 20, { fontSize: 13, bold: hi, color: hi ? K.w : K.sub });
    T(s, k[1], x + 28, 226, cw - 40, 70, { fontSize: 40, bold: true, color: hi ? K.w : (k[3] === 'red' ? K.red : K.tg) });
    T(s, k[2], x + 28, 318, cw - 56, 36, { fontSize: 12, bold: hi, color: hi ? K.w : K.sub });
  });
  card(s, 64, 396, 1152, 260, K.bg2, null);
  T(s, '33분 48초는 어디에 쓰였나', 100, 420, 600, 24, { fontSize: 16, bold: true });
  T(s, '결재 대기 + 직접 제출 = 78%', 700, 420, 480, 24, { fontSize: 14, bold: true, color: K.tg, align: 'right' });
  let x = 100; const W = 1080;
  [['57%', '아날로그 내부 결재 대기', '19분 24초', 1164, K.g8, K.w], ['21%', '직접 제출 (이동)', '7분 00초', 420, K.g7, K.w], ['22%', '작성 · 출력 · 서명 · 취합 등', '7분 24초', 444, K.gb, K.tg]].forEach((g, i) => {
    const w = g[3] / 2028 * W - (i < 2 ? 6 : 0);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, 464, w, 72), fill: { color: g[4] }, line: { type: 'none' }, rectRadius: p(8), objectName: nm() });
    T(s, g[0], x + 20, 464, w - 30, 72, { fontSize: 24, bold: true, color: g[5], valign: 'middle' });
    T(s, g[1], x, 552, w, 20, { fontSize: 13, bold: true });
    T(s, g[2], x, 576, w, 18, { fontSize: 12, color: K.sub });
    x += w + 6;
  });
  s.addNotes('[1:15–2:00] 현황 파악\n\n2025년 신규 입사자 106명을 전수조사한 결과, 18건(17%)이 누락되어 있었습니다. 특정 시기가 아니라 12개월 중 8개월에서 반복됐습니다.\n\n특별안전보건교육 1건의 서류 처리에 33분 48초. 그중 결재 대기 57%, 걸어서 제출 21% — 78%가 아무 가치도 만들지 않는 시간이었습니다.\n\n전환: "문제는 사람이 아니라 방식이었습니다."');
}

// ───── 06 03 ROOT CAUSE & GOAL [2:00–2:40]
{
  const s = slide();
  chrome(s, '03   ROOT CAUSE & GOAL  ·  원인과 목표', 0, '원인은 사람이 아니라 프로세스에 있었습니다', '업무 흐름도 · Fishbone 분석으로 기여요인(CF) 도출 → 4개 핵심지표 목표 설정', 6);
  card(s, 64, 168, 664, 488);
  T(s, 'Fishbone 원인 분류 (4M)', 92, 188, 600, 22, { fontSize: 15, bold: true });
  T(s, '굵은 글씨 = 핵심 원인', 460, 190, 240, 20, { fontSize: 11, color: K.tg, bold: true, align: 'right' });
  [['방법', 'Method', [['종이 결재 대기 19.4분', 1], [' · 제출 채널 이원화 · 명단 수작업 대조', 0]]],
   ['환경', 'Environment', [['부서 ↔ 보건관리실 이동 평균 7분', 1]]],
   ['장비 · 시스템', 'Machine', [['실시간 이수 모니터링 불가', 1], [' · 인트라넷 검색 어려움 · 반복 수기 작성', 0]]],
   ['재료 · 매체', 'Material', [['종이 분실 위험 · 바인더 검색 한계 · MSDS 최신본 비치 확인 미흡', 0]]]].forEach((r, i) => {
    const y = 230 + i * 104;
    if (i) hl(s, 92, y - 12, 608, K.card);
    T(s, r[0], 92, y, 150, 24, { fontSize: 15, bold: true, color: K.tg });
    T(s, r[1], 92, y + 26, 150, 18, { fontSize: 10, color: K.sub });
    s.addText(r[2].map(([t, b]) => ({ text: t, options: { bold: !!b, color: b ? K.ink : K.ink2 } })), { ...B(252, y, 448, 76), isTextBox: true, margin: 0, fontFace: KO, fontSize: 13, valign: 'top', lineSpacingMultiple: 1.4, objectName: nm() });
  });
  T(s, '기준값 → 목표', 748, 172, 468, 18, { fontSize: 12, bold: true, color: K.sub });
  const gw = (468 - 16) / 2;
  [['GOAL 01', '법정교육 누락률', '17% → 0%'], ['GOAL 02', '행정 리드타임', '70%↓', '33분 48초 → 10분 이내'], ['GOAL 03', '종이 교육일지', '0건', '연 94건 → 0건'], ['GOAL 04', '6대 모듈 구축', '100%', '0 → 6 / 6']].forEach((g, i) => {
    const x = 748 + (i % 2) * (gw + 16), y = 196 + Math.floor(i / 2) * 238;
    card(s, x, y, gw, 222, K.g40, K.gb);
    T(s, g[0], x + 20, y + 18, gw - 40, 16, { fontSize: 10, bold: true, color: K.tg, charSpacing: 1 });
    T(s, g[1], x + 20, y + 40, gw - 40, 22, { fontSize: 14, bold: true });
    T(s, g[2], x + 20, y + 80, gw - 30, 60, { fontSize: g[2].length > 6 ? 26 : 38, bold: true, color: K.tg });
    if (g[3]) T(s, g[3], x + 20, y + 160, gw - 40, 40, { fontSize: 11, color: K.sub, lineSpacingMultiple: 1.3 });
  });
  s.addNotes('[2:00–2:40] 원인과 목표\n\nFishbone으로 4M 원인을 분류했고, 핵심은 결재 대기 · 방문 제출 · 실시간 모니터링 부재 세 가지였습니다.\n\n현황을 기준값으로 목표를 숫자로 정했습니다: 누락률 0%, 리드타임 70% 이상 단축, 종이 일지 0건, 6대 모듈 100%.\n\n전환: "원인마다 대응 기능을 하나씩 설계했습니다."');
}

// ───── 07 04 SOLUTION [2:40–3:20]
{
  const s = slide();
  chrome(s, '04   SOLUTION  ·  플랫폼 구성', 1, '원인마다 하나씩, 6개 모듈로 설계했습니다', null, 7);
  const fw = (1152 - 3 * 24) / 4;
  ['모바일 작성', '전자서명', '비대면 승인', '자동 보관 · 대시보드'].forEach((t, i) => {
    const x = 64 + i * (fw + 24);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, 140, fw, 40), fill: { color: K.g50 }, line: { color: K.gb, width: 1 }, rectRadius: p(20), objectName: nm() });
    T(s, t, x, 140, fw, 40, { fontSize: 14, bold: true, color: K.tg, align: 'center', valign: 'middle' });
    if (i < 3) T(s, '›', x + fw, 140, 24, 40, { fontSize: 20, bold: true, color: K.g5, align: 'center', valign: 'middle' });
  });
  const cw = (1152 - 2 * 20) / 3;
  [['01', 'CF1', 'MSDS 디지털 라이브러리', '부서별 유해화학물질 조회\n경고표지 · 응급조치 · 원본 PDF 원클릭'],
   ['02', 'CF2 · 3', '모바일 교육일지 · 전자서명', '인사 DB 연동, 이름 검색만으로\n일지 자동 작성 · 서명 즉시 DB 저장'],
   ['03', 'CF4 · 5', '비대면 전자결재', '서명 완료 즉시 관리자 알림\n장소 제약 없는 원격 승인'],
   ['04', 'CF7', '보호구 · 법정서류 전산화', '보호구 수령증, 유해화학물질\n신규 · 중단 보고서 단일 창구'],
   ['05', 'CF6', '관리자 모니터링 대시보드', '부서 · 기간 · 근로자별\n교육 이수 현황 실시간 조회'],
   ['06', 'AI', 'AI 수시 위험성평가', '현장 사진 · 작업 입력 → 위험요인 ·\n위험도 산출 → 평가서 자동 작성']].forEach((m, i) => {
    const x = 64 + (i % 3) * (cw + 20), y = 192 + Math.floor(i / 3) * 216, ai = i === 5;
    card(s, x, y, cw, 200, ai ? K.g5 : K.w, ai ? null : K.card);
    T(s, m[0], x + 24, y + 20, 60, 24, { fontSize: 18, bold: true, color: ai ? K.w : K.tg });
    chip(s, x + cw - 104, y + 20, 80, m[1], ai ? K.w : K.g50, K.tg);
    T(s, m[2], x + 24, y + 60, cw - 48, 28, { fontSize: 17, bold: true, color: ai ? K.w : K.ink });
    T(s, m[3], x + 24, y + 100, cw - 48, 80, { fontSize: 13, color: ai ? K.w : K.ink2, lineSpacingMultiple: 1.45 });
  });
  T(s, 'AppSheet · Google Sheets · Apps Script · EmailJS · GitHub 기반 자체 개발   |   4월 프로토타입 → 61병동 시범 → 전 부서 확대', 64, 630, 1152, 18, { fontSize: 11, color: K.sub });
  s.addNotes('[2:40–3:20] 플랫폼 구성\n\n흐름은 단순합니다: 모바일 작성 → 전자서명 → 비대면 승인 → 자동 보관.\n\n각 카드의 태그는 앞에서 찾은 원인 코드입니다. 원인 하나에 기능 하나.\n\n초록 카드, AI 수시 위험성평가는 사진과 작업 내용만으로 위험도를 산출하고 평가서 초안을 만듭니다.\n\n담당자가 직접 개발했고, 61병동 시범 운영 피드백을 반영해 전 부서로 확대했습니다.\n\n전환: "실제 화면으로 보여드리겠습니다."');
}

// ───── 08 05 DEMO [3:20–4:50 · 영상 1:30]
{
  const s = slide();
  chrome(s, '05   DEMO  ·  시연', 1, '신규 입사자 1명의 교육, 작성부터 승인까지 한 화면에서', '프로그램 시연 영상 1분 30초', 8);
  card(s, 64, 168, 784, 441, K.bg2, K.card);
  T(s, '시연 영상 (16:9) — 삽입 → 비디오', 64, 168, 784, 441, { fontSize: 16, bold: true, color: K.sub, align: 'center', valign: 'middle' });
  T(s, '시연 순서', 880, 172, 336, 20, { fontSize: 13, bold: true, color: K.sub });
  vl(s, 896, 210, 380, K.gb);
  [['0:00', 'MSDS 라이브러리 조회'], ['0:20', '교육일지 작성 · 전자서명'], ['0:50', '관리자 알림 · 비대면 승인'], ['1:10', '대시보드 · AI 위험성평가']].forEach((t, i) => {
    const y = 204 + i * 100;
    circ(s, 888, y, 16, i === 2 ? K.g5 : K.gb3);
    T(s, t[0], 920, y - 2, 120, 20, { fontSize: 12, bold: true, color: K.tg });
    T(s, t[1], 920, y + 22, 296, 26, { fontSize: 16, bold: true });
  });
  T(s, '"이 1분이 예전의 19분 24초입니다"', 880, 608, 336, 20, { fontSize: 12, bold: true, color: K.tg });
  s.addNotes('[3:20–4:50] 시연 영상 · 1분 30초\n\n시작: "신규 입사자 한 명이 들어왔을 때를 가정해 보겠습니다."\n\n- MSDS — 부서 선택 → 경고표지·응급조치·원본 PDF\n- 교육일지 — 이름 검색 → 자동 입력 → 현장 서명\n- 승인 — "이 1분이 예전의 19분 24초입니다."\n- 대시보드·AI — 미이수자가 한 화면에\n\n종료 후: "AI는 어디에 쓰였는지 짚고 가겠습니다."');
}

// ───── 09 05+ AI 적용 [4:50–5:15]
{
  const s = slide();
  chrome(s, '05+   SOLUTION  ·  AI 적용', 1, 'AI는 초안을 만들고, 최종 판단은 사람이 합니다', 'AI 수시 위험성평가 — 근로자 참여형 · 2026.05–07 시범 운영', 9);
  const fw = (1152 - 3 * 24) / 4;
  ['현장 사진 · 작업 입력', '위험요인 자동 분석', '위험도 산출 (5×4)', '개선대책 · 평가서 작성'].forEach((t, i) => {
    const x = 64 + i * (fw + 24);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, 168, fw, 44), fill: { color: i === 3 ? K.g5 : K.g50 }, line: { type: 'none' }, rectRadius: p(22), objectName: nm() });
    T(s, t, x, 168, fw, 44, { fontSize: 14, bold: true, color: i === 3 ? K.w : K.tg, align: 'center', valign: 'middle' });
    if (i < 3) T(s, '›', x + fw, 168, 24, 44, { fontSize: 20, bold: true, color: K.g5, align: 'center', valign: 'middle' });
  });
  card(s, 64, 232, 664, 376, K.bg2, K.card);
  T(s, 'AI 위험성평가 결과 화면 캡처\n(현장 사진 → 위험요인 → 위험도 → 개선대책 실제 사례 1건)', 64, 232, 664, 376, { fontSize: 14, bold: true, color: K.sub, align: 'center', valign: 'middle', lineSpacingMultiple: 1.4 });
  card(s, 748, 232, 468, 150, K.w, K.card);
  T(s, '운영 실적  ·  2026.05–07', 776, 252, 400, 18, { fontSize: 12, color: K.sub });
  T(s, '[실적]건', 776, 278, 400, 56, { fontSize: 38, bold: true, color: K.tg });
  T(s, '수시 위험성평가 시행 건수', 776, 344, 400, 18, { fontSize: 11, color: K.sub });
  card(s, 748, 398, 468, 210, K.g8, null);
  T(s, '검토 원칙', 776, 418, 400, 18, { fontSize: 12, bold: true, color: K.gb3 });
  T(s, 'AI는 초안, 확정은 안전관리자', 776, 444, 420, 28, { fontSize: 19, bold: true, color: K.w });
  T(s, '점수 · 등급은 서버가 다시 계산하고, 법령 조문 번호를 대조합니다. 현재 시범 단계이며 정확도 고도화가 다음 과제입니다.', 776, 484, 412, 110, { fontSize: 13, color: K.w, lineSpacingMultiple: 1.45 });
  T(s, '그 밖의 AI 기능   AI MSDS 검토 (규제 항목 제안 → 관리자 승인)  ·  AI 도우미 (출처와 함께 답변)', 64, 626, 1152, 18, { fontSize: 11, color: K.sub });
  s.addNotes('[4:50–5:15] AI 적용 — 25초\n\n- "AI는 어디에?" 질문에 대한 답 — 6번 모듈의 실제 사례\n- 사진 한 장으로 위험요인 · 위험도 · 개선대책 초안이 나오는 흐름만 짚기\n- AI는 초안, 확정은 사람 — 이 문장은 꼭 말하기\n- 시범 단계라는 한계를 먼저 밝히기 (AI MSDS 검토 · AI 도우미는 질문이 나오면 답변)\n\n전환: "이 과정이 만든 차이를 숫자로 보여드리겠습니다."');
}

// ───── 10 06 RESULT [5:15–6:05]
{
  const s = slide();
  chrome(s, '06   RESULT  ·  개선 전후 비교', 2, '같은 교육, 행정 시간은 93% 줄었습니다', '특별안전보건교육 1건당 행정 소요시간 (교육 진행시간 제외) · 세부 행동 단위 실측', 10);
  card(s, 64, 168, 1152, 488);
  T(s, 'BEFORE · 종이 수기', 96, 184, 300, 18, { fontSize: 11, bold: true, color: K.sub });
  T(s, '33분 48초', 96, 204, 300, 44, { fontSize: 30, bold: true, color: K.sub });
  T(s, 'AFTER · AI 통합 플랫폼', 420, 184, 300, 18, { fontSize: 11, bold: true, color: K.tg });
  T(s, '2분 22초', 420, 204, 300, 44, { fontSize: 30, bold: true, color: K.tg });
  card(s, 980, 184, 204, 64, K.g5, null);
  T(s, '−93%', 980, 184, 204, 64, { fontSize: 30, bold: true, color: K.w, align: 'center', valign: 'middle' });
  const steps = [['1. 대상자 확인 · 일지 작성', 138, 18], ['2. 부서장에게 일지 전송', 60, 50], ['3. 출력 · 수기 서명', 90, 9], ['4. 보건관리실 직접 제출', 420, 0], ['5. 사진 촬영 · 전송', 30, 6], ['6. 취합 · 인쇄 · 파일링', 126, 0], ['7. 내부 결재', 1164, 60]];
  const bx = 340, bw = 720;
  steps.forEach((st, i) => {
    const y = 272 + i * 52;
    hl(s, 96, y - 6, 1088, K.card);
    T(s, st[0], 96, y, 240, 40, { fontSize: 13, bold: i === 3 || i === 6, valign: 'middle' });
    const wb = Math.max(4, st[1] / 1164 * bw), wa = Math.max(4, st[2] / 1164 * bw);
    rect(s, bx, y + 4, wb, 14, K.before);
    T(s, fmt(st[1]).replace('0분 ', ''), bx + wb + 8, y, 120, 20, { fontSize: 11, color: K.sub, valign: 'middle' });
    if (st[2]) { rect(s, bx, y + 22, wa, 14, K.g5); T(s, fmt(st[2]).replace('0분 ', ''), bx + wa + 8, y + 18, 120, 20, { fontSize: 11, bold: true, color: K.tg, valign: 'middle' }); }
    else T(s, '단계 삭제', bx, y + 18, 160, 20, { fontSize: 11, bold: true, color: K.tg, valign: 'middle' });
  });
  T(s, '회색 = 개선 전 · 초록 = 개선 후   |   MSDS 교육 4분 18초 → 49초 (−81%)', 96, 632, 1088, 18, { fontSize: 11, color: K.sub });
  s.addNotes('[5:15–6:05] 개선 전후 비교 · 핵심 장면\n\n왼쪽이 예전, 오른쪽이 지금 — 같은 특별안전보건교육 한 건입니다.\n\n가장 긴 막대, 내부 결재 19분 24초가 1분이 되었습니다.\n\n직접 제출 7분과 취합·파일링은 단계 자체가 사라졌습니다.\n\n합계 33분 48초 → 2분 22초, 93% 단축. 목표 70% 초과, MSDS 교육도 81% 단축.\n\n전환: "시간뿐만 아니라, 네 가지 지표 모두 목표를 넘었습니다."');
}

// ───── 11 07 IMPACT [6:05–6:45]
{
  const s = slide();
  chrome(s, '07   IMPACT  ·  성과와 ESG', 2, '4개 지표 모두 목표를 달성했습니다', '개선 전 2025 → 개선 후 2026.05–07', 11);
  const cw = (1152 - 3 * 20) / 4;
  [['0건', '법정교육 누락', '18건 → 0건 · 준수율 100%', '대상 33건 전원 기한 내 이수'], ['93%', '행정시간 단축', '33분 48초 → 2분 22초', '목표 70% 초과 · MSDS −81%'],
   ['100%', 'Paperless 전환', '월 17.3건 전량 디지털', '5–7월 52건 · 종이 출력 0건'], ['6/6', '플랫폼 구축', '핵심 모듈 100% 가동', '현장 피드백 반영 완료']].forEach((k, i) => {
    const x = 64 + i * (cw + 20), hi = i === 1;
    card(s, x, 168, cw, 232, hi ? K.g5 : K.w, hi ? null : K.card);
    T(s, k[0], x + 24, 188, cw - 48, 70, { fontSize: 46, bold: true, color: hi ? K.w : K.tg });
    T(s, k[1], x + 24, 268, cw - 48, 22, { fontSize: 15, bold: true, color: hi ? K.w : K.ink });
    T(s, k[2], x + 24, 300, cw - 48, 20, { fontSize: 12, bold: true, color: hi ? K.w : K.ink2 });
    T(s, k[3], x + 24, 344, cw - 48, 36, { fontSize: 11, color: hi ? K.w : K.sub });
  });
  [['E', 'ESG · 환경', '종이 교육일지 · MSDS 출력 · 바인더 보관 0건 — 연 94건의 종이 행정을 데이터 자산으로'],
   ['S', 'ESG · 사회', '직원 안전보건교육 누락 0건 · 근로자 참여형 AI 수시 위험성평가'],
   ['G', 'ESG · 지배구조', '중대재해처벌법 · 인증 10.7 · 11.3 증빙 무결성 · 표준 업무지침 제정']].forEach((r, i) => {
    const y = 420 + i * 80;
    card(s, 64, y, 1152, 68, K.g40, K.gb);
    circ(s, 88, y + 14, 40, K.g8, r[0], K.w, 16);
    T(s, r[1], 144, y, 170, 68, { fontSize: 14, bold: true, color: K.tg, valign: 'middle' });
    T(s, r[2], 320, y, 880, 68, { fontSize: 14, valign: 'middle' });
  });
  s.addNotes('[6:05–6:45] 성과와 ESG\n\n네 개 지표 모두 목표를 달성했습니다.\n\n누락 18건 → 0건 — 효율이 아니라 법적 리스크의 문제입니다.\n\n행정시간 93% 단축, Paperless 100%, 6개 모듈 모두 정상 가동.\n\nESG로도 이어집니다: E 종이 행정의 데이터화, S 누락 없는 직원 보호, G 언제든 준비된 법적 증빙.\n\n전환: "이 성과를 어떻게 지속하고 넓힐지 말씀드리겠습니다."');
}

// ───── 12 08 NEXT + 감사합니다 [6:45–7:30]
{
  const s = slide();
  chrome(s, '08   NEXT  ·  표준화 · 고도화', 3, '검증된 프로세스를 병원의 표준으로', null, 12);
  [['표준화', '완료', '「법정 안전보건 관리 표준 업무지침」 제정 · 정식 업무 전환'], ['지속 관리', '운영', '관리자 대시보드로 신규 입사자 이수 현황 실시간 점검 (CQI 연계)'],
   ['고도화', '완료', '미이수자 부서장 자동 알림 · AI MSDS 검토 · AI 도우미 (2026.08–10)'], ['확장', '제언', '건강상담(연 819건) · 기타 서류 확대 · AI 위험성평가 참여 확대']].forEach((r, i) => {
    const y = 140 + i * 128;
    card(s, 64, y, 704, 112, i === 3 ? K.g40 : K.w, i === 3 ? K.gb : K.card);
    T(s, r[0], 92, y + 20, 200, 26, { fontSize: 18, bold: true, color: K.tg });
    chip(s, 640, y + 22, 104, r[1], r[1] === '제언' ? K.g5 : K.g50, r[1] === '제언' ? K.w : K.tg);
    T(s, r[2], 92, y + 58, 640, 40, { fontSize: 14, color: K.ink2, lineSpacingMultiple: 1.35 });
  });
  card(s, 788, 140, 428, 496, K.g8, null);
  T(s, 'OUR PROMISE', 820, 168, 360, 18, { fontSize: 11, bold: true, color: K.gb3, charSpacing: 2 });
  T(s, '기다리지 않는 행정,\n누락 없는 안전.', 820, 200, 380, 110, { fontSize: 30, bold: true, color: K.w, lineSpacingMultiple: 1.2 });
  T(s, '안전보건 데이터가 실시간으로 흐를 때, 관리자는 서류가 아닌 현장에 집중합니다.', 820, 330, 364, 80, { fontSize: 14, color: K.gb, lineSpacingMultiple: 1.45 });
  hl(s, 820, 520, 364, K.g7);
  T(s, '감사합니다', 820, 540, 364, 44, { fontSize: 28, bold: true, color: K.w });
  T(s, '문의 · 보건관리실 내선 3293 · 3034', 820, 590, 364, 20, { fontSize: 11, color: K.gb });
  s.addNotes('[6:45–7:30] 결론과 제언\n\n시범사업에 그치지 않고 표준 업무지침으로 정식 전환했습니다.\n\n대시보드와 현장 피드백 점검을 통해 CQI로 이어 가고 있습니다. 미이수자 부서장 자동 알림과 AI MSDS 검토 · 도우미까지 고도화를 마쳤고, 다음은 연 819건의 건강상담과 다른 안전보건 서류로의 확대입니다.\n\n마무리: "기다리지 않는 행정, 누락 없는 안전. 감사합니다."');
}

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log('ok'); })();
