// QI 경진대회 발표 — 애플 스타일(감색) · 바탕 흰색 · 이모지/아이콘 없음
// node build_apple.js OUT [variants]  — variants 를 주면 5장 시안 B·C 만 따로 만든다
const pptxgen = require('pptxgenjs');
const { applyTheme } = require('/root/.claude/skills/synced/1ea9381d-2a82-4a7a-9b65-5d5f63853442_6dadc0fa-bda1-4702-b74a-0df3a28208f5/pptx/scripts/apply_theme.js');
const OUT = process.argv[2];
const MODE = process.argv[3] || 'deck';

const KO = 'Noto Sans KR', NUM = 'Inter';
const THEME = { name: 'Bumin Navy', headFontFace: KO, bodyFontFace: KO,
  colors: { dk1: '1D1D1F', lt1: 'FFFFFF', dk2: '636366', lt2: 'F5F5F7', accent1: '0A1F44', accent2: '0066CC', accent3: 'C7C7CC',
    accent4: 'F56300', accent5: '4F6A96', accent6: '8FA6C9', hlink: '0066CC', folHlink: '0A1F44' } };
const K = { navy: '0A1F44', ink: '1D1D1F', sub: '636366', surf: 'F5F5F7', blue: '0066CC', grey: 'C7C7CC', line: 'E5E5EA',
  orange: 'F56300', orangeT: 'B33F00', w: 'FFFFFF', onNavy: 'A9B8D0', n1: '8FA6C9', n2: '4F6A96', n3: '2A4472' };

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: KO, bodyFontFace: KO };
pres.title = 'AI 통합 안전보건 플랫폼 구축을 통한 업무 효율성 증가';
pres.author = '총무-안전보건팀 (Safety First)';
pres.defineSlideMaster({ title: 'APPLE', background: { color: K.w }, objects: [] });

const p = v => v / 96;
const B = (x, y, w, h) => ({ x: p(x), y: p(y), w: p(w), h: p(h) });
let n = 0; const nm = () => 'o' + (++n);
function T(s, t, x, y, w, h, o = {}) {
  s.addText(t, { ...B(x, y, w, h), isTextBox: true, margin: 0, fontFace: KO, fontSize: 14, color: K.ink, valign: 'top', objectName: nm(), ...o });
}
function card(s, x, y, w, h, fill = K.surf, r = 20) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, y, w, h), fill: { color: fill }, line: { type: 'none' }, rectRadius: p(r), objectName: nm() });
}
function rect(s, x, y, w, h, fill) {
  s.addShape(pres.shapes.RECTANGLE, { ...B(x, y, w, h), fill: { color: fill }, line: { type: 'none' }, objectName: nm() });
}
function bar(s, x, y, w, h, fill) { // 둥근 막대
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...B(x, y, w, h), fill: { color: fill }, line: { type: 'none' }, rectRadius: p(Math.min(h / 2, 6)), objectName: nm() });
}
function hl(s, x, y, w, c = K.line) { s.addShape(pres.shapes.LINE, { ...B(x, y, w, 0), line: { color: c, width: 0.75 }, objectName: nm() }); }
function vl(s, x, y, h, c = K.line) { s.addShape(pres.shapes.LINE, { ...B(x, y, 0, h), line: { color: c, width: 0.5 }, objectName: nm() }); }
function dot(s, x, y, d, fill, label, fs = 14, color = K.w) {
  s.addShape(pres.shapes.OVAL, { ...B(x, y, d, d), fill: { color: fill }, line: { type: 'none' }, objectName: nm() });
  if (label) T(s, label, x, y, d, d, { fontSize: fs, bold: true, color, align: 'center', valign: 'middle' });
}
const PH = ['Plan', 'Do', 'Check', 'Act'];
function chrome(s, eyebrow, phase, title, sub, page) {
  T(s, eyebrow, 64, 44, 700, 18, { fontSize: 10, bold: true, color: K.navy, charSpacing: 1 });
  if (phase >= 0) PH.forEach((ph, i) => {
    const x = 940 + i * 70;
    T(s, ph, x, 44, 64, 18, { fontFace: NUM, fontSize: 9, bold: i === phase, color: i === phase ? K.navy : K.sub, align: 'center' });
    if (i === phase) rect(s, x + 14, 64, 36, 2, K.navy);
  });
  T(s, title, 64, 70, 1152, 48, { fontSize: 28, bold: true, color: K.ink });
  if (sub) T(s, sub, 64, 122, 1152, 26, { fontSize: 15, color: K.sub });
  T(s, '부산부민병원 안전보건 통합관리 시스템', 64, 680, 600, 16, { fontSize: 9, color: K.sub });
  T(s, String(page).padStart(2, '0'), 1116, 680, 100, 16, { fontFace: NUM, fontSize: 9, color: K.sub, align: 'right' });
}
const fmt = v => Math.floor(v / 60) + ':' + String(v % 60).padStart(2, '0');
const slide = () => pres.addSlide({ masterName: 'APPLE' });

function deck14() {
  // ───── 01 표지 (확정 구성: 왼쪽 제목 + 오른쪽 결과 3개)
  {
    const s = slide();
    T(s, '2026 QI 경진대회 · 구연 발표', 64, 64, 600, 20, { fontSize: 12, bold: true, color: K.navy, charSpacing: 1 });
    T(s, 'AI 통합 안전보건 플랫폼\n구축을 통한 업무 효율성 증가', 64, 226, 680, 140, { fontSize: 34, bold: true, lineSpacingMultiple: 1.12 });
    T(s, 'Development of an AI-integrated occupational safety and health platform for improved work efficiency', 64, 392, 600, 46, { fontFace: NUM, fontSize: 13, color: K.sub });
    T(s, '총무-안전보건팀 (Safety First)', 64, 560, 640, 26, { fontSize: 16, bold: true });
    T(s, '발표 장은선 · 팀장 김광현 · 팀원 장종문 · 인당의료재단 부민병원', 64, 592, 640, 22, { fontSize: 13, color: K.sub });
    card(s, 760, 64, 456, 592, K.surf, 28);
    T(s, '개선 전(2025) → 개선 후(2026.05–07)', 800, 100, 380, 20, { fontSize: 12, bold: true, color: K.sub });
    const R = [['특별교육 행정시간', '−93%', '33:48 → 2:22 (분:초)', true], ['법정교육 누락', '18 → 0건', '준수율 83% → 100%', false], ['종이 출력', '94 → 0건', '페이퍼리스 0% → 100%', false]];
    R.forEach((r, i) => {
      const y = 150 + i * 166;
      if (i) hl(s, 800, y - 14, 376);
      T(s, r[0], 800, y, 376, 20, { fontSize: 13, color: K.sub });
      T(s, r[1], 800, y + 26, 376, 72, { fontFace: NUM, fontSize: 50, bold: true, color: r[3] ? K.navy : K.ink });
      T(s, r[2], 800, y + 104, 376, 20, { fontSize: 12, color: K.sub });
    });
    s.addNotes('[도입부 30초]\n- "작년 11월, 신규 입사자가 한꺼번에 17명 들어왔습니다. 그달 교육일지 7건이 누락됐습니다."\n- "종이로는 누가 빠졌는지 볼 방법이 없었기 때문입니다."\n- "올해 6월, 그보다 많은 24건이 몰렸지만 누락은 0건이었습니다. 무엇이 달라졌는지 말씀드리겠습니다."');
  }
  // ───── 02 목차 (확정 구성: P·D·C·A 카드 4개)
  {
    const s = slide();
    chrome(s, 'CONTENTS', -1, '목차', 'PDCA 4단계 순서로 말씀드립니다', 2);
    const cw = (1152 - 3 * 20) / 4;
    [['P', 'Plan', '주제 선정 배경\n현황 분석\n지표 정의 · 목표\n원인 분석 · 개선 전략', 'Slide 04 – 07'], ['D', 'Do', '6대 모듈 구축\n플랫폼 시연', 'Slide 08'],
      ['C', 'Check', '법정교육 준수율\n페이퍼리스 전환율\n행정 소요시간', 'Slide 09 – 10'], ['A', 'Act', '표준화\nAI 통합 플랫폼 고도화\n결론 및 제언', 'Slide 11 – 12']].forEach((c, i) => {
      const x = 64 + i * (cw + 20);
      card(s, x, 176, cw, 470);
      T(s, c[0], x + 32, 206, 120, 90, { fontFace: NUM, fontSize: 64, bold: true, color: K.navy });
      T(s, c[1], x + 34, 302, cw - 64, 22, { fontFace: NUM, fontSize: 14, bold: true, color: K.sub });
      T(s, c[2], x + 30, 352, cw - 48, 200, { fontSize: 15, lineSpacingMultiple: 1.7 });
      T(s, c[3], x + 34, 606, cw - 64, 18, { fontFace: NUM, fontSize: 10, color: K.sub });
    });
    s.addNotes('- "PDCA 순서로 말씀드리겠습니다." (5초, 바로 넘김)');
  }
  // ───── 03 팀 · 일정 (확정 구성: 팀원 카드 + 간트)
  {
    const s = slide();
    chrome(s, 'TEAM', -1, '3명의 팀이 외부 개발 없이 기획부터 표준화까지 직접 수행했다', '팀 구성 · 1–10월 활동 일정', 3);
    const cw = (1152 - 2 * 20) / 3;
    [['김', '김광현', '팀장', '활동 총괄'], ['장', '장은선', '팀원 · 보건관리자', '플랫폼 기획 · 직접 개발'], ['장', '장종문', '팀원 · 안전관리자', 'MSDS · AI 위험성평가']].forEach((m, i) => {
      const x = 64 + i * (cw + 20);
      card(s, x, 176, cw, 112);
      dot(s, x + 24, 204, 56, K.navy, m[0], 18);
      T(s, m[1], x + 100, 202, cw - 120, 26, { fontSize: 18, bold: true });
      T(s, m[2], x + 100, 230, cw - 120, 18, { fontSize: 11, bold: true, color: K.navy });
      T(s, m[3], x + 100, 250, cw - 120, 18, { fontSize: 11, color: K.sub });
    });
    card(s, 64, 308, 1152, 348);
    const lx = 96, gx = 330, gw = 1184 - gx, mw = gw / 10;
    for (let m = 0; m < 10; m++) {
      T(s, (m + 1) + '월', gx + m * mw, 330, mw, 16, { fontFace: NUM, fontSize: 10, bold: true, color: K.sub, align: 'center' });
      vl(s, gx + m * mw, 356, 280);
    }
    vl(s, gx + 10 * mw, 356, 280);
    [['Plan', '현황 · 원인 분석', 1, 3, K.n1], ['Do', '플랫폼 개발', 3, 4, K.n2], ['Do', '현장 시범 적용', 5, 7, K.n2],
      ['Check', '결과 확인', 7, 7, K.n3], ['Act', '표준화 · AI 고도화', 8, 10, K.navy], ['', 'QI 경진대회', 9, 9, K.blue]].forEach((r, i) => {
      const y = 370 + i * 44;
      T(s, r[0], lx, y, 60, 26, { fontFace: NUM, fontSize: 10, bold: true, color: K.navy, valign: 'middle' });
      T(s, r[1], lx + 62, y, 170, 26, { fontSize: 12, valign: 'middle' });
      bar(s, gx + (r[2] - 1) * mw + 5, y + 4, (r[3] - r[2] + 1) * mw - 10, 18, r[4]);
    });
    s.addNotes('- 김광현: 총괄 / 장은선(보건관리자): 플랫폼 기획 · 직접 코딩 / 장종문(안전관리자): MSDS · AI 위험성평가\n- 4월 1주 프로토타입 → 고위험 부서 우선 → 전 부서 확대\n- 8~10월 표준화 이후 AI 고도화 (Slide 11)');
  }
  // ───── 04 배경 (확정 구성: 3카드 + 결론 바)
  {
    const s = slide();
    chrome(s, 'PLAN · 주제 선정 배경', 0, '법은 실시간 증빙을 요구했지만, 안전보건 기록은 100% 종이였다', '주제 선정 배경 — 법 · 현장 · 경영 세 측면', 4);
    const cw = (1152 - 2 * 20) / 3;
    [['01', '법 · 인증', '산업안전보건법 제29조\n중대재해처벌법\n의료기관 인증기준 10.7 · 11.3', '실시간 증빙 요구'],
      ['02', '현장', '교육일지 100% 수기 작성\n대면 서명 · 오프라인 결재\n바인더 편철 보관', '건당 최대 33:48'],
      ['03', '경영 (ESG)', '연 94건 종이 출력 · 보관\n친환경 경영 방향과 불일치\nMSDS 개정본 비치 확인 어려움', '연 94건 출력']].forEach((c, i) => {
      const x = 64 + i * (cw + 20);
      card(s, x, 176, cw, 340);
      T(s, c[0], x + 32, 204, 80, 24, { fontFace: NUM, fontSize: 16, bold: true, color: K.navy });
      T(s, c[1], x + 32, 236, cw - 64, 34, { fontSize: 22, bold: true });
      T(s, c[2], x + 32, 292, cw - 64, 120, { fontSize: 15, color: '3A3A3C', lineSpacingMultiple: 1.65 });
      hl(s, x + 32, 430, cw - 64);
      T(s, c[3], x + 32, 448, cw - 64, 40, { fontFace: NUM, fontSize: 22, bold: true, color: i === 1 ? K.navy : K.ink });
    });
    card(s, 64, 540, 1152, 116, K.navy, 22);
    T(s, '모바일 기반 통합 디지털 안전보건 체계가 필요했습니다', 108, 540, 960, 116, { fontSize: 24, bold: true, color: K.w, valign: 'middle' });
    T(s, '→', 1120, 540, 60, 116, { fontFace: NUM, fontSize: 26, bold: true, color: '4DA3FF', valign: 'middle', align: 'right' });
    s.addNotes('- 교육 미이행 시 과태료, 중대재해 시 교육 이행 기록이 핵심 증빙\n- 인증 조사 때 바인더를 넘기며 누락자를 찾던 경험\n- 근거: 산안법 제29조, 인증기준 10.7 · 11.3');
  }
}

// ───── 05 현황 — 시안 B · C
const MONTHS = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
const DONE = [5, 8, 16, 7, 6, 9, 11, 6, 4, 1, 10, 5], MISS = [3, 1, 2, 0, 1, 0, 2, 0, 1, 1, 7, 0];
const H5 = '행정시간의 78%는 이동 · 대기였고, 교육 6건 중 1건은 누락됐다';
function monthly(s, x, y, w, h) {
  s.addChart(pres.charts.BAR, [{ name: '이행', labels: MONTHS, values: DONE }, { name: '누락', labels: MONTHS, values: MISS }],
    { ...B(x - 6, y, w + 6, h), barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 60, chartColors: [K.navy, K.orange], showLegend: false,
      valAxisMaxVal: 20, valAxisMajorUnit: 10, valGridLine: { color: K.line, size: 0.5 }, catGridLine: { style: 'none' },
      catAxisLabelColor: K.sub, valAxisLabelColor: K.sub, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelFontFace: '+mn-lt', valAxisLabelFontFace: '+mn-lt',
      valAxisLineShow: false, catAxisLineColor: K.line, objectName: nm() });
}
const STEPS = [['① 초안 작성', 138, 0], ['② 부서장 전송', 60, 0], ['③ 출력 · 서명', 90, 0], ['④ 보건실 직접 제출', 420, 1], ['⑤ 사진 전송', 30, 0], ['⑥ 취합 · 파일링', 126, 0], ['⑦ 아날로그 결재', 1164, 1]];
function slide5B(tag) {
  const s = slide();
  chrome(s, 'PLAN · 현황 분석', 0, H5, tag || '2025년 법정교육 누락(좌) · 특별안전보건교육 1건의 행정시간(우)', 5);
  card(s, 64, 176, 660, 480); card(s, 744, 176, 472, 480);
  T(s, '17%', 96, 200, 200, 70, { fontFace: NUM, fontSize: 52, bold: true });
  T(s, '2025년 신규 입사자 법정교육\n106건 중 18건 누락 (이행률 83%)', 290, 212, 400, 50, { fontSize: 13, color: K.sub, lineSpacingMultiple: 1.35 });
  T(s, '이행', 600, 286, 90, 16, { fontSize: 11, bold: true, color: K.navy, align: 'right' });
  T(s, '누락', 600, 304, 90, 16, { fontSize: 11, bold: true, color: K.orangeT, align: 'right' });
  monthly(s, 96, 300, 600, 290);
  T(s, '11월: 17건 중 7건 누락 (이행률 59%) — 입사가 몰린 달일수록 누락이 커졌다', 96, 608, 600, 22, { fontSize: 12, bold: true, color: K.orangeT });
  T(s, '78%', 776, 200, 170, 70, { fontFace: NUM, fontSize: 52, bold: true, color: K.navy });
  T(s, '특별안전보건교육 1건 33:48\n이동 · 결재 대기 26:24', 946, 212, 250, 50, { fontSize: 13, color: K.sub, lineSpacingMultiple: 1.35 });
  STEPS.forEach((r, i) => {
    const y = 300 + i * 42;
    T(s, r[0], 776, y, 150, 26, { fontSize: 11, bold: !!r[2], valign: 'middle' });
    const w = Math.max(4, r[1] / 1164 * 190);
    bar(s, 930, y + 6, w, 14, r[2] ? K.navy : K.grey);
    T(s, fmt(r[1]), 938 + w, y, 60, 26, { fontFace: NUM, fontSize: 11, bold: true, color: r[2] ? K.navy : K.sub, valign: 'middle' });
  });
  T(s, 'MSDS 교육 1건도 4:18 (검색 · 인쇄 · 서명 · 편철)', 776, 608, 420, 22, { fontSize: 11, color: K.sub });
  s.addNotes('- 자료수집: 입사자 명단 ↔ 교육일지 대조 전수조사(2025.01~12), 단계별 직접 관찰(Time Study)(2025.01)\n- 1월 · 10월 · 11월 이행률 63 · 50 · 59% — 문제는 사람이 아니라 확인 체계\n- 이동 7분 + 결재 대기 19분 24초 = 26분 24초 (78%)');
}
function slide5C() {
  const s = slide();
  chrome(s, 'PLAN · 현황 분석', 0, H5, '시안 C — 교육일지 1건의 경로(상) · 월별 누락(하)', 5);
  const nodes = [['작성', '2:18', 0], ['전송', '1:00', 0], ['출력 · 서명', '1:30', 0], ['직접 제출', '7:00', 1], ['사진 전송', '0:30', 0], ['취합 · 편철', '2:06', 0], ['결재', '19:24', 1]];
  const nw = (1152 - 6 * 16) / 7;
  nodes.forEach((nd, i) => {
    const x = 64 + i * (nw + 16);
    card(s, x, 176, nw, 112, nd[2] ? K.navy : K.surf, 18);
    T(s, String(i + 1).padStart(2, '0'), x + 16, 190, nw - 24, 16, { fontFace: NUM, fontSize: 10, bold: true, color: nd[2] ? K.onNavy : K.sub });
    T(s, nd[0], x + 16, 210, nw - 24, 22, { fontSize: 13, bold: true, color: nd[2] ? K.w : K.ink });
    T(s, nd[1], x + 16, 240, nw - 24, 36, { fontFace: NUM, fontSize: 22, bold: true, color: nd[2] ? K.w : K.ink });
  });
  T(s, '합계 33:48 · 이동 + 대기 26:24 (78%) · MSDS 교육 1건 4:18', 64, 300, 1152, 20, { fontSize: 12, bold: true, color: K.sub });
  card(s, 64, 336, 744, 320);
  T(s, '월별 신규 입사자 교육 건수 (건) — 주황 = 누락', 96, 356, 680, 18, { fontSize: 12, bold: true, color: K.sub });
  monthly(s, 96, 380, 680, 260);
  const kw = (388 - 16) / 2;
  [['누락률', '17%', '18 / 106건'], ['이행률', '83%', '88 / 106건'], ['11월 이행률', '59%', '17건 중 7건 누락'], ['비부가가치', '78%', '26:24 / 33:48']].forEach((k, i) => {
    const x = 828 + (i % 2) * (kw + 16), y = 336 + Math.floor(i / 2) * 168;
    card(s, x, y, kw, 152);
    T(s, k[0], x + 22, y + 22, kw - 44, 18, { fontSize: 12, color: K.sub });
    T(s, k[1], x + 22, y + 50, kw - 44, 50, { fontFace: NUM, fontSize: 36, bold: true, color: i === 3 ? K.navy : K.ink });
    T(s, k[2], x + 22, y + 108, kw - 44, 18, { fontSize: 11, color: K.sub });
  });
}


// 5장 시안 D — 월별 이행률(목표선) + 파레토
function slide5D() {
  const s = slide();
  chrome(s, 'PLAN · 현황 분석', 0, H5, '시안 D — 월별 이행률(좌) · 행정시간 파레토(우)', 5);
  card(s, 64, 176, 660, 480); card(s, 744, 176, 472, 480);
  // 왼쪽: 월별 이행률 + 100% 목표선
  T(s, '83%', 96, 200, 200, 70, { fontFace: NUM, fontSize: 52, bold: true });
  T(s, '2025년 법정교육 이행률 (106건 중 88건)\n12개월 중 8개월에서 누락 발생', 280, 212, 420, 50, { fontSize: 13, color: K.sub, lineSpacingMultiple: 1.35 });
  const ox = 132, oy = 580, ch = 250, cw = 560, slot = cw / 12, bw = slot * 0.6;
  [0, 50, 100].forEach(v => { const y = oy - v / 100 * ch; hl(s, ox, y, cw, K.line); T(s, v + '%', ox - 40, y - 8, 34, 16, { fontFace: NUM, fontSize: 9, color: K.sub, align: 'right' }); });
  s.addShape(pres.shapes.LINE, { ...B(ox, oy - ch, cw, 0), line: { color: K.navy, width: 1, dashType: 'dash' }, objectName: nm() });
  T(s, '목표 100%', ox + cw - 80, oy - ch - 20, 80, 16, { fontSize: 10, bold: true, color: K.navy, align: 'right' });
  DONE.forEach((d, i) => {
    const t = d + MISS[i], r = d / t, h = r * ch, x = ox + i * slot + (slot - bw) / 2, gap = MISS[i] > 0;
    bar(s, x, oy - h, bw, h, gap ? K.orange : K.grey);
    if (gap) T(s, Math.round(r * 100) + '', x - 8, oy - h - 18, bw + 16, 14, { fontFace: NUM, fontSize: 9, bold: true, color: K.orangeT, align: 'center' });
    T(s, MONTHS[i], x - 8, oy + 6, bw + 16, 14, { fontSize: 9, color: K.sub, align: 'center' });
  });
  T(s, '주황 = 누락이 있었던 달 — 특정 시기가 아니라 1년 내내 반복됐다', 96, 616, 600, 20, { fontSize: 12, bold: true, color: K.orangeT });
  // 오른쪽: 파레토
  T(s, '78%', 776, 200, 170, 70, { fontFace: NUM, fontSize: 52, bold: true, color: K.navy });
  T(s, '두 단계(결재 · 직접 제출)가\n33:48 중 26:24를 차지', 946, 212, 250, 50, { fontSize: 13, color: K.sub, lineSpacingMultiple: 1.35 });
  const P = [['결재', 1164], ['직접 제출', 420], ['초안', 138], ['파일링', 126], ['출력·서명', 90], ['전송', 60], ['사진', 30]];
  const px = 800, py = 580, ph = 250, pw = 380, ps = pw / 7, pb = ps * 0.62, max = 1164;
  hl(s, px, py, pw, K.line);
  let cum = 0; const pts = [];
  P.forEach((r, i) => {
    const h = r[1] / max * ph, x = px + i * ps + (ps - pb) / 2, key = i < 2;
    bar(s, x, py - h, pb, h, key ? K.navy : K.grey);
    T(s, fmt(r[1]), x - 10, py - h - 16, pb + 20, 14, { fontFace: NUM, fontSize: 9, bold: true, color: key ? K.navy : K.sub, align: 'center' });
    T(s, r[0], x - 12, py + 6, pb + 24, 14, { fontSize: 8, color: K.sub, align: 'center' });
    cum += r[1]; pts.push([x + pb / 2, py - cum / 2028 * ph]);
  });
  for (let i = 0; i < pts.length - 1; i++) {
    const [x1, y1] = pts[i], [x2, y2] = pts[i + 1];
    s.addShape(pres.shapes.LINE, { ...B(x1, Math.min(y1, y2), x2 - x1, Math.abs(y2 - y1)), flipV: y2 < y1, line: { color: K.blue, width: 1.5 }, objectName: nm() });
  }
  pts.forEach((pt, i) => dot(s, pt[0] - 4, pt[1] - 4, 8, K.blue));
  T(s, '누적 78%', pts[1][0] + 10, pts[1][1] + 8, 80, 16, { fontFace: NUM, fontSize: 10, bold: true, color: K.blue });
  T(s, '누적 100% = 33:48', px + pw - 120, py - ph - 20, 120, 16, { fontFace: NUM, fontSize: 9, color: K.sub, align: 'right' });
  T(s, '파레토: 상위 2단계 개선 = 행정시간 78% 해결', 776, 616, 420, 20, { fontSize: 12, bold: true, color: K.navy });
}

if (MODE === 'v2') { slide5B('시안 B — 현행 (월별 건수 · 공정 순서 막대)'); slide5D(); }
else if (MODE === 'variants') { slide5B('시안 B — 누락(좌) · 시간(우) 두 패널'); slide5C(); }
else { deck14(); }

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log('ok', MODE); })();
