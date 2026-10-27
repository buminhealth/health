// 시안: IBM Carbon · 숫자 강조형(한 장 = 큰 숫자 하나 + 근거 3줄 + 핵심 요약 한 줄)
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const KO = 'IBM Plex Sans KR';
const K = { blue: '0F62FE', ink: '161616', ink2: '393939', sub: '525252', cap: '6F6F6F', bg2: 'F4F4F4', line: 'E0E0E0', red: 'DA1E28', grey: 'C6C6C6', w: 'FFFFFF' };
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: KO, bodyFontFace: KO };
const p = v => v / 96, B = (x, y, w, h) => ({ x: p(x), y: p(y), w: p(w), h: p(h) });
let n = 0; const nm = () => 'o' + (++n);
const T = (s, t, x, y, w, h, o = {}) => s.addText(t, { ...B(x, y, w, h), isTextBox: true, margin: 0, fontFace: KO, fontSize: 14, color: K.ink, valign: 'top', objectName: nm(), ...o });
const R = (s, x, y, w, h, c) => s.addShape(pres.shapes.RECTANGLE, { ...B(x, y, w, h), fill: { color: c }, line: { type: 'none' }, objectName: nm() });
const HL = (s, x, y, w, c = K.line) => s.addShape(pres.shapes.LINE, { ...B(x, y, w, 0), line: { color: c, width: 0.75 }, objectName: nm() });
const PH = ['Plan', 'Do', 'Check', 'Act'];
function frame(s, eyebrow, phase, head, page) {
  PH.forEach((ph, i) => { const x = 64 + i * 72; T(s, ph, x, 36, 64, 16, { fontSize: 9, bold: i === phase, color: i === phase ? K.blue : K.cap }); if (i === phase) R(s, x, 54, 40, 2, K.blue); });
  T(s, eyebrow, 64, 84, 900, 18, { fontSize: 12, color: K.sub });
  T(s, head, 64, 108, 1152, 50, { fontSize: 26, color: K.ink });
  HL(s, 64, 664, 1152);
  T(s, '부산부민병원 총무-안전보건팀  ·  2026 QI 경진대회', 64, 672, 700, 18, { fontSize: 9, color: K.cap });
  T(s, String(page).padStart(2, '0'), 1116, 672, 100, 18, { fontSize: 9, color: K.cap, align: 'right' });
}
function summary(s, text, y = 572) {
  R(s, 64, y, 1152, 76, K.bg2); R(s, 64, y, 4, 76, K.blue);
  T(s, '핵심 요약', 92, y, 120, 76, { fontSize: 12, bold: true, color: K.blue, valign: 'middle' });
  T(s, text, 212, y, 990, 76, { fontSize: 17, valign: 'middle' });
}
function facts(s, x, y, w, rows) {
  rows.forEach((r, i) => {
    const yy = y + i * 116;
    HL(s, x, yy, w);
    T(s, r[0], x, yy + 14, w, 18, { fontSize: 12, color: K.sub });
    T(s, r[1], x, yy + 38, w, 44, { fontSize: 24, bold: true, color: r[2] || K.ink });
    if (r[3]) T(s, r[3], x, yy + 84, w, 18, { fontSize: 11, color: K.cap });
  });
}
const slide = () => pres.addSlide();

// 1 표지
{
  const s = slide();
  s.addImage({ path: __dirname + '/assets/bumin_logo.jpg', ...B(64, 48, 140, 36), objectName: nm() });
  T(s, '2026 가을 학술대회 · QI 경진대회 구연', 64, 168, 640, 20, { fontSize: 13, color: K.sub });
  T(s, 'AI 통합 안전보건 플랫폼\n구축을 통한 업무 효율성 증가', 64, 198, 700, 120, { fontSize: 32, color: K.ink, lineSpacingMultiple: 1.15 });
  T(s, '종이 · 수기 · 대면 결재 → 모바일 · 전자서명 · 비대면 승인', 64, 330, 700, 24, { fontSize: 15, color: K.sub });
  T(s, '장은선 · 김광현 · 장종문', 64, 520, 640, 22, { fontSize: 15, bold: true });
  T(s, '인당의료재단 부민병원 총무-안전보건팀 (Safety First)', 64, 548, 640, 20, { fontSize: 12, color: K.cap });
  R(s, 800, 0, 480, 720, K.bg2);
  T(s, '특별안전보건교육 1건 행정시간', 840, 168, 400, 20, { fontSize: 13, color: K.sub });
  T(s, '−93%', 836, 200, 440, 130, { fontSize: 88, color: K.blue });
  T(s, '33분 48초 → 2분 22초', 840, 356, 400, 30, { fontSize: 20, bold: true });
  HL(s, 840, 420, 376, K.grey);
  [['법정교육 누락', '18 → 0건'], ['종이 출력', '94 → 0건']].forEach((r, i) => {
    T(s, r[0], 840 + i * 196, 440, 180, 18, { fontSize: 12, color: K.sub });
    T(s, r[1], 840 + i * 196, 462, 190, 40, { fontSize: 26, bold: true });
  });
}
// 5 현황
{
  const s = slide();
  frame(s, '02  현황 파악  ·  2025년 1–12월 전수조사 · 단계별 시간 실측', 0, '교육 1건 33분 48초 — 그중 78%는 기다림과 이동이었습니다', 5);
  T(s, '기다림 + 이동', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '78%', 58, 220, 600, 190, { fontSize: 150, color: K.blue });
  T(s, '결재 대기 19분 24초 + 직접 제출 7분 = 26분 24초', 64, 420, 600, 22, { fontSize: 14, color: K.ink2 });
  // 얇은 구성 막대
  [[1164, K.blue], [420, '4589FF'], [444, K.grey]].reduce((x, g) => { const w = g[0] / 2028 * 600; R(s, x, 470, w - 3, 14, g[1]); return x + w; }, 64);
  T(s, '결재 대기 57%', 64, 492, 200, 18, { fontSize: 11, color: K.sub });
  T(s, '이동 21%', 410, 492, 120, 18, { fontSize: 11, color: K.sub });
  T(s, '작성 등 22%', 560, 492, 120, 18, { fontSize: 11, color: K.sub });
  facts(s, 760, 196, 456, [['특별안전보건교육 1건 행정시간', '33분 48초', K.ink, '교육 진행시간 제외'], ['법정교육 누락', '18건 (17%)', K.red, '106명 중 · 12개월 중 8개월 발생'], ['종이 기반 처리', '100%', K.ink, '연 94건 출력 · 바인더 보관']]);
  summary(s, '문제는 사람이 아니라 방식 — 기다림과 이동은 체계를 바꿔야만 없어집니다');
}
// 10 결과
{
  const s = slide();
  frame(s, '06  개선 전후 비교  ·  특별안전보건교육 1건 · 세부 행동 단위 실측', 2, '같은 교육, 행정 시간은 93% 줄었습니다', 10);
  T(s, '행정시간 단축', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '−93%', 58, 220, 660, 190, { fontSize: 140, color: K.blue });
  T(s, '목표 70% 초과 달성  ·  MSDS 교육도 −81% (4분 18초 → 49초)', 64, 420, 640, 22, { fontSize: 14, color: K.ink2 });
  T(s, '개선 전', 64, 470, 100, 18, { fontSize: 11, color: K.sub }); R(s, 140, 472, 480, 14, K.grey); T(s, '33분 48초', 628, 468, 120, 20, { fontSize: 12, bold: true, color: K.sub });
  T(s, '개선 후', 64, 500, 100, 18, { fontSize: 11, color: K.sub }); R(s, 140, 502, 480 * 142 / 2028, 14, K.blue); T(s, '2분 22초', 190, 498, 120, 20, { fontSize: 12, bold: true, color: K.blue });
  facts(s, 760, 196, 456, [['내부 결재', '19분 24초 → 1분', K.ink, '비대면 전자결재'], ['보건관리실 직접 제출', '7분 → 단계 삭제', K.ink, '모바일 제출'], ['취합 · 인쇄 · 파일링', '2분 06초 → 단계 삭제', K.ink, '자동 보관']]);
  summary(s, '가장 긴 두 단계(결재 · 제출)를 없애 33분 48초가 2분 22초가 되었습니다');
}
// 11 성과
{
  const s = slide();
  frame(s, '07  성과  ·  개선 전 2025 → 개선 후 2026.05–07', 2, '4개 지표 모두 목표를 달성했습니다', 11);
  const cw = (1152 - 3 * 16) / 4;
  [['법정교육 누락', '0건', '18건 → 0건', '목표 0건'], ['행정시간 단축', '93%', '33분 48초 → 2분 22초', '목표 70%↑'], ['페이퍼리스 전환', '100%', '월 17.3건 전량 디지털', '목표 100%'], ['플랫폼 구축', '6/6', '핵심 모듈 100% 가동', '목표 6/6']].forEach((k, i) => {
    const x = 64 + i * (cw + 16), hi = i === 1;
    R(s, x, 196, cw, 330, hi ? K.blue : K.bg2);
    T(s, k[0], x + 24, 220, cw - 48, 20, { fontSize: 14, color: hi ? K.w : K.sub });
    T(s, k[1], x + 20, 262, cw - 30, 100, { fontSize: 60, color: hi ? K.w : K.ink });
    T(s, k[2], x + 24, 400, cw - 40, 40, { fontSize: 12, bold: true, color: hi ? K.w : K.ink });
    T(s, k[3], x + 24, 470, cw - 48, 20, { fontSize: 12, color: hi ? K.w : K.cap });
  });
  summary(s, 'E 종이 0건 · S 교육 누락 0건 · G 법적 증빙 무결성 — 효율을 넘어 ESG 성과로', 548);
}
(async () => { await pres.writeFile({ fileName: OUT }); console.log('ok'); })();
