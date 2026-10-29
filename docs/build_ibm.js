// QI 경진대회 발표 — IBM Carbon 숫자 강조형 12장 (한 장 = 큰 숫자 하나 + 근거 3줄 + 핵심 요약 한 줄)
// node build_ibm.js out.pptx [대본.md] — 발표자 노트는 QI_발표대본.md의 '## N.' 절에서 읽음
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
const fs = require('fs');
const SCRIPT = process.argv[3] || __dirname + '/QI_발표대본.md';
const NOTES = {};
fs.readFileSync(SCRIPT, 'utf8').split(/^## /m).slice(1).forEach(sec => {
  const m = sec.match(/^(\d+)\. /); if (!m) return;
  NOTES[m[1]] = sec.replace(/\*\*/g, '').replace(/^> /gm, '').replace(/^---\s*$/gm, '').trim();
});
let sn = 0;
const slide = () => { const s = pres.addSlide(); s.background = { color: K.w }; const t = NOTES[++sn]; if (t) s.addNotes(t); return s; };
const LOGO = fs.existsSync(__dirname + '/assets/bumin_logo.jpg') ? __dirname + '/assets/bumin_logo.jpg' : __dirname + '/logo/img0.jpg';

// 1 표지
{
  const s = slide();
  s.addImage({ path: LOGO, ...B(64, 48, 140, 36), objectName: nm() });
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
// 2 목차
{
  const s = slide();
  frame(s, 'CONTENTS  ·  목차', -1, '무엇이 달라졌는지, PDCA 순서로 말씀드립니다', 2);
  const cw = (1152 - 3 * 32) / 4;
  [['Plan', '17%', '2025년 교육 누락률', '01  WHY · 활동 배경\n팀 구성\n02  PROBLEM · 현황\n03  원인 · 목표', '03 – 06'],
   ['Do', '6개', '플랫폼 모듈', '04  SOLUTION · 플랫폼\n05  DEMO · 시연\nAI 적용', '07 – 09'],
   ['Check', '−93%', '행정시간', '06  RESULT · 전후 비교\n07  IMPACT · 성과 · ESG', '10 – 11'],
   ['Act', '표준화', '정식 업무 전환', '08  NEXT · 결론과 제언', '12']].forEach((c, i) => {
    const x = 64 + i * (cw + 32);
    R(s, x, 196, cw, 2, i === 2 ? K.blue : K.ink);
    T(s, c[0], x, 216, cw, 20, { fontSize: 14, bold: true, color: K.blue });
    T(s, c[1], x - 4, 248, cw + 8, 90, { fontSize: 60, color: i === 2 ? K.blue : K.ink });
    T(s, c[2], x, 344, cw, 20, { fontSize: 12, color: K.sub });
    HL(s, x, 392, cw);
    T(s, c[3], x, 412, cw, 180, { fontSize: 14, color: K.ink2, lineSpacingMultiple: 1.7 });
    T(s, 'Slide ' + c[4], x, 616, cw, 16, { fontSize: 10, color: K.cap });
  });
}
// 3 WHY
{
  const s = slide();
  frame(s, '01  WHY · 활동 배경  ·  주제 선정 평가 22.5점 · 병원 중점과제 정합성 검토 후 선정', 0, '법은 실시간 증빙을 요구했지만, 기록은 모두 종이였습니다', 3);
  T(s, '종이 기반 처리', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '100%', 58, 220, 640, 190, { fontSize: 150, color: K.blue });
  T(s, '연 94건 출력  →  대면 서명  →  방문 제출  →  바인더 보관', 64, 420, 640, 22, { fontSize: 14, color: K.ink2 });
  facts(s, 760, 196, 456, [['법 · 제도', '실시간 증빙 의무', K.ink, '중대재해처벌법 · 산업안전보건법 · 인증기준 10.7 · 11.3'], ['현장', '이수 누락 즉시 확인 불가', K.ink, '출력 · 서명 · 제출 · 결재 4단계 수기 행정'], ['경영', '업무 프로세스 개선', K.ink, '병원 중점과제 · ESG 환경(E) 부문']]);
  summary(s, '법은 "했다는 증거"를 요구합니다 — 모바일 기반 디지털 체계가 필요했습니다');
}
// 4 팀
{
  const s = slide();
  frame(s, 'PLAN  ·  팀 구성 · 2026년 1–10월 활동 일정', 0, '3명이 외부 개발 없이 기획부터 표준화까지 직접 수행했습니다', 4);
  T(s, '팀원 · 활동 기간', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '3명 · 10개월', 60, 224, 680, 100, { fontSize: 64, color: K.blue });
  const gx = 160, gw = 540, mw = gw / 10;
  for (let m = 0; m < 10; m++) T(s, (m + 1) + '월', gx + m * mw, 356, mw, 14, { fontSize: 9, color: K.cap, align: 'center' });
  HL(s, gx, 376, gw);
  [['Plan', '현황 · 원인 · 목표', 1, 3, K.grey], ['Do', '개발 · 전 부서 적용', 3, 7, '4589FF'], ['Check', '결과 확인', 7, 7, K.blue], ['Act', '표준화 · AI 고도화', 8, 10, K.ink]].forEach((r, i) => {
    const y = 388 + i * 40;
    T(s, r[0], 64, y, 90, 24, { fontSize: 12, bold: true, color: K.ink, valign: 'middle' });
    R(s, gx + (r[2] - 1) * mw + 2, y + 4, (r[3] - r[2] + 1) * mw - 4, 16, r[4]);
    const after = gx + r[3] * mw + 8;
    T(s, r[1], after < gx + gw - 120 ? after : gx + (r[2] - 1) * mw - 168, y, 164, 24, { fontSize: 11, color: K.sub, valign: 'middle', align: after < gx + gw - 120 ? 'left' : 'right' });
  });
  facts(s, 760, 196, 456, [['팀장', '김광현', K.ink, '활동 총괄'], ['팀원 · 보건관리자', '장은선', K.blue, '플랫폼 기획 · 직접 개발'], ['팀원 · 안전관리자', '장종문', K.ink, 'MSDS · AI 위험성평가']]);
  summary(s, '외부 개발비 없이, 현장을 아는 담당자가 직접 만들고 직접 표준화했습니다');
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
// 6 원인 · 목표
{
  const s = slide();
  frame(s, '03  ROOT CAUSE & GOAL  ·  업무 흐름도 · Fishbone(4M) → 기여요인(CF) 도출', 0, '원인은 사람이 아니라 프로세스 — 현황을 기준값으로 목표를 정했습니다', 6);
  T(s, '핵심 원인', 64, 196, 520, 20, { fontSize: 14, color: K.sub });
  [['01', '결재 대기', '19분 24초', 'Method · 종이 결재 · 명단 수작업 대조'], ['02', '방문 제출', '평균 7분', 'Environment · 부서 ↔ 보건관리실 이동'], ['03', '실시간 모니터링 부재', '확인 불가', 'Machine · 바인더 검색 · 반복 수기 작성']].forEach((r, i) => {
    const y = 228 + i * 108;
    HL(s, 64, y, 560);
    T(s, r[0], 64, y + 16, 50, 24, { fontSize: 14, bold: true, color: K.blue });
    T(s, r[1], 116, y + 14, 260, 28, { fontSize: 20, bold: true });
    T(s, r[2], 376, y + 14, 248, 28, { fontSize: 20, color: K.red, align: 'right' });
    T(s, r[3], 116, y + 56, 508, 18, { fontSize: 11, color: K.cap });
  });
  T(s, '지표', 704, 196, 200, 20, { fontSize: 14, color: K.sub });
  T(s, '기준값', 920, 196, 130, 20, { fontSize: 14, color: K.sub });
  T(s, '목표', 1066, 196, 150, 20, { fontSize: 14, color: K.sub, align: 'right' });
  [['법정교육 누락률', '17%', '0%'], ['행정 리드타임', '33분 48초', '70%↓'], ['종이 교육일지', '연 94건', '0건'], ['6대 모듈 구축', '0 / 6', '6 / 6']].forEach((g, i) => {
    const y = 228 + i * 81;
    HL(s, 704, y, 512);
    T(s, g[0], 704, y + 22, 210, 36, { fontSize: 16, bold: true, valign: 'middle' });
    T(s, g[1], 920, y + 22, 140, 36, { fontSize: 16, color: K.sub, valign: 'middle' });
    T(s, g[2], 1056, y + 14, 160, 52, { fontSize: 32, color: K.blue, align: 'right', valign: 'middle' });
  });
  summary(s, '원인 하나에 기능 하나 — 세 원인을 없애는 방향으로 4개 목표를 숫자로 정했습니다');
}
// 7 SOLUTION
{
  const s = slide();
  frame(s, '04  SOLUTION  ·  AppSheet · Google Sheets · Apps Script 기반 자체 개발', 1, '원인마다 하나씩, 6개 모듈로 설계했습니다', 7);
  const cw = (1152 - 2 * 16) / 3;
  [['01', 'CF1', 'MSDS 디지털 라이브러리', '부서별 유해화학물질 조회 · 경고표지 ·\n응급조치 · 원본 PDF 원클릭'],
   ['02', 'CF2 · 3', '모바일 교육일지 · 전자서명', '인사 DB 연동, 이름 검색만으로\n자동 작성 · 서명 즉시 저장'],
   ['03', 'CF4 · 5', '비대면 전자결재', '서명 완료 즉시 관리자 알림 ·\n장소 제약 없는 원격 승인'],
   ['04', 'CF7', '보호구 · 법정서류 전산화', '보호구 수령증 · 유해화학물질\n신규 · 중단 보고서 단일 창구'],
   ['05', 'CF6', '관리자 모니터링 대시보드', '부서 · 기간 · 근로자별\n교육 이수 현황 실시간 조회'],
   ['06', 'AI', 'AI 수시 위험성평가', '현장 사진 · 작업 입력 → 위험요인 ·\n위험도 산출 → 평가서 초안']].forEach((m, i) => {
    const x = 64 + (i % 3) * (cw + 16), y = 196 + Math.floor(i / 3) * 180, ai = i === 5;
    R(s, x, y, cw, 164, ai ? K.blue : K.bg2);
    T(s, m[0], x + 24, y + 20, 60, 20, { fontSize: 14, bold: true, color: ai ? K.w : K.blue });
    T(s, m[1], x + cw - 124, y + 20, 100, 20, { fontSize: 11, color: ai ? K.w : K.cap, align: 'right' });
    T(s, m[2], x + 24, y + 52, cw - 48, 28, { fontSize: 18, bold: true, color: ai ? K.w : K.ink });
    T(s, m[3], x + 24, y + 92, cw - 48, 56, { fontSize: 12, color: ai ? K.w : K.ink2, lineSpacingMultiple: 1.4 });
  });
  summary(s, '모바일 작성 → 전자서명 → 비대면 승인 → 자동 보관  ·  61병동 시범 후 전 부서 확대', 568);
}
// 8 DEMO
{
  const s = slide();
  frame(s, '05  DEMO  ·  프로그램 시연 영상 1분 30초', 1, '신규 입사자 1명의 교육, 작성부터 승인까지 한 화면에서', 8);
  R(s, 64, 196, 784, 441, K.bg2);
  T(s, '시연 영상 (16:9)', 64, 196, 784, 441, { fontSize: 16, color: K.cap, align: 'center', valign: 'middle' });
  [['0:00', 'MSDS 라이브러리 조회'], ['0:20', '교육일지 작성 · 전자서명'], ['0:50', '관리자 알림 · 비대면 승인'], ['1:10', '대시보드 · AI 위험성평가']].forEach((t, i) => {
    const y = 196 + i * 92;
    HL(s, 888, y, 328);
    T(s, t[0], 888, y + 14, 120, 30, { fontSize: 22, color: i === 2 ? K.blue : K.ink });
    T(s, t[1], 888, y + 50, 328, 22, { fontSize: 14, bold: i === 2, color: i === 2 ? K.blue : K.ink2 });
  });
  R(s, 888, 568, 4, 69, K.blue);
  T(s, '"이 1분이 예전의\n19분 24초입니다"', 908, 568, 308, 69, { fontSize: 15, bold: true, valign: 'middle' });
}
// 9 AI
{
  const s = slide();
  frame(s, '05+  AI 적용  ·  AI 수시 위험성평가 — 근로자 참여형 · 2026.05–07 시범 운영', 1, 'AI는 초안을 만들고, 최종 판단은 사람이 합니다', 9);
  R(s, 64, 196, 640, 360, K.bg2);
  T(s, 'AI 위험성평가 결과 화면 캡처\n현장 사진 → 위험요인 → 위험도(5×4) → 개선대책', 64, 196, 640, 360, { fontSize: 14, color: K.cap, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });
  facts(s, 760, 196, 456, [['AI 수시 위험성평가 · 운영 실적', '[실적]건', K.blue, '사진 한 장 → 위험요인 · 위험도 · 개선대책 초안'], ['AI MSDS 검토', '규제 항목 제안', K.ink, '관리자 승인 후 반영'], ['AI 도우미', '출처와 함께 답변', K.ink, '법령 · 지침 질의 응답']]);
  summary(s, 'AI는 초안, 확정은 사람 — 점수 · 등급은 서버가 다시 계산하고 법령 조문을 대조합니다');
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
// 12 NEXT
{
  const s = slide();
  frame(s, '08  NEXT  ·  표준화 · 지속 관리 · 고도화 · 제언', 3, '검증된 프로세스를 병원의 표준으로', 12);
  [['표준화', '완료', '법정 안전보건 관리 표준 업무지침', '시범사업을 정식 업무로 전환'], ['지속 관리', '운영', '대시보드 실시간 점검 · CQI 연계', '신규 입사자 이수 현황 · 현장 피드백'],
   ['고도화', '완료', '미이수자 부서장 자동 알림 · AI 2종', 'AI MSDS 검토 · AI 도우미 (2026.08–10)'], ['확장', '제언', '건강상담 연 819건 · 기타 안전보건 서류', 'AI 위험성평가 참여 확대']].forEach((r, i) => {
    const y = 196 + i * 112, hi = r[1] === '제언';
    HL(s, 64, y, 656);
    T(s, r[0], 64, y + 16, 140, 22, { fontSize: 13, bold: true, color: K.sub });
    T(s, r[1], 620, y + 16, 100, 22, { fontSize: 13, bold: true, color: hi ? K.blue : K.ink, align: 'right' });
    T(s, r[2], 64, y + 44, 656, 30, { fontSize: 19, bold: true, color: hi ? K.blue : K.ink });
    T(s, r[3], 64, y + 78, 656, 20, { fontSize: 12, color: K.cap });
  });
  R(s, 760, 196, 456, 444, K.bg2); R(s, 760, 196, 456, 4, K.blue);
  T(s, '맺음말', 792, 228, 300, 20, { fontSize: 13, color: K.sub });
  T(s, '기다리지 않는 행정,\n누락 없는 안전.', 792, 260, 400, 110, { fontSize: 30, color: K.ink, lineSpacingMultiple: 1.2 });
  T(s, '안전보건 데이터가 실시간으로 흐를 때,\n관리자는 서류가 아닌 현장에 집중합니다.', 792, 390, 400, 60, { fontSize: 14, color: K.ink2, lineSpacingMultiple: 1.45 });
  HL(s, 792, 520, 392, K.grey);
  T(s, '감사합니다', 792, 540, 392, 48, { fontSize: 32, bold: true, color: K.blue });
  T(s, '보건관리실 내선 3293 · 3034', 792, 596, 392, 20, { fontSize: 11, color: K.cap });
}
(async () => { await pres.writeFile({ fileName: OUT }); console.log('ok', sn); })();
