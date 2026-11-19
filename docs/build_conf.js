// QI 학술대회 발표 B버전 13장 · IBM Carbon 숫자 강조형
// node build_conf.js out.pptx [대본.md] : 발표자 노트는 학술대회_발표대본.md의 'B. 학술 보고 버전' 절 '### N.'에서 읽음
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
const PH = ['서론', '방법', '결과', '고찰'];
function frame(s, eyebrow, phase, head, page) {
  PH.forEach((ph, i) => { const x = 64 + i * 72; T(s, ph, x, 36, 64, 16, { fontSize: 9, bold: i === phase, color: i === phase ? K.blue : K.cap }); if (i === phase) R(s, x, 54, 40, 2, K.blue); });
  T(s, eyebrow, 64, 84, 900, 18, { fontSize: 12, color: K.sub });
  T(s, head, 64, 108, 1152, 50, { fontSize: 26, color: K.ink });
  HL(s, 64, 664, 1152);
  T(s, '부산부민병원 총무-안전보건팀  ·  2026 QI 학술대회', 64, 672, 700, 18, { fontSize: 9, color: K.cap });
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
const SCRIPT = process.argv[3] || __dirname + '/학술대회_발표대본.md';
const NOTES = {};
fs.readFileSync(SCRIPT, 'utf8').split(/^## B\. /m)[1].split(/^---\s*$/m)[0].split(/^### /m).slice(1).forEach(sec => {
  const m = sec.match(/^(\d+)\. /); if (!m) return;
  NOTES[m[1]] = sec.trim();
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
  T(s, '서류 대신 환자 곁으로, 사무실 대신 현장으로', 64, 330, 700, 24, { fontSize: 16, color: K.blue });
  T(s, '장은선 · 김광현 · 장종문', 64, 520, 640, 22, { fontSize: 15, bold: true });
  T(s, '인당의료재단 부민병원 총무-안전보건팀 (Safety First)', 64, 548, 640, 20, { fontSize: 12, color: K.cap });
  R(s, 800, 0, 480, 720, K.bg2);
  T(s, '특별안전보건교육 1건 행정시간', 840, 168, 400, 20, { fontSize: 13, color: K.sub });
  T(s, '−93%', 836, 200, 440, 130, { fontSize: 88, color: K.blue });
  T(s, '33분 48초 → 2분 22초', 840, 356, 400, 30, { fontSize: 20, bold: true });
  HL(s, 840, 420, 376, K.grey);
  [['법정교육 준수율', '83 → 100%'], ['종이 출력', '94 → 0건']].forEach((r, i) => {
    T(s, r[0], 840 + i * 196, 440, 180, 18, { fontSize: 12, color: K.sub });
    T(s, r[1], 840 + i * 196, 462, 190, 40, { fontSize: 24, bold: true });
  });
}
// 2 배경 및 주제 선정
{
  const s = slide();
  frame(s, '배경 및 주제 선정  ·  실현 가능성 × 파급력 평가 → 병원 중점과제 정합성 검토', 0, '법은 교육 이행 증빙을 요구했고, 기록은 수기와 대면 결재였습니다', 2);
  T(s, '주제 선정 평가 점수', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '22.5점', 58, 220, 660, 190, { fontSize: 130, color: K.blue });
  T(s, '동점 후보 가운데 병원 중점과제 · ESG 경영 방향과의 정합성으로 최종 선정', 64, 420, 640, 22, { fontSize: 14, color: K.ink2 });
  facts(s, 760, 196, 456, [['법 · 제도', '교육 이행 증빙 의무', K.ink, '산업안전보건법 · 중대재해처벌법 · 인증기준 10.7 · 11.3'], ['병원 중점과제', '업무 프로세스 개선', K.ink, '수기 작성 · 대면 결재의 디지털 전환'], ['ESG 경영 방향', '종이 행정의 데이터화', K.ink, '연 94건 출력 · 바인더 보관']]);
  summary(s, '활동 목적: 법정 안전보건교육 준수율 100%, 행정 소요시간 단축');
}
// 3 활동 개요
{
  const s = slide();
  frame(s, '활동 개요  ·  팀 구성 · 2026년 1–10월 PDCA 일정', 1, '2026년 1~10월, 3명이 PDCA 방법으로 수행했습니다', 3);
  T(s, '팀원 · 활동 기간', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '3명 · 10개월', 60, 224, 680, 100, { fontSize: 64, color: K.blue });
  const gx = 160, gw = 540, mw = gw / 10;
  for (let m = 0; m < 10; m++) T(s, (m + 1) + '월', gx + m * mw, 356, mw, 14, { fontSize: 9, color: K.cap, align: 'center' });
  HL(s, gx, 376, gw);
  [['Plan', '현황 · 원인 · 지표', 1, 3, K.grey], ['Do', '개발 · 전 부서 적용', 3, 7, '4589FF'], ['Check', '결과 측정', 7, 7, K.blue], ['Act', '표준화 · AI 고도화', 8, 10, K.ink]].forEach((r, i) => {
    const y = 388 + i * 40;
    T(s, r[0], 64, y, 90, 24, { fontSize: 12, bold: true, color: K.ink, valign: 'middle' });
    R(s, gx + (r[2] - 1) * mw + 2, y + 4, (r[3] - r[2] + 1) * mw - 4, 16, r[4]);
    const after = gx + r[3] * mw + 8, fit = after < gx + gw - 120;
    T(s, r[1], fit ? after : gx + (r[2] - 1) * mw - 168, y, 164, 24, { fontSize: 11, color: K.sub, valign: 'middle', align: fit ? 'left' : 'right' });
  });
  facts(s, 760, 196, 456, [['팀장', '김광현', K.ink, '활동 총괄'], ['팀원 · 보건관리자', '장은선', K.blue, '플랫폼 기획 · 직접 개발'], ['팀원 · 안전관리자', '장종문', K.ink, 'MSDS · AI 위험성평가']]);
  summary(s, '플랫폼은 외주 개발 없이 보건관리자가 직접 개발했습니다');
}
// 4 현황 분석
{
  const s = slide();
  frame(s, '현황 분석  ·  2025.01–12 입사자 명단 ↔ 교육일지 전수조사 · 7단계 업무 직접 관찰', 1, '2025년 준수율 83.0%, 교육 1건 행정에 33분 48초가 걸렸습니다', 4);
  T(s, '법정교육 준수율 (2025)', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '83.0%', 58, 216, 640, 160, { fontSize: 110, color: K.blue });
  T(s, '106명 중 88명 기한 내 이수  ·  95% CI 74.5–89.6%', 64, 394, 640, 22, { fontSize: 14, color: K.ink2 });
  const rate = [63, 89, 89, 100, 86, 100, 85, 100, 80, 50, 59, 100], bx = 64, bw = 40, base = 530, bh = 90;
  HL(s, bx, base - bh, 12 * 48 - 8, K.grey);
  T(s, '100%', bx + 12 * 48, base - bh - 8, 60, 16, { fontSize: 9, color: K.cap });
  rate.forEach((v, i) => {
    const h = bh * v / 100;
    R(s, bx + i * 48, base - h, bw, h, v === 100 ? K.blue : K.grey);
    T(s, (i + 1) + '월', bx + i * 48, base + 4, bw, 14, { fontSize: 9, color: K.cap, align: 'center' });
  });
  facts(s, 760, 196, 456, [['특별안전보건교육 1건 행정시간', '33분 48초', K.ink, '교육 진행시간 제외 · 7단계 직접 관찰'], ['비부가가치 시간', '78%', K.ink, '결재 대기 57% + 방문 제출 21%'], ['종이 기반 처리', '100%', K.ink, '연 94건 출력 · 바인더 보관']]);
  summary(s, '12개월 중 4개월은 100%, 나머지 달은 50~89%로 달마다 차이가 컸습니다');
}
// 5 지표 정의 및 측정 방법
{
  const s = slide();
  frame(s, '지표 정의 및 측정 방법  ·  개선 전 2025.01–12 / 개선 후 2026.05–07', 1, '4개 지표의 정의와 측정 방법을 미리 정했습니다', 5);
  const hd = ['구분', '지표', '정의 · 산출식', '측정 방법', '기준값', '목표'];
  const rows = [['과정', '법정교육 준수율', '기한 내 이수 인원 ÷ 신규 입사자 수 × 100', '입사자 명단 ↔ 교육일지 전수조사', '83.0%\n(88/106)', '100%'],
    ['효율', '행정 소요시간 단축률', '(전 − 후) ÷ 전 × 100\n교육 진행시간 제외', '7단계 업무 직접 관찰\n(Time Study)', '33분 48초', '70% 이상'],
    ['결과', '페이퍼리스 전환율', '전산 저장 완결 건수 ÷ 대상 업무 건수 × 100', '출력 대장 · 시스템 저장 기록', '0%\n(연 94건 출력)', '100%'],
    ['인프라', '플랫폼 구축률', '정상 가동 모듈 ÷ 계획 모듈', '모듈별 운영 확인', '0 / 6', '6 / 6']];
  const cell = (t, o = {}) => ({ text: t, options: { fontFace: KO, fontSize: 12, color: K.ink, valign: 'middle', border: [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: K.line }, { type: 'none' }], ...o } });
  const data = [hd.map(h => cell(h, { fontSize: 11, bold: true, color: K.sub, fill: { color: K.bg2 } }))]
    .concat(rows.map(r => r.map((c, j) => cell(c, j === 1 ? { bold: true } : j === 5 ? { bold: true, color: K.blue, fontSize: 16 } : j === 0 ? { color: K.sub } : j === 4 ? { color: K.ink2 } : {}))));
  s.addTable(data, { ...B(64, 196, 1152, 360), colW: [72, 200, 330, 260, 160, 130].map(p), rowH: [36, 81, 81, 81, 81].map(p), objectName: nm() });
  summary(s, '행정 소요시간은 교육 진행 시간을 빼고 문서 작성부터 결재까지만 측정했습니다', 572);
}
// 6 원인 분석
{
  const s = slide();
  frame(s, '원인 분석  ·  업무 흐름도 · Fishbone(4M) → 기여요인(CF) 7개 · 효과 × 실행 난이도 우선순위', 1, '기다림과 이동은 체계를 바꿔야 줄어드는 시간이었습니다', 6);
  T(s, 'Fishbone 4M', 64, 196, 520, 20, { fontSize: 14, color: K.sub });
  [['방법', 'Method', '종이 결재 대기 · 제출 채널 이원화 · 명단 수작업 대조'], ['환경', 'Environment', '부서 ↔ 보건관리실 이동'], ['장비 · 시스템', 'Machine', '실시간 이수 현황 확인 불가 · 반복 수기 작성'], ['재료 · 매체', 'Material', '바인더 분산 보관 · MSDS 최신본 비치 확인 미흡']].forEach((r, i) => {
    const y = 228 + i * 81;
    HL(s, 64, y, 600);
    T(s, r[0], 64, y + 14, 140, 22, { fontSize: 15, bold: true });
    T(s, r[1], 64, y + 40, 140, 16, { fontSize: 10, color: K.cap });
    T(s, r[2], 210, y + 14, 454, 50, { fontSize: 13, color: K.ink2, lineSpacingMultiple: 1.3 });
  });
  T(s, '핵심 원인', 724, 196, 300, 20, { fontSize: 14, color: K.sub });
  [['01', '결재 대기', '19분 24초'], ['02', '방문 제출', '7분'], ['03', '실시간 확인 불가', '수작업 대조']].forEach((r, i) => {
    const y = 228 + i * 108;
    HL(s, 724, y, 492);
    T(s, r[0], 724, y + 18, 50, 24, { fontSize: 14, bold: true, color: K.blue });
    T(s, r[1], 776, y + 14, 260, 30, { fontSize: 20, bold: true });
    T(s, r[2], 1016, y + 14, 200, 30, { fontSize: 20, color: K.blue, align: 'right' });
  });
  summary(s, '효과가 크고 실행이 쉬운 반복 수기 · 이동 · 모니터링부터 개선했습니다');
}
// 7 개선 활동
{
  const s = slide();
  frame(s, '개선 활동  ·  Google Sheets · AppSheet · Apps Script 기반 자체 개발 · 외주 개발 없음', 1, '원인마다 하나씩, 6개 모듈로 설계했습니다', 7);
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
  summary(s, '2026.04 시제품 → 병동 1곳 시범 운영 → 전 부서 확대', 568);
}
// 8 시연
{
  const s = slide();
  frame(s, '개선 활동  ·  프로그램 시연 영상 1분 15초', 1, '신규 입사자 1명의 교육, 작성부터 승인까지 한 화면에서', 8);
  R(s, 64, 196, 784, 441, K.bg2);
  T(s, '시연 영상 (16:9)', 64, 196, 784, 441, { fontSize: 16, color: K.cap, align: 'center', valign: 'middle' });
  [['0:00', 'MSDS 라이브러리 조회'], ['0:15', '교육일지 작성 · 전자서명'], ['0:45', '관리자 알림 · 비대면 승인'], ['1:00', '대시보드 · 이수 현황']].forEach((t, i) => {
    const y = 196 + i * 92;
    HL(s, 888, y, 328);
    T(s, t[0], 888, y + 14, 120, 30, { fontSize: 22, color: i === 2 ? K.blue : K.ink });
    T(s, t[1], 888, y + 50, 328, 22, { fontSize: 14, bold: i === 2, color: i === 2 ? K.blue : K.ink2 });
  });
  R(s, 888, 568, 4, 69, K.blue);
  T(s, '결재 대기 19분 24초 → 1분', 908, 568, 308, 69, { fontSize: 15, bold: true, valign: 'middle' });
}
// 9 AI
{
  const s = slide();
  frame(s, '개선 활동  ·  AI 수시 위험성평가 · 근로자 참여형 · 2026.05–07 시범 운영', 1, 'AI는 초안을 만들고, 최종 판단은 사람이 합니다', 9);
  R(s, 64, 196, 640, 360, K.bg2);
  T(s, 'AI 위험성평가 결과 화면 캡처\n현장 사진 → 위험요인 → 위험도(5×4) → 개선대책', 64, 196, 640, 360, { fontSize: 14, color: K.cap, align: 'center', valign: 'middle', lineSpacingMultiple: 1.5 });
  facts(s, 760, 196, 456, [['AI 수시 위험성평가 · 운영 실적', '[실적]건', K.blue, '사진 한 장 → 위험요인 · 위험도 · 개선대책 초안'], ['AI MSDS 검토', '규제 항목 제안', K.ink, '관리자 승인 후 반영'], ['AI 도우미', '출처와 함께 답변', K.ink, '법령 · 지침 질의 응답']]);
  summary(s, '점수 · 등급은 서버가 다시 계산하고, 법령 조문 번호를 대조합니다');
}
// 10 결과 ① 준수율 · 종이
{
  const s = slide();
  frame(s, '결과 ①  ·  개선 전 2025.01–12 (n=106) / 개선 후 2026.05–07 (n=33)', 2, '개선 후 3개월, 신규 입사자 33명 전원이 기한 내 이수했습니다', 10);
  T(s, '법정교육 준수율 (개선 후)', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '100%', 58, 216, 660, 170, { fontSize: 120, color: K.blue });
  T(s, '33명 중 33명  ·  95% CI 89.4–100%', 64, 402, 640, 22, { fontSize: 14, color: K.ink2 });
  T(s, '개선 전', 64, 446, 100, 18, { fontSize: 11, color: K.sub }); R(s, 140, 448, 480 * 0.83, 14, K.grey); T(s, '83.0%', 140 + 480 * 0.83 + 8, 444, 100, 20, { fontSize: 12, bold: true, color: K.sub });
  T(s, '개선 후', 64, 476, 100, 18, { fontSize: 11, color: K.sub }); R(s, 140, 478, 480, 14, K.blue); T(s, '100%', 628, 474, 100, 20, { fontSize: 12, bold: true, color: K.blue });
  T(s, 'Fisher 정확검정 양측 p = 0.007  ·  미이수 18/106 vs 0/33', 64, 516, 640, 20, { fontSize: 13, bold: true, color: K.ink });
  facts(s, 760, 196, 456, [['준수율', '83.0% → 100%', K.ink, 'Fisher 정확검정 p = 0.007'], ['페이퍼리스 전환율', '0% → 100%', K.ink, '교육일지 출력 연 94건 → 0건'], ['업무가 가장 많던 달', '2026.06 · 24건', K.ink, '준수율 100% 유지']]);
  summary(s, '관찰 기간 차이(12개월 vs 3개월)는 한계로 함께 보고합니다');
}
// 11 결과 ② 행정시간 · 되찾은 시간
{
  const s = slide();
  frame(s, '결과 ②  ·  특별안전보건교육 1건 · 7단계 세부 행동 단위 실측', 2, '행정시간 93.0% 단축, 되찾은 시간은 환자와 현장으로', 11);
  T(s, '행정시간 단축', 64, 196, 560, 20, { fontSize: 14, color: K.sub });
  T(s, '−93%', 58, 220, 660, 190, { fontSize: 140, color: K.blue });
  T(s, '목표 70% 초과  ·  MSDS 교육 −81.0%  ·  6개 모듈 6/6 가동', 64, 420, 640, 22, { fontSize: 14, color: K.ink2 });
  T(s, '개선 전', 64, 470, 100, 18, { fontSize: 11, color: K.sub }); R(s, 140, 472, 480, 14, K.grey); T(s, '33분 48초', 628, 468, 120, 20, { fontSize: 12, bold: true, color: K.sub });
  T(s, '개선 후', 64, 500, 100, 18, { fontSize: 11, color: K.sub }); R(s, 140, 502, 480 * 142 / 2028, 14, K.blue); T(s, '2분 22초', 190, 498, 120, 20, { fontSize: 12, bold: true, color: K.blue });
  facts(s, 760, 196, 456, [['내부 결재', '19분 24초 → 1분', K.ink, '비대면 전자결재'], ['보건관리실 직접 제출', '7분 → 단계 삭제', K.ink, '모바일 제출'], ['되찾은 시간 · 교육 1건당', '31분 26초', K.blue, '부서는 환자 간호로 · 안전보건팀은 현장으로']]);
  summary(s, '부서는 환자 간호에, 안전보건팀은 현장의 직원 안전보건 활동에 시간을 씁니다');
}
// 12 고찰
{
  const s = slide();
  frame(s, '고찰  ·  한계 · 다른 기관 적용 시 고려사항', 3, '단일 기관 · 3개월 결과이며, 하반기까지 같은 지표로 추적합니다', 12);
  const col = (x, title, rows) => {
    T(s, title, x, 196, 540, 20, { fontSize: 14, color: K.sub });
    rows.forEach((r, i) => {
      const y = 228 + i * 108;
      HL(s, x, y, 544);
      T(s, r[0], x, y + 16, 544, 28, { fontSize: 19, bold: true });
      T(s, r[1], x, y + 52, 544, 40, { fontSize: 13, color: K.ink2, lineSpacingMultiple: 1.3 });
    });
  };
  col(64, '한계', [['단일 기관 결과', '부민병원 1곳의 신규 입사자 교육 자료'], ['관찰 기간 차이', '개선 전 12개월 vs 개선 후 3개월 · 하반기(8~12월) 추적 중'], ['행정시간 검정 미시행', '직접 관찰 측정값으로 제시 · 측정 건수 각주 표기']]);
  col(672, '다른 기관 적용 시 확인할 것', [['인사 데이터 연동 방식', '신규 입사자 명단을 자동으로 받을 경로'], ['개인정보 접근 권한', '부서 · 관리자별 열람 범위와 접근 기록'], ['전자서명 증빙 효력', '교육 기록 · 인증 조사 증빙으로 인정되는지']]);
  summary(s, '상용 노코드 도구로 만들어 별도 개발 인력 없이도 적용을 검토할 수 있습니다');
}
// 13 결론 및 제언
{
  const s = slide();
  frame(s, '결론 및 제언  ·  표준화 · 고도화 · 추적 계획', 3, '검증된 프로세스를 병원의 표준 업무로 정착시켰습니다', 13);
  [['결론', '결과', '준수율 100% · 행정시간 93.0% 단축', '페이퍼리스 100% · 6개 모듈 6/6 가동'], ['표준화', '완료', '법정 안전보건 관리 표준 업무지침 제정', '시범사업을 정식 업무로 전환'],
   ['고도화', '완료', '미이수자 부서장 자동 알림 · AI 2종 추가', 'AI MSDS 검토 · AI 도우미 (2026.08–10)'], ['제언', '계획', '하반기까지 같은 지표로 효과 추적', 'AI 위험성평가 근로자 참여 확대']].forEach((r, i) => {
    const y = 196 + i * 112, hi = i === 3;
    HL(s, 64, y, 656);
    T(s, r[0], 64, y + 16, 140, 22, { fontSize: 13, bold: true, color: K.sub });
    T(s, r[1], 620, y + 16, 100, 22, { fontSize: 13, bold: true, color: hi ? K.blue : K.ink, align: 'right' });
    T(s, r[2], 64, y + 44, 656, 30, { fontSize: 19, bold: true, color: hi ? K.blue : K.ink });
    T(s, r[3], 64, y + 78, 656, 20, { fontSize: 12, color: K.cap });
  });
  R(s, 760, 196, 456, 444, K.bg2); R(s, 760, 196, 456, 4, K.blue);
  T(s, '맺음말', 792, 228, 300, 20, { fontSize: 13, color: K.sub });
  T(s, '기다리지 않는 행정,\n빈틈없는 안전.', 792, 260, 400, 110, { fontSize: 30, color: K.ink, lineSpacingMultiple: 1.2 });
  T(s, '서류 대신 환자 곁으로,\n사무실 대신 현장으로.', 792, 390, 400, 60, { fontSize: 15, color: K.ink2, lineSpacingMultiple: 1.45 });
  HL(s, 792, 520, 392, K.grey);
  T(s, '감사합니다', 792, 540, 392, 48, { fontSize: 32, bold: true, color: K.blue });
  T(s, '인당의료재단 부민병원 총무-안전보건팀', 792, 596, 392, 20, { fontSize: 11, color: K.cap });
}
(async () => { await pres.writeFile({ fileName: OUT }); console.log('ok', sn, Object.keys(NOTES).length); })();
