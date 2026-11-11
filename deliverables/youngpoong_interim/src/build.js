// 중간보고 발표자료 생성 (착수회의 양식 재현)
const pptxgen = require('pptxgenjs');
const path = require('path');
const IMG = p => path.join(__dirname, 'img', p);
const OUT = process.argv[2] || 'out.pptx';

const pptx = new pptxgen();
pptx.defineLayout({ name: 'K', width: 11.93, height: 8.5 });
pptx.layout = 'K';
const W = 11.93, H = 8.5;
const F = '맑은 고딕';
const C = {
  navy: '0B2F5B', navy2: '17375E', blue: '2E75B6', blue2: '1F5AA6', sky: 'DCE6F2', sky2: 'EEF3F9',
  gray: '595959', lgray: 'D9D9D9', bg: 'F2F2F2', text: '262626', orange: 'E8590C', peach: 'FDEBDD',
  green: '2E7D32', ph: '8C8C8C',
};
const PROJ = 'AI 기반 제조 지식 자산화 및 전자회로기판 제조 공정 관리 자동화';
let slideCount = 0;
const addSlide = () => { slideCount++; return pptx.addSlide(); };

// ---------- 공통 요소 ----------
function footer(s) {
  s.addText(String(slideCount), { x: W / 2 - 0.4, y: 8.13, w: 0.8, h: 0.25, fontFace: F, fontSize: 9, color: '7F7F7F', align: 'center' });
  s.addImage({ path: IMG('yp.png'), x: 10.1, y: 8.1, w: 1.6, h: 0.22 });
}
function header(s, num, crumb, title) {
  s.background = { color: 'FFFFFF' };
  s.addImage({ path: IMG('hdr_bg.png'), x: 0, y: 0, w: W, h: 1.1 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.28, y: 0, w: 0.9, h: 1.05, fill: { color: C.navy }, line: { color: C.navy } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.28, y: 1.05, w: 0.9, h: 0.05, fill: { color: '3FA9F5' }, line: { color: '3FA9F5' } });
  s.addText(num, { x: 0.28, y: 0.25, w: 0.9, h: 0.65, fontFace: F, fontSize: 24, color: 'FFFFFF', align: 'center', valign: 'middle' });
  s.addText(crumb, { x: 1.38, y: 0.14, w: 8, h: 0.25, fontFace: F, fontSize: 9, color: C.blue });
  s.addText(title, { x: 1.38, y: 0.4, w: 9.5, h: 0.55, fontFace: F, fontSize: 22, bold: true, color: C.text });
  s.addShape(pptx.shapes.LINE, { x: 0.28, y: 1.1, w: W - 0.56, h: 0, line: { color: 'E3E8EF', width: 1 } });
  footer(s);
}
// 소제목(■ 표시)
function sub(s, text, x, y, w) {
  s.addShape(pptx.shapes.RECTANGLE, { x, y: y + 0.07, w: 0.1, h: 0.18, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText(text, { x: x + 0.17, y, w: w || 8, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: C.text, valign: 'middle' });
}
// 하단 출처/주석
function note(s, text, y) {
  s.addText(text, { x: 0.45, y: y || 7.72, w: W - 0.9, h: 0.3, fontFace: F, fontSize: 8.5, color: C.blue, italic: false });
}
// 진척 막대
function bar(s, x, y, w, h, pct, opts = {}) {
  s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: 'E7ECF3' }, line: { color: 'E7ECF3' } });
  if (pct > 0) {
    const c = opts.color || (pct >= 100 ? C.navy2 : C.blue);
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: Math.min(w * pct / 100, w), h, fill: { color: c }, line: { color: c } });
  }
}
// 표 헬퍼
const TH = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: 'FFFFFF', fill: { color: C.navy2 }, align: 'center', valign: 'middle' }, o) });
const TD = (t, o = {}) => ({ text: t, options: Object.assign({ color: C.text, valign: 'middle' }, o) });
const TK = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true, color: C.navy, fill: { color: C.sky }, align: 'center', valign: 'middle' }, o) });
const PH = (t = '[확인 필요]', o = {}) => ({ text: t, options: Object.assign({ color: C.ph, italic: true, align: 'center', valign: 'middle' }, o) });
function table(s, rows, o) {
  s.addTable(rows, Object.assign({ fontFace: F, fontSize: 10, border: { type: 'solid', pt: 0.75, color: 'BFBFBF' }, margin: [2, 5, 2, 5] }, o));
}

// ---------- 표지 / 목차 / 간지 ----------
function cover(label, dateText) {
  const s = addSlide();
  s.addImage({ path: IMG('cover.png'), x: 0, y: 0, w: W, h: H });
  s.addImage({ path: IMG('mss.png'), x: 0.5, y: 0.3, w: 1.55, h: 0.66 });
  s.addImage({ path: IMG('tipa.png'), x: 2.2, y: 0.4, w: 1.75, h: 0.49 });
  s.addText('과제번호 : SF26179244', { x: 0.55, y: 1.2, w: 5, h: 0.35, fontFace: F, fontSize: 12, color: '7F7F7F' });
  s.addText('제조분야 「AI 응용제품 신속 상용화 지원사업」', { x: 0.5, y: 1.8, w: 10, h: 0.6, fontFace: F, fontSize: 26, bold: true, color: '1F4E99' });
  s.addText('(과제명 : ' + PROJ + ')', { x: 0.5, y: 2.4, w: 10, h: 0.45, fontFace: F, fontSize: 15, color: C.text });
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.55, y: 3.12, w: 1.25, h: 0.36, rectRadius: 0.18, fill: { color: 'FFFFFF' }, line: { color: '7F7F7F', width: 1 } });
  s.addText(label, { x: 0.55, y: 3.12, w: 1.25, h: 0.36, fontFace: F, fontSize: 13, bold: true, color: 'D32F2F', align: 'center', valign: 'middle' });
  s.addText(dateText, { x: 1.95, y: 3.12, w: 3, h: 0.36, fontFace: F, fontSize: 12, color: '595959', valign: 'middle' });
  const ly = 7.85;
  s.addImage({ path: IMG('yp.png'), x: 0.35, y: ly, w: 1.35, h: 0.18 });
  s.addImage({ path: IMG('konan.png'), x: 2.05, y: ly - 0.06, w: 1.05, h: 0.26 });
  s.addImage({ path: IMG('abacus.png'), x: 3.4, y: ly - 0.03, w: 0.95, h: 0.2 });
  s.addImage({ path: IMG('cwnu.png'), x: 4.6, y: ly - 0.03, w: 1.2, h: 0.21 });
  s.addImage({ path: IMG('gntp.png'), x: 6.05, y: ly - 0.05, w: 1.5, h: 0.25 });
  return s;
}
const SECTIONS = [
  ['Ⅰ', '사업 개요', ['사업 개요', '추진 경과']],
  ['Ⅱ', '추진 현황', ['전체 기성 현황', '세부과제별 추진 현황', '단계별 추진 현황']],
  ['Ⅲ', '세부 추진 실적', ['분석 단계 실적', '설계 단계 실적(서비스·인프라)', '설계 단계 실적(지식 기반)', '개발·구축 진행 현황']],
  ['Ⅳ', '산출물 현황', ['중간보고 산출물 작성 현황', '프로젝트 관리 현황']],
  ['Ⅴ', '사업비 집행 현황', ['사업비 집행 총괄', '세부 항목별 진행 현황']],
  ['Ⅵ', '향후 추진 계획', ['추진 일정(계획 대비 실적)', '단계별 향후 추진 계획', '성과지표 관리 계획']],
  ['#', '이슈 및 협의사항', ['주요 이슈 및 대응 방안', '착수회의 후속조치 이행 점검']],
];
function toc() {
  const s = addSlide();
  s.addImage({ path: IMG('toc_bg.png'), x: 0, y: 0, w: W, h: H });
  s.addText('목차', { x: 6.3, y: 0.45, w: 1.6, h: 0.8, fontFace: F, fontSize: 40, bold: true, color: 'C9D9EE' });
  s.addText(PROJ, { x: 7.75, y: 0.62, w: 4.1, h: 0.5, fontFace: F, fontSize: 9.5, color: 'FFFFFF', valign: 'middle' });
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.45, y: 1.55, w: 5.1, h: 6.3, fill: { color: 'FFFFFF' }, line: { color: 'A9C4E6', width: 6 } });
  SECTIONS.forEach((sec, i) => {
    const y = 1.95 + i * 0.8;
    s.addShape(pptx.shapes.RECTANGLE, { x: 6.85, y, w: 0.5, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(sec[0], { x: 6.85, y, w: 0.5, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    s.addShape(pptx.shapes.RECTANGLE, { x: 7.4, y, w: 3.8, h: 0.5, fill: { color: 'F2F2F2' }, line: { color: 'F2F2F2' } });
    s.addText(sec[1], { x: 7.55, y, w: 3.6, h: 0.5, fontFace: F, fontSize: 15, bold: true, color: C.text, valign: 'middle' });
  });
}
function divider(idx) {
  const [num, title, items] = SECTIONS[idx];
  const s = addSlide();
  s.addImage({ path: IMG('div_bg.png'), x: 0, y: 0, w: W, h: H });
  s.addText(PROJ, { x: 5.0, y: 2.25, w: 6.5, h: 0.35, fontFace: F, fontSize: 12, color: '7F7F7F' });
  s.addShape(pptx.shapes.RECTANGLE, { x: 5.0, y: 2.7, w: 0.75, h: 0.75, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText(num, { x: 5.0, y: 2.7, w: 0.75, h: 0.75, fontFace: F, fontSize: 28, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  s.addText(title, { x: 5.95, y: 2.7, w: 5.6, h: 0.75, fontFace: F, fontSize: 30, bold: true, color: C.text, valign: 'middle' });
  s.addShape(pptx.shapes.LINE, { x: 5.0, y: 3.55, w: 5.4, h: 0, line: { color: '7F7F7F', width: 0.75 } });
  s.addShape(pptx.shapes.LINE, { x: 3.7, y: 4.0, w: 0, h: 0.9, line: { color: '595959', width: 1.5 } });
  items.forEach((t, i) => {
    s.addText(`${i + 1}.  ${t}`, { x: 5.9, y: 3.85 + i * 0.47, w: 5.5, h: 0.42, fontFace: F, fontSize: 16, color: C.text, valign: 'middle' });
  });
}

// =====================================================================
cover('중간보고', '2026. 11.');
toc();

// ---------------- Ⅰ. 사업 개요 ----------------
divider(0);
{
  const s = addSlide();
  header(s, '01', 'Ⅰ. 사업 개요  ▸  1. 사업 개요', '사업 개요');
  const k = (t) => TK(t, { fill: { color: C.navy2 }, color: 'FFFFFF', fontSize: 11 });
  const rows = [
    [k('과제명'), TD(PROJ, { align: 'center', bold: true, fontSize: 12, colspan: 2 })],
    [k('목 표'), TD([
      { text: '다품종·소량 생산의 복잡·다양한 ', options: { breakLine: false } }, { text: '회로기판(PBA) 제조 현장의 숙련자 암묵지', options: { color: C.orange, bold: true, breakLine: false } },
      { text: '를 ', options: { breakLine: false } }, { text: '온톨로지', options: { color: C.orange, bold: true, breakLine: false } }, { text: ' 기반의 구조화된 디지털 자산으로 전환하고, ', options: { breakLine: false } },
      { text: 'AI Agent 기반 업무 자동화', options: { color: C.orange, bold: true, breakLine: false } }, { text: ' 서비스 구축으로 제조공정 혁신 및 확산', options: { breakLine: false } },
    ], { align: 'center', colspan: 2 })],
    [k('주관/책임자'), TD('영풍전자 주식회사 / 이홍근 부장', { align: 'center', colspan: 2 })],
    [k('사업기간'), TD("2026. 07. 27 ~ 2027. 07. 26 (12개월)", { align: 'center', colspan: 2, fill: { color: C.peach } })],
    [k('사업비'), TD('총 1,882,900천원  (정부지원금 1,318,000천원 / 도입기업 현금 282,450천원 · 현물 282,450천원)', { align: 'center', colspan: 2, fill: { color: C.peach } })],
    [k('세부과제'), TD([
      { text: '① (기반 구축) 제조 데이터 표준화 및 온톨로지 지식 구축', options: { breakLine: true } },
      { text: '② (서비스 개발) LLM + RAG + Graph-RAG 기반 핵심 AI 서비스 구축', options: { breakLine: true } },
      { text: '③ (인프라) 인프라 구축 및 거버넌스(LLMOps, AgentOps) 환경 구축' },
    ], { colspan: 2, fontSize: 10.5 })],
    [k('도입기업'), TD('영풍전자 주식회사', { align: 'center', bold: true, color: C.navy }), TD('데이터 수집·정제 / 현장 전문가 암묵지 제공 · 기존 시스템 연동 인터페이스 지원 · 서비스 기획 및 검증', { fontSize: 9.5 })],
    [k('공급기업'), TD('㈜코난테크놀로지(대표)\n㈜애버커스\n국립창원대학교 산학협력단', { align: 'center', bold: true, color: C.navy }), TD([
      { text: '· AAS 기반 제조 데이터 표준화 체계 구축', options: { breakLine: true } },
      { text: '· 회로기판(PBA) 온톨로지 설계 및 지식그래프 구축, RAG 벡터DB 구축', options: { breakLine: true } },
      { text: '· AX 서비스 플랫폼 구축(LLM/RAG/Graph-RAG/Agent Builder/Admin)', options: { breakLine: true } },
      { text: '· AI 서비스 개발(고장탐구 어시스턴트, 전원 비인가/인가 검사 Agent, 유사 고장사례 매칭 Agent, Q&A 채팅)' },
    ], { fontSize: 9.5 })],
    [k('기획기관'), TD('(재)경남테크노파크', { align: 'center', bold: true, color: C.navy }), TD('과제기획, 진도관리, 위기관리, 성과창출·관리 / 결과물 확산계획 수립, 후속 확산체계 신규사업 기획', { fontSize: 9.5 })],
  ];
  table(s, rows, { x: 0.45, y: 1.4, w: 11.03, colW: [1.5, 2.6, 6.93], rowH: [0.42, 0.62, 0.42, 0.42, 0.42, 0.95, 0.55, 1.15, 0.55], fontSize: 11 });
  note(s, '※ 출처: 중간보고서(기성고 확인조서, 2026.10.26 제출) 및 착수회의 자료(2026.09.17)', 7.75);
}
{
  const s = addSlide();
  header(s, '02', 'Ⅰ. 사업 개요  ▸  2. 추진 경과', '추진 경과');
  sub(s, '주요 추진 경과', 0.45, 1.35);
  const ev = [
    ["'26.07.27", '사업 개시', '협약 사업기간 시작', false],
    ["'26.09.17", '사업 착수회의', '수행계획·추진체계·협의사항 공유', false],
    ["'26.10.23", '기성고 검사', '기성률 30% 확인', false],
    ["'26.10.26", '기성고 확인조서 제출', '인공지능혁신추진단 제출', false],
    ["'26.11", '중간보고', '추진 실적 및 향후 계획 보고', true],
    ["'27.07.26", '사업 종료', '시범운영·성과 검증 완료', false],
  ];
  const x0 = 0.9, x1 = 11.0, ty = 2.75;
  s.addShape(pptx.shapes.LINE, { x: x0, y: ty, w: x1 - x0, h: 0, line: { color: '9DB7D5', width: 3 } });
  ev.forEach((e, i) => {
    const cx = x0 + (x1 - x0) * i / (ev.length - 1);
    const col = e[3] ? C.orange : C.navy2;
    s.addShape(pptx.shapes.OVAL, { x: cx - 0.13, y: ty - 0.13, w: 0.26, h: 0.26, fill: { color: e[3] ? C.orange : 'FFFFFF' }, line: { color: col, width: 2.5 } });
    s.addText(e[0], { x: cx - 0.9, y: ty - 0.68, w: 1.8, h: 0.35, fontFace: F, fontSize: 12, bold: true, color: col, align: 'center' });
    s.addText(e[1], { x: cx - 0.95, y: ty + 0.25, w: 1.9, h: 0.32, fontFace: F, fontSize: 11.5, bold: true, color: C.text, align: 'center' });
    s.addText(e[2], { x: cx - 0.95, y: ty + 0.57, w: 1.9, h: 0.45, fontFace: F, fontSize: 9, color: C.gray, align: 'center', valign: 'top' });
  });
  s.addText('현재', { x: x0 + (x1 - x0) * 0.8 - 0.4, y: 1.72, w: 0.8, h: 0.28, fontFace: F, fontSize: 9, bold: true, color: 'FFFFFF', align: 'center', fill: { color: C.orange } });

  sub(s, '단계별 진행 요약 (기성고 검사 기준, 2026.10.23)', 0.45, 4.05);
  const st = [
    ['분석', '완료', '100%', '현행 업무·공정 분석, 요구사항 101건 도출(기능 63·비기능 38), 추적체계 수립, 도메인 분석(CQ 목록)', C.navy2],
    ['설계', '진행', '90%', '아키텍처(서버 9종), DB(188 테이블), 화면(41본), 인터페이스·보안·AI 통합, 온톨로지·KG·AAS 설계', C.blue],
    ['개발·구축', '착수', '20~50%', '온톨로지·KG 프로토타입 구현, HW·네트워크 도입·설치 및 K8s 플랫폼 구성 진행', '5B9BD5'],
    ['테스트', '예정', '0%', '단위·통합·성능·보안 테스트, AI 품질시험·사용자 인수시험(UAT) 예정', 'A6A6A6'],
  ];
  st.forEach((r, i) => {
    const x = 0.45 + i * 2.78, y = 4.5, w = 2.62;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.5, fill: { color: r[4] }, line: { color: r[4] } });
    s.addText(r[0], { x: x + 0.12, y, w: 1.4, h: 0.5, fontFace: F, fontSize: 13, bold: true, color: 'FFFFFF', valign: 'middle' });
    s.addText(r[1], { x: x + 1.5, y, w: 1.0, h: 0.5, fontFace: F, fontSize: 10, color: 'FFFFFF', align: 'right', valign: 'middle' });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: y + 0.5, w, h: 2.45, fill: { color: 'FFFFFF' }, line: { color: r[4], width: 1 } });
    s.addText(r[2], { x, y: y + 0.6, w, h: 0.6, fontFace: F, fontSize: 24, bold: true, color: r[4] === 'A6A6A6' ? '7F7F7F' : r[4], align: 'center' });
    s.addText(r[3], { x: x + 0.12, y: y + 1.25, w: w - 0.24, h: 1.6, fontFace: F, fontSize: 10, color: C.text, valign: 'top' });
  });
  note(s, '※ 개발·구축 기성률: 세부과제① 구현 20%, 세부과제③ 구축 50% (기성고 확인조서 세부과제별 기성률)', 7.75);
}

// ---------------- Ⅱ. 추진 현황 ----------------
divider(1);
{
  const s = addSlide();
  header(s, '01', 'Ⅱ. 추진 현황  ▸  1. 전체 기성 현황', '전체 기성 현황');
  // 도넛
  sub(s, '전체 기성률', 0.45, 1.35, 4);
  s.addChart(pptx.charts.DOUGHNUT, [{ name: '기성', labels: ['기성', '잔여'], values: [30, 70] }], {
    x: 0.6, y: 1.8, w: 3.0, h: 3.0, holeSize: 68, chartColors: [C.navy2, 'E7ECF3'], showLegend: false, showValue: false, showPercent: false, showLabel: false, dataBorder: { pt: 0.5, color: 'FFFFFF' },
  });
  s.addText([{ text: '30', options: { fontSize: 38, bold: true, color: C.navy2 } }, { text: '%', options: { fontSize: 18, bold: true, color: C.navy2 } }], { x: 0.6, y: 2.85, w: 3.0, h: 0.8, fontFace: F, align: 'center', valign: 'middle' });
  s.addText('기성고 검사일 2026.10.23', { x: 0.6, y: 4.85, w: 3.0, h: 0.3, fontFace: F, fontSize: 10, color: C.gray, align: 'center' });
  // 지표 타일
  const tiles = [
    ['요구사항 도출', '101건', '기능 63 · 비기능 38'],
    ['요구사항 변경관리', '148건', '변경 이력 관리(추적표)'],
    ['설계 산출물', '90%', '아키텍처설계서 100%'],
    ['사업비 집행률', '0%', '집행 0원 / 총 1,882,900천원'],
  ];
  tiles.forEach((t, i) => {
    const x = 4.15 + (i % 2) * 3.7, y = 1.8 + Math.floor(i / 2) * 1.62;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 3.5, h: 1.45, fill: { color: C.sky2 }, line: { color: 'C9D6E8', width: 1 } });
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 0.08, h: 1.45, fill: { color: i === 3 ? C.orange : C.blue }, line: { color: i === 3 ? C.orange : C.blue } });
    s.addText(t[0], { x: x + 0.25, y: y + 0.12, w: 3.1, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: C.gray });
    s.addText(t[1], { x: x + 0.25, y: y + 0.42, w: 3.1, h: 0.6, fontFace: F, fontSize: 26, bold: true, color: i === 3 ? C.orange : C.navy2 });
    s.addText(t[2], { x: x + 0.25, y: y + 1.02, w: 3.1, h: 0.3, fontFace: F, fontSize: 9.5, color: C.gray });
  });
  // 일정 대비 진척
  sub(s, '단계별 소요일수(계획 대비) 및 기성률', 0.45, 5.2, 8);
  const rows = [
    [TH('단계'), TH('소요 / 계획(일)'), TH('일정 경과율'), TH(''), TH('기성률'), TH('비고')],
    ...[
      ['분석', '88 / 62', 142, '100%', '계획 대비 26일 초과 소요 후 완료'],
      ['설계', '58 / 92', 63, '90%', '세부과제①·②·③ 설계 90%'],
      ['개발·구축', '27 / 153', 18, '20~50%', 'AI 서비스·Agent 구현은 M5부터 착수'],
      ['테스트', '0 / 90', 0, '0%', '테스트 계획서 작성 예정'],
      ['프로젝트 관리', '89 / 365', 24, '-', '계획서류 V1.0 수립, 변경·위험 관리'],
    ].map(r => [TK(r[0]), TD(r[1], { align: 'center' }), TD(r[2] + '%', { align: 'center', bold: true, color: r[2] > 100 ? C.orange : C.text }), TD(''), TD(r[3], { align: 'center', bold: true, color: C.navy2 }), TD(r[4], { fontSize: 9.5 })]),
  ];
  table(s, rows, { x: 0.45, y: 5.6, w: 11.03, colW: [1.5, 1.4, 1.1, 2.6, 1.1, 3.33], rowH: 0.33, fontSize: 10 });
  [142, 63, 18, 0, 24].forEach((p, i) => bar(s, 4.55, 5.6 + 0.33 * (i + 1) + 0.1, 2.4, 0.13, Math.min(p, 100), { color: p > 100 ? C.orange : C.blue }));
  note(s, '※ 일정 경과율 = 소요일수 ÷ 계획일수 (기성고 확인조서 붙임 "추진단계별 세부현황" 기준)', 7.65);
}
{
  const s = addSlide();
  header(s, '02', 'Ⅱ. 추진 현황  ▸  2. 세부과제별 추진 현황', '세부과제별 추진 현황');
  const cols = [
    ['세부과제 ①', '(기반 구축) 제조 데이터 표준화 및\n온톨로지 지식 구축', [
      ['분석', 100, ['요구사항정의서 v1.0', '요구사항추적표 v1.0', '도메인분석보고서 v1.0']],
      ['설계', 90, ['시스템구성도 v0.9', '인터페이스설계서 v0.5', '온톨로지설계서 v0.5', 'OWL2 파일 v0.5', '지식그래프스키마정의서 v0.5', '추론엔진설계서 v0.5', 'AAS서브모델명세서 v0.3', '용어사전구조정의서 v0.5']],
      ['구현', 20, ['프로토타입 구현보고서 v0.3', '소스코드 v0.3', '데모 UI(정적)']],
    ]],
    ['세부과제 ②', '(서비스 개발) LLM+RAG+Graph-RAG\n기반 핵심 AI 서비스 구축', [
      ['분석', 100, ['요구사항정의서 V1.0', '요구사항추적표 V1.0', 'AI 분석서 V1.0']],
      ['설계', 90, ['아키텍처설계서 V1.0', '데이터베이스설계서 V0.9', '화면설계서 V1.0', '인터페이스설계서 V1.0', '보안설계서 V1.0', 'AI 통합설계서 V0.9']],
    ]],
    ['세부과제 ③', '(인프라) 인프라 구축 및 거버넌스\n(LLMOps, AgentOps) 환경 구축', [
      ['분석', 100, ['요구사항정의서 V1.0', '요구사항추적표 V1.0', 'AI 분석서 V1.0']],
      ['설계', 90, ['아키텍처설계서 V1.0', '인터페이스설계서 V1.0', '보안설계서 V1.0']],
      ['구축', 50, ['HW·네트워크 도입·설치', 'K8s 플랫폼 구성']],
    ]],
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.72, w = 3.55;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 1.35, w, h: 0.95, fill: { color: C.navy2 }, line: { color: C.navy2 } });
    s.addText(c[0], { x: x + 0.1, y: 1.38, w: w - 0.2, h: 0.3, fontFace: F, fontSize: 10, bold: true, color: '9FC5F0' });
    s.addText(c[1], { x: x + 0.1, y: 1.66, w: w - 0.2, h: 0.6, fontFace: F, fontSize: 11, bold: true, color: 'FFFFFF', valign: 'top' });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 2.3, w, h: 5.3, fill: { color: 'FFFFFF' }, line: { color: 'BFCDE0', width: 1 } });
    let y = 2.45;
    c[2].forEach(st => {
      s.addText(st[0], { x: x + 0.12, y, w: 0.9, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: C.navy });
      bar(s, x + 0.85, y + 0.08, 1.75, 0.14, st[1]);
      s.addText(st[1] + '%', { x: x + 2.65, y, w: 0.8, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: st[1] >= 100 ? C.navy2 : C.blue, align: 'right' });
      y += 0.32;
      const txt = st[2].map((t, k) => ({ text: '· ' + t, options: { breakLine: k < st[2].length - 1 } }));
      const h = st[2].length * 0.215 + 0.05;
      s.addShape(pptx.shapes.RECTANGLE, { x: x + 0.12, y, w: w - 0.24, h, fill: { color: 'F5F7FA' }, line: { color: 'F5F7FA' } });
      s.addText(txt, { x: x + 0.18, y, w: w - 0.3, h, fontFace: F, fontSize: 9.5, color: C.text, valign: 'middle', paraSpaceAfter: 0 });
      y += h + 0.18;
    });
  });
  note(s, '※ 기성고 확인조서 "검사사항" 기준 세부과제·추진단계별 산출물 및 기성률 (2026.10.23 검사)', 7.72);
}
{
  const s = addSlide();
  header(s, '03', 'Ⅱ. 추진 현황  ▸  3. 단계별 추진 현황', '단계별 추진 현황');
  const R = (t, o) => TD(t, Object.assign({ align: 'center' }, o));
  const rows = [[TH('단계'), TH('주요 내용'), TH('소요/계획(일)'), TH('결과물'), TH('기성률')]];
  const add = (stage, span, days, items) => items.forEach((it, k) => {
    const row = [];
    if (k === 0) { row.push(TK(stage, { rowspan: span })); }
    row.push(TD(it[0]));
    if (k === 0) row.push(R(days, { rowspan: span, bold: true }));
    row.push(TD(it[1]));
    row.push(typeof it[2] === 'string' && it[2].startsWith('[') ? PH(it[2]) : R(it[2], { bold: true, color: it[2] === '100%' ? C.navy2 : C.text }));
    rows.push(row);
  });
  add('분석', 4, '88 / 62', [
    ['현행 업무·공정 분석(제조 10단계, 고장탐구 AS-IS/TO-BE)', 'AI 분석서', '100%'],
    ['요구사항 101건 도출(기능 63 · 비기능 38)', '요구사항 정의서', '100%'],
    ['요구사항–설계–시험 추적 체계 수립', '요구사항 추적표', '100%'],
    ['도메인 요구사항 분석(CQ 목록)', '도메인분석보고서', '100%'],
  ]);
  add('설계', 7, '58 / 92', [
    ['시스템 아키텍처(서버 9종 · SW · N/W) 설계', '아키텍처설계서', '100%'],
    ['DB 설계(테이블 188개)', '데이터베이스설계서', '90%'],
    ['화면 설계(화면 41본)', '화면설계서', '90%'],
    ['인터페이스 · API 계약 설계', '인터페이스설계서', '90%'],
    ['인증 · 권한 · 접근통제 · 로그 보안 설계', '보안설계서', '90%'],
    ['AI 서비스 · 모델 · 데이터 · 인프라 통합 설계', 'AI 통합설계서', '90%'],
    ['OWL2 온톨로지 · KG 스키마 · AAS 서브모델 설계', '온톨로지 설계서 외', '90%'],
  ]);
  add('개발·구축', 3, '27 / 153', [
    ['온톨로지 · KG 프로토타입 구현', '프로토타입 구현보고서 V0.3', '25%'],
    ['서버 · 네트워크 도입 · 설치, K8s 플랫폼 구성', '설치·구성 결과서(작성 중)', '-'],
    ['AI 서비스 · Agent 구현', '착수 전(M5~)', '-'],
  ]);
  add('테스트', 2, '0 / 90', [
    ['단위 · 통합 · 성능 · 보안 테스트', '테스트 계획서(예정)', '0%'],
    ['AI 품질시험 · 사용자 인수 테스트(UAT)', '시험 결과서(예정)', '0%'],
  ]);
  table(s, rows, { x: 0.45, y: 1.32, w: 11.03, colW: [1.25, 4.6, 1.35, 2.83, 1.0], rowH: 0.335, fontSize: 9.5 });
  note(s, '※ 출처: 기성고 확인조서 붙임 "추진단계별 세부현황" (프로젝트 관리 항목은 Ⅳ-2 참조)', 7.72);
}

// ---------------- Ⅲ. 세부 추진 실적 ----------------
divider(2);
{
  const s = addSlide();
  header(s, '01', 'Ⅲ. 세부 추진 실적  ▸  1. 분석 단계 실적', '분석 단계 실적');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y: 1.32, w: 11.03, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('분석 단계 완료(기성률 100%) — 현행 업무 분석부터 요구사항·추적 체계·도메인 분석까지 확정', { x: 0.45, y: 1.32, w: 11.03, h: 0.5, fontFace: F, fontSize: 13, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  const cards = [
    ['현행 업무·공정 분석', 'AI 분석서', ['제조 10단계 공정 분석', '고장탐구 업무 AS-IS / TO-BE 정의', 'AI 적용 대상 업무 도출']],
    ['요구사항 정의', '요구사항 정의서', ['요구사항 101건 도출', '기능 63건 · 비기능 38건', '세부과제별 요구사항 정의']],
    ['요구사항 추적 체계', '요구사항 추적표', ['요구사항–설계–시험 추적 체계 수립', '변경 이력 148건 관리', '산출물 연계 관리']],
    ['도메인 분석', '도메인분석보고서', ['PBA 고장탐구 도메인 요구 분석', '역량 질의(CQ) 목록 작성', '온톨로지 설계 입력 확보']],
  ];
  cards.forEach((c, i) => {
    const x = 0.45 + i * 2.78, y = 2.05, w = 2.62;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 2.7, fill: { color: 'FFFFFF' }, line: { color: 'BFCDE0', width: 1 } });
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.48, fill: { color: C.sky }, line: { color: C.sky } });
    s.addText(c[0], { x, y, w, h: 0.48, fontFace: F, fontSize: 12, bold: true, color: C.navy, align: 'center', valign: 'middle' });
    s.addText(c[2].map((t, k) => ({ text: t, options: { bullet: { code: '25AA' }, breakLine: k < c[2].length - 1 } })), { x: x + 0.1, y: y + 0.6, w: w - 0.2, h: 1.45, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top', paraSpaceAfter: 4 });
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 2.12, w: w - 0.4, h: 0.4, rectRadius: 0.08, fill: { color: C.peach }, line: { color: 'F4B183', width: 1 } });
    s.addText('▶ ' + c[1], { x: x + 0.2, y: y + 2.12, w: w - 0.4, h: 0.4, fontFace: F, fontSize: 10.5, bold: true, color: C.orange, align: 'center', valign: 'middle' });
  });
  // 요구사항 구성
  sub(s, '요구사항 구성 (총 101건)', 0.45, 5.0, 5);
  s.addChart(pptx.charts.DOUGHNUT, [{ name: '요구사항', labels: ['기능', '비기능'], values: [63, 38] }], {
    x: 0.5, y: 5.4, w: 2.2, h: 2.2, holeSize: 60, chartColors: [C.navy2, '8FB4E3'], showLegend: false, showValue: false, showPercent: false, dataBorder: { pt: 1, color: 'FFFFFF' },
  });
  s.addText('101건', { x: 0.5, y: 6.25, w: 2.2, h: 0.5, fontFace: F, fontSize: 16, bold: true, color: C.navy2, align: 'center' });
  [['기능 요구사항', '63건', C.navy2, 62], ['비기능 요구사항', '38건', '8FB4E3', 38]].forEach((r, i) => {
    const y = 5.75 + i * 0.7;
    s.addShape(pptx.shapes.RECTANGLE, { x: 2.9, y: y + 0.07, w: 0.2, h: 0.2, fill: { color: r[2] }, line: { color: r[2] } });
    s.addText(r[0], { x: 3.18, y, w: 1.6, h: 0.34, fontFace: F, fontSize: 11, color: C.text, valign: 'middle' });
    s.addText(r[1], { x: 4.7, y, w: 0.8, h: 0.34, fontFace: F, fontSize: 13, bold: true, color: C.navy2, valign: 'middle', align: 'right' });
  });
  sub(s, '분석 단계 시사점', 6.05, 5.0, 5);
  s.addText([
    { text: '분석 단계 소요 88일(계획 62일)로 26일 초과 — 요구사항 정합성 확보에 추가 기간 투입', options: { bullet: { code: '25AA' }, breakLine: true } },
    { text: '요구사항 변경 148건을 추적표로 이력 관리하여 설계·시험 단계 추적성 확보', options: { bullet: { code: '25AA' }, breakLine: true } },
    { text: 'CQ 목록을 온톨로지·지식그래프 설계의 검증 기준으로 활용', options: { bullet: { code: '25AA' } } },
  ], { x: 6.05, y: 5.45, w: 5.4, h: 2.0, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top', paraSpaceAfter: 6 });
}
{
  const s = addSlide();
  header(s, '02', 'Ⅲ. 세부 추진 실적  ▸  2. 설계 단계 실적(서비스·인프라)', '설계 단계 실적 (서비스·인프라)');
  const tiles = [
    ['아키텍처 설계', '서버 9종', '서버·SW·N/W 구성 설계', 100],
    ['DB 설계', '188개', '테이블 설계', 90],
    ['화면 설계', '41본', '사용자·관리자 화면', 90],
    ['인터페이스 설계', 'API 계약', '연계 시스템·데이터 정의', 90],
    ['보안 설계', '4대 영역', '인증·권한·접근통제·로그', 90],
    ['AI 통합 설계', '4대 영역', '서비스·모델·데이터·인프라', 90],
  ];
  tiles.forEach((t, i) => {
    const x = 0.45 + (i % 3) * 3.72, y = 1.35 + Math.floor(i / 3) * 1.52, w = 3.55;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 1.38, fill: { color: C.sky2 }, line: { color: 'C9D6E8', width: 1 } });
    s.addText(t[0], { x: x + 0.15, y: y + 0.1, w: 2.2, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: C.navy });
    s.addText(t[1], { x: x + 0.15, y: y + 0.42, w: 2.3, h: 0.55, fontFace: F, fontSize: 22, bold: true, color: C.navy2 });
    s.addText(t[2], { x: x + 0.15, y: y + 0.97, w: 2.6, h: 0.3, fontFace: F, fontSize: 9.5, color: C.gray });
    s.addShape(pptx.shapes.OVAL, { x: x + 2.6, y: y + 0.32, w: 0.78, h: 0.78, fill: { color: t[3] >= 100 ? C.navy2 : 'FFFFFF' }, line: { color: C.navy2, width: 1.5 } });
    s.addText(t[3] + '%', { x: x + 2.6, y: y + 0.32, w: 0.78, h: 0.78, fontFace: F, fontSize: 12, bold: true, color: t[3] >= 100 ? 'FFFFFF' : C.navy2, align: 'center', valign: 'middle' });
  });
  sub(s, '서버 구성(9종) — 아키텍처설계서 기준 신규 도입 대상', 0.45, 4.5, 9);
  const sv = ['AI 추론서버', 'RAG 서버', 'OCR 서버', '온톨로지 서버', 'LLMOps /\nAgentOps 서버', 'WEB/WAS 서버', 'K8s · Kafka ·\nRedis 서버', 'DB 서버 /\nMinIO 저장소', 'STT 서버'];
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y: 4.9, w: 11.03, h: 2.05, fill: { color: 'F7F9FC' }, line: { color: 'BFCDE0', width: 1, dashType: 'dash' } });
  sv.forEach((t, i) => {
    const x = 0.65 + i * 1.205;
    s.addShape(pptx.shapes.RECTANGLE, { x, y: 5.1, w: 1.08, h: 1.0, fill: { color: i === 0 ? C.navy2 : 'FFFFFF' }, line: { color: C.navy2, width: 1 } });
    s.addText(t, { x, y: 5.1, w: 1.08, h: 1.0, fontFace: F, fontSize: 9, bold: true, color: i === 0 ? 'FFFFFF' : C.navy, align: 'center', valign: 'middle' });
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.65, y: 6.3, w: 10.65, h: 0.45, fill: { color: C.sky }, line: { color: C.sky } });
  s.addText('L4 스위치 · 네트워크 스위치 · Rack/PDU/IP-KVM — 단일 네트워크 영역 구성  |  Kubernetes Platform 기반 운영', { x: 0.65, y: 6.3, w: 10.65, h: 0.45, fontFace: F, fontSize: 10, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  note(s, '※ 서버 명칭은 사업비 사용실적(H/W 구입비) 항목 기준. 주 추론서버 GPU 사양은 착수회의 자료(RTX6000 Pro × 2) 참조', 7.15);
  note(s, '※ 설계 90% 산출물(DB·화면·인터페이스·보안·AI 통합)은 V1.0 확정 후 구현 단계로 이관 예정', 7.45);
}
{
  const s = addSlide();
  header(s, '03', 'Ⅲ. 세부 추진 실적  ▸  3. 설계 단계 실적(지식 기반)', '설계 단계 실적 (지식 기반)');
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y: 1.32, w: 11.03, h: 0.5, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText('세부과제 ① 제조 데이터 표준화 및 온톨로지 지식 구축 — 설계 산출물 버전 현황 (목표 V1.0)', { x: 0.45, y: 1.32, w: 11.03, h: 0.5, fontFace: F, fontSize: 13, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  const items = [
    ['시스템구성도', 0.9, '지식 기반 계층 시스템 구성'],
    ['인터페이스설계서', 0.5, 'Graph-RAG · 서비스 연동 인터페이스'],
    ['온톨로지설계서', 0.5, 'PBA 고장탐구 도메인 클래스 · 속성 계층'],
    ['OWL2 파일', 0.5, 'PBA 도메인 온톨로지(OWL2)'],
    ['지식그래프스키마정의서', 0.5, '그래프 DB 스키마 · 인과관계 경로'],
    ['추론엔진설계서', 0.5, '논리 일관성 검사 · 추론 규칙'],
    ['AAS서브모델명세서', 0.3, 'IEC 63278-1 기반 PBA AAS 서브모델'],
    ['용어사전구조정의서', 0.5, '한/영/약어 3중 대조 용어사전 구조'],
  ];
  const rows = [[TH('산출물'), TH('주요 내용'), TH('현재 버전'), TH('v0.0 ─────── 성숙도 ─────── v1.0')]];
  items.forEach(it => rows.push([TK(it[0], { align: 'left' }), TD(it[2]), TD('v' + it[1].toFixed(1), { align: 'center', bold: true, color: C.navy2 }), TD('')]));
  table(s, rows, { x: 0.45, y: 2.0, w: 11.03, colW: [2.4, 4.0, 1.1, 3.53], rowH: 0.42, fontSize: 10 });
  items.forEach((it, i) => bar(s, 8.1, 2.0 + 0.42 * (i + 1) + 0.13, 3.2, 0.16, it[1] * 100, { color: it[1] >= 0.9 ? C.navy2 : C.blue }));
  sub(s, '연계 계획', 0.45, 5.95, 4);
  s.addText([
    { text: '구현 단계(프로토타입 v0.3)와 병행하여 온톨로지·KG 스키마를 반복 검증 후 V1.0 확정', options: { bullet: { code: '25AA' }, breakLine: true } },
    { text: '현장 전문가 인터뷰·사례 검증 결과를 온톨로지 클래스·관계 및 용어사전에 반영', options: { bullet: { code: '25AA' }, breakLine: true } },
    { text: 'AAS 서브모델(v0.3)은 표준 코드 매핑 테이블과 함께 우선 보완', options: { bullet: { code: '25AA' } } },
  ], { x: 0.55, y: 6.32, w: 10.9, h: 1.1, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top', paraSpaceAfter: 3 });
  note(s, '※ 버전은 기성고 확인조서 기준. 성숙도 막대는 버전 번호를 시각화한 것으로 별도 평가 수치가 아님', 7.72);
}
{
  const s = addSlide();
  header(s, '04', 'Ⅲ. 세부 추진 실적  ▸  4. 개발·구축 진행 현황', '개발·구축 진행 현황');
  const cols = [
    ['프로토타입 구현', '세부과제 ①', '20~25%', C.blue, [
      '온톨로지·지식그래프 프로토타입 구현', '프로토타입 구현보고서 v0.3', '소스코드 v0.3', '데모 UI(정적) 구성'],
      '구현보고서·소스코드 고도화, 1차 개발(Prototype) 완료'],
    ['인프라 구축', '세부과제 ③', '50%', C.navy2, [
      '서버·네트워크(HW·N/W) 도입·설치', 'Kubernetes 플랫폼 구성', '설치·구성 결과서 작성 중'],
      'K8s 기반 개발·운영 환경 구성 완료 및 결과서 확정'],
    ['AI 서비스·Agent', '세부과제 ②', '착수 전', 'A6A6A6', [
      '고장탐구 AI 어시스턴트', '전원 비인가/인가 검사 Agent · 유사 고장사례 매칭 Agent', '보고서 자동 생성 · Q&A 챗봇'],
      'M5(사업 5개월차)부터 구현 착수'],
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.72, w = 3.55, y = 1.35;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 4.6, fill: { color: 'FFFFFF' }, line: { color: c[3], width: 1.25 } });
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.95, fill: { color: c[3] }, line: { color: c[3] } });
    s.addText(c[1], { x: x + 0.15, y: y + 0.08, w: 2, h: 0.28, fontFace: F, fontSize: 9.5, color: 'FFFFFF' });
    s.addText(c[0], { x: x + 0.15, y: y + 0.38, w: 2.4, h: 0.45, fontFace: F, fontSize: 13, bold: true, color: 'FFFFFF', valign: 'middle' });
    s.addText(c[2], { x: x + 2.2, y: y + 0.2, w: 1.25, h: 0.6, fontFace: F, fontSize: c[2].length > 4 ? 15 : 20, bold: true, color: 'FFFFFF', align: 'right', valign: 'middle' });
    s.addText('추진 현황', { x: x + 0.15, y: y + 1.1, w: 2, h: 0.3, fontFace: F, fontSize: 10.5, bold: true, color: C.navy });
    s.addText(c[4].map((t, k) => ({ text: t, options: { bullet: { code: '25AA' }, breakLine: k < c[4].length - 1 } })), { x: x + 0.15, y: y + 1.42, w: w - 0.3, h: 1.9, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top', paraSpaceAfter: 4 });
    s.addShape(pptx.shapes.RECTANGLE, { x: x + 0.15, y: y + 3.5, w: w - 0.3, h: 0.95, fill: { color: C.sky2 }, line: { color: C.sky2 } });
    s.addText([{ text: '다음 단계', options: { bold: true, color: C.navy, breakLine: true } }, { text: c[5] }], { x: x + 0.25, y: y + 3.5, w: w - 0.5, h: 0.95, fontFace: F, fontSize: 10, color: C.text, valign: 'middle' });
  });
  sub(s, '테스트 단계 준비', 0.45, 6.15, 5);
  const rows = [
    [TH('구분'), TH('범위'), TH('산출물'), TH('상태')],
    [TK('기능·비기능 시험'), TD('단위 · 통합 · 성능 · 보안 테스트'), TD('테스트 계획서', { align: 'center' }), TD('작성 예정', { align: 'center' })],
    [TK('품질·인수 시험'), TD('AI 품질시험 · 사용자 인수 테스트(UAT)'), TD('시험 결과서', { align: 'center' }), TD('작성 예정', { align: 'center' })],
  ];
  table(s, rows, { x: 0.45, y: 6.52, w: 11.03, colW: [2.0, 5.0, 2.2, 1.83], rowH: 0.32, fontSize: 10 });
  note(s, '※ 프로토타입 기성률: 검사사항 표 20%, 세부현황 표 25%로 기재(원문 표기 유지)', 7.72);
}

// ---------------- Ⅳ. 산출물 현황 ----------------
divider(3);
{
  const s = addSlide();
  header(s, '01', 'Ⅳ. 산출물 현황  ▸  1. 중간보고 산출물 작성 현황', '중간보고 산출물 작성 현황');
  const ST = {
    done: TD('작성 완료', { align: 'center', bold: true, color: C.navy2 }),
    run: TD('운영 중', { align: 'center', bold: true, color: C.blue }),
    prog: TD('보완 중', { align: 'center', bold: true, color: C.orange }),
    alt: TD('대체 관리', { align: 'center', color: C.gray }),
    chk: PH('확인 필요'),
  };
  const data = [
    ['프로젝트\n관리', 9, [
      ['AI-PM1', '사업수행계획서', '필수', '사업수행계획서 V1.0', 'done'],
      ['AI-PM2', 'WBS', '필수', 'WBS V1.0', 'done'],
      ['AI-PM3', '투입인력 계획서', '필수', '인력투입계획서 V1.0', 'done'],
      ['AI-PM4', '품질관리계획서', '필수', '품질관리계획서 V1.0', 'done'],
      ['AI-PM5', '형상관리계획서', '필수', '형상관리계획서 V1.0', 'done'],
      ['AI-PM6', '위험관리대장', '필수', '위험관리내역서 V1.0', 'done'],
      ['AI-PM7', '주간/월간 보고서', '필수', '월간 보고서', 'run'],
      ['AI-PM8', '회의록', '필수', '회의록', 'run'],
      ['AI-PM9', '변경관리내역서', '선택', '요구사항추적표 변경이력(148건)', 'alt'],
    ]],
    ['시스템\n분석', 4, [
      ['AI-AL1', '요구사항 정의서', '필수', '요구사항정의서 V1.0', 'done'],
      ['AI-AL2', '업무프로세스 분석서', '필수', '(AI 분석서 V1.0 내 현행 업무·공정 분석)', 'chk'],
      ['AI-AL3', '요구사항 추적표', '필수', '요구사항추적표 V1.0', 'done'],
      ['AI-AL4', 'AI 분석서', '필수', 'AI 분석서 V1.0', 'done'],
    ]],
    ['시스템\n설계', 6, [
      ['AI-SD1', '시스템 아키텍처 설계서', '필수', '아키텍처설계서 V1.0 / 시스템구성도 v0.9', 'done'],
      ['AI-SD2', 'DB 설계서(ERD 포함)', '필수', '데이터베이스설계서 V0.9', 'prog'],
      ['AI-SD3', '화면 설계서(UI/UX)', '필수', '화면설계서 V1.0 (세부현황 90%)', 'prog'],
      ['AI-SD4', '인터페이스 설계서', '필수', '인터페이스설계서 V1.0 / v0.5(세부①)', 'prog'],
      ['AI-SD5', '보안 설계서', '필수', '보안설계서 V1.0 (세부현황 90%)', 'prog'],
      ['AI-SD6', 'AI 통합설계서', '필수', 'AI 통합설계서 V0.9', 'prog'],
    ]],
  ];
  const rows = [[TH('단계'), TH('양식번호'), TH('주요 산출물'), TH('구분'), TH('작성 산출물(버전)'), TH('상태')]];
  data.forEach(g => g[2].forEach((r, k) => {
    const row = [];
    if (k === 0) row.push(TK(g[0], { rowspan: g[1] }));
    row.push(TD(r[0], { align: 'center', fontSize: 9 }), TD(r[1]), TD(r[2], { align: 'center', color: r[2] === '선택' ? C.gray : C.text }), TD(r[3]), ST[r[4]]);
    rows.push(row);
  }));
  table(s, rows, { x: 0.45, y: 1.3, w: 11.03, colW: [1.1, 1.0, 2.6, 0.75, 4.33, 1.25], rowH: 0.3, fontSize: 9.5 });
  note(s, '※ 기준: 「AI 응용제품 신속 상용화 지원사업 중간보고 산출물 목록」 19종(필수 18 · 선택 1) 대비 기성고 확인조서 기재 산출물 매핑', 7.62);
  note(s, '※ "보완 중" = 기성률 90% 또는 V1.0 미만 / "확인 필요" = 별도 문서 작성 여부 확인 대상', 7.85);
}
{
  const s = addSlide();
  header(s, '02', 'Ⅳ. 산출물 현황  ▸  2. 프로젝트 관리 현황', '프로젝트 관리 현황');
  const cards = [
    ['계획 수립', ['사업수행계획서 V1.0', 'WBS V1.0', '인력투입계획서 V1.0', '품질관리계획서 V1.0', '형상관리계획서 V1.0']],
    ['변경 관리', ['요구사항 변경 148건 이력 관리', '요구사항추적표 변경이력으로 관리', '요구사항–설계–시험 추적성 유지']],
    ['의사소통 관리', ['월간 보고서 작성', '회의록 작성·공유', '착수회의(2026.09.17) 결과 후속조치']],
    ['위험 관리', ['위험관리내역서 V1.0', '일정·데이터·검증·계약 위험 식별', '대응 방안·담당 지정 관리']],
  ];
  cards.forEach((c, i) => {
    const x = 0.45 + i * 2.78, y = 1.35, w = 2.62;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 2.75, fill: { color: 'FFFFFF' }, line: { color: 'BFCDE0', width: 1 } });
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 0.48, fill: { color: C.navy2 }, line: { color: C.navy2 } });
    s.addText(c[0], { x, y, w, h: 0.48, fontFace: F, fontSize: 12, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    s.addText(c[1].map((t, k) => ({ text: t, options: { bullet: { code: '25AA' }, breakLine: k < c[1].length - 1 } })), { x: x + 0.1, y: y + 0.6, w: w - 0.2, h: 2.05, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top', paraSpaceAfter: 4 });
  });
  sub(s, '프로젝트 관리 소요 현황', 0.45, 4.35, 6);
  s.addText('89일 / 365일', { x: 0.6, y: 4.75, w: 2.4, h: 0.45, fontFace: F, fontSize: 18, bold: true, color: C.navy2 });
  bar(s, 3.0, 4.92, 8.2, 0.18, 24);
  s.addText('24%', { x: 3.0, y: 5.12, w: 8.2, h: 0.25, fontFace: F, fontSize: 9, color: C.gray });
  sub(s, '착수회의 이후 관리 체계 운영', 0.45, 5.55, 6);
  const rows = [
    [TH('관리 항목'), TH('운영 방식'), TH('근거 산출물')],
    [TK('일정·진척'), TD('WBS 기준 단계별 소요/계획 일수 및 기성률 관리'), TD('WBS V1.0, 월간 보고서')],
    [TK('품질·형상'), TD('산출물 검토·버전 관리, 변경 통제'), TD('품질관리계획서 · 형상관리계획서 V1.0')],
    [TK('요구사항 변경'), TD('변경 요청 → 영향 분석 → 추적표 반영'), TD('요구사항추적표 변경이력(148건)')],
  ];
  table(s, rows, { x: 0.45, y: 5.92, w: 11.03, colW: [1.8, 5.4, 3.83], rowH: 0.36, fontSize: 10 });
}

// ---------------- Ⅴ. 사업비 ----------------
divider(4);
{
  const s = addSlide();
  header(s, '01', 'Ⅴ. 사업비 집행 현황  ▸  1. 사업비 집행 총괄', '사업비 집행 총괄');
  const tiles = [['총 사업비', '1,882,900', '천원'], ['집행 사업비', '0', '천원'], ['잔여 사업비', '1,882,900', '천원'], ['집행률', '0', '%']];
  tiles.forEach((t, i) => {
    const x = 0.45 + i * 2.78, w = 2.62, y = 1.35;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h: 1.15, fill: { color: i === 3 ? C.peach : C.sky2 }, line: { color: i === 3 ? 'F4B183' : 'C9D6E8', width: 1 } });
    s.addText(t[0], { x, y: y + 0.08, w, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: C.gray, align: 'center' });
    s.addText([{ text: t[1], options: { fontSize: 22, bold: true } }, { text: ' ' + t[2], options: { fontSize: 11 } }], { x, y: y + 0.42, w, h: 0.6, fontFace: F, color: i === 3 ? C.orange : C.navy2, align: 'center', valign: 'middle' });
  });
  sub(s, '비목별 예산 및 집행 현황', 0.45, 2.75, 6);
  s.addText('(단위: 천원)', { x: 9.5, y: 2.78, w: 2, h: 0.3, fontFace: F, fontSize: 9, color: C.gray, align: 'right' });
  const N = v => v.toLocaleString('en-US');
  const data = [
    ['S/W 개발비(A)', 0, 0, 761320], ['H/W 구입비(B)', 282450, 0, 109330], ['S/W 구입비(C)', 0, 0, 300000], ['N/W 구축비(D)', 0, 0, 80000],
    ['기타(할인)비용(E)', 0, 0, 2350], ['도입기업 인력 인건비(F)', 0, 282450, 0], ['기획기관 운영비(I)', 0, 0, 65000],
  ];
  const rows = [
    [TH('구분', { rowspan: 2 }), TH('도입기업 부담금', { colspan: 2 }), TH('정부지원금', { rowspan: 2 }), TH('합계', { rowspan: 2 }), TH('집행액', { rowspan: 2 }), TH('집행률', { rowspan: 2 })],
    [TH('현금'), TH('현물')],
  ];
  const R = (v, o) => TD(N(v), Object.assign({ align: 'right' }, o));
  data.forEach(d => rows.push([TK(d[0], { align: 'left' }), R(d[1]), R(d[2]), R(d[3]), R(d[1] + d[2] + d[3], { bold: true }), R(0), TD('0%', { align: 'center' })]));
  rows.push([TK('총 사업비', { fill: { color: 'BDD7EE' } }), R(282450, { bold: true, fill: { color: 'EAF1F8' } }), R(282450, { bold: true, fill: { color: 'EAF1F8' } }), R(1318000, { bold: true, fill: { color: 'EAF1F8' } }), R(1882900, { bold: true, fill: { color: 'EAF1F8' } }), R(0, { bold: true, fill: { color: 'EAF1F8' } }), TD('0%', { align: 'center', bold: true, fill: { color: 'EAF1F8' } })]);
  table(s, rows, { x: 0.45, y: 3.12, w: 11.03, colW: [2.6, 1.4, 1.4, 1.5, 1.6, 1.3, 1.23], rowH: 0.33, fontSize: 10 });
  note(s, '※ 출처: 기성고 확인조서(2026.10.26) "사업비 사용실적" 및 착수회의 자료 "총괄 예산". 클라우드 서비스 이용료(G)·연동 기구축장비(H)는 0원으로 생략', 7.05);
  note(s, '※ 비목별 세부 항목은 전부 발주 또는 운용 단계이며, 기성고 검사 시점 집행 실적 없음(Ⅴ-2 참조)', 7.32);
}
{
  const s = addSlide();
  header(s, '02', 'Ⅴ. 사업비 집행 현황  ▸  2. 세부 항목별 진행 현황', '세부 항목별 진행 현황');
  const mk = (rowsData) => {
    const rows = [[TH('구분'), TH('내역'), TH('현황'), TH('계획(천원)')]];
    rowsData.forEach(g => g[1].forEach((r, k) => {
      const row = [];
      if (k === 0) row.push(TK(g[0], { rowspan: g[1].length }));
      row.push(TD(r[0]), TD(r[1], { align: 'center', color: r[1] === '발주' ? C.blue : C.gray, bold: true }));
      if (r[2] !== null) row.push(TD(r[2], Object.assign({ align: 'right' }, r[3] ? { rowspan: r[3] } : {})));
      rows.push(row);
    }));
    return rows;
  };
  const left = [
    ['S/W\n개발비', [['LLM/RAG 지식관리', '발주', '228,970'], ['제조AI통합플랫폼', '발주', '442,350'], ['온톨로지 설계 및 지식베이스 시스템', '발주', '90,000']]],
    ['S/W\n구입비', [['RAG 엔진SW(Konan RAG)', '발주', '100,000'], ['DocuAnalyzer(문서OCR)', '발주', '40,000'], ['LLM Model SW(Konan LLM)/LLM Studio', '발주', '120,000'], ['STT 엔진SW', '발주', '40,000']]],
    ['N/W', [['서버실 전원설비(UPS) 공사', '운용', '80,000', 3], ['서버실 냉방설비(항온항습기) 공사', '운용', null], ['네트워크 케이블 및 기타 인프라 공사', '운용', null]]],
    ['기타', [['회계정산수수료', '-', '1,900'], ['기술임치비', '-', '450']]],
    ['인건비', [['도입기업 인건비(현물)', '-', '282,450']]],
    ['운영비', [['기획기관 운영비(현금)', '-', '65,000']]],
  ];
  const right = [
    ['H/W', [['AI 추론서버', '운용', '106,280'], ['RAG 서버', '운용', '67,910'], ['OCR 서버', '운용', '67,910'], ['온톨로지 서버', '운용', '20,000'], ['LLMOps/AgentOps 서버', '운용', '20,000'], ['WEB/WAS 서버', '운용', '20,000'], ['K8s, Kafka, Redis 서버', '운용', '20,000'], ['DB서버/MinIO 객체 저장소', '운용', '20,000'], ['STT 서버', '운용', '20,000'], ['L4·네트워크 스위치, Rack/PDU/IP-KVM', '운용', '14,080'], ['검사용 노트북 5대·음성인식용 마이크 10대', '발주', '15,600']]],
  ];
  table(s, mk(left), { x: 0.45, y: 1.32, w: 5.9, colW: [0.85, 3.05, 0.7, 1.3], rowH: 0.36, fontSize: 9 });
  table(s, mk(right), { x: 6.5, y: 1.32, w: 4.98, colW: [0.7, 2.68, 0.7, 0.9], rowH: 0.36, fontSize: 9 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.5, y: 5.85, w: 4.98, h: 1.2, fill: { color: C.sky2 }, line: { color: 'C9D6E8' } });
  s.addText([
    { text: '집행 현황 요약', options: { bold: true, color: C.navy, breakLine: true } },
    { text: '· 전 항목 집행액 0원(기성고 검사 시점)', options: { breakLine: true } },
    { text: '· S/W 개발·구입비 및 노트북·마이크: 발주 단계', options: { breakLine: true } },
    { text: '· 서버·네트워크·서버실 공사: 운용 단계' },
  ], { x: 6.65, y: 5.85, w: 4.7, h: 1.2, fontFace: F, fontSize: 9.5, color: C.text, valign: 'middle' });
  note(s, '※ 현황 구분(발주/운용)은 기성고 확인조서 원문 표기. 서버실 공사 3개 항목 계획액은 합산 80,000천원', 7.72);
}

// ---------------- Ⅵ. 향후 추진 계획 ----------------
divider(5);
{
  const s = addSlide();
  header(s, '01', 'Ⅵ. 향후 추진 계획  ▸  1. 추진 일정(계획 대비 실적)', '추진 일정 (계획 대비 실적)');
  const months = ['8', '9', '10', '11', '12', '1', '2', '3', '4', '5', '6', '7'];
  const x0 = 2.2, cw = 0.77, y0 = 1.32;
  // 날짜→x (2026-07-27 시작 기준, 월 단위 근사)
  const mx = (m, d) => { // m: 2026 기준 월(8..19), d: 일
    const idx = m - 8 + (d - 1) / 30; return x0 + idx * cw;
  };
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y: y0, w: 1.75, h: 0.7, fill: { color: C.navy2 }, line: { color: 'FFFFFF' } });
  s.addText('구분', { x: 0.45, y: y0, w: 1.75, h: 0.7, fontFace: F, fontSize: 11, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  s.addShape(pptx.shapes.RECTANGLE, { x: x0, y: y0, w: cw * 5, h: 0.35, fill: { color: C.navy2 }, line: { color: 'FFFFFF' } });
  s.addText('2026년', { x: x0, y: y0, w: cw * 5, h: 0.35, fontFace: F, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  s.addShape(pptx.shapes.RECTANGLE, { x: x0 + cw * 5, y: y0, w: cw * 7, h: 0.35, fill: { color: C.navy2 }, line: { color: 'FFFFFF' } });
  s.addText('2027년', { x: x0 + cw * 5, y: y0, w: cw * 7, h: 0.35, fontFace: F, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  months.forEach((m, i) => {
    s.addShape(pptx.shapes.RECTANGLE, { x: x0 + i * cw, y: y0 + 0.35, w: cw, h: 0.35, fill: { color: '2F5597' }, line: { color: 'FFFFFF' } });
    s.addText(m + '월', { x: x0 + i * cw, y: y0 + 0.35, w: cw, h: 0.35, fontFace: F, fontSize: 9.5, color: 'FFFFFF', align: 'center', valign: 'middle' });
  });
  const rowsG = [
    ['요구분석', [[8, 1, 9, 26], [8, 1, 10, 23]], '완료(10.23)'],
    ['설계', [[8, 26, 11, 26], [8, 26, 11, 11]], '90%'],
    ['개발·구축', [[9, 26, 14, 26], [9, 26, 11, 11]], '20~50%'],
    ['HW·NW 구축', [[8, 1, 9, 30], [8, 1, 11, 11]], '50%'],
    ['테스트', [[16, 27, 19, 26], null], '예정'],
    ['프로젝트 관리', [[7, 27, 19, 26], [7, 27, 11, 11]], '진행'],
  ];
  const rh = 0.62, gy = y0 + 0.7;
  rowsG.forEach((r, i) => {
    const y = gy + i * rh;
    s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y, w: 1.75, h: rh, fill: { color: C.sky }, line: { color: 'FFFFFF' } });
    s.addText(r[0], { x: 0.45, y, w: 1.75, h: rh, fontFace: F, fontSize: 10.5, bold: true, color: C.navy, align: 'center', valign: 'middle' });
    s.addShape(pptx.shapes.RECTANGLE, { x: x0, y, w: cw * 12, h: rh, fill: { color: i % 2 ? 'FFFFFF' : 'F7F9FC' }, line: { color: 'E1E6EE' } });
    const [p, a] = r[1];
    const px1 = Math.max(mx(p[0], p[1]), x0), px2 = Math.min(mx(p[2], p[3]), x0 + cw * 12);
    s.addShape(pptx.shapes.RECTANGLE, { x: px1, y: y + 0.1, w: px2 - px1, h: 0.17, fill: { color: 'BDD7EE' }, line: { color: '9DC3E6' } });
    if (a) {
      const ax1 = Math.max(mx(a[0], a[1]), x0), ax2 = mx(a[2], a[3]);
      const col = (r[0] === '요구분석') ? C.orange : C.navy2;
      s.addShape(pptx.shapes.RECTANGLE, { x: ax1, y: y + 0.33, w: ax2 - ax1, h: 0.17, fill: { color: col }, line: { color: col } });
      s.addText(r[2], { x: ax2 + 0.05, y: y + 0.27, w: 1.2, h: 0.28, fontFace: F, fontSize: 8.5, bold: true, color: col });
    } else {
      s.addText(r[2], { x: px1 + 0.05, y: y + 0.27, w: 1.2, h: 0.28, fontFace: F, fontSize: 8.5, color: C.gray });
    }
  });
  // 현재선
  const nx = mx(11, 11);
  s.addShape(pptx.shapes.LINE, { x: nx, y: gy - 0.05, w: 0, h: rh * rowsG.length + 0.1, line: { color: 'D32F2F', width: 1.5, dashType: 'dash' } });
  s.addText('작성 기준일(11.11)', { x: nx - 0.8, y: gy + rh * rowsG.length + 0.05, w: 1.6, h: 0.25, fontFace: F, fontSize: 8.5, bold: true, color: 'D32F2F', align: 'center' });
  // 범례
  const ly = gy + rh * rowsG.length + 0.4;
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.5, y: ly + 0.07, w: 0.4, h: 0.15, fill: { color: 'BDD7EE' }, line: { color: '9DC3E6' } });
  s.addText('계획', { x: 0.95, y: ly, w: 0.6, h: 0.3, fontFace: F, fontSize: 9, color: C.text });
  s.addShape(pptx.shapes.RECTANGLE, { x: 1.6, y: ly + 0.07, w: 0.4, h: 0.15, fill: { color: C.navy2 }, line: { color: C.navy2 } });
  s.addText('실적(진행)', { x: 2.05, y: ly, w: 1.0, h: 0.3, fontFace: F, fontSize: 9, color: C.text });
  s.addShape(pptx.shapes.RECTANGLE, { x: 3.1, y: ly + 0.07, w: 0.4, h: 0.15, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText('계획 초과 소요', { x: 3.55, y: ly, w: 1.3, h: 0.3, fontFace: F, fontSize: 9, color: C.text });
  // 마일스톤 표
  const rows = [
    [TH('주요 마일스톤'), TH('요구분석 완료'), TH('설계 완료'), TH('Prototype 완료'), TH('개발 완료'), TH('단위시험'), TH('연동시험'), TH('성능시험')],
    [TK('착수 계획'), TD("'26.9월", { align: 'center' }), TD("'26.10월", { align: 'center' }), TD("'26.12월", { align: 'center' }), TD("'27.3월", { align: 'center' }), TD("'27.4월", { align: 'center' }), TD("'27.5월", { align: 'center' }), TD("'27.6~7월", { align: 'center' })],
    [TK('현행/전망'), TD("'26.10.23 완료", { align: 'center', bold: true, color: C.orange }), TD("'26.11월(목표)", { align: 'center' }), PH('[확인 필요]'), PH('[확인 필요]'), PH('[확인 필요]'), PH('[확인 필요]'), PH('[확인 필요]')],
  ];
  table(s, rows, { x: 0.45, y: 6.05, w: 11.03, colW: [1.55, 1.4, 1.35, 1.4, 1.35, 1.3, 1.3, 1.38], rowH: 0.36, fontSize: 9.5 });
  note(s, '※ 계획 막대: 착수회의 WBS 및 기성고 확인조서 소요/계획 일수로 환산(추정). 실적 막대는 기성고 검사(10.23) 이후 현재까지 진행 중으로 표시', 7.2);
}
{
  const s = addSlide();
  header(s, '02', 'Ⅵ. 향후 추진 계획  ▸  2. 단계별 향후 추진 계획', '단계별 향후 추진 계획');
  const cols = [
    ["'26.11 ~ 12월", '설계 확정 · 1차 개발', [
      '설계 산출물 V1.0 확정(DB·화면·인터페이스·보안·AI 통합)', '온톨로지·OWL2·KG 스키마·AAS 서브모델 v1.0 확정', 'HW·N/W 설치 및 K8s 플랫폼 구성 완료, 설치·구성 결과서', 'Prototype(1차 개발) 완료', 'AI 서비스·Agent 구현 착수(M5~)', '제조 암묵지 수집·지식베이스 구축 착수', '발주 건 계약·집행 착수']],
    ["'27.1 ~ 3월", '2차 개발 · 연동', [
      'Agent / 지식베이스 구축(온톨로지 포함)', '사용자 화면 / 관리자 화면 개발', '영풍전자 기존 시스템 연동(ERP·APS·SCM 등)', 'LLM/RAG/Graph-RAG/STT 기능 구현', '1차 기능 업그레이드', '테스트 계획서 작성']],
    ["'27.4 ~ 7월", '시험 · 실증 · 성과 검증', [
      '단위시험 · 연동시험', '통합 · 성능 · 보안 시험', 'AI 품질시험 · 사용자 인수시험(UAT)', '시범운영 및 서비스 릴리스', 'KPI 측정 및 성과 검증', '시험 결과서 · 최종 산출물 확정']],
  ];
  cols.forEach((c, i) => {
    const x = 0.45 + i * 3.72, w = 3.55, y = 1.35;
    s.addShape(pptx.shapes.CHEVRON, { x, y, w, h: 0.75, fill: { color: [C.navy2, C.blue, '5B9BD5'][i] }, line: { color: 'FFFFFF' } });
    s.addText([{ text: c[0], options: { fontSize: 10, breakLine: true } }, { text: c[1], options: { fontSize: 13, bold: true } }], { x: x + 0.3, y, w: w - 0.5, h: 0.75, fontFace: F, color: 'FFFFFF', align: 'center', valign: 'middle' });
    s.addShape(pptx.shapes.RECTANGLE, { x, y: y + 0.9, w, h: 4.4, fill: { color: 'FFFFFF' }, line: { color: 'BFCDE0', width: 1 } });
    s.addText(c[2].map((t, k) => ({ text: t, options: { bullet: { code: '25AA' }, breakLine: k < c[2].length - 1 } })), { x: x + 0.12, y: y + 1.0, w: w - 0.24, h: 4.2, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top', paraSpaceAfter: 7 });
  });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y: 6.85, w: 11.03, h: 0.62, fill: { color: C.peach }, line: { color: 'F4B183' } });
  s.addText('중점 관리: 분석 단계 지연(26일)을 설계·구현 병행으로 흡수하고, 11월 내 설계 V1.0 확정 및 인프라 구축 완료로 구현 일정 확보', { x: 0.6, y: 6.85, w: 10.8, h: 0.62, fontFace: F, fontSize: 11, bold: true, color: C.orange, valign: 'middle' });
  note(s, '※ 착수회의 WBS(개발·실증 로드맵) 기준으로 작성. 세부 일정은 WBS 재기준선 확정 시 갱신', 7.72);
}
{
  const s = addSlide();
  header(s, '03', 'Ⅵ. 향후 추진 계획  ▸  3. 성과지표 관리 계획', '성과지표 관리 계획');
  const rows = [
    [TH('구분'), TH('성과지표'), TH('목표'), TH('측정 기준(산식)'), TH('담당기관')],
    [TK('공통', { rowspan: 2 }), TD('고장탐구 시간당 생산성 향상률'), TD('30% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('(도입 후 시간당 완료건수 − 도입 전) ÷ 도입 전 × 100'), TD('코난테크놀로지', { align: 'center' })],
    [TD('납기 단축률'), TD('20% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('(도입 전 평균 납기일수 − 도입 후) ÷ 도입 전 × 100'), TD('코난테크놀로지', { align: 'center' })],
    [TK('핵심특화'), TD('AI 적용률'), TD('80% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('AI 처리·지원 고장탐구 건수 ÷ 전체 고장탐구 건수 × 100'), TD('코난테크놀로지', { align: 'center' })],
    [TK('자율', { rowspan: 4 }), TD('AI 기반 불량 예측 정확도'), TD('90% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('원인·부품 일치 건수 ÷ 전체 AI 예측 건수 × 100 (최소 100건)'), TD('코난테크놀로지', { align: 'center' })],
    [TD('온톨로지 도메인 커버리지율'), TD('90% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('500건 중 클래스·관계로 완전 표현 가능한 건수 비율'), TD('국립창원대학교\n산학협력단', { align: 'center' })],
    [TD('제조데이터 표준(AAS) 수용도'), TD('99% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('필수 항목 충족 건수 ÷ 전체 대상 데이터 건수 × 100'), TD('국립창원대학교\n산학협력단', { align: 'center' })],
    [TD('수집률'), TD('95% 이상', { align: 'center', bold: true, color: C.navy2 }), TD('최종 저장 수량 ÷ 원천 발생 수량 × 100'), TD('애버커스', { align: 'center' })],
  ];
  table(s, rows, { x: 0.45, y: 1.35, w: 11.03, colW: [1.2, 2.7, 1.1, 4.33, 1.7], rowH: [0.38, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5], fontSize: 10 });
  sub(s, '중간 점검 항목', 0.45, 5.6, 5);
  const r2 = [
    [TH('점검 항목'), TH('내용'), TH('현황')],
    [TK('기준값(Baseline)'), TD('도입 전 6개월 이상 생산성·납기 기준자료 확보'), PH()],
    [TK('평가 데이터'), TD('고장 사례 100건(정확도) · 500건(커버리지) 확보 및 정답 기준 정의'), PH()],
    [TK('KPI 정의서'), TD('판정기준·표본·제외조건·측정담당·현장 승인자 기록'), PH()],
  ];
  table(s, r2, { x: 0.45, y: 5.98, w: 11.03, colW: [2.0, 7.0, 2.03], rowH: 0.34, fontSize: 10 });
}

// ---------------- # 이슈 및 협의사항 ----------------
divider(6);
{
  const s = addSlide();
  header(s, '01', '# 이슈 및 협의사항  ▸  1. 주요 이슈 및 대응 방안', '주요 이슈 및 대응 방안');
  const rows = [
    [TH('No'), TH('구분'), TH('이슈 현황'), TH('영향'), TH('대응 방안(안)')],
    [TD('1', { align: 'center' }), TK('일정'), TD('분석 단계 소요 88일(계획 62일), 26일 초과'), TD('설계·구현 착수 지연, 후속 일정 압박'), TD('설계·구현 병행 추진, WBS 재기준선 수립 및 주간 진척 점검')],
    [TD('2', { align: 'center' }), TK('설계 확정'), TD('설계 산출물 90%, 지식 기반 산출물 v0.3~0.9'), TD('구현 단계 기준 문서 미확정'), TD('11월 내 V1.0 확정 목표, 미확정 항목은 구현과 병행 반복 검증')],
    [TD('3', { align: 'center' }), TK('인프라'), TD('HW·N/W 구축 50% (착수 계획: 9월 설치 완료)'), TD('개발·통합 환경 확보 지연'), TD('설치·구성 결과서 완료, K8s 개발환경 우선 전환')],
    [TD('4', { align: 'center' }), TK('사업비'), TD('집행률 0% (기성률 30% 대비)'), TD('대금 지급·증빙 일정 집중 위험'), TD('발주 건 계약 체결·집행 일정표 수립, 건별 증빙 사전 준비')],
    [TD('5', { align: 'center' }), TK('지식 구축'), TD('현장 전문가 인터뷰(계획 200시간) 연계 필요'), TD('온톨로지·KG 품질 및 KPI 검증 영향'), TD('인터뷰 일정·검증자 확정, 사례카드 기반 수집·검증')],
  ];
  table(s, rows, { x: 0.45, y: 1.35, w: 11.03, colW: [0.5, 1.2, 3.2, 2.6, 3.53], rowH: [0.42, 0.8, 0.8, 0.8, 0.8, 0.8], fontSize: 10.5 });
  note(s, '※ 이슈 1~4는 기성고 확인조서 수치 기반, 대응 방안은 협의를 위한 (안)임', 5.95);
  sub(s, '협의 요청 사항', 0.45, 6.35, 5);
  s.addText([
    { text: '설계 V1.0 확정 일정 및 WBS 재기준선 승인', options: { bullet: { code: '25AA' }, breakLine: true } },
    { text: '사업비 집행 일정(계약·검수·대금 지급) 및 증빙 준비 역할 분담', options: { bullet: { code: '25AA' } } },
  ], { x: 0.55, y: 6.7, w: 10.9, h: 0.75, fontFace: F, fontSize: 10.5, color: C.text, valign: 'top' });
}
{
  const s = addSlide();
  header(s, '02', '# 이슈 및 협의사항  ▸  2. 착수회의 후속조치 이행 점검', '착수회의 후속조치 이행 점검');
  const rows = [
    [TH('ID'), TH('실행할 일'), TH('담당 역할(안)'), TH('완료 기준'), TH('이행 상태'), TH('비고')],
    ...[
      ['A01', '최종 기간·통합 WBS 확정', '코난 PM / 전체', '기준 일정본'],
      ['A02', '계약 쟁점·체결일 합의', '계약 당사자', '검토본 · 체결일'],
      ['A03', '인터뷰 배정·1차 일정', '영풍 · 코난 · 창원대', '참여표 · 질문지'],
      ['A04', '1차 자료·보안조건 확정', '영풍 / 공급 · 참여', '자료목록 · 샘플'],
      ['A05', 'KPI 정의·기준자료 정리', '지표 담당 / 영풍', '정의서 · 원천목록'],
      ['A06', '감리 산출물·증빙 점검', '코난 · 경남TP · 사업비', '준비 · 조치목록'],
    ].map(r => [TD(r[0], { align: 'center', bold: true }), TD(r[1]), TD(r[2], { align: 'center' }), TD(r[3], { align: 'center' }), PH('[기입]'), PH('')]),
  ];
  table(s, rows, { x: 0.45, y: 1.35, w: 11.03, colW: [0.7, 3.0, 2.2, 1.9, 1.5, 1.73], rowH: 0.5, fontSize: 10.5 });
  note(s, '※ 착수회의(2026.09.17) 부록 "후속조치 기록표" 기준. 이행 상태는 미착수 · 진행 · 완료 · 보류로 기입', 5.05);
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.45, y: 5.5, w: 11.03, h: 1.6, fill: { color: C.sky2 }, line: { color: 'C9D6E8' } });
  s.addText([
    { text: '참고 — 기성고 확인조서 기준 확인 가능한 이행 근거', options: { bold: true, color: C.navy, breakLine: true } },
    { text: '· A01: WBS V1.0 수립(프로젝트 관리 산출물)', options: { breakLine: true } },
    { text: '· A06: 사업수행계획서 · 품질 · 형상 · 위험관리 계획서 V1.0 등 감리 대상 산출물 작성', options: { breakLine: true } },
    { text: '· A02 ~ A05: 기성고 확인조서에 이행 근거 미기재 → 담당 기관 확인 후 기입 필요' },
  ], { x: 0.65, y: 5.55, w: 10.7, h: 1.5, fontFace: F, fontSize: 10.5, color: C.text, valign: 'middle', paraSpaceAfter: 3 });
}

// 감사합니다
{
  const s = addSlide();
  s.addImage({ path: IMG('cover.png'), x: 0, y: 0, w: W, h: H });
  s.addText(PROJ, { x: 0.5, y: 3.0, w: 8, h: 0.35, fontFace: F, fontSize: 12, color: '595959' });
  s.addText('감사합니다', { x: 0.5, y: 3.35, w: 6, h: 1.0, fontFace: F, fontSize: 44, bold: true, color: '1F4E99' });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.55, y: 4.42, w: 2.6, h: 0.28, fill: { color: 'DCE6F2' }, line: { color: 'DCE6F2' } });
  s.addText('THANK YOU', { x: 0.65, y: 4.42, w: 2.4, h: 0.28, fontFace: F, fontSize: 9, color: C.navy, valign: 'middle' });
}

pptx.writeFile({ fileName: OUT }).then(f => console.log('saved', f));
