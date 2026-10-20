// 경남인공지능혁신본부 입주신청 사업계획 발표자료 (15분)
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const META = JSON.parse(fs.readFileSync('img/meta.json'));

const F = '페이퍼로지 6 SemiBold';
// v7(경남 지사 설립 방안) 팔레트
const C = { title: '0E2240', sub: '4D6E8E', body: '1F2A37', card: 'F7FAFD', line: 'C9D9E8', gold: 'B8860B', gold2: 'E2B23A',
  hl: 'B8860B', navy: '0E2240', steel: '3C6791', blue: '4D6E8E', pale: 'E9F2FB', white: 'FFFFFF' };
const FOOT = '경남인공지능혁신본부 입주신청 · 사업계획 발표 · ㈜코난테크놀로지 · 2026.10';

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 in
pptx.title = '경남인공지능혁신본부 입주신청 사업계획';
pptx.company = '㈜코난테크놀로지';
pptx.theme = { headFontFace: F, bodyFontFace: F };

let pageNo = 0;
const T = (s, text, x, y, w, h, o = {}) => s.addText(text, { x, y, w, h, fontFace: F, fontSize: 9.5, color: C.body, valign: 'top', margin: 0, paraSpaceAfter: 2, ...o });
const icon = (s, n, x, y, sz = 0.42) => s.addImage({ path: `icons/${n}.png`, x, y, w: sz, h: sz });
const card = (s, x, y, w, h, o = {}) => {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: o.hl ? C.pale : (o.fill || 'FFFFFF') }, line: { color: o.line || C.line, width: 0.75 } });
  if (o.hl) s.addShape(pptx.ShapeType.rect, { x, y, w: 0.04, h, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
};
// v7 섹션 라벨: 금색 세로 바 + 굵은 남색 제목
const sec = (s, text, x, y, w = 4, note = '') => {
  s.addShape(pptx.ShapeType.rect, { x, y: y + 0.04, w: 0.035, h: 0.2, fill: { color: C.gold }, line: { color: C.gold, width: 0 } });
  s.addText([{ text, options: { bold: true, color: C.title, fontSize: 10.5 } }, ...(note ? [{ text: '  ' + note, options: { color: C.sub, fontSize: 8 } }] : [])],
    { x: x + 0.1, y, w, h: 0.28, fontFace: F, valign: 'middle', margin: 0 });
};
// {강조} 구간은 금색으로
const rich = (str, base, hi) => str.split(/(\{[^}]*\})/).filter(Boolean).map(t => t.startsWith('{') ? { text: t.slice(1, -1), options: { color: hi } } : { text: t, options: { color: base } });
const bullets = (arr, o = {}) => arr.map((t, i) => {
  if (Array.isArray(t)) return { text: t[0], options: { bullet: { indent: 10 }, ...(t[1] || {}), breakLine: true } };
  return { text: t, options: { bullet: o.plain ? false : { indent: 10 }, breakLine: i < arr.length - 1 } };
});
function body(title, sub, notes) {
  const s = pptx.addSlide(); pageNo++;
  s.background = { color: 'FFFFFF' };
  T(s, rich(title, C.title, C.gold), 0.5, 0.32, 9.2, 0.45, { fontSize: 18, bold: true, valign: 'middle' });
  if (sub) T(s, sub, 0.5, 0.78, 9.2, 0.26, { fontSize: 9.75, color: C.sub, valign: 'middle' });
  T(s, FOOT, 0.5, 5.3, 6, 0.18, { fontSize: 6.5, color: '8E99A5', valign: 'middle' });
  T(s, String(pageNo).padStart(2, '0'), 9.0, 5.3, 0.5, 0.18, { fontSize: 6.5, color: '8E99A5', align: 'right', valign: 'middle' });
  if (notes) s.addNotes(notes);
  return s;
}
// place a cropped 3D image; returns fn mapping anchor -> slide inches
function img3d(s, key, x, y, w) {
  const m = META[key], h = w * m.h / m.w;
  s.addImage({ path: `img/${key}_c.png`, x, y, w, h });
  return (a) => [x + m.a[a][0] / m.w * w, y + m.a[a][1] / m.h * h];
}
const tableOpts = (colW, o = {}) => ({ fontFace: F, fontSize: 8, color: C.body, colW, border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: C.line }, { type: 'none' }], valign: 'middle', margin: [2, 4, 2, 4], ...o });
const th = (t) => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: C.steel }, align: 'center' } });

// =============== 1. COVER ===============
{
  const s = pptx.addSlide(); pageNo++;
  s.background = { path: 'img/cover_bg.jpg' };
  T(s, '경남인공지능혁신본부 입주신청 · 사업계획 발표 · 2026. 10', 0.55, 0.55, 6, 0.25, { fontSize: 9, color: C.gold2 });
  T(s, rich('경남권 제조AX\n{Konan Maestro}', 'FFFFFF', C.gold2), 0.55, 1.0, 6.5, 1.5, { fontSize: 36, bold: true, valign: 'middle', lineSpacingMultiple: 1.0 });
  T(s, '입주신청 사업계획', 0.55, 2.75, 6, 0.3, { fontSize: 12.5, bold: true, color: 'FFFFFF' });
  T(s, '제조 특화 온프레미스 생성형AI · 암묵지 지식화 플랫폼\n기반 경남권 제조AX 서비스', 0.55, 3.08, 5.0, 0.5, { fontSize: 10.5, bold: true, color: C.gold2 });
  T(s, '입주 희망 : 경남인공지능혁신본부 입주기업동 4동 703호 (93.79㎡) · 지사', 0.55, 3.68, 5.0, 0.26, { fontSize: 9.5, color: 'B7C3D3' });
  s.addShape(pptx.ShapeType.rect, { x: 0.55, y: 4.2, w: 0.6, h: 0.03, fill: { color: C.gold2 }, line: { color: C.gold2, width: 0 } });
  T(s, '㈜코난테크놀로지 국방 · 전략사업부', 0.55, 4.42, 6, 0.28, { fontSize: 11, color: 'FFFFFF' });
  T(s, 'Business Plan for Gyeongnam AI Innovation HQ', 0.55, 4.72, 6, 0.22, { fontSize: 8, color: '8E99A5' });
  s.addNotes('표지. 코난테크놀로지의 경남인공지능혁신본부 입주신청 사업계획을 15분간 발표드리겠습니다. 사업화 품목은 제조 특화 온프레미스 생성형AI·암묵지 지식화 플랫폼, Konan Maestro 기반 경남권 제조AX 서비스입니다.');
}

// =============== 2. AGENDA ===============
{
  const s = body('발표 개요 — {20주 안에 숫자로 증명}하는 경남 제조AX 거점', '경남 거점에서 무엇을, 누구와, 어떻게 증명할 것인가',
    '발표 개요입니다. 핵심 메시지는 네 가지입니다. 국산 Full-Stack AI 기업이고, 경남 현장에서 이미 정부 실증 3건을 수행 중이며, 20주 안에 숫자로 증명하는 Quick Win 전략을 갖고 있고, 지역 청년을 제조AI 인력으로 키우겠다는 것입니다. 순서는 오른쪽과 같습니다.');
  const items = [
    ['chip', '국산 Full-Stack AI 기업', '자체 LLM · RAG · LLMOps · STT를 모두 보유한 27년 업력의 코스닥 상장 AI 전문기업'],
    ['globe', '경남 현장에서 이미 수행 중', '영풍전자(창원) · KAI(사천) · 경남 뿌리기업 13개사 등 정부 실증사업 3건 진행'],
    ['gauge', '20주 안에 숫자로 증명', 'Quick Win A · B로 정확도 · 탐색시간 · 대조시간 KPI를 현장 기준선 대비 실측'],
    ['people', '지역 청년 · 제조AI 인력 양성', '2027년 현지 채용 3명, 본사 교육 후 창원 배치하는 순환 육성 모델'],
  ];
  card(s, 0.5, 1.15, 9.0, 0.62, { hl: true });
  T(s, '핵심 메시지', 0.65, 1.15, 0.95, 0.62, { fontSize: 9.5, bold: true, color: C.gold, valign: 'middle' });
  T(s, '"2026.11 경남사무소 1인 개설, 정부 실증 3건을 현장 상주로 대응하고 Quick Win으로 20주 안에 성과를 수치로 검증 — 2027.7 5인 체제 제조AX기술센터@경남으로 전환한다."', 1.6, 1.15, 7.8, 0.62, { fontSize: 9.5, color: C.title, valign: 'middle' });
  items.forEach(([ic, h, d], i) => {
    const y = 2.0 + i * 0.8;
    icon(s, ic, 0.55, y, 0.48);
    T(s, h, 1.2, y - 0.02, 4.8, 0.28, { fontSize: 11, bold: true, color: C.title });
    T(s, d, 1.2, y + 0.27, 4.8, 0.42, { fontSize: 9 });
  });
  card(s, 6.3, 1.95, 3.2, 3.15, { fill: C.card });
  sec(s, '발표 순서', 6.45, 2.05, 2.8);
  const ag = ['입주 신청 개요', '신청기업 현황 · 실적', '사업화 품목 및 기술 경쟁력', '시장 · 수요처 · 경쟁 현황', '개발 및 사업화 추진계획', '매출 · 판매전략', '청년고용 · 소요자금'];
  ag.forEach((t, i) => {
    const y = 2.42 + i * 0.37;
    s.addShape(pptx.ShapeType.ellipse, { x: 6.58, y: y + 0.02, w: 0.24, h: 0.24, fill: { color: C.navy }, line: { color: C.navy, width: 0.75 } });
    T(s, String(i + 1), 6.58, y + 0.02, 0.24, 0.24, { fontSize: 8, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    T(s, t, 6.95, y, 2.4, 0.28, { fontSize: 9.5, color: C.body, valign: 'middle' });
  });
}

// =============== 3. 입주 신청 개요 ===============
{
  const s = body('입주 신청 개요 — 4동 703호, {경남 제조AX 현장 거점}', '경남인공지능혁신본부 입주기업동 4동 703호 — 경남 제조AX 현장 거점',
    '입주 신청 개요입니다. 2025년 매출 339억 원, 종업원 194명, 등록특허 18건, GS 1등급 14개 제품을 보유하고 있습니다. 입주 희망 공간은 4동 703호 93.79제곱미터, 지사 형태이며, 입주 시 1명으로 시작해 2027년 파견 2명과 현지 채용 3명으로 5명 체제로 확대합니다.');
  const k = [['1999.04', '법인 설립 (업력 27년)'], ['339.8억 원', '2025년 매출액'], ['194명', '종업원 (AI 인력 139명)'], ['18건', '등록특허'], ['14개', 'GS 1등급 인증 제품'], ['2022.07', '코스닥 상장 (402030)']];
  k.forEach(([v, l], i) => {
    const x = 0.5 + i * 1.52;
    card(s, x, 1.3, 1.4, 1.05, { fill: 'FFFFFF' });
    T(s, v, x, 1.43, 1.4, 0.45, { fontSize: 17, bold: true, color: C.title, align: 'center', valign: 'middle' });
    T(s, l, x, 1.92, 1.4, 0.3, { fontSize: 8, color: C.sub, align: 'center', valign: 'middle' });
  });
  const rows = [
    ['입주 희망 공간', '경남인공지능혁신본부 입주기업동 4동 703호  /  93.79㎡ (28.4평)'],
    ['입주 형태', '지사  (1단계 「코난테크놀로지 경남사무소」 → 2단계 「(가칭) 제조AX기술센터@경남」 지점 등기 전환)'],
    ['동반 입주 인력', '1명 (2026년 입주 시)  →  5명 (2027년, 파견 2 · 현지 채용 3)'],
    ['사업화 품목', '제조 특화 온프레미스 생성형AI · 암묵지 지식화 플랫폼(Konan Maestro) 기반 경남권 제조AX 서비스'],
    ['사업 분야', '인공지능 소프트웨어 (제조AX · 생성형AI · RAG · 온톨로지/지식그래프 · 음성인식)'],
  ];
  s.addTable(rows.map(([a, b]) => [{ text: a, options: { bold: true, color: C.title, fill: { color: C.card }, align: 'center' } }, { text: b }]),
    { x: 0.5, y: 2.6, ...tableOpts([1.6, 7.4], { fontSize: 9.5, rowH: 0.42 }) });
}

// =============== 4. 경영진 · 주주 ===============
{
  const s = body('Ⅰ. 기업 현황 ① — {박사급 경영진}과 전략적 투자자(SK텔레콤 · KAI)', '연구소 출신 창업자와 박사급 경영진, 전략적 투자자(SK텔레콤 · KAI) 참여',
    '신청기업 현황입니다. 대표이사는 한양대 전산학 박사로 ETRI와 미국 Bellcore를 거쳐 1999년 창업했습니다. COO, CTO 모두 박사급이며, 주주 구성에 SK텔레콤과 한국항공우주산업이 전략적 투자자로 참여하고 있어, KAI 등 경남 앵커기업과의 협력 기반이 있습니다.');
  card(s, 0.5, 1.25, 4.4, 1.55, { hl: true });
  icon(s, 'usercheck', 0.65, 1.38, 0.5);
  T(s, '대표이사  김영섬', 1.25, 1.38, 3.5, 0.3, { fontSize: 12, bold: true, color: C.title });
  T(s, '한양대학교 대학원 전산학 박사 (1989)', 1.25, 1.68, 3.5, 0.22, { fontSize: 9, color: C.sub });
  T(s, bullets(['1999.04 ~ 현재  ㈜코난테크놀로지 대표이사 (창업자)', '1989.09 ~ 1997.01  한국전자통신연구원 선임연구원', '1991.01 ~ 1996.03  미국 Bellcore 연구소 객원연구원', '벤처산업 대통령 표창, 대한민국 S/W기업대상 등 수상']), 0.65, 2.0, 4.15, 0.8, { fontSize: 8.5 });
  const ex = [['COO', '양승현', '서울대 컴퓨터공학 박사', 'ETRI · 서울대IBM 자연어처리연구실'], ['CFO', '홍순철', '고려대 경제학 학사', '현대카드 회계팀 · KTB네트워크'], ['CTO', '최정주', '포항공대 컴퓨터공학 박사', 'ETRI · 디지털아리아 CTO · 아주대 교수']];
  ex.forEach(([r, n, e, c], i) => {
    const y = 2.95 + i * 0.72;
    card(s, 0.5, y, 4.4, 0.62);
    T(s, r, 0.6, y, 0.6, 0.62, { fontSize: 11, bold: true, color: C.gold, valign: 'middle' });
    T(s, `${n}   ${e}`, 1.25, y + 0.08, 3.55, 0.24, { fontSize: 9.5, bold: true, color: C.title });
    T(s, c, 1.25, y + 0.33, 3.55, 0.22, { fontSize: 8.5 });
  });
  sec(s, '주주 구성', 5.3, 1.25, 4);
  const sh = [['김영섬 (CEO)', 22.88], ['SK텔레콤㈜', 18.85], ['한국항공우주산업㈜', 7.1], ['양승현 (COO)', 6.72], ['윤덕호 (연구위원)', 4.84]];
  const etc = +(100 - sh.reduce((a, b) => a + b[1], 0)).toFixed(2);
  s.addChart(pptx.ChartType.doughnut, [{ name: '지분율', labels: [...sh.map(a => a[0]), '기타(주주명부)'], values: [...sh.map(a => a[1]), etc] }], {
    x: 5.2, y: 1.55, w: 2.2, h: 2.2, holeSize: 62, showLegend: false, showValue: false, showPercent: false, dataLabelColor: 'FFFFFF',
    chartColors: ['0E2240', 'B8860B', '3C6791', '4D6E8E', '9FB4CC', 'DDE9F4'], dataBorder: { pt: 1, color: 'FFFFFF' } });
  T(s, '코스닥\n402030', 5.2, 2.4, 2.2, 0.5, { fontSize: 9, color: C.title, align: 'center', valign: 'middle', bold: true });
  const cols = ['0E2240', 'B8860B', '3C6791', '4D6E8E', '9FB4CC', 'DDE9F4'];
  [...sh, ['기타 (주주명부 별첨)', etc]].forEach(([n, v], i) => {
    const y = 1.62 + i * 0.33;
    s.addShape(pptx.ShapeType.rect, { x: 7.55, y: y + 0.07, w: 0.14, h: 0.14, fill: { color: cols[i] }, line: { color: cols[i] } });
    T(s, n, 7.78, y, 1.3, 0.28, { fontSize: 8.5, valign: 'middle' });
    T(s, v.toFixed(2) + '%', 8.9, y, 0.6, 0.28, { fontSize: 8.5, bold: true, color: C.title, align: 'right', valign: 'middle' });
  });
  card(s, 5.3, 3.95, 4.2, 1.1, { fill: 'FFFFFF' });
  icon(s, 'transfer', 5.42, 4.1, 0.42);
  T(s, '전략적 투자자 = 경남 협력 채널', 5.95, 4.07, 3.5, 0.26, { fontSize: 10, bold: true, color: C.title });
  T(s, 'KAI(사천)와 고정익 · 회전익 PHM DB 구축(2022.12~2028.12) 수행 중. SK텔레콤 · KAI 협력과 연계한 동반진출 추진', 5.95, 4.35, 3.45, 0.65, { fontSize: 8.5 });
}

// =============== 5. 연혁 · 인력 · 인프라 ===============
{
  const s = body('Ⅰ. 기업 현황 ② — {27년} 원천기술 축적, 석 · 박사 65명 · DGX H100 9대', '자연어처리 원천기술에서 제조AX 전문기업으로 — 27년의 축적',
    '연혁입니다. 1999년 자연어처리·검색·영상AI 원천기술로 출발해, 2007년부터 실시간 AI 분석 서비스 펄스K로 250억 건 이상의 한국어 데이터를 확보했습니다. 2022년 코스닥 상장, 2025년 AI+ 인증, 2026년에는 M.AX 전문기업 등록과 정부 실증사업 3건 선정이 이어졌습니다. 기술인력은 석·박사 65명, 인프라는 NVIDIA DGX H100 9대를 보유하고 있습니다.');
  const ms = [['1999.04', '법인 설립', '자연어처리·검색·영상AI'], ['2007~', '펄스K 운영', '한국어 데이터 250억 건+'], ['2022.07', '코스닥 상장', 'GS 1등급 Konan Search6'], ['2023.05', '메인비즈 인증', 'SP 2등급(2023.11)'],
    ['2025.08', 'Konan RAG-X V3', 'GS 1등급 인증'], ['2025.12', 'AI+ 인증', 'Konan-LLM-ENT-11'], ['2026.07', 'M.AX 전문기업', '포항 AX 실증산단 선정'], ['2026', '정부 실증 선정', '중기부 신속상용화 · 피지컬AI']];
  s.addShape(pptx.ShapeType.line, { x: 0.6, y: 1.72, w: 8.8, h: 0, line: { color: C.line, width: 2 } });
  ms.forEach(([d, h, t], i) => {
    const x = 0.5 + i * 1.13, last = i >= 6;
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.44, y: 1.63, w: 0.18, h: 0.18, fill: { color: last ? C.gold2 : C.navy }, line: { color: 'FFFFFF', width: 1.5 } });
    T(s, d, x, 1.3, 1.06, 0.26, { fontSize: 9.5, bold: true, color: last ? C.gold : C.title, align: 'center' });
    T(s, h, x, 1.92, 1.06, 0.24, { fontSize: 9, bold: true, color: C.title, align: 'center' });
    T(s, t, x, 2.16, 1.06, 0.38, { fontSize: 7.5, color: C.sub, align: 'center' });
  });
  // 인력
  card(s, 0.5, 2.75, 4.4, 2.3);
  icon(s, 'people', 0.62, 2.85, 0.42);
  T(s, '기술인력 현황', 1.12, 2.9, 3, 0.3, { fontSize: 11, bold: true, color: C.title });
  const ed = [['박사', 9], ['석사', 56], ['학사', 117], ['전문대', 11], ['고졸 이하', 2]];
  s.addChart(pptx.ChartType.bar, [{ name: '학력별', labels: ed.map(a => a[0]), values: ed.map(a => a[1]) }], {
    x: 0.6, y: 3.3, w: 2.2, h: 1.7, barDir: 'bar', chartColors: ['3C6791'], catAxisLabelFontSize: 7.5, catAxisLabelFontFace: F, valAxisHidden: true, valGridLine: { style: 'none' },
    showValue: true, dataLabelFontSize: 7.5, dataLabelFontFace: F, dataLabelColor: C.title, catAxisLabelColor: C.body, barGapWidthPct: 45, catAxisOrientation: 'maxMin' });
  T(s, bullets(['종업원 194명 (정규 188 · 계약 6)', 'AI 인력 139명, 청년(만15~34세) 34%', 'CTO 최정주 — 연구개발 총괄', '김만수 상무 — PHM · 디지털트윈 · 이상탐지', '강현수 이사 — AI 엔진 · LLM 연구개발', '제조AX 전담 5명 (책임 · 선임 · 전임)', '동업계 15~34년 경력 연구인력 다수']), 2.85, 3.3, 2.0, 1.7, { fontSize: 7.5 });
  // 인프라
  card(s, 5.1, 2.75, 4.4, 2.3, { hl: true });
  icon(s, 'server', 5.22, 2.85, 0.42);
  T(s, '연구 · 개발 인프라', 5.72, 2.9, 3, 0.3, { fontSize: 11, bold: true, color: C.title });
  T(s, 'NVIDIA DGX H100', 5.3, 3.35, 2.5, 0.3, { fontSize: 10, bold: true, color: C.title });
  T(s, '9대', 7.9, 3.25, 1.4, 0.45, { fontSize: 22, bold: true, color: C.gold, align: 'right' });
  const inf = [['본사 9F', '1대', '473,900천원'], ['가산 IDC', '8대', '3,667,450천원']];
  s.addTable([[th('설치장소'), th('수량'), th('구입금액'), th('용도')], ...inf.map(r => [...r, '연구 · 개발'])], { x: 5.3, y: 3.75, ...tableOpts([1.0, 0.6, 1.3, 1.1], { fontSize: 8, rowH: 0.27, align: 'center' }) });
  T(s, '2023년 도입 · 자체 Foundation LLM 학습 및 제조 특화 모델 미세조정에 활용', 5.3, 4.65, 4.1, 0.3, { fontSize: 8, color: C.sub });
}

// =============== 6. 기술개발 및 사업화 실적 ===============
{
  const s = body('Ⅰ. 기업 현황 ③ — 실적은 이미 {경남 제조 현장}에 있다', '경남(창원 · 사천) 제조 현장에서 이미 진행 중인 실증 · 구축 사업',
    '기술개발 및 사업화 실적입니다. 창원 영풍전자와 방산 전자회로기판 고장탐구 AI를 대표 공급기업으로 수행 중이고, 포항 스마트그린산단 AX 실증, 경남 뿌리기업 13개사 피지컬AI, 사천 KAI PHM DB 구축, 창원 세아창원특수강 이상탐지까지, 경남 현장 경험이 축적돼 있습니다.');
  const pr = [
    ['중기부 AI 응용제품 신속상용화', '영풍전자 컨소시엄 · 애버커스 · 국립창원대 · 경남TP', '2026~2027.07', '대표 공급기업', '방산 전자회로기판(PBA) 고장탐구 암묵지 온톨로지 · Graph-RAG · 온프레미스 LLM 자산화', '고장 원인 분석시간 50% · 보고서 작성시간 70% 단축 목표, 종료 후 2년 내 방산 전자 신규 고객 5개사', '창원', true],
    ['산업부 스마트그린산단 AX 실증산단', '경북 포항 · 주관 RIST', '2026.07~2029.12', '참여기업', '철강 숙련공 조업 암묵지 지식그래프 · AI Agent 자산화(Steel Intelligence), 예지보전 · 암묵지 Agent', '카세트형 모듈 구독형 SaaS로 산단 카탈로그 등재, 5극3특 권역 확산 계획', '포항'],
    ['과기정통부 피지컬AI 연구개발 세부3', '물리데이터 품질관리 · ETRI 컨소시엄 · 메가존 · KTL', '2026~2030.12', '참여기업', '경남 뿌리기업 13개사 실증, 제조 온톨로지 자동설계, 3D 공정 시나리오 자동 생성 · Sim2Real 정합', '온톨로지 자동설계SW 사이트 라이선스 사업화 추진', '경남'],
    ['KAI 고정익 · 회전익 PHM 시스템용 DB 구축', '한국항공우주산업㈜', '2022.12~2028.12', '수행기업', '비행 · 정비 · 센서 · 부품교체 데이터 통합 PHM DB, ETL · Feature Engineering · 시계열 처리', 'T-50 및 회전익 · 고정익 PHM DB 구축, 수명예측 알고리즘 검증 완료', '사천'],
    ['세아창원특수강 Pump 장치 AI 이상탐지', '㈜세아창원특수강', '2024.11~2024.12', '수행기업', '2단조 프레스 main pump 고장 15일 전 이상탐지 알고리즘 · 데이터 플랫폼 연동', '이상탐지 알고리즘 2건 개발, 한국철도공사 KTX 모터 감속기 PoC 연계', '창원'],
  ];
  pr.forEach(([n, org, per, role, what, res, loc, hl], i) => {
    const y = 1.22 + i * 0.79;
    card(s, 0.5, y, 9.0, 0.71, { hl, fill: hl ? 'FFFFFF' : C.card });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.6, y: y + 0.17, w: 0.62, h: 0.36, rectRadius: 0.05, fill: { color: loc === '포항' ? '77869A' : C.navy }, line: { color: 'FFFFFF', width: 0 } });
    T(s, loc, 0.6, y + 0.17, 0.62, 0.36, { fontSize: 9, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    T(s, n, 1.35, y + 0.07, 3.0, 0.24, { fontSize: 9.5, bold: true, color: C.title });
    T(s, `${org}\n${per}  ·  ${role}`, 1.35, y + 0.32, 3.0, 0.36, { fontSize: 7.5, color: C.sub });
    T(s, what, 4.4, y + 0.08, 2.65, 0.58, { fontSize: 8, valign: 'middle' });
    T(s, res, 7.15, y + 0.08, 2.25, 0.58, { fontSize: 8, color: C.title, valign: 'middle' });
  });
}

// =============== 7. 정책자금 수혜 현황 ===============
{
  const s = body('Ⅰ. 기업 현황 ④ — 정부과제 {9건} 수행 중, 2026년 신규 3건', '2026년 신규 선정 3건 포함, 국방 · 제조 · AI 원천 분야 정부과제 9건 수행 중  (단위 : 천원)',
    '정책자금 수혜 현황입니다. 2026년에만 중기부 18억, 산업부 12억, 과기정통부 피지컬AI 80억 규모의 과제가 신규 선정됐고, 국방 인공지능 핵심기술, 무인기 탑재 AI 파일럿, 드론 통합임무 등 국방 분야 과제도 함께 수행하고 있습니다. 신용보증 수혜는 없습니다.');
  const rows = [
    ['AI 응용제품 신속상용화', '방산 전자PBA 고장탐구 암묵지AI 실증 (영풍전자 컨소시엄)', '중소벤처기업부', '1,830,000', '2026.07', '~2027.07'],
    ['스마트그린산단 AX 실증산단 구축', '철강 암묵지 지식그래프 · AI Agent 디지털 자산화 (동국제강)', '산업통상부 / RIST', '1,200,000', '2026.07', '~2029.12'],
    ['피지컬AI 연구개발사업 세부3', '물리데이터 품질관리 (ETRI 컨소시엄, 경남 뿌리기업 13개사)', '과학기술정보통신부', '8,000,000', '2026.08', '~2030.12'],
    ['AI반도체 데이터센터 컴퓨팅 SW', '이종 AI반도체용 분산 추론 및 모델 최적화', '과학기술정보통신부', '100,000 (26년도분)', '2025', '~2029.12'],
    ['국방인공지능핵심기술개발', '유·무인 복합체계 상황인지 · 복합AI 분석 모델 자동 협업', '과학기술정보통신부', '51,664 (26년도분)', '2025', '~2028.12'],
    ['국산 SoC 기반 온디바이스 AI', '대화형 에이전트 탑재 키오스크 시스템 개발 및 실증', '한국산업기술기획평가원', '466,000 (26년도분)', '2024', '~2027.05'],
    ['디지털트윈 활용 AI 파일럿', '인공지능 파일럿 개발 및 무인기 탑재 실증', '국방기술진흥연구소', '907,500 (26년도분)', '2024', '~2028.10'],
    ['드론 통합임무 모듈 · 통합 GCS', '드론 통합임무 모듈 및 통합 GCS 기술 개발', '국방기술진흥연구소', '310,000 (26년도분)', '2025', '~2028.12'],
    ['XAI+ 플러그앤플레이', '설명가능성 제공 인공지능 기술 및 설명 제공 검증', '과학기술정보통신부', '626,000 (26년도분)', '2022', '~2026.12'],
  ];
  const hlRow = (r, i) => r.map((c, j) => ({ text: c, options: { fill: { color: i < 3 ? 'FBF5E6' : 'FFFFFF' }, bold: i < 3 && (j === 0 || j === 3), color: i < 3 && j === 3 ? C.gold : C.body, align: j >= 3 ? 'center' : 'left' } }));
  s.addTable([[th('사업명'), th('과제명'), th('지원기관'), th('금액'), th('선정'), th('진행')], ...rows.map(hlRow)],
    { x: 0.5, y: 1.22, ...tableOpts([2.05, 3.05, 1.35, 1.2, 0.68, 0.67], { fontSize: 8, rowH: 0.37 }) });
  T(s, '※ 음영 = 2026년 신규 선정 경남 · 제조 실증 과제, 9건 모두 수행 중. 신용보증 수혜 없음', 0.5, 5.0, 9, 0.22, { fontSize: 7.5, color: C.sub });
}

// =============== 8. 사업화 품목 ===============
{
  const s = body('Ⅱ. 사업화 품목 — {Konan Maestro}, 국산 Full-Stack 제조AI', '제조 특화 온프레미스 생성형AI · 암묵지 지식화 플랫폼 (국산 Full-Stack)',
    '사업화 품목 Konan Maestro입니다. 자체 Foundation LLM, 자체 형태소 분석기 기반 RAG, LLMOps·RAGOps 통합 플랫폼, 음성인식 엔진을 모두 자체 보유한 Full-Stack 구조입니다. 그 위에 설비 매뉴얼, 정비이력, 작업일지, 도면, 숙련자의 암묵지를 온톨로지와 지식그래프로 자산화하고, 폐쇄망에서 생성형AI로 활용합니다.');
  const at = img3d(s, 'stack', 0.35, 1.2, 3.55);
  const lay = [['Konan LLM', 'ENT-13(32B) · ENT-20 · PRO-10 · PRO-20 자체 Foundation 모델, 온프레미스 · 폐쇄망'],
    ['Konan RAG', '자체 형태소 분석기 · 10억 건+ 벡터검색 · 출처 표시 · 접근권한 제어'],
    ['STT · 멀티모달', '종단형 한국어 음성인식(음절 인식률 90%+), 도면 · 영상 · 음성 통합'],
    ['Konan LLM Studio', 'LLMOps + RAGOps, Agent Builder, NL2SQL, REST API · MCP로 MES · PLM · ERP · CMMS 연계'],
    ['온톨로지 · Graph-RAG', '설비 · 정비 · 도면 · 숙련자 암묵지(Tacit Knowledge)를 지식그래프로 자산화']];
  lay.forEach(([h, d], i) => {
    const [ax, ay] = at('R' + i); const ly = 4.25 - i * 0.68;
    s.addShape(pptx.ShapeType.line, { x: ax, y: Math.min(ay, ly + 0.15), w: 4.15 - ax, h: Math.abs(ly + 0.15 - ay), flipV: ay > ly + 0.15, line: { color: 'A9B8C8', width: 0.75, dashType: 'dash' } });
    s.addShape(pptx.ShapeType.ellipse, { x: ax - 0.04, y: ay - 0.04, w: 0.08, h: 0.08, fill: { color: C.gold2 }, line: { color: C.gold2 } });
    T(s, h, 4.2, ly, 1.7, 0.3, { fontSize: 10.5, bold: true, color: i === 4 ? C.gold : C.title, valign: 'middle' });
    T(s, d, 4.2, ly + 0.28, 2.0, 0.4, { fontSize: 7.5, color: C.sub });
  });
  card(s, 6.45, 1.25, 3.05, 3.8, { hl: true });
  T(s, '핵심 기술', 6.6, 1.35, 2.8, 0.28, { fontSize: 11, bold: true, color: C.title });
  T(s, '설비 매뉴얼 · 정비이력 · 작업일지 · 도면 · 품질 리포트 등 방대한 비정형 데이터와 숙련자 암묵지를 온톨로지 · 지식그래프(Graph-RAG)로 자산화하고, 폐쇄망(Air-Gapped)에서 생성형AI로 활용', 6.6, 1.65, 2.8, 1.05, { fontSize: 8.5 });
  T(s, '경남 거점의 역할', 6.6, 2.8, 2.8, 0.28, { fontSize: 11, bold: true, color: C.title });
  T(s, '수행 중인 실증사업과 등대고객(영풍전자 · KAI · 사천 항공 협력사 · 거제 조선 협력사)이 모두 창원 · 사천 · 거제 권역에 소재', 6.6, 3.1, 2.8, 0.75, { fontSize: 8.5 });
  T(s, '→ 거점에서 현장 데이터 수집 · 검증 · 고객 대면, 실증 결과를 경남 제조기업으로 수평 확산', 6.6, 3.9, 2.8, 0.6, { fontSize: 8.5, bold: true, color: C.gold });
  T(s, '자체 개발 · GS 1등급 · AI+ 인증', 0.5, 4.95, 3.5, 0.22, { fontSize: 7.5, color: C.sub });
}

// =============== 9. 경쟁제품 대비 개선점 · 기술/품질 경쟁력 ===============
{
  const s = body('Ⅱ. 기술 경쟁력 — 외산 · 클라우드 LLM 대비 {4가지 개선점}', '외산 · 클라우드 LLM이 해결하지 못하는 제조 현장의 4가지 요구',
    '경쟁제품 대비 개선점은 네 가지입니다. 기밀 데이터를 외부로 내보내지 않는 온프레미스 운영, 출처 표시로 안전·품질사고로 이어질 수 있는 환각을 방지, 자체 형태소 분석기로 설비명·부품명 같은 제조 복합명사를 정확히 인식, 그리고 토큰은 많이 파라미터는 작게 전략으로 중소·중견기업도 도입 가능한 비용입니다. 하단은 이를 뒷받침하는 정량 지표와 품질 인증입니다.');
  const pts = [['lock', '① 기밀 데이터 외부 유출 없음', '외산 · 클라우드 LLM과 달리 온프레미스 · 폐쇄망 운영'],
    ['docsearch', '② 출처 표시(Citation) 기반 답변', '안전 · 품질사고로 직결될 수 있는 할루시네이션 방지'],
    ['tags', '③ 제조 복합명사 정확 인식', '자체 형태소 분석기로 설비명 · 부품명 · 공정용어 인식'],
    ['gauge', '④ 「토큰은 많이, 파라미터는 작게」', '32B 이하 고효율 모델로 엣지 GPU 1~2기 수준에서 운영']];
  pts.forEach(([ic, h, d], i) => {
    const x = 0.5 + i * 2.27;
    card(s, x, 1.25, 2.13, 1.55, { hl: i === 0 });
    icon(s, ic, x + 0.12, 1.37, 0.46);
    T(s, h, x + 0.12, 1.9, 1.95, 0.3, { fontSize: 9.5, bold: true, color: C.title });
    T(s, d, x + 0.12, 2.2, 1.95, 0.55, { fontSize: 8 });
  });
  sec(s, '기술 경쟁력', 0.5, 2.98, 3);
  const kp = [['250억 건+', '펄스K 확보 데이터\n(양질 약 20억 건 학습)'], ['600만 개', '자체 형태소 분석기\n엔트리'], ['10억 건+', '대규모 문서\n벡터검색'], ['5GB / 시간', '2 Core 서버 인덱싱\n초당 50건 질의 처리'], ['1.5~10배', '타 엔진 대비\n서버당 처리량']];
  kp.forEach(([v, l], i) => {
    const x = 0.5 + i * 1.22;
    T(s, v, x, 3.3, 1.15, 0.38, { fontSize: 14, bold: true, color: i === 4 ? C.gold : C.title, align: 'center', valign: 'middle' });
    T(s, l, x, 3.7, 1.15, 0.42, { fontSize: 7.5, color: C.sub, align: 'center' });
  });
  T(s, bullets(['제조 적용 실적 : KAI T-50 · 회전익 · 고정익 PHM DB 및 수명예측, 세아창원특수강 이상탐지 2건, 발전3사 AI Agent · 업무자료 자산화']), 0.5, 4.3, 5.9, 0.6, { fontSize: 8.5 });
  card(s, 6.75, 2.98, 2.75, 2.05, { fill: 'FFFFFF', hl: true });
  icon(s, 'shield', 6.87, 3.08, 0.42);
  T(s, '품질 경쟁력', 7.35, 3.13, 2, 0.3, { fontSize: 11, bold: true, color: C.title });
  T(s, bullets(['GS 1등급 인증 14개 제품 (Konan RAG-X V3, Konan Search6 등)', 'AI+ 인증 Konan-LLM-ENT-11 (2025.12~2027.12)', 'SP 2등급 (CMMI · SPICE 3레벨 준형)', '등록특허 18건 (설명가능AI, 딥러닝 성능평가 · 고속 실행, AI 부품 검색 등)', '개인정보보호 가이드라인 준수 데이터 사전학습 · 필터링']), 6.87, 3.55, 2.55, 1.45, { fontSize: 7.5 });
}

// =============== 10. 국내외 동향 · 포지셔닝 ===============
{
  const s = body('Ⅲ. 관련기술 현황 — 제조AI의 무게중심이 {현장 지식 · 암묵지}로 이동', '제조AI의 무게중심이 「현장 문서 · 암묵지의 지식화」와 「에이전트 기반 업무 자동화」로 이동',
    '국내외 동향입니다. 제조AI는 대기업의 예지보전·품질관리와 중소기업의 기초 디지털화로 양극화돼 있고, 생성형AI 적용은 현장 문서와 암묵지의 지식화, 에이전트 기반 자동화로 이동하고 있습니다. 국내에서는 산업부 M.AX, 과기정통부 피지컬AI, 경남도 마스터플랜과 창원시의 제조AI 전환 원년 선언이 이어지고 있습니다. 당사는 설비·클라우드·범용LLM 사이의 빈자리인 지식·암묵지 계층에 집중합니다.');
  card(s, 0.5, 1.25, 4.6, 1.75);
  icon(s, 'globe', 0.62, 1.35, 0.42);
  T(s, '글로벌 동향', 1.12, 1.4, 3, 0.3, { fontSize: 11, bold: true, color: C.title });
  T(s, bullets(['예지보전 · 품질관리 고도화(대기업)와 MES/ERP 기초 디지털화(중소기업)로 양극화, 생성형AI 적용은 현장 문서 · 암묵지 지식화와 에이전트 기반 자동화로 이동 (삼일PwC, 2026.3)', 'LLM은 범용 초거대 모델에서 도메인 특화 sLLM · 온프레미스 · Agentic AI · Physical AI로 분화 — 폐쇄망 운영 가능한 모델 · RAG 수요 급증']), 0.62, 1.8, 4.4, 1.2, { fontSize: 8 });
  card(s, 0.5, 3.1, 4.6, 1.95);
  icon(s, 'flag', 0.62, 3.2, 0.42);
  T(s, '국내 동향', 1.12, 3.25, 3, 0.3, { fontSize: 11, bold: true, color: C.title });
  T(s, bullets(['산업통상부 「제조업AI 대전환(M.AX)」: 얼라이언스 1,000여 개 산학연, 2030년까지 AI 팩토리 500개', '과기정통부 피지컬AI 연구개발(총 6,763억 원, 2026~2030), 창원 착수 · 경남 12개 제조사 실증', '경남도 「2026~2030 AI 산업 마스터플랜」: 제조AI 특화기업 100곳, AI 혁신공장 4,000개, 2026 경남 제조AI 데이터센터', '창원시 「2026년 제조AI 전환 원년」, 창원국가산단 구조고도화(피지컬AI 실증 선도 제조거점)']), 0.62, 3.62, 4.4, 1.4, { fontSize: 8 });
  // positioning
  sec(s, '국내 AI 기업 포지셔닝', 5.4, 1.25, 4);
  s.addImage({ path: 'img/graph_c.png', x: 7.65, y: 1.5, w: 1.85, h: 1.85 * META.graph.h / META.graph.w });
  const pos = [['설비 · 공정 데이터 기반 최적화', '마키나락스 등', C.pale], ['클라우드 · 시스템 통합', '메가존 · LG CNS 등', C.pale], ['범용 LLM', '솔트룩스 · 와이즈넛 등', C.pale], ['지식 · 암묵지 계층', '코난테크놀로지 — 자체 LLM + RAG + 온톨로지/지식그래프 + 온프레미스', 'FBF5E6']];
  pos.forEach(([a, b, f], i) => {
    const y = 1.62 + i * 0.62, me = i === 3;
    s.addShape(pptx.ShapeType.roundRect, { x: 5.4, y, w: me ? 4.1 : 2.15, h: 0.52, rectRadius: 0.05, fill: { color: f }, line: { color: me ? C.hl : C.line, width: me ? 1.25 : 0.5 } });
    T(s, a, 5.52, y + 0.05, me ? 3.9 : 2.0, 0.22, { fontSize: 8.5, bold: true, color: me ? C.gold : C.title });
    T(s, b, 5.52, y + 0.26, me ? 3.9 : 2.0, 0.22, { fontSize: 7.5, color: C.sub });
  });
  card(s, 5.4, 4.2, 4.1, 0.85, { fill: 'FFFFFF' });
  T(s, '→ 설비 데이터 · 클라우드 · 범용 LLM 사이의 빈자리, 「현장 지식 · 암묵지」를 폐쇄망에서 활용 가능하게 만드는 계층에 집중', 5.55, 4.27, 3.85, 0.7, { fontSize: 9, bold: true, color: C.title, valign: 'middle' });
}

// =============== 11. 시장규모 ===============
{
  const s = body('Ⅲ. 사업성 ① — 제조AX 예산이 {영남 · 경남}에 집중된다', '2026~2030년 영남권에 조 단위 제조AI 공적 예산 집행 — 경남 · 창원이 핵심 실증 무대',
    '시장 규모입니다. 글로벌 제조AI 시장은 2025년 85.7억 달러에서 2035년 2,872억 달러로 연평균 40%대 성장이 전망됩니다. 국내는 M.AX 정책과 피지컬AI, 실증산단, 거제·부산 공모 등 영남권에 2030년까지 조 단위 공적 예산이 집행되며, 경남·창원은 그 핵심 실증 무대입니다. 실증사업은 앵커기업, 지역 대학·TP, AI 공급기업 컨소시엄 구조이고, 보안 요건으로 온프레미스 수요가 높습니다.');
  const at = img3d(s, 'mkt', 0.75, 1.62, 1.9);
  const [x0, y0] = at('T0'), [x1, y1] = at('T1');
  T(s, '85.7억$', x0 - 0.5, y0 - 0.32, 1.0, 0.26, { fontSize: 10, bold: true, color: C.title, align: 'center' });
  T(s, '2,872억$', x1 - 0.6, y1 - 0.35, 1.2, 0.3, { fontSize: 13, bold: true, color: C.gold, align: 'center' });
  const [bx0, by0] = at('B0'), [bx1] = at('B1');
  T(s, '2025', bx0 - 0.4, by0 + 0.22, 0.8, 0.22, { fontSize: 9, bold: true, color: C.sub, align: 'center' });
  T(s, '2035', bx1 - 0.4, by0 + 0.22, 0.8, 0.22, { fontSize: 9, bold: true, color: C.sub, align: 'center' });
  T(s, '글로벌 제조AI 시장 — 향후 10년 연평균 40%대 성장\n(삼일PwC 「제조업, AI로 다시 설계되다」, 2026.3)', 0.4, by0 + 0.6, 2.8, 0.5, { fontSize: 7.5, color: C.sub, align: 'center' });
  const blk = [
    ['flag', '국내', ['M.AX 정책으로 2030년까지 AI 팩토리 500개, 150조 원 국민성장펀드의 40% 이상 지역 우선 투자', '과기정통부 피지컬AI 6,763억 원(2026~2030), 산업부 AX 실증산단(구미 · 포항 481억 원)', '거제 조선해양 M.AX 얼라이언스 250억 원, 부산TP 제조AI 특화 스마트공장 800억 원(2026)']],
    ['globe', '경남 · 창원', ['경남도 마스터플랜 : 지원기업 1만 개, 제조AI 전환율 30%', '창원시 : 초거대 제조AI 서비스 226.9억 · PINN 제조 융합데이터 실증 320억 · 스마트그린AX 실증산단 222억 · AI 팩토리 70억 · 기계방산 특화DX 지원센터 258.6억 원(2026)', '창원국가산단 : 기계 49% · 전기전자 20% · 운송장비 8% — 지능형 기계 · 첨단 방산 · 원자력 주기기가 핵심']],
    ['shield', '시장 특성', ['실증사업 공통 구조 = 「지역 앵커기업 + 지역 대학 · TP + AI 공급기업」 컨소시엄, 현장 데이터 · 시스템 연동 · 현장 검증이 성패 좌우', '보안 요건(데이터 외부 반출 불가)으로 온프레미스 솔루션 수요 높음', '수출입 : 내수 중심, 향후 SK텔레콤 · KAI 협력 및 영풍전자 방산 수출망(폴란드 K2PL · 호주 · 미국) 연계 동반진출']],
  ];
  let y = 1.25;
  blk.forEach(([ic, h, arr], i) => {
    const hh = i === 1 ? 1.4 : 1.15;
    card(s, 3.4, y, 6.1, hh, { hl: i === 1 });
    icon(s, ic, 3.5, y + 0.1, 0.38);
    T(s, h, 3.95, y + 0.13, 2, 0.28, { fontSize: 10.5, bold: true, color: C.title });
    T(s, bullets(arr), 3.95, y + 0.42, 5.45, hh - 0.47, { fontSize: 7.8 });
    y += hh + 0.1;
  });
}

// =============== 12. 주요 수요처 ===============
{
  const s = body('Ⅲ. 사업성 ② — 등대고객과 실증 현장은 {거점 반경 내}에 있다', '창원 · 사천 · 거제 · 포항 · 부산 — 거점 반경 내에 등대고객과 실증 현장이 집중',
    '주요 수요처입니다. 창원의 영풍전자, 세아창원특수강, 두산에너빌리티 등 피지컬AI 실증 기업, 사천의 KAI, 거제의 조선 협력사, 포항 동국제강, 그리고 발전사와 서연이화까지, 거점 반경 내에 등대고객과 실증 현장이 모여 있습니다. 경남 거점은 이들을 출장이 아니라 상주 인력으로 대응하기 위한 것입니다.');
  const at = img3d(s, 'hub', 0.4, 1.5, 4.1);
  const [cx, cy] = at('C');
  s.addShape(pptx.ShapeType.roundRect, { x: cx - 0.75, y: cy + 0.12, w: 1.5, h: 0.3, rectRadius: 0.05, fill: { color: C.navy }, line: { color: C.gold2, width: 0.75 } });
  T(s, '경남 거점 (창원)', cx - 0.75, cy + 0.12, 1.5, 0.3, { fontSize: 9, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
  const reg = ['창원 국가산단', '사천', '거제', '포항', '부산 · 전국'];
  reg.forEach((r, i) => {
    const [px, py] = at('P' + i);
    const w = 0.22 + r.length * 0.11;
    s.addShape(pptx.ShapeType.roundRect, { x: px - w / 2, y: py + 0.2, w, h: 0.24, rectRadius: 0.04, fill: { color: 'FFFFFF' }, line: { color: C.line, width: 0.5 } });
    T(s, r, px - w / 2, py + 0.2, w, 0.24, { fontSize: 8, bold: true, color: C.title, align: 'center', valign: 'middle' });
  });
  const rows = [
    ['영풍전자㈜', '방산 전자회로기판(PBA)', 'PBA 고장탐구 암묵지AI Agent · 검사Agent (~2027.07), 협력사 200여 개사 확산', '창원'],
    ['한국항공우주산업㈜', '항공기', '고정익 · 회전익 PHM DB 구축 및 데이터 처리(2022.12~2028.12), 협력사 확산', '사천'],
    ['㈜세아창원특수강', '특수강', '프레스 펌프AI 이상탐지(2024), 피지컬AI 실증 참여기업', '창원'],
    ['두산에너빌리티 · 범한메카텍 · CTR 등', '발전설비 · 기계 · 뿌리산업', '피지컬AI 세부3 물리데이터 품질관리 실증(~2030), 경남 뿌리기업 13개사', '창원 · 경남'],
    ['한화오션 · 삼성중공업 1차 협력사', '조선 블록 · 용접 · 의장', '용접WPS · 트러블슈팅 지식 어시스턴트, 선급 규격 대조', '거제'],
    ['동국제강㈜ 및 포항 철강 밸류체인', '철강', '선도공장 예지보전 · 암묵지 Agent, 연주 공정 온톨로지(~2029)', '포항'],
    ['한국남부 · 서부 · 동서발전, 서연이화', '발전 · 자동차부품', 'AI Agent · 업무자료AI 자산화 · AIOps · LLM 인프라(2024~2026)', '부산 · 전국'],
  ];
  s.addTable([[th('수요처'), th('주생산품'), th('수요 내용'), th('소재지')], ...rows.map(r => r.map((c, j) => ({ text: c, options: { bold: j === 0, color: j === 0 ? C.title : C.body, align: j === 3 ? 'center' : 'left' } })))],
    { x: 4.7, y: 1.22, ...tableOpts([1.5, 0.95, 1.85, 0.5], { fontSize: 7.3, rowH: 0.47 }) });
}

// =============== 12-2. 접근성 맵 (v7 자료 재사용) ===============
{
  const s = body('Ⅲ. 사업성 ③ — 창원(봉암동) 거점에서 {2시간 이내} 방산 · 국가산단 접근', '경남인공지능혁신본부(창원 봉암동) 기준 자차 이동시간 30분 단위 권역  (출처 : 경남(창원) 지사 설립 방안, 2026.09)',
    '거점 접근성입니다. 경남인공지능혁신본부가 있는 창원 봉암동을 기준으로, 창원 국가산단 방산 클러스터는 10~15분, KAI 사천과 거제 한화오션, 부산, 포항 철강 국가산단까지 모두 2시간 이내입니다. 서울에서 출장으로 대응하던 현장을 당일 왕복 권역으로 바꾸는 것이 입주의 핵심 효과입니다.');
  s.addImage({ path: 'img/map.png', x: 0.95, y: 1.12, w: 8.1, h: 8.1 * 1880 / 3700 });
}

// =============== 13. 경쟁 · 가격 ===============
{
  const s = body('Ⅲ. 사업성 ④ — 경쟁사는 창원에 왔다, 차별화는 {국산 Full-Stack}', '대형 경쟁사의 창원 진출 속에서 「국산 Full-Stack · 온프레미스 · 합리적 비용」으로 차별화',
    '동종업체 현황입니다. 마키나락스, 메가존, LG CNS 등 대형 기업들이 2026년 창원에 거점을 개소하거나 설립 예정입니다. 당사는 가격 측면에서 토큰 과금 없는 온프레미스 라이선스, 엣지 GPU 1~2기 수준의 운영 비용, 고효율 RAG, 단계형 패키지와 정부 매칭, 그리고 GS 1등급 인증 국산 Full-Stack으로 공공 수의계약이 가능하다는 점에서 경쟁력이 있습니다.');
  const rows = [
    ['마키나락스', '설비 · 공정 데이터 기반 산업AI(런웨이)', '-', '서울 / 2026.08 「AI스튜디오@창원」 개소'],
    ['메가존(메가존클라우드)', '클라우드 · AI 시스템 통합', '-', '서울 / 2026.08 경남대 피지컬AX연구협력센터 개소'],
    ['LG CNS · LG AI연구원', '엑사원LLM, 제조DX', '-', '서울 / 피지컬AI 협력기업, LG전자 연구센터 창원대 설립 예정'],
    ['㈜코난테크놀로지 (당사)', '자체LLM · RAG · 온톨로지/지식그래프 · 온프레미스 생성형AI', '33,979백만원\n(2025)', '서울 본사 / 경남 창원 거점 신설 (본 신청)'],
  ];
  s.addTable([[th('업체명'), th('주생산품'), th('매출액 규모'), th('소재지 · 경남권 활동')], ...rows.map((r, i) => r.map((c, j) => ({ text: c, options: { bold: j === 0, color: i === 3 ? C.gold : (j === 0 ? C.title : C.body), fill: { color: i === 3 ? 'FBF5E6' : 'FFFFFF' }, align: j === 2 ? 'center' : 'left' } })))],
    { x: 0.5, y: 1.22, ...tableOpts([1.9, 2.75, 1.15, 3.2], { fontSize: 8, rowH: 0.36 }) });
  const pr = [['transfer', '외산 · 클라우드LLM 대비', '토큰 과금 없는 온프레미스 라이선스, 사용량 증가에 따른 비용 폭증 · 기밀 유출 리스크 비용 없음'],
    ['chip', '「토큰은 많이, 파라미터는 작게」', '32B 이하 고효율 모델로 엣지 GPU 1~2기 수준 인프라 운영, 중소 · 중견기업 도입 가능'],
    ['server', '고성능 · 고효율 RAG', '2 Core 서버 시간당 5GB 인덱싱 · 초당 50건, 타 엔진 대비 1.5~10배 서버당 처리량'],
    ['calendar', '단계형 패키지', 'Quick Win(소규모 PoC) → Standard(통합 플랫폼) → 카세트형 모듈 구독(SaaS), 정부 실증 매칭'],
    ['shield', '국산 Full-Stack', '외산 모델 라이선스 · 로열티 없음, GS 1등급 인증 제품으로 공공 · 공기업 수의계약(조달) 가능']];
  sec(s, '가격 경쟁력', 0.5, 3.2, 3);
  pr.forEach(([ic, h, d], i) => {
    const x = 0.5 + i * 1.82;
    card(s, x, 3.52, 1.72, 1.55, { hl: i === 1 });
    icon(s, ic, x + 0.1, 3.6, 0.36);
    T(s, h, x + 0.1, 4.0, 1.55, 0.3, { fontSize: 8.5, bold: true, color: C.title });
    T(s, d, x + 0.1, 4.3, 1.55, 0.75, { fontSize: 7.3 });
  });
}

// =============== 14. 추진 체계 · 단계별 계획 ===============
{
  const s = body('Ⅳ. 추진계획 ① — {2026.11} 1인 개설 → {2027.7} 제조AX기술센터 전환', '1단계 「경남사무소」 → 2단계 「(가칭) 제조AX기술센터@경남」 지점 등기 전환 (2027.7 목표)',
    '추진 계획입니다. 2026년 4분기에 1인 경남사무소로 개설하고, 2027년 6월 성과 검증 후 7월에 5인 체제 제조AX기술센터로 지점 등기 전환합니다. 전환 조건은 유상계약 누적 3.7억 원, 파이프라인 15억 원, 지역 공모 제안 2건, 등대고객 3개사 PoC 중 3개 충족입니다. 2027년 거점 증분 매출 8.5억, 2028년 18억 원이 목표입니다.');
  const at = img3d(s, 'steps', 0.35, 1.25, 3.4);
  const st = [['2026 4Q', '개설 (1인)'], ['2027', '확장 (5인)'], ['2028', '정착 (5인+)']];
  st.forEach(([a, b], i) => { let [x, y] = at('F' + i); if (!i) x += 0.1; T(s, a, x - 0.5, y - 0.2, 1.0, 0.22, { fontSize: i ? 10 : 8.5, bold: true, color: i === 2 ? 'F3D49B' : 'FFFFFF', align: 'center' }); T(s, b, x - 0.5, y + 0.03, 1.0, 0.2, { fontSize: i ? 8 : 7, color: 'D6E2EF', align: 'center' }); });
  card(s, 0.5, 3.95, 3.25, 1.1, { hl: true, fill: 'FFFFFF' });
  T(s, '2단계 전환 조건 (2027.6 성과 검증, 3개 충족)', 0.62, 4.02, 3.05, 0.25, { fontSize: 8.5, bold: true, color: C.gold });
  T(s, bullets(['유상 계약 누적 3.7억 원', '거점 발굴 파이프라인 15억 원', '지역 소재 요건 공모 제안 2건', '등대고객 3개사 PoC 진입']), 0.62, 4.3, 3.05, 0.75, { fontSize: 8 });
  const ph = [
    ['2026년 4분기 — 개설 (1인)', ['경남인공지능혁신본부 4동 703호 입주, 거점장(영업/PM) 부임', '유상 PoC 2건 계약(매출 인식 2027), 창원대 국가연구소 연계 2건(AI 지식서비스 · 도면이해)', '정부 실증 3건 현장 착수 요건 확인(SOP 30건 · 숙련자 인터뷰 · 데이터 반출 협의 · 방산 보안 절차)', '경남도 · 창원시 · 경남TP · 거제시 접점 구축, 등대고객 진단']],
    ['2027년 — 확장 (5인)', ['1분기 수석 엔지니어 파견, 연내 현지 채용 3인(엔지니어 2인은 피지컬AI 과제 청년 채용)', '영풍전자 PBA 고장탐구AI · 검사Agent 완성(~2027.07, 분석시간 50% · 보고서 70% 단축 KPI 검증)', '피지컬AI 3D 공정 시나리오 자동 생성 · Sim2Real 정합(뿌리기업 13개사), 포항 모듈 2차년도', '거점 증분 매출 8.5억 원(유상 PoC 5건 2.5억 + 본구축 2건 6억), 창원 · 김해 · 진해 · 거제 확산']],
    ['2028년 — 정착 (5인+)', ['거점 증분 매출 18억 원 목표 (PoC · 본구축 · 신규 공모)', '실증 결과의 카세트형 모듈 · 구독형 SaaS 표준 패키지로 경남 제조기업 수평 확산', '포항 본격 진입(분소 검토), 여수 · 광양(철강 · 화학) 탐색']],
  ];
  let y = 1.22;
  ph.forEach(([h, arr], i) => {
    const hh = i === 2 ? 0.95 : 1.33;
    card(s, 3.95, y, 5.55, hh, { hl: i === 1 });
    T(s, h, 4.08, y + 0.07, 5.3, 0.25, { fontSize: 10, bold: true, color: i === 2 ? C.gold : C.title });
    T(s, bullets(arr), 4.08, y + 0.34, 5.3, hh - 0.38, { fontSize: 7.6 });
    y += hh + 0.08;
  });
}

// =============== 15. Quick Win · 기술개발 · 지역협력 ===============
{
  const s = body('Ⅳ. 추진계획 ② — Quick Win A · B로 {20주 안에 숫자로 증명}', 'Quick Win A · B로 20주 안에 숫자로 증명 — 한 번 만든 코어를 업종 · 지역으로 수평 확산',
    '시장 공략 방안입니다. Quick Win A는 12주 동안 숙련자 노하우·SOP 지식 어시스턴트로 정확도 85%, 탐색시간 40% 단축을 목표로 하고, Quick Win B는 7주차에 병행 착수해 20주 안에 원청 규격·도면 대조 검토를 추출 재현율 90%, 검토시간 50% 단축으로 검증합니다. 착수 2주 안에 현행 기준선을 실측하고 같은 방법으로 재측정합니다. 가드레일은 틀릴 때 틀렸다고 말하는 설계입니다.');
  // gantt
  const gx = 0.5, gw = 5.4, wk = gw / 20;
  T(s, '0주', gx, 1.22, 0.4, 0.2, { fontSize: 7, color: C.sub });
  [7, 12, 20].forEach(w => T(s, `${w}주`, gx + w * wk - 0.2, 1.22, 0.4, 0.2, { fontSize: 7, color: C.sub, align: 'center' }));
  s.addShape(pptx.ShapeType.line, { x: gx, y: 1.45, w: gw, h: 0, line: { color: C.line, width: 0.75 } });
  const bars = [['Quick Win A  |  숙련자 노하우 · SOP 지식 어시스턴트', 0, 12, C.navy, 'SOP · 작업표준 · 숙련자 노하우 · 이력을 근거로 출처와 함께 답하는 현장 Q&A', '정확도 85% 이상 · 탐색시간 40% 단축 · 반장 문의 30% 감소 · 활성률 60% 이상'],
    ['Quick Win B  |  원청 규격 · 도면 대조 검토', 6, 20, 'B8860B', '원청 도면 · 사양서와 사내 검사기준 자동 대조, 최종 판단은 담당자 승인', '추출 재현율 90% 이상 · 탐지율 85% 이상 · 오탐 30% 이하 · 검토시간 50% 단축']];
  bars.forEach(([n, a, b, col, d, k], i) => {
    const y = 1.6 + i * 1.05;
    s.addShape(pptx.ShapeType.roundRect, { x: gx + a * wk, y, w: (b - a) * wk, h: 0.3, rectRadius: 0.05, fill: { color: col }, line: { color: col } });
    T(s, n, gx + a * wk + 0.08, y, (b - a) * wk, 0.3, { fontSize: 8.5, bold: true, color: 'FFFFFF', valign: 'middle' });
    T(s, d, gx, y + 0.36, gw, 0.22, { fontSize: 8 });
    T(s, '목표 KPI  ' + k, gx, y + 0.6, gw, 0.22, { fontSize: 8, bold: true, color: i ? C.gold : C.title });
  });
  T(s, '※ 착수 2주 내 현행 기준선 실측 후 12주차 동일 방법으로 재측정', gx, 3.7, gw, 0.2, { fontSize: 7.5, color: C.sub });
  card(s, 0.5, 3.98, 5.4, 1.07);
  icon(s, 'warning', 0.6, 4.07, 0.36);
  T(s, '가드레일 — 「틀릴 때 틀렸다고 말하는 설계」를 상품 가치로', 1.02, 4.08, 4.8, 0.3, { fontSize: 9.5, bold: true, color: C.title });
  T(s, '출처 필수(문서명 · 개정일 · 페이지), 추측 금지, 안전 질문은 규정 원문 인용, 확신도 분리, 저신뢰 항목 「미확인」 강제', 1.02, 4.4, 4.8, 0.6, { fontSize: 8 });
  // right column
  card(s, 6.1, 1.22, 3.4, 1.55, { hl: true });
  T(s, '공통 코어 + 업종 · 지역 확장 모듈', 6.22, 1.3, 3.2, 0.26, { fontSize: 9.5, bold: true, color: C.title });
  T(s, bullets(['창원 (기계 · 방산) : KS · ASTM · KDS, 재질-공구-절삭', '사천 (항공) : AS9100 · NADCAP, 복합재 · FAI · 개정 추적', '거제 (조선) : 선급 KR · ABS · DNV, 강종-두께-자세-WPS', '한 번 만든 코어를 수평 확산']), 6.22, 1.6, 3.2, 1.15, { fontSize: 7.6 });
  card(s, 6.1, 2.87, 3.4, 1.2);
  T(s, '기술개발 계획 (거점 연계)', 6.22, 2.95, 3.2, 0.26, { fontSize: 9.5, bold: true, color: C.title });
  T(s, bullets(['온톨로지 반자동 구축 · Graph-RAG 하이브리드 검색 → 영풍전자 · KAI · 뿌리기업 검증', '멀티모달 문서AI · VLM 품질 · 안전 상황인지 → 거점 폐쇄망 워크스테이션', '엣지 sLLM · PINN + 생성형AI 자율제조 → 창원대 · 경남대 · ETRI · KTL']), 6.22, 3.23, 3.2, 0.84, { fontSize: 7.3 });
  card(s, 6.1, 4.17, 3.4, 0.88);
  T(s, '지역 협력', 6.22, 4.24, 3.2, 0.26, { fontSize: 9.5, bold: true, color: C.title });
  T(s, '국립창원대(SMR² 국가연구소 · 영풍전자 과제) · 경남대 · 경남TP · 중소조선연구원 · 부산TP 컨소시엄, 창원대 국가연구소 연계 2건(5~10억 원 규모) 추진', 6.22, 4.5, 3.2, 0.52, { fontSize: 7.3 });
}

// =============== 16. 매출 · 판매전략 ===============
{
  const s = body('Ⅴ. 매출 · 판매전략 — 거점 증분 매출 {2027년 8.5억 → 2028년 18억}', '정부 실증 → 카탈로그 등재 → 앵커기업 협력사 수평 확산',
    '매출 실적과 판매 계획입니다. 전년도 339.8억 원, 올해 330억 원, 내년 400억 원을 추정합니다. 경남 거점 증분 매출은 2027년 8.5억, 2028년 18억 원이 목표입니다. 판매 경로는 정부 실증 컨소시엄으로 진입하고, 실증 결과를 카세트형 모듈과 구독형 SaaS로 카탈로그에 등재한 뒤, 영풍전자 협력사 200여 개사, KAI·조선 1차 협력사로 수평 확산하는 3단계입니다.');
  const at = img3d(s, 'bars', 0.4, 1.35, 3.6);
  const lb = [['33,979', '전년도 (2025)'], ['33,000', '당해년 (추정)'], ['40,000', '차년도 (추정)']];
  lb.forEach(([v, l], i) => { const [x, y] = at('T' + i); T(s, v, x - 0.55, y - 0.3, 1.1, 0.24, { fontSize: 10.5, bold: true, color: i === 2 ? C.gold : C.title, align: 'center' }); T(s, l, x - 0.55, 4.15, 1.1, 0.2, { fontSize: 7.5, color: C.sub, align: 'center' }); });
  T(s, '내수 매출 (백만원) · 수출 0 · 회사 전체 매출 기준', 0.4, 4.38, 3.6, 0.2, { fontSize: 7.5, color: C.sub, align: 'center' });
  card(s, 0.5, 4.62, 3.5, 0.45, { hl: true, fill: 'FFFFFF' });
  T(s, '경남 거점 증분 매출 목표   2027년 8.5억 원  →  2028년 18억 원', 0.6, 4.62, 3.3, 0.45, { fontSize: 8.5, bold: true, color: C.gold, valign: 'middle', align: 'center' });
  // right: 판매경로 flow
  sec(s, '판매 경로', 4.3, 1.22, 3);
  const fl = [['①', '정부 실증사업 컨소시엄', '지역 앵커기업 + 대학 · TP + 당사'], ['②', '카탈로그 등재', '카세트형 모듈 · 구독형 SaaS (AX 실증산단)'], ['③', '앵커기업 협력사 수평 확산', '영풍전자 협력사 200여 개사, KAI · 조선 1차 협력사']];
  fl.forEach(([n, h, d], i) => {
    const x = 4.3 + i * 1.75;
    s.addShape(pptx.ShapeType.homePlate, { x, y: 1.52, w: 1.68, h: 0.92, fill: { color: i === 2 ? 'B8860B' : (i ? C.steel : C.navy) }, line: { color: 'FFFFFF', width: 0 } });
    T(s, `${n} ${h}`, x + 0.1, 1.56, 1.35, 0.42, { fontSize: 8, bold: true, color: 'FFFFFF', valign: 'middle' });
    T(s, d, x + 0.1, 1.96, 1.35, 0.44, { fontSize: 7, color: 'E1EBF7' });
  });
  const st = [['판로 확보', ['경남TP 입주기업 · 경남도 제조AI 특화기업 생태계 참여, 2027 제조AI · 스마트공장 공모(부산TP 400개 과제, 거제 M.AX 250억 원 등) 지역 거점 기반 대응', '스마트제조혁신 공급기업 풀, 경남 뿌리기업 13개사 실증 네트워크를 영업 채널로 활용', '국립창원대(국가연구소 기업 멤버십 25개사 · 원전 부품기업 340여 개사) · 경남대 · 중소조선연구원 · RIST · KTL 산학연 공동 제안']],
    ['차별화', ['지식 · 암묵지 · 온톨로지 계층 특화, 온프레미스 · 폐쇄망 운영(방산 · 발전 레퍼런스), GS 1등급 · AI+ 인증 국산 Full-Stack', '현장 상주 인력의 신속 대응 — 출장 의존 대비 착수 · 검증 리드타임 단축']],
    ['가격 · 계약', ['Quick Win 소규모 PoC → Standard 통합 구축 → 구독형 모듈 단계형 패키지, 정부 매칭(스마트공장 50% 등)으로 초기 도입 부담 완화']]];
  let y = 2.58;
  st.forEach(([h, arr], i) => {
    const hh = [1.05, 0.72, 0.6][i];
    card(s, 4.3, y, 5.2, hh);
    T(s, h, 4.4, y + 0.06, 0.9, 0.25, { fontSize: 9, bold: true, color: C.title });
    T(s, bullets(arr), 5.3, y + 0.06, 4.12, hh - 0.1, { fontSize: 7.2 });
    y += hh + 0.07;
  });
}

// =============== 17. 청년고용 · 노동시간 ===============
{
  const s = body('Ⅵ. 청년고용 · 노동시간 — 지역 청년을 {제조AI 전문인력}으로', '현지 채용 → 본사 4~8주 교육 → 창원 배치 : 지역 청년을 제조AI 전문인력으로',
    '청년고용입니다. 당사는 AI 인력 139명, 청년 비율 34%로 청년 고용 비중이 높은 기업입니다. 입주 후 2027년 현지 채용 3명 중 엔지니어 2명은 피지컬AI 과제의 청년채용 인력으로 최소 4년간 인건비가 충당되며, 현장 PM 1명을 더합니다. 창원대, 경남대, 창원시 이노베이션 아카데미 인재풀에서 우선 채용합니다. 노동시간 측면에서는 주 52시간 상한을 준수하고, 출장 기반 대응을 상주 거점으로 전환해 장거리 이동시간을 줄입니다.');
  const at = img3d(s, 'phases', 0.45, 1.3, 4.1);
  const [ax, ay] = at('A'), [bx, by] = at('B');
  T(s, '2026 입주 시  1명', ax - 0.9, ay + 0.02, 1.8, 0.26, { fontSize: 10, bold: true, color: C.title, align: 'center' });
  T(s, '거점장 (영업 / PM)', ax - 0.9, ay + 0.27, 1.8, 0.22, { fontSize: 8, color: C.sub, align: 'center' });
  T(s, '2027  5명', bx - 0.9, by + 0.02, 1.8, 0.26, { fontSize: 10, bold: true, color: C.gold, align: 'center' });
  T(s, '파견 2 · 현지 채용 3', bx - 0.9, by + 0.27, 1.8, 0.22, { fontSize: 8, color: C.sub, align: 'center' });
  card(s, 0.5, 3.55, 4.1, 1.5, { hl: true });
  T(s, '청년고용 현황', 0.62, 3.62, 3, 0.26, { fontSize: 10, bold: true, color: C.title });
  T(s, bullets(['청년고용 우수업체 지정 : 해당 없음', '전사 194명(정규 188 · 계약 6), AI 인력 139명 — 소프트웨어 개발 중심의 청년 고용 비중 높은 기업', '청년(만15~34세) 비율 34% (66명)']), 0.62, 3.92, 3.9, 1.1, { fontSize: 8 });
  card(s, 4.85, 1.22, 4.65, 2.2);
  icon(s, 'people', 4.95, 1.3, 0.38);
  T(s, '입주 후 청년고용 계획', 5.4, 1.35, 3.5, 0.26, { fontSize: 10, bold: true, color: C.title });
  T(s, bullets(['2027년 : 현지 채용 3명 — 엔지니어 2명(과기정통부 피지컬AI 과제 청년채용, 인건비 2027년부터 최소 4년 과제 충당) + 현장 PM/영업 1명', '채용 풀 : 국립창원대(2027년 원자력기계공학과 신설) · 경남대(피지컬AX연구협력센터 · RISE AX 인재양성) · 창원시 이노베이션 아카데미(2031년까지 AI 인재 1,200명) · 경남TP 인재풀', '2028년 : 제조AX기술센터 5인 이상 확대 시 추가 1~2명(누적 4~5명)', '인턴 · 산학 : 창원대 · 경남대 재학생 현장실습 · 인턴십, 피지컬AI · 영풍전자 과제 참여']), 5.0, 1.68, 4.4, 1.72, { fontSize: 7.6 });
  card(s, 4.85, 3.52, 4.65, 1.53);
  icon(s, 'calendar', 4.95, 3.6, 0.38);
  T(s, '노동시간 단축 이행계획', 5.4, 3.65, 3.5, 0.26, { fontSize: 10, bold: true, color: C.title });
  T(s, bullets(['현황 : 주 40시간 법정근로 준수, 유연근로제(시차출퇴근제)로 주 52시간 상한 준수', '출장 기반 대응을 상주 거점으로 전환 — 서울↔창원 주 2~3회 출장의 이동시간 단축, 파견자 주 1.5회 본사 왕복 교통비 · 주거 지원', '본사(플랫폼 · 모델 개발) / 거점(현장 데이터 · 검증 · 고객 대면) 분업으로 야간 · 휴일 작업 억제, SP 2등급 표준 프로세스 적용']), 5.0, 4.02, 4.4, 1.02, { fontSize: 7.3 });
}

// =============== 18. 소요자금 ===============
{
  const s = body('Ⅶ. 소요자금 — 3개년 {1,294백만원}, 외부 차입 없이 조달', '3개년 소요자금 1,294백만원 — 정책자금(과제비) 520 + 본사 자기자금 774  (단위 : 백만원)',
    '소요자금 조달계획입니다. 1차년도 2027년 4.18억 원, 2028년과 2029년 각각 4.38억 원이 소요되며, 운전자금이 가장 큰 비중입니다. 조달은 피지컬AI 과제비로 현지 엔지니어 2인 인건비를 충당하고, 나머지는 본사 자기자금으로 충당합니다. 2026년 상반기 본사 자기자본 총계는 295억 원 수준으로 조달 여력이 충분합니다.');
  const yrs = [['1차년도 (2027)', { 사업장: 32, 유형고정자산: 31, 운전자금: 355 }, 418, 120, 298, '사업장: 운영 22 · 시설 6 · 사무실임대 4 / 유형고정자산: 차량운반구 31 / 운전자금: 기타운전자금 355'],
    ['2차년도 (2028)', { 사업장: 26, 유형고정자산: 31, 운전자금: 381 }, 438, 200, 238, ''],
    ['3차년도 (2029)', { 사업장: 26, 유형고정자산: 31, 운전자금: 381 }, 438, 200, 238, '']];
  s.addChart(pptx.ChartType.bar, ['사업장', '유형고정자산', '운전자금'].map(k => ({ name: k, labels: yrs.map(y => y[0]), values: yrs.map(y => y[1][k]) })), {
    x: 0.4, y: 1.2, w: 4.4, h: 3.0, barDir: 'col', barGrouping: 'stacked', chartColors: ['E2B23A', '3C6791', '0E2240'],
    showLegend: true, legendPos: 'b', legendFontSize: 8, legendFontFace: F, catAxisLabelFontSize: 8.5, catAxisLabelFontFace: F, valAxisHidden: true, valGridLine: { style: 'none' },
    showValue: true, dataLabelFontSize: 8, dataLabelFontFace: F, dataLabelColor: 'FFFFFF', barGapWidthPct: 70 });
  yrs.forEach((y, i) => T(s, `합계 ${y[2]}`, 0.95 + i * 1.33, 1.15, 1.2, 0.24, { fontSize: 9.5, bold: true, color: C.gold, align: 'center' }));
  const rows = yrs.map(y => [y[0], String(y[2]), String(y[3]), String(y[4]), String(y[2])]);
  s.addTable([[th('구분'), th('소요자금'), th('정책자금지원\n(과제비)'), th('B/S상 가용자금\n(본사 자기자금)'), th('조달 합계')],
    ...rows.map(r => r.map((c, j) => ({ text: c, options: { align: 'center', bold: j === 0 || j === 4, color: j === 4 ? C.gold : C.body } }))),
    ['3개년 합계', '1,294', '520', '774', '1,294'].map((c, j) => ({ text: c, options: { align: 'center', bold: true, color: C.title, fill: { color: C.pale } } }))],
  { x: 5.0, y: 1.22, ...tableOpts([1.2, 0.75, 0.85, 1.0, 0.7], { fontSize: 8.5, rowH: 0.36 }) });
  card(s, 5.0, 3.25, 4.5, 1.8, { hl: true });
  icon(s, 'books', 5.1, 3.33, 0.38);
  T(s, '조달 근거', 5.55, 3.38, 3, 0.26, { fontSize: 10, bold: true, color: C.title });
  T(s, bullets(['1차년도 정책자금 120 : 피지컬AI 과제비 — 현지 엔지니어 2인 인건비', '2 · 3차년도 정책자금 각 200 (과제비)', 'B/S상 가용자금 : 본사 자기자금 (2026년 상반기 자기자본 총계 29,558백만원)', '1차년도 내역 — 운영자금 22 · 시설자금 6 · 사무실임대 4 · 차량운반구 31 · 기타운전자금 355', '외부 차입 · 회사채 · VC 투자 계획 없음']), 5.1, 3.7, 4.3, 1.33, { fontSize: 7.8 });
  T(s, '※ 3개년 합계는 사업계획서 연도별 금액의 단순 합산', 0.5, 4.98, 4.3, 0.2, { fontSize: 7, color: C.sub });
}

// =============== CLOSING ===============
{
  const s = pptx.addSlide(); pageNo++;
  s.background = { path: 'img/end_bg.jpg' };
  T(s, 'THANK YOU', 0.6, 1.25, 4, 0.3, { fontSize: 10.5, color: C.gold2, charSpacing: 3 });
  T(s, rich('감사{합니다}', 'FFFFFF', C.gold2), 0.6, 1.6, 6, 1.0, { fontSize: 46, bold: true, valign: 'middle' });
  s.addShape(pptx.ShapeType.rect, { x: 0.6, y: 2.75, w: 0.6, h: 0.03, fill: { color: C.gold2 }, line: { color: C.gold2, width: 0 } });
  T(s, [{ text: '"제조의 지식을 가장 깊이 이해하는 AI 기업"', options: { color: 'FFFFFF', bold: true } }, { text: '\n그 첫걸음은 경남(창원)의 현장에서 시작합니다', options: { color: 'B7C3D3' } }], 0.6, 2.95, 5.5, 0.6, { fontSize: 11, lineSpacingMultiple: 1.2 });
  T(s, '㈜코난테크놀로지', 0.6, 4.05, 4.9, 0.3, { fontSize: 13, color: 'FFFFFF', bold: true });
  T(s, [{ text: '염종학 이사 (국방 · 전략사업부)   ' }, { text: 'jonghak.yeom@konantech.com', options: { hyperlink: { url: 'mailto:jonghak.yeom@konantech.com' }, color: 'C0D1E6' } }, { text: '   010-9082-1522' }], 0.6, 4.4, 7, 0.28, { fontSize: 10, color: 'C0D1E6' });
  T(s, '경남인공지능혁신본부 입주신청 · 사업계획 발표  |  2026. 10', 0.6, 4.95, 6, 0.22, { fontSize: 8, color: '8E99A5' });
  s.addNotes('감사합니다. 질문 받겠습니다.');
}

pptx.writeFile({ fileName: 'out/Konan_GyeongnamAI_Tenancy_Pitch_15min.pptx' }).then(f => console.log('saved', f));
