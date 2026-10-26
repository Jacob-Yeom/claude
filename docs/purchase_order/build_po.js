const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  ShadingType, BorderStyle, VerticalAlign } = require('docx');

const FONT = '맑은 고딕', NAVY = '1F2A44', LIGHT = 'E8ECF3', HINT = '8A8F99';
const W = 9638;
const bd = { style: BorderStyle.SINGLE, size: 4, color: '7F8796' };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: o.size || 18, bold: o.bold, color: o.color, italics: o.italics });
const para = (items, o = {}) => new Paragraph({ alignment: o.align || AlignmentType.LEFT, spacing: { before: o.before || 0, after: o.after || 0, line: 276 },
  children: (Array.isArray(items) ? items : [items]).map(x => typeof x === 'string' ? run(x, o) : x) });
// 셀: text 배열 허용, hint=true면 회색 안내문
function cell(text, width, o = {}) {
  const lines = Array.isArray(text) ? text : [text];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders, columnSpan: o.span, rowSpan: o.rowSpan,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    children: lines.map(l => para(l, { align: o.align, bold: o.bold, color: o.color || (o.hint ? HINT : undefined), italics: o.hint, size: o.size })),
  });
}
const head = (t, w, o = {}) => cell(t, w, { fill: NAVY, color: 'FFFFFF', bold: true, align: AlignmentType.CENTER, ...o });
const label = (t, w, o = {}) => cell(t, w, { fill: LIGHT, bold: true, align: AlignmentType.CENTER, ...o });
const table = (cols, rows) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cols, rows });
const section = t => new Paragraph({ spacing: { before: 240, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
  children: [run(t, { bold: true, size: 22, color: NAVY })] });

// ── 1. 발주 기본정보
const c4 = [1500, 3319, 1500, 3319];
const info = table(c4, [
  new TableRow({ children: [label('발주번호', c4[0]), cell('KT-PO-2026-____', c4[1]), label('발주일자', c4[2]), cell('2026년    월    일', c4[3])] }),
  new TableRow({ children: [label('건    명', c4[0]), cell('Synap DocuAnalyzer 소프트웨어 라이선스 구매', c4[1], { span: 3 })] }),
  new TableRow({ children: [label('사용목적', c4[0]), cell('(적용 사업/과제명 기재)', c4[1], { span: 3, hint: true })] }),
]);

// ── 2. 수신(공급자) / 발주자
const c5 = [700, 1300, 2819, 1300, 3519];
const kv = (k, v, hint) => [label(k, 1300), cell(v, hint === 'w' ? 3519 : 2819, { hint: !!hint && v.startsWith('(') })];
const party = table([700, 1300, 3019, 1300, 3319], [
  new TableRow({ children: [head('구분', 700), head('공급자 (수신)', 4319, { span: 2 }), head('발주자 (발신)', 4619, { span: 2 })] }),
  ...[
    ['상    호', '주식회사 사이냅소프트', '주식회사 코난테크놀로지'],
    ['대 표 자', '(기재)', '(기재)'],
    ['사업자번호', '(기재)', '(기재)'],
    ['주    소', '(기재)', '(기재)'],
    ['담 당 자', '(성명 / 부서)', '염종학 이사 / 국방·전략사업부'],
    ['연 락 처', '(전화 / 이메일)', '(전화) / konan_claude33@konantech.com'],
  ].map(([k, a, b], i) => new TableRow({ children: [
    ...(i === 0 ? [label(['업체', '정보'], 700, { rowSpan: 6 })] : []),
    label(k, 1300), cell(a, 3019, { hint: a.startsWith('(') }),
    label(k, 1300), cell(b, 3319, { hint: b.startsWith('(') }),
  ] })),
]);

// ── 3. 발주 내역
const ci = [500, 1900, 2238, 800, 600, 1200, 1400, 1000];
const itemHead = new TableRow({ tableHeader: true, children:
  ['No', '품명', '규격 / 내용', '단위', '수량', '단가(원)', '금액(원)', '비고'].map((t, i) => head(t, ci[i])) });
const items = [
  ['1', 'Synap DocuAnalyzer', ['소프트웨어 라이선스', '(버전·라이선스 유형 기재)'], 'Copy', '', '', '', ''],
  ['2', 'DocuAnalyzer 유지보수', ['기술지원·업데이트', '(기간: 12개월)'], '식', '', '', '', ''],
  ['3', '설치 및 기술지원', ['구축·연동 지원', '(Man-Day 기재)'], '식', '', '', '', ''],
  ['4', '', '', '', '', '', '', ''],
].map(r => new TableRow({ children: r.map((t, i) => cell(t, ci[i], {
  align: [0, 3, 4].includes(i) ? AlignmentType.CENTER : [5, 6].includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
})) }));
// 규격 셀의 괄호 안내문은 회색 처리
const sumW = ci.slice(0, 6).reduce((a, b) => a + b, 0);
const sums = [['공급가액', ''], ['부가가치세 (10%)', ''], ['합계금액 (VAT 포함)', '']].map(([k], i) =>
  new TableRow({ children: [label(k, sumW, { span: 6, fill: i === 2 ? 'D3DAE6' : LIGHT }), cell('', ci[6], { align: AlignmentType.RIGHT }), cell('', ci[7])] }));
const itemTable = table(ci, [itemHead, ...items, ...sums]);
const amountLine = para([run('합계금액 :  금 ', { bold: true, size: 20 }), run('                                              ', { size: 20 }),
  run('원정 (₩                          , VAT 포함)', { bold: true, size: 20 })], { before: 120 });

// ── 4. 거래 조건
const ct = [1900, 7738];
const terms = table(ct, [
  ['납품기한', '2026년    월    일까지'],
  ['납품장소', '(기재) 주식회사 코난테크놀로지 지정 장소'],
  ['납품물', '라이선스 증서(키), 설치 파일, 설치·운영·API 매뉴얼, 제품 사양서'],
  ['라이선스 범위', '(기재) 사용 기간(영구/기간제), 서버·CPU·코어 수, 사용 사업 범위, 재판매·납품 허용 여부'],
  ['검수조건', '납품 후 발주자 검수(설치 및 기능 확인) 완료 시 납품 완료로 간주'],
  ['대금지급조건', '(기재) 검수 완료 후 세금계산서 수령일로부터 __일 이내 현금 지급'],
  ['하자보증', '(기재) 검수 완료일로부터 __개월 무상 하자보수'],
  ['기타', '본 발주서에 명시되지 않은 사항은 상호 협의하여 정하며, 별도 계약서가 있는 경우 계약서가 우선한다.'],
].map(([k, v]) => new TableRow({ children: [label(k, ct[0]), cell(v, ct[1], { hint: v.startsWith('(') })] })));

// ── 5. 서명
const sign = [
  para('위와 같이 발주합니다.', { align: AlignmentType.CENTER, before: 360, size: 20 }),
  para('2026년        월        일', { align: AlignmentType.CENTER, before: 200, size: 20 }),
  para([run('발주자 :  주식회사 코난테크놀로지', { bold: true, size: 22 })], { align: AlignmentType.RIGHT, before: 280 }),
  para([run('대표이사                                      (인)', { size: 22 })], { align: AlignmentType.RIGHT, before: 120 }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 18 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [run('발    주    서', { bold: true, size: 40, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [run('PURCHASE ORDER', { size: 18, color: HINT })] }),
      para([run('수신 :  주식회사 사이냅소프트  귀중', { bold: true, size: 22 })], { after: 40 }),
      section('1. 발주 개요'), info,
      section('2. 공급자 및 발주자'), party,
      section('3. 발주 내역'), itemTable, amountLine,
      section('4. 거래 조건'), terms,
      ...sign,
    ],
  }],
});
Packer.toBuffer(doc).then(b => fs.writeFileSync(process.argv[2], b));
