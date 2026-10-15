// 2027 사업계획 사전조사 (제조 AX) — docx 렌더러
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, HeadingLevel, LevelFormat, ShadingType, BorderStyle, PageBreak,
  VerticalAlign, TableLayoutType, Footer, PageNumber, Header,
} = require('docx');

const FONT = '맑은 고딕';
const NAVY = '1F3864';
const GRAY = '595959';
const LIGHT = 'EEF2F8';
const CONTENT_W = 11906 - 1440 - 1440; // 9026 DXA

const TOC_PAGES = process.env.TOC_PAGES ? JSON.parse(process.env.TOC_PAGES) : {};
const content = [
  ...require('./ch1.js'),
  ...require('./ch2.js'),
  ...require('./ch3.js'),
  ...require('./ch4.js'),
];

// ---- inline markup: **bold** ----
function runs(text, base = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  return parts.map((seg) => {
    const m = seg.match(/^\*\*([^*]+)\*\*$/);
    return new TextRun({ text: m ? m[1] : seg, bold: m ? true : base.bold, font: FONT, size: base.size || 20, color: base.color, italics: base.italics });
  });
}

function heading(level, text) {
  const sizes = { 1: 26, 2: 23, 3: 21 };
  const hl = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 }[level];
  return new Paragraph({
    heading: hl,
    numbering: { reference: 'toc', level: level - 1 },
    spacing: { before: level === 1 ? 160 : level === 2 ? 110 : 80, after: 30 },
    keepNext: true,
    pageBreakBefore: level === 1 && !process.env.NO_H1_BREAK,
    children: [new TextRun({ text, bold: true, font: FONT, size: sizes[level], color: level === 1 ? NAVY : '000000' })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 100, line: 276 },
    alignment: AlignmentType.BOTH,
    indent: { left: opts.indent == null ? 0 : opts.indent },
    children: runs(text, opts),
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bul', level },
    spacing: { after: 24, line: 245 },
    alignment: AlignmentType.BOTH,
    children: runs(text, { size: 18 }),
  });
}

function note(text) {
  return new Paragraph({
    spacing: { before: 40, after: 120 },
    children: runs('※ ' + text, { size: 17, color: GRAY }),
  });
}

function box(text) {
  // 음영 박스(핵심 요약)
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: LIGHT, color: 'auto' },
    border: {
      left: { style: BorderStyle.SINGLE, size: 18, color: NAVY, space: 6 },
    },
    indent: { left: 120, right: 120 },
    spacing: { before: 80, after: 160, line: 276 },
    alignment: AlignmentType.BOTH,
    children: runs(text, { size: 18 }),
  });
}

function cell(text, { head = false, width, align } = {}) {
  const lines = String(text).split('\n');
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: head ? { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' } : undefined,
    margins: { top: 30, bottom: 30, left: 60, right: 60 },
    children: lines.map((ln) => new Paragraph({
      alignment: head ? AlignmentType.CENTER : (align || AlignmentType.LEFT),
      spacing: { after: 0, line: 230 },
      children: runs(ln, { size: 15, bold: head ? true : undefined, color: head ? 'FFFFFF' : undefined }),
    })),
  });
}

function table(spec) {
  const total = spec.cols.reduce((a, b) => a + b, 0);
  const widths = spec.cols.map((c) => Math.round((c / total) * CONTENT_W));
  const diff = CONTENT_W - widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += diff;
  const border = { style: BorderStyle.SINGLE, size: 4, color: '8A99B3' };
  const rows = [];
  if (spec.head) rows.push(new TableRow({ tableHeader: true, cantSplit: true, children: spec.head.map((h, i) => cell(h, { head: true, width: widths[i] })) }));
  spec.rows.forEach((r) => rows.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, { width: widths[i], align: spec.align && spec.align[i] })) })));
  return [
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: widths,
      layout: TableLayoutType.FIXED,
      borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
      rows,
    }),
    new Paragraph({ spacing: { after: 20 }, children: [] }),
  ];
}

function render(items) {
  const out = [];
  for (const it of items) {
    if (it.h) out.push(heading(it.h, it.t));
    else if (it.p != null) out.push(para(it.p, it.o));
    else if (it.b) it.b.forEach((x) => out.push(typeof x === 'string' ? bullet(x) : bullet(x.t, x.l == null ? 1 : x.l)));
    else if (it.t) out.push(...table(it.t));
    else if (it.note) out.push(note(it.note));
    else if (it.box) out.push(box(it.box));
    else if (it.raw) out.push(...it.raw);
    else if (it.pb) out.push(new Paragraph({ children: [new PageBreak()] }));
  }
  return out;
}

function tocPage() {
  const { PositionalTab, PositionalTabAlignment, PositionalTabLeader, PositionalTabRelativeTo } = require('docx');
  const out = [new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: '목 차', font: FONT, size: 28, bold: true })] })];
  let n1 = 0, n2 = 0; const ga = '가나다라마바사아자차카타파하';
  for (const it of content) {
    if (!it.h || it.h > 2) continue;
    let label;
    if (it.h === 1) { n1++; n2 = 0; label = `${n1}. ${it.t}`; } else { n2++; label = `${ga[n2 - 1]}. ${it.t}`; }
    const key = `${it.h}:${it.t}`; const pg = TOC_PAGES[key] || '';
    const { TabStopType, LeaderType } = require('docx');
    out.push(new Paragraph({ spacing: { after: it.h === 1 ? 60 : 20, before: it.h === 1 ? 120 : 0 }, indent: { left: it.h === 1 ? 0 : 400 },
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W, leader: LeaderType.DOT }],
      children: [new TextRun({ text: label, font: FONT, size: it.h === 1 ? 21 : 19, bold: it.h === 1 }),
        new TextRun({ text: '\t' + String(pg), font: FONT, size: 19 })] }));
  }
  out.push(new Paragraph({ children: [new PageBreak()] }));
  return out;
}

// ---- title block (양식 준용) ----
function titleBlock() {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80 }, children: [new TextRun({ text: '2027년 사업계획 사전조사 보고서 (사업부용)', font: FONT, size: 32, bold: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: '제조 AX 분야 — KONAN MAESTRO · 영남권 제조 AX 거점', font: FONT, size: 22, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: '국방·전략사업부', font: FONT, size: 20 })] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 200 }, children: [new TextRun({ text: '2026. 10. 15.', font: FONT, size: 22 })] }),
  ];
}

const doc = new Document({
  creator: '국방·전략사업부',
  title: '2027년 사업계획 사전조사 보고서 (제조 AX)',
  styles: {
    default: { document: { run: { font: FONT, size: 20 }, paragraph: { spacing: { after: 100, line: 276 } } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 26, bold: true, color: NAVY }, paragraph: { outlineLevel: 0, spacing: { before: 360, after: 100 } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 23, bold: true }, paragraph: { outlineLevel: 1, spacing: { before: 240, after: 100 } } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 21, bold: true }, paragraph: { outlineLevel: 2, spacing: { before: 160, after: 100 } } },
    ],
  },
  numbering: {
    config: [
      {
        reference: 'toc',
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 425, hanging: 425 } } } },
          { level: 1, format: LevelFormat.GANADA, text: '%2.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 850, hanging: 425 } } } },
          { level: 2, format: LevelFormat.DECIMAL_FULL_WIDTH, text: '%3)', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1275, hanging: 425 } } } },
        ],
      },
      {
        reference: 'bul',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '□', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 320, hanging: 300 } } } },
          { level: 1, format: LevelFormat.BULLET, text: 'ㅇ', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 640, hanging: 280 } } } },
          { level: 2, format: LevelFormat.BULLET, text: '-', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 920, hanging: 220 } } } },
        ],
      },
    ],
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1701, right: 1440, bottom: 1440, left: 1440, header: 851, footer: 992 } },
    },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: '2027년 사업계획 사전조사 — 제조 AX 분야 · 국방·전략사업부 · 내부 보고용', font: FONT, size: 16, color: GRAY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GRAY })] })] }) },
    children: [...titleBlock(), ...render(require('./ch0.js')), ...tocPage(), ...render(content.filter((it) => !it._front))],
  }],
});

const outPath = path.resolve(__dirname, '../out/2027_BizPlan_PreSurvey_MfgAX_v1.3.docx');
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(outPath, buf); console.log('written', outPath, buf.length); });
