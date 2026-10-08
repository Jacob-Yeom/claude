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

const content = [
  ...require('./ch0.js'),
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
    spacing: { before: level === 1 ? 360 : level === 2 ? 240 : 160, after: 100 },
    keepNext: true,
    pageBreakBefore: level === 1,
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
    spacing: { after: 60, line: 276 },
    alignment: AlignmentType.BOTH,
    children: runs(text),
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
    children: runs(text, { size: 19 }),
  });
}

function cell(text, { head = false, width, align } = {}) {
  const lines = String(text).split('\n');
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: head ? { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: lines.map((ln) => new Paragraph({
      alignment: head ? AlignmentType.CENTER : (align || AlignmentType.LEFT),
      spacing: { after: 0, line: 240 },
      children: runs(ln, { size: 17, bold: head ? true : undefined, color: head ? 'FFFFFF' : undefined }),
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
    new Paragraph({ spacing: { after: 60 }, children: [] }),
  ];
}

function render(items) {
  const out = [];
  for (const it of items) {
    if (it.h) out.push(heading(it.h, it.t));
    else if (it.p != null) out.push(para(it.p, it.o));
    else if (it.b) it.b.forEach((x) => out.push(typeof x === 'string' ? bullet(x) : bullet(x.t, x.l || 1)));
    else if (it.t) out.push(...table(it.t));
    else if (it.note) out.push(note(it.note));
    else if (it.box) out.push(box(it.box));
    else if (it.raw) out.push(...it.raw);
    else if (it.pb) out.push(new Paragraph({ children: [new PageBreak()] }));
  }
  return out;
}

// ---- title block (양식 준용) ----
function titleBlock() {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80 }, children: [new TextRun({ text: '2027년 사업계획 사전조사 보고서 (사업부용)', font: FONT, size: 32, bold: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: '제조 AX 분야 — KONAN MAESTRO · 영남권 제조 AX 거점', font: FONT, size: 22, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: '국방·전략사업부', font: FONT, size: 20 })] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 200 }, children: [new TextRun({ text: '2026. 10. 08.', font: FONT, size: 22 })] }),
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
          { level: 0, format: LevelFormat.BULLET, text: 'ㅇ', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '-', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 760, hanging: 240 } } } },
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
    children: [...titleBlock(), ...render(content)],
  }],
});

const outPath = path.resolve(__dirname, '../out/2027_BizPlan_PreSurvey_MfgAX_v1.0.docx');
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(outPath, buf); console.log('written', outPath, buf.length); });
