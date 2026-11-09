const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, PageBreak,
  VerticalAlign, HeightRule, TableLayoutType,
} = require("docx");

const FONT = "맑은 고딕";
const OUT = process.argv[2];

// A4 = 11906 x 16838 DXA, margins 2.0cm (1134) sides, 2.5cm (1418) top/bottom
const CONTENT_W = 11906 - 1134 * 2; // 9638

const thin = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const thick = { style: BorderStyle.SINGLE, size: 12, color: "000000" };
const cellBorders = { top: thin, bottom: thin, left: thin, right: thin };

const run = (text, opts = {}) => new TextRun({ text, font: FONT, size: 20, ...opts });
const para = (children, opts = {}) =>
  new Paragraph({ children: Array.isArray(children) ? children : [children], spacing: { before: 0, after: 0 }, ...opts });
const empty = (n = 1) => Array.from({ length: n }, () => para([run("")]));

function cell(children, width, opts = {}) {
  const { shade, span, vMerge, align } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders,
    columnSpan: span,
    verticalMerge: vMerge,
    verticalAlign: VerticalAlign.CENTER,
    shading: shade ? { type: ShadingType.CLEAR, fill: shade, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: (Array.isArray(children) ? children : [children]).map((c) =>
      c instanceof Paragraph ? c : para(c, { alignment: align || AlignmentType.CENTER })),
  });
}
const HEAD_FILL = "D9D9D9";
const label = (text, width, span) => cell(run(text, { bold: true }), width, { shade: HEAD_FILL, span });
const blank = (width, span) => cell(run(""), width, { span });
const hint = (text) => run(text, { italics: true, color: "595959", size: 18 });

// ---------- Page 1: Cover ----------
const cover = [
  ...empty(9),
  para(run("도입기업명", { bold: true, size: 32 }), { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 360 } }),
  new Table({
    alignment: AlignmentType.CENTER,
    width: { size: 8400, type: WidthType.DXA },
    columnWidths: [8400],
    layout: TableLayoutType.FIXED,
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 8400, type: WidthType.DXA },
      borders: { top: thin, bottom: thin, left: thin, right: thin },
      shading: { type: ShadingType.CLEAR, fill: HEAD_FILL, color: "auto" },
      margins: { top: 240, bottom: 240, left: 200, right: 200 },
      children: [
        para(run("AI응용제품 신속상용화 지원사업", { size: 32 }), { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 } }),
        para(run("변경관리내역서", { bold: true, size: 44 }), { alignment: AlignmentType.CENTER }),
      ],
    })] })],
  }),
  ...empty(11),
  para(run("202  .    .", { size: 32 }), { alignment: AlignmentType.CENTER }),
  ...empty(11),
  para(run("공급기업명", { size: 44 }), { alignment: AlignmentType.CENTER }),
  para([new PageBreak()]),
];

// ---------- Page 2: 변경관리대장 ----------
const ledgerW = [600, 1700, 2200, 900, 900, 1000, 1538, 800]; // sum 9638
const ledgerHeads = ["번호", "변경요청 제목", "변경요청 공문번호", "요청\n일자", "요청자", "회신번호", "회신내용", "회신\n일자"];
const headCell = (t, w) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders: cellBorders, verticalAlign: VerticalAlign.CENTER,
  shading: { type: ShadingType.CLEAR, fill: HEAD_FILL, color: "auto" },
  margins: { top: 40, bottom: 40, left: 60, right: 60 },
  children: t.split("\n").map((line) => para(run(line, { bold: true }), { alignment: AlignmentType.CENTER })),
});
const ledgerRows = [
  new TableRow({ tableHeader: true, children: ledgerHeads.map((t, i) => headCell(t, ledgerW[i])) }),
];
for (let r = 0; r < 8; r++) {
  ledgerRows.push(new TableRow({
    height: { value: 360, rule: HeightRule.ATLEAST },
    children: ledgerW.map((w, i) =>
      r === 0 && i === 2 ? cell(hint("발송된 공문번호를 기록"), w) : blank(w)),
  }));
}
const ledger = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: ledgerW, layout: TableLayoutType.FIXED, rows: ledgerRows,
});

// ---------- Page 2: 변경요청서 ----------
const reqW = [1900, 2919, 1900, 2919]; // sum 9638
const pair = (l1, l2) => new TableRow({
  height: { value: 520, rule: HeightRule.ATLEAST },
  children: [label(l1, reqW[0]), blank(reqW[1]), label(l2, reqW[2]), blank(reqW[3])],
});
const wide = (l, hintText, h) => new TableRow({
  height: { value: h, rule: HeightRule.ATLEAST },
  children: [label(l, reqW[0]), cell(hintText ? hint(hintText) : run(""), reqW[1] + reqW[2] + reqW[3], { span: 3 })],
});
const request = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: reqW, layout: TableLayoutType.FIXED,
  rows: [
    pair("사업명", "도입기업명"),
    pair("과제번호", "공급기업명"),
    pair("변경요청 제목", "요청 일자"),
    wide("변경요청 사유", "변경 사유에 대해 객관적인 자료를 첨부할 필요가 있음", 1500),
    wide("변경요청 내용", "변경 전 및 후 내용을 명확히 제시할 필요가 있음.", 1500),
    wide("파급효과", "변경에 따른 영향도 및 파급 효과를 명확히 제시할 필요가 있음.", 900),
    wide("비고", "", 520),
  ],
});

// ---------- 작성 방법 ----------
const BLUE = "1F4EA8";
const note = (text, indent = 0, hanging = 0) =>
  para(run(text, { color: BLUE, size: 18 }), { indent: { left: indent, hanging }, spacing: { before: 0, after: 40 } });
const notes = [
  para(run("※ 작성 방법", { bold: true, color: BLUE, size: 20 }), { spacing: { before: 240, after: 60 } }),
  note("- 변경관리대장은 수행기업 등이 일정, 인력뿐만 아니라 요구사항, 장비 규격 등에 대한 변경 사항이 발생될 경우 변경 사유, 승인 여부, 관련 공문, 관련 증빙자료를 작성 관리한다.", 200, 200),
  note("1. 변경관리대장 : 문서 목차로 승인 요청한 문서 정보 및 승인(회신)된 문서 정보를 기술함.", 560, 280),
  note("2. 사업변경신청서(변경요청서)", 560, 280),
  note("3. 사업변경 사유 및 내용(도입기업, 공급기업이 서명한 회의록), 관련 증빙자료", 560, 280),
  note("4. 전담(운영)기관 승인 공문(사업관리시스템에서 캡처한 승인 또는 확인 화면 증적)", 560, 280),
];

const title = (t) => para(run(t, { bold: true, size: 36 }), { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 } });

const doc = new Document({
  creator: "Konan Technology",
  title: "AI응용제품 신속상용화 지원사업 변경관리내역서",
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1134, right: 1134 } } },
    children: [
      ...cover,
      title("변경관리대장"),
      ledger,
      ...empty(1),
      title("변경요청서"),
      request,
      ...notes,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("written", OUT, buf.length); });
