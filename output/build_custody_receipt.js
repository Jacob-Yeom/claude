// 보관증 (HW 납품 보관증) 생성 스크립트 — docx (npm)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, VerticalAlign,
  Header, Footer, PageNumber, PageBreak, LevelFormat, TabStopType,
} = require("docx");

const OUT = process.argv[2] || "custody_receipt_YoungPoong_HW.docx";

// ---------- 공통 스타일 ----------
const FONT = { ascii: "맑은 고딕", hAnsi: "맑은 고딕", eastAsia: "맑은 고딕", cs: "맑은 고딕" };
const NAVY = "1F2A44";
const NAVY_LIGHT = "E8ECF4";
const GRAY_LIGHT = "F3F4F6";
const BORDER_GRAY = "9CA3AF";

const PAGE_W = 11906, PAGE_H = 16838; // A4 DXA
const MARGIN = 1000;                    // 약 1.76cm
const CONTENT_W = PAGE_W - MARGIN * 2;  // 9906

const border = (color = BORDER_GRAY, size = 4) => ({ style: BorderStyle.SINGLE, size, color });
const allBorders = (color, size) => ({
  top: border(color, size), bottom: border(color, size), left: border(color, size), right: border(color, size),
});
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function run(text, opt = {}) {
  return new TextRun({ text, font: FONT, size: opt.size ?? 20, bold: opt.bold, color: opt.color, italics: opt.italics });
}
function para(text, opt = {}) {
  const runs = Array.isArray(text) ? text : [run(text, opt)];
  return new Paragraph({
    children: runs,
    alignment: opt.align ?? AlignmentType.LEFT,
    spacing: { before: opt.before ?? 0, after: opt.after ?? 0, line: opt.line ?? 276 },
    indent: opt.indent,
    numbering: opt.numbering,
    keepNext: opt.keepNext,
  });
}
function cell(content, opt = {}) {
  const children = Array.isArray(content)
    ? content
    : [para(content, { size: opt.size ?? 18, bold: opt.bold, color: opt.color, align: opt.align ?? AlignmentType.LEFT, line: 252 })];
  return new TableCell({
    children,
    width: { size: opt.width, type: WidthType.DXA },
    columnSpan: opt.span,
    rowSpan: opt.rowSpan,
    verticalAlign: opt.vAlign ?? VerticalAlign.CENTER,
    shading: opt.fill ? { type: ShadingType.CLEAR, fill: opt.fill, color: "auto" } : undefined,
    margins: { top: opt.pad ?? 60, bottom: opt.pad ?? 60, left: 90, right: 90 },
    borders: opt.borders,
  });
}
const labelCell = (text, width, opt = {}) =>
  cell(text, { width, fill: NAVY_LIGHT, bold: true, align: AlignmentType.CENTER, size: opt.size ?? 18, span: opt.span, rowSpan: opt.rowSpan });
const headCell = (text, width, span) =>
  cell(text, { width, span, fill: NAVY, bold: true, color: "FFFFFF", align: AlignmentType.CENTER, size: 17 });
const blank = (n = 1) => Array.from({ length: n }, () => para("", { size: 8, line: 200 }));

// ---------- 데이터: 견적서(SWSYUN260910-01) HW 내역 ----------
const items = [
  { no: 1, name: "추론서버", spec: "Dell PowerEdge R770 / Xeon 6 6520P 2.4GHz 24C ×2 / 256GB / 480GB SSD ×2 / 1.2TB HDD ×6 / 4×1GbE / RTX PRO 6000 96GB ×2", qty: 1, unit: "대" },
  { no: 2, name: "RAG서버", spec: "Dell PowerEdge R770 / Xeon 6 6515P 2.3GHz 16C ×2 / 128GB / 480GB SSD ×2 / 1.2TB HDD ×6 / 4×1GbE / RTX PRO 6000 96GB ×1", qty: 1, unit: "대" },
  { no: 3, name: "OCR서버", spec: "Dell PowerEdge R770 / Xeon 6 6515P 2.3GHz 16C ×2 / 128GB / 480GB SSD ×2 / 1.2TB HDD ×6 / 4×1GbE / RTX PRO 6000 96GB ×1", qty: 1, unit: "대" },
  { no: 4, name: "On-Tology서버", spec: "Dell PowerEdge R660xs / Xeon Silver 4524Y 2.0GHz 16C ×1 / 32GB / 480GB SSD ×2 / 1.2TB HDD ×4 / 4×1GbE", qty: 1, unit: "대" },
  { no: 5, name: "LLMOps관리서버", spec: "Dell PowerEdge R660xs / Xeon Silver 4524Y 2.0GHz 16C ×1 / 32GB / 480GB SSD ×2 / 1.2TB HDD ×4 / 4×1GbE", qty: 1, unit: "대" },
  { no: 6, name: "WEB/WAS서버", spec: "Dell PowerEdge R660xs / Xeon Silver 4524Y 2.0GHz 16C ×1 / 32GB / 480GB SSD ×2 / 1.2TB HDD ×4 / 4×1GbE", qty: 1, unit: "대" },
  { no: 7, name: "K8s 등 통합서버", spec: "Dell PowerEdge R660xs / Xeon Silver 4524Y 2.0GHz 16C ×1 / 32GB / 480GB SSD ×2 / 1.2TB HDD ×4 / 4×1GbE", qty: 1, unit: "대" },
  { no: 8, name: "DB서버", spec: "Dell PowerEdge R660xs / Xeon Silver 4524Y 2.0GHz 16C ×1 / 32GB / 480GB SSD ×2 / 1.2TB HDD ×4 / 4×1GbE", qty: 1, unit: "대" },
  { no: 9, name: "STT서버", spec: "Dell PowerEdge R660xs / Xeon Silver 4524Y 2.0GHz 16C ×1 / 32GB / 480GB SSD ×2 / 1.2TB HDD ×4 / 4×1GbE", qty: 1, unit: "대" },
  { no: 10, name: "L4 Switch", spec: "Radware Alteon 1G 8Port L4 Switch", qty: 1, unit: "대" },
  { no: 11, name: "Network Switch", spec: "HPE Aruba CX 6000 24G, 24Port 1GbE Switch", qty: 1, unit: "대" },
  { no: 12, name: "Rack / KVM", spec: "42U Rack ×1, 2×32A 3상 PDU(Defog) ×1식, IP-KVM 16Port(ATEN) ×1", qty: 1, unit: "식" },
];

// 붙임: 세부 사양 (견적서 구성 항목 기준)
const detailGroups = [
  {
    title: "A. 추론서버 — Dell PowerEdge R770 (1대)",
    rows: [
      ["CPU", "Intel Xeon 6 6520P 2.4GHz 24C/48T 144MB, 210W", "2"],
      ["Memory", "256GB (8× 32GB DDR5 6400MT/s RDIMM)", "1"],
      ["GPU", "NVIDIA RTX PRO 6000 Blackwell Server Edition 96GB, PCIe", "2"],
      ["Disk-1", "480GB SSD SATA 6Gbps", "2"],
      ["Disk-2", "1.2TB SAS HDD 12Gbps 10K rpm", "6"],
      ["NIC", "Broadcom 5719 4Port 1GbE-RJ45, OCP NIC", "1"],
      ["Power Supply", "3200W Redundant(1+1)", "2"],
      ["Support", "OS 설치(Ubuntu), 장애 시 익일방문 서비스, H/W 보증 36개월", "1"],
    ],
  },
  {
    title: "B. RAG서버 · OCR서버 — Dell PowerEdge R770 (각 1대, 총 2대)",
    rows: [
      ["CPU", "Intel Xeon 6 6515P 2.3GHz 16C/32T 72MB, 150W", "2"],
      ["Memory", "128GB (8× 16GB DDR5 6400MT/s RDIMM)", "1"],
      ["GPU", "NVIDIA RTX PRO 6000 Blackwell Server Edition 96GB, PCIe", "1"],
      ["Disk-1", "480GB SSD SATA 6Gbps", "2"],
      ["Disk-2", "1.2TB SAS HDD 12Gbps 10K rpm", "6"],
      ["NIC", "Broadcom 5719 4Port 1GbE-RJ45, OCP NIC", "1"],
      ["Power Supply", "3200W Redundant(1+1)", "2"],
      ["Support", "OS 설치(Ubuntu), 장애 시 익일방문 서비스, H/W 보증 36개월", "1"],
    ],
  },
  {
    title: "C. On-Tology · LLMOps관리 · WEB/WAS · K8s 등 통합 · DB · STT 서버 — Dell PowerEdge R660xs (각 1대, 총 6대)",
    rows: [
      ["CPU", "Intel Xeon Silver 4524Y 2.0GHz 16C/32T 30MB, 150W", "1"],
      ["Memory", "32GB DDR5 5600MT/s RDIMM", "1"],
      ["GPU", "N/A", "-"],
      ["Disk-1", "480GB SSD SATA 6Gbps", "2"],
      ["Disk-2", "1.2TB SAS HDD 12Gbps 10K rpm", "4"],
      ["NIC", "Broadcom 5719 4Port 1GbE-RJ45, OCP NIC", "1"],
      ["Power Supply", "800W Redundant(1+1)", "2"],
      ["Support", "OS 설치(Ubuntu), 장애 시 익일방문 서비스, H/W 보증 36개월", "1"],
    ],
  },
  {
    title: "D. 네트워크 · 랙 장비",
    rows: [
      ["L4 Switch", "Radware Alteon 1G 8Port L4 Switch (보증 1년)", "1"],
      ["Network Switch", "HPE Aruba CX 6000 24G, 24Port 1GbE (보증 1년)", "1"],
      ["Rack", "42U Rack (보증 1년)", "1"],
      ["PDU", "2× 32A 3상 PDU, Defog (보증 1년)", "1식"],
      ["IP-KVM", "ATEN IP-KVM 16Port (보증 1년)", "1"],
    ],
  },
];

// ---------- 표 빌더 ----------
function infoTable() {
  const W = [1500, 3453, 1500, 3453];
  const r = (l1, v1, l2, v2) => new TableRow({
    children: [labelCell(l1, W[0]), cell(v1, { width: W[1] }), labelCell(l2, W[2]), cell(v2, { width: W[3] })],
  });
  const rFull = (l, v) => new TableRow({
    children: [labelCell(l, W[0]), cell(v, { width: W[1] + W[2] + W[3], span: 3 })],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: W,
    borders: allBorders(BORDER_GRAY, 4),
    rows: [
      r("보관증 번호", "KT-YP-HW-2026-____", "작성일자", "2026년    월    일"),
      rFull("프로젝트명", "영풍전자 과제용 서버 시스템 구축"),
      rFull("관련 계약(발주)", "계약(발주)번호: ______________________   /   계약(발주)일자: 20    .    .    ."),
      rFull("보관 장소", "영풍전자 ________________ (건물 · 층 · 전산실/서버실 호수)"),
      rFull("보관 기간", "20    .    .    .  ~  20    .    .    .   (설치 및 검수 완료 시까지)"),
      rFull("보관 사유", "☐ 설치 전 임시보관     ☐ 검수 대기     ☐ 기타 (                          )"),
    ],
  });
}

function partyTable() {
  const W = [1300, 3653, 1300, 3653];
  const row = (label, v1, v2) => new TableRow({
    children: [labelCell(label, W[0]), cell(v1, { width: W[1] }), labelCell(label, W[2]), cell(v2, { width: W[3] })],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: W,
    borders: allBorders(BORDER_GRAY, 4),
    rows: [
      new TableRow({
        children: [
          headCell("공  급  자 (납품)", W[0] + W[1], 2),
          headCell("확  인  자 (보관)", W[2] + W[3], 2),
        ],
      }),
      row("회 사 명", "(주)코난테크놀로지", "영풍전자"),
      row("사업자번호", "", ""),
      row("주    소", "", ""),
      row("부서/직위", "", ""),
      row("담 당 자", "", ""),
      row("연 락 처", "", ""),
    ],
  });
}

function itemTable() {
  // NO | 품명 | 규격 | 수량 | 단위 | 제조번호(S/N) | 비고
  const W = [480, 1500, 4076, 560, 560, 1650, 1080];
  const header = new TableRow({
    tableHeader: true,
    children: [
      headCell("NO", W[0]), headCell("품    명", W[1]), headCell("규격 (모델 · 주요 사양)", W[2]),
      headCell("수량", W[3]), headCell("단위", W[4]), headCell("제조번호 (S/N)", W[5]), headCell("비고", W[6]),
    ],
  });
  const rows = items.map((it) => new TableRow({
    cantSplit: true,
    children: [
      cell(String(it.no), { width: W[0], align: AlignmentType.CENTER, size: 17 }),
      cell(it.name, { width: W[1], size: 17, bold: true }),
      cell(it.spec, { width: W[2], size: 16 }),
      cell(String(it.qty), { width: W[3], align: AlignmentType.CENTER, size: 17 }),
      cell(it.unit, { width: W[4], align: AlignmentType.CENTER, size: 17 }),
      cell("", { width: W[5], size: 17 }),
      cell("", { width: W[6], size: 17 }),
    ],
  }));
  const total = new TableRow({
    cantSplit: true,
    children: [
      cell("합    계", { width: W[0] + W[1] + W[2], span: 3, fill: GRAY_LIGHT, bold: true, align: AlignmentType.CENTER, size: 18 }),
      cell("12", { width: W[3], fill: GRAY_LIGHT, bold: true, align: AlignmentType.CENTER, size: 18 }),
      cell("종", { width: W[4], fill: GRAY_LIGHT, bold: true, align: AlignmentType.CENTER, size: 18 }),
      cell("서버 9대, 스위치 2대, 랙/KVM 1식", { width: W[5] + W[6], span: 2, fill: GRAY_LIGHT, size: 15, align: AlignmentType.CENTER }),
    ],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: W,
    borders: allBorders(BORDER_GRAY, 4),
    rows: [header, ...rows, total],
  });
}

function signTable() {
  const W = [1400, 3553, 1400, 3553];
  const row = (label, v1, v2, opt = {}) => new TableRow({
    height: opt.h ? { value: opt.h, rule: "atLeast" } : undefined,
    children: [labelCell(label, W[0]), cell(v1, { width: W[1], ...opt }), labelCell(label, W[2]), cell(v2, { width: W[3], ...opt })],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: W,
    borders: allBorders(BORDER_GRAY, 4),
    rows: [
      new TableRow({
        children: [
          headCell("공  급  자 (인도)", W[0] + W[1], 2),
          headCell("확  인  자 (보관 · 인수)", W[2] + W[3], 2),
        ],
      }),
      row("회 사 명", "(주)코난테크놀로지", "영풍전자", { bold: true }),
      row("직위 / 성명", "                         /                      (인)", "                         /                      (인)", { h: 700 }),
      row("서명일자", "2026년      월      일", "2026년      월      일"),
    ],
  });
}

function detailTable(group) {
  const W = [1700, 7106, 1100];
  const header = new TableRow({
    tableHeader: true,
    children: [headCell("구  성", W[0]), headCell("사    양", W[1]), headCell("수량(대당)", W[2])].map((c) => c),
  });
  const rows = group.rows.map(([k, v, q]) => new TableRow({
    cantSplit: true,
    children: [
      cell(k, { width: W[0], fill: NAVY_LIGHT, bold: true, size: 17, align: AlignmentType.CENTER, pad: 30 }),
      cell(v, { width: W[1], size: 17, pad: 30 }),
      cell(q, { width: W[2], size: 17, align: AlignmentType.CENTER, pad: 30 }),
    ],
  }));
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: W,
    borders: allBorders(BORDER_GRAY, 4),
    rows: [header, ...rows],
  });
}

function sectionTitle(text) {
  return new Paragraph({
    children: [run(text, { size: 22, bold: true, color: NAVY })],
    spacing: { before: 240, after: 100 },
    keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 2 } },
  });
}

// ---------- 본문 ----------
const conditions = [
  "본 보관증은 상기 물품이 공급자로부터 확인자에게 인도되어 확인자의 관리 하에 보관 중임을 증명한다. 인도 시 양 당사자는 수량·외관·제조번호(S/N)를 상호 확인하였다.",
  "물품의 소유권 이전 시점은 계약(발주)서에서 정한 바에 따르며, 계약서에 별도의 정함이 없는 경우 검수 완료 시 확인자에게 이전된다.",
  "확인자는 선량한 관리자의 주의로써 물품을 보관하며, 보관 중 확인자의 고의 또는 과실로 인하여 발생한 멸실·훼손·도난에 대하여 책임을 진다. 단, 천재지변 등 불가항력으로 인한 손해는 양 당사자가 협의하여 처리한다.",
  "확인자는 물품을 지정된 보관 장소 외로 반출·이동하거나 임의로 개봉·설치·개조하지 않으며, 필요한 경우 공급자와 사전에 협의한다.",
  "확인자는 서버·GPU 등 정밀 전산장비의 특성을 고려하여 실내의 적정 온도·습도, 출입 통제 등 보관 환경을 유지한다.",
  "설치·검수 착수 또는 보관기간 만료 시, 양 당사자 입회 하에 물품의 수량·외관·제조번호를 재확인한 후 인수인계한다.",
  "본 보관증은 물품의 인도 및 보관 사실을 확인하는 문서로서, 검수 결과는 별도의 검수확인서에 따른다.",
  "본 보관증은 2부를 작성하여 공급자와 확인자가 서명(날인) 후 각 1부씩 보관한다.",
];

const body = [
  // 제목
  new Paragraph({
    children: [run("보  관  증", { size: 44, bold: true, color: NAVY })],
    alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
  }),
  new Paragraph({
    children: [run("(HW 납품 물품 보관 확인서)", { size: 20, color: "4B5563" })],
    alignment: AlignmentType.CENTER, spacing: { after: 240 },
  }),
  infoTable(),
  ...blank(),
  sectionTitle("1. 당사자"),
  partyTable(),
  ...blank(),
  sectionTitle("2. 보관 물품 내역"),
  para("※ 제조번호(S/N)는 인도 시 실물 확인 후 기재한다. 각 물품의 세부 구성 사양은 붙임 1과 같다.", { size: 16, color: "4B5563", after: 80 }),
  itemTable(),
  ...blank(),
  sectionTitle("3. 보관 조건"),
  ...conditions.map((t) => new Paragraph({
    children: [run(t, { size: 18 })],
    numbering: { reference: "cond", level: 0 },
    spacing: { after: 60, line: 264 },
  })),
  ...blank(),
  para("위와 같이 물품을 인도·인수하고 보관함을 확인합니다.", { size: 21, bold: true, align: AlignmentType.CENTER, before: 200, after: 200 }),
  signTable(),
  ...blank(),
  para("붙임  1. 보관 물품 세부 사양 1부.  끝.", { size: 18, before: 120 }),

  // ---- 붙임 1 ----
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({
    children: [run("[붙임 1]  보관 물품 세부 사양", { size: 26, bold: true, color: NAVY })],
    spacing: { after: 60 },
  }),
  para("프로젝트: 영풍전자 과제용 서버 시스템 구축   |   보관증 번호: KT-YP-HW-2026-____", { size: 17, color: "4B5563", after: 160 }),
  ...detailGroups.flatMap((g, i) => [
    new Paragraph({ children: [run(g.title, { size: 19, bold: true })], spacing: { before: i === 0 ? 0 : 160, after: 60 }, keepNext: true }),
    detailTable(g),
  ]),
  ...blank(),
  para("※ 서버(R770/R660xs) H/W 보증기간은 36개월, L4·네트워크 스위치·IP-KVM·랙은 1년이다. 케이블링/라벨링은 본 물품 내역에 포함되지 않는다.", { size: 16, color: "4B5563", before: 120 }),
];

// ---------- 문서 ----------
const doc = new Document({
  creator: "코난테크놀로지",
  title: "보관증 - 영풍전자 과제용 서버 시스템 구축 HW",
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
  },
  numbering: {
    config: [{
      reference: "cond",
      levels: [{
        level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 400 } }, run: { font: FONT, size: 18 } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, bottom: 900, left: MARGIN, right: MARGIN, header: 500, footer: 450 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [run("영풍전자 과제용 서버 시스템 구축 — HW 납품 보관증", { size: 15, color: "6B7280" })],
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB", space: 2 } },
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            run("(주)코난테크놀로지  ·  영풍전자          - ", { size: 15, color: "6B7280" }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 15, color: "6B7280" }),
            run(" / ", { size: 15, color: "6B7280" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 15, color: "6B7280" }),
            run(" -", { size: 15, color: "6B7280" }),
          ],
        })],
      }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("written:", OUT, buf.length, "bytes");
});
