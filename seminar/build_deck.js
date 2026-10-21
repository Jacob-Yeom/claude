// 제조 AX 세미나 강의 덱 (코난테크놀로지, 60분) — pptxgenjs
const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pptx.title = "Manufacturing AX Seminar - Konan Technology";

const F = "맑은 고딕";
const C = {
  bg: "0B1A2E", panel: "12263F", panel2: "183253", line: "2C4A70",
  text: "E8EEF5", muted: "9FB3CC", accent: "4CC3EA", accent2: "F2B544",
  green: "6BD4A8", dim: "6F86A3", white: "FFFFFF",
};
const W = 13.333;
let pageNo = 0;

function base(section, title, sub) {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  pageNo++;
  if (section) {
    s.addText(section, { x: 0.5, y: 0.28, w: 8, h: 0.3, fontFace: F, fontSize: 11, color: C.accent, bold: true });
  }
  if (title) {
    s.addText(title, { x: 0.5, y: 0.55, w: 12.3, h: 0.6, fontFace: F, fontSize: 24, color: C.white, bold: true });
  }
  if (sub) {
    s.addText(sub, { x: 0.5, y: 1.13, w: 12.3, h: 0.4, fontFace: F, fontSize: 13, color: C.muted });
  }
  s.addShape(pptx.ShapeType.line, { x: 0.5, y: 1.58, w: 12.33, h: 0, line: { color: C.line, width: 1 } });
  s.addText("Konan Technology  |  제조 AX 세미나 2026", { x: 0.5, y: 7.08, w: 6, h: 0.25, fontFace: F, fontSize: 8, color: C.dim });
  s.addText(String(pageNo), { x: 12.3, y: 7.08, w: 0.53, h: 0.25, fontFace: F, fontSize: 8, color: C.dim, align: "right" });
  return s;
}

// 카드: 제목 + 본문 bullet
function card(s, x, y, w, h, title, lines, o = {}) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: o.fill || C.panel }, line: { color: o.border || C.line, width: 1 } });
  if (o.bar !== false) s.addShape(pptx.ShapeType.rect, { x, y, w: 0.07, h, fill: { color: o.barColor || C.accent }, line: { color: o.barColor || C.accent, width: 0 } });
  let ty = y + 0.12;
  if (title) {
    s.addText(title, { x: x + 0.2, y: ty, w: w - 0.3, h: 0.4, fontFace: F, fontSize: o.titleSize || 14, bold: true, color: o.titleColor || C.white, valign: "top" });
    ty += o.titleH || 0.45;
  }
  if (lines && lines.length) {
    const arr = lines.map((t) => {
      if (typeof t === "object") return t;
      return { text: t, options: { bullet: o.bullet === false ? false : { code: "25AA" }, breakLine: true } };
    });
    s.addText(arr, { x: x + 0.2, y: ty, w: w - 0.32, h: y + h - ty - 0.08, fontFace: F, fontSize: o.size || 11.5, color: o.color || C.text, valign: "top", paraSpaceAfter: o.gap ?? 4, lineSpacingMultiple: 1.05 });
  }
}

// 하단 핵심 메시지 바
function keyMsg(s, text, y = 6.45) {
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.33, h: 0.5, fill: { color: C.panel2 }, line: { color: C.accent, width: 1 } });
  s.addText([{ text: "핵심  ", options: { bold: true, color: C.accent } }, { text, options: { color: C.white } }],
    { x: 0.7, y, w: 12, h: 0.5, fontFace: F, fontSize: 13, valign: "middle" });
}

function source(s, text) {
  s.addText("출처: " + text, { x: 0.5, y: 6.97, w: 12.3, h: 0.15, fontFace: F, fontSize: 7.5, color: C.dim });
}

// KPI 박스
function kpi(s, x, y, w, h, value, label, o = {}) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: o.fill || C.panel }, line: { color: o.border || C.line, width: 1 } });
  s.addText(value, { x, y: y + 0.08, w, h: h * 0.52, fontFace: F, fontSize: o.vs || 24, bold: true, color: o.vc || C.accent, align: "center", valign: "middle" });
  s.addText(label, { x: x + 0.1, y: y + h * 0.55, w: w - 0.2, h: h * 0.42, fontFace: F, fontSize: o.ls || 10.5, color: C.text, align: "center", valign: "top" });
}

function arrowR(s, x, y, w = 0.35, h = 0.3, color = C.accent) {
  s.addShape(pptx.ShapeType.rightArrow, { x, y, w, h, fill: { color }, line: { color, width: 0 } });
}

function box(s, x, y, w, h, text, o = {}) {
  s.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: o.fill || C.panel2 }, line: { color: o.border || C.line, width: 1 } });
  s.addText(text, { x: x + 0.05, y, w: w - 0.1, h, fontFace: F, fontSize: o.size || 11, color: o.color || C.text, bold: !!o.bold, align: "center", valign: "middle" });
}

const TH = (t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.panel2 } } });

// ───────────────────────── 1. 표지
{
  const s = pptx.addSlide();
  pageNo++;
  s.background = { color: C.bg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText("제조 AX 세미나  |  2026. 11.", { x: 0.9, y: 1.2, w: 10, h: 0.4, fontFace: F, fontSize: 14, color: C.accent, bold: true });
  s.addText("명장의 판단을 회사의 자산으로", { x: 0.9, y: 1.85, w: 11.5, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: C.white });
  s.addText("제조 암묵지 기반 지식베이스 구축과 제조 AX 서비스 실행 전략", { x: 0.9, y: 2.8, w: 11.5, h: 0.6, fontFace: F, fontSize: 22, color: C.text });
  s.addText("서연이화 · 영풍전자 · 포항 AX실증산단 사례 중심", { x: 0.9, y: 3.45, w: 11.5, h: 0.5, fontFace: F, fontSize: 16, color: C.muted });
  s.addShape(pptx.ShapeType.line, { x: 0.9, y: 4.4, w: 5, h: 0, line: { color: C.line, width: 1 } });
  s.addText("㈜코난테크놀로지  국방·전략사업부  이사  염종학", { x: 0.9, y: 4.6, w: 10, h: 0.4, fontFace: F, fontSize: 14, color: C.text });
  s.addNotes("[0:00~1:00] 인사. 오늘 60분 동안 '제조 암묵지'를 어떻게 회사 자산으로 바꾸는지, 실제 진행 중인 세 사례와 함께 말씀드리고, 마지막에 우리 회사는 어디부터 시작해야 하는지 숫자로 정하는 방법을 드리겠습니다.");
}

// ───────────────────────── 2. 목차
{
  const s = base("CONTENTS", "오늘 이야기할 순서", "60분 · 5개 파트");
  const items = [
    ["Ⅰ", "왜 지금 제조 AX인가", "현장의 3대 위기, 제조 AI와 제조 AX의 차이", "5분"],
    ["Ⅱ", "제조 AX 서비스의 구성요소", "5-Layer AX Stack: 데이터 · 지식 · AI 엔진 · Agent · 운영", "12분"],
    ["Ⅲ", "제조 암묵지: 개념과 구현", "명장의 '감'을 근거–조건–결과 구조로 꺼내는 방법", "15분"],
    ["Ⅳ", "구축 사례", "서연이화 · 영풍전자 · 포항 AX실증산단", "15분"],
    ["Ⅴ", "AX 도입 우선순위 결정", "정량 스코어카드 · 효과 산정 · 도입 로드맵", "10분"],
  ];
  items.forEach((it, i) => {
    const y = 1.9 + i * 0.92;
    s.addShape(pptx.ShapeType.rect, { x: 0.8, y, w: 11.7, h: 0.78, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
    s.addText(it[0], { x: 0.9, y, w: 0.9, h: 0.78, fontFace: F, fontSize: 26, bold: true, color: C.accent, align: "center", valign: "middle" });
    s.addText(it[1], { x: 1.95, y: y + 0.05, w: 6, h: 0.4, fontFace: F, fontSize: 17, bold: true, color: C.white });
    s.addText(it[2], { x: 1.95, y: y + 0.42, w: 8.5, h: 0.32, fontFace: F, fontSize: 11.5, color: C.muted });
    s.addText(it[3], { x: 10.9, y, w: 1.4, h: 0.78, fontFace: F, fontSize: 14, color: C.accent2, align: "right", valign: "middle" });
  });
  s.addText("※ 남은 3분은 Q&A로 둡니다.", { x: 0.8, y: 6.55, w: 8, h: 0.3, fontFace: F, fontSize: 10.5, color: C.dim });
  s.addNotes("[1:00~2:00] 순서 안내. Ⅱ는 '무엇을 갖춰야 하나', Ⅲ은 '암묵지를 어떻게 꺼내나', Ⅳ는 '실제로 어떻게 했나', Ⅴ는 '우리 회사는 무엇부터 하나'입니다.");
}

// ───────────────────────── 3. Ⅰ-1 현장의 3대 위기
{
  const s = base("Ⅰ. 왜 지금 제조 AX인가", "제조 현장의 3대 위기: 사람 · 지식 · 성과", "AI가 없어서가 아니라, 판단의 근거가 회사에 남지 않는 것이 문제입니다");
  const cols = [
    ["사람: 숙련 인력의 단절", ["숙련공 고령화·은퇴로 조업 노하우 유실 위기", "신규 인력이 숙련 수준에 이르는 데 약 5년 (포항 철강 현재 수준 진단)", "인력 수급이 어렵고 인원 변동이 잦아 숙련도 편차 심화 (영풍전자)"]],
    ["지식: 머릿속과 메모에 갇힌 노하우", ["경험·직관·판단이 구전과 개인 메모에만 존재 → 디지털화 0% (포항 진단)", "고장탐구 이력·시정조치 데이터가 흩어져 사례 재활용 곤란", "회로도·데이터시트·시험절차서가 부서별로 산재"]],
    ["성과: 사람마다 다른 결과", ["조업 판단 표준편차 ±30% (포항 진단)", "PBA 수리 공수 연 700MD, 7명 투입 · 수리 불가 폐기 77개 (영풍전자, 2025)", "보고서 수기·반수기 작성, 작성자별 일관성 부족"]],
  ];
  cols.forEach((c, i) => card(s, 0.5 + i * 4.16, 1.85, 4.0, 4.35, c[0], c[1], { size: 12.5, gap: 10, barColor: i === 2 ? C.accent2 : C.accent }));
  keyMsg(s, "숙련자가 떠나면 판단도 함께 떠납니다. 제조 AX의 출발점은 '판단의 자산화'입니다.");
  source(s, "포항 AX실증산단 착수회의(2026.10) · 영풍전자 신속상용화 발표평가(2026.05) · 코난테크놀로지 회사소개서");
  s.addNotes("[2:00~4:00] 세 가지 숫자를 강조: 숙련 도달 5년, 판단 편차 ±30%, 수리공수 700MD. 이 숫자들은 실제 과제에서 진단한 현재 수준입니다. 객석에 질문: '우리 회사에서 5년 안에 은퇴하는 핵심 기술자가 몇 명입니까?'");
}

// ───────────────────────── 4. Ⅰ-2 제조 AI vs 제조 AX
{
  const s = base("Ⅰ. 왜 지금 제조 AX인가", "제조 AI에서 제조 AX로: 탐지·예측을 넘어 '설명'과 '실행'으로", "AX(AI Transformation) = AI로 업무 방식 자체를 바꾸는 것");
  const rows = [
    [TH("구분"), TH("기존 제조 AI"), TH("제조 AX")],
    ["목표", "불량 탐지 · 고장 예측 (단일 과제)", "업무 프로세스 전환 (접수 → 분석 → 보고 → 승인)"],
    ["결과물", "점수 · 경보 (\"이상 있음\")", "근거와 함께 설명 (\"왜 그렇게 판단하는가\", XAI)"],
    ["핵심 자산", "학습 데이터 · 모델", "지식베이스 (온톨로지 · 지식그래프) + Agent"],
    ["사용자", "데이터 과학자 · IT 부서", "현장 엔지니어 · 조업 담당자 · 관리자"],
    ["운영", "모델 납품으로 종료", "LLMOps · RAGOps · AgentOps로 지속 고도화"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.85, w: 7.6, colW: [1.3, 2.9, 3.4], fontFace: F, fontSize: 11.5, color: C.text, fill: { color: C.panel }, border: { type: "solid", color: C.line, pt: 1 }, rowH: 0.62, valign: "middle" });
  s.addText("AI Agent 성숙도 3단계", { x: 8.5, y: 1.85, w: 4.3, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: C.white });
  const steps = [
    ["STEP 1  정보제공", "조업 지식 검색 · 판단 근거 설명"],
    ["STEP 2  판단지원", "공정 이상 원인 추론 · 품질·설비 예측 제시"],
    ["STEP 3  자율실행", "정비 우선순위 도출 · 실시간 경보 · 제어시스템 연계"],
  ];
  steps.forEach((st, i) => {
    const y = 2.35 + i * 1.25;
    s.addShape(pptx.ShapeType.rect, { x: 8.5, y, w: 4.33, h: 1.0, fill: { color: i === 0 ? C.panel2 : C.panel }, line: { color: i === 0 ? C.accent : C.line, width: 1 } });
    s.addText(st[0], { x: 8.7, y: y + 0.08, w: 4, h: 0.4, fontFace: F, fontSize: 13, bold: true, color: C.accent });
    s.addText(st[1], { x: 8.7, y: y + 0.48, w: 4, h: 0.45, fontFace: F, fontSize: 11, color: C.text });
    if (i < 2) s.addShape(pptx.ShapeType.downArrow, { x: 10.5, y: y + 1.0, w: 0.3, h: 0.25, fill: { color: C.dim }, line: { color: C.dim, width: 0 } });
  });
  keyMsg(s, "대부분의 기업은 STEP 1(정보제공)부터 시작합니다. 단계를 건너뛰면 현장 신뢰를 잃습니다.");
  source(s, "포항 AX실증산단 착수회의(2026.10) 자율수행형 Agent 고도화 로드맵");
  s.addNotes("[4:00~7:00] 기존 AI 과제는 '불량입니다'에서 끝났지만, AX는 '왜 불량인지, 무엇을 점검할지, 보고서까지' 이어집니다. 3단계 로드맵은 포항 과제의 실제 계획입니다. 자율실행은 마지막 단계이고, 처음부터 목표로 삼지 않습니다.");
}

// ───────────────────────── 5. Ⅱ-1 5-Layer AX Stack
{
  const s = base("Ⅱ. 제조 AX 서비스의 구성요소", "제조 AX 서비스의 전체 구조: 5-Layer AX Stack", "AI 모델은 한 층일 뿐입니다. 성패는 아래 두 층(데이터·지식)에서 갈립니다");
  const layers = [
    ["L5", "사용자 인터페이스", "현장 어시스턴트 · 범용 챗봇 · 대시보드 · 보고서 화면", C.panel],
    ["L4", "서비스 로직 (AI Agent)", "고장탐구 · 검사 · 유사사례 매칭 · 보고서 생성 Agent, 오케스트레이션", C.panel],
    ["L3", "검색 / 생성", "온프레미스 LLM · RAG · Graph-RAG · Query Router(질문 의도별 분기)", C.panel],
    ["L2", "지식 / 데이터", "벡터DB · 지식그래프(온톨로지) · 정형 DB · 제조 용어사전", C.panel2],
    ["L1", "데이터 수집 · 표준화", "ERP · MES · SCM · DCS · 문서 · 영상 · 음성 → AAS 표준 모델", C.panel2],
  ];
  layers.forEach((l, i) => {
    const y = 1.8 + i * 0.86;
    const hl = i >= 3;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 9.0, h: 0.74, fill: { color: l[3] }, line: { color: hl ? C.accent : C.line, width: hl ? 1.5 : 1 } });
    s.addText(l[0], { x: 0.6, y, w: 0.7, h: 0.74, fontFace: F, fontSize: 18, bold: true, color: hl ? C.accent2 : C.accent, valign: "middle" });
    s.addText(l[1], { x: 1.35, y, w: 2.6, h: 0.74, fontFace: F, fontSize: 13.5, bold: true, color: C.white, valign: "middle" });
    s.addText(l[2], { x: 3.95, y, w: 5.5, h: 0.74, fontFace: F, fontSize: 11, color: C.text, valign: "middle" });
  });
  card(s, 9.75, 1.8, 3.08, 2.05, "운영 거버넌스", ["LLMOps: 모델 배포·교체·성능 관리", "RAGOps: 데이터 소스·색인 관리", "AgentOps: 실행 관측·실패 감지"], { size: 10.5, barColor: C.green });
  card(s, 9.75, 3.98, 3.08, 2.12, "보안 · 통제", ["온프레미스 운영(데이터 외부 유출 차단)", "사용자별 권한(RBAC)", "가드레일 · 감사·승인 로그"], { size: 10.5, barColor: C.green });
  keyMsg(s, "L1·L2(데이터와 지식)는 회사 고유 자산으로 남고, L3(모델)는 언제든 더 좋은 것으로 교체할 수 있습니다.");
  source(s, "코난테크놀로지 AX 플랫폼 구축 사례(2026.08) 5-Layer AX Stack · 영풍전자 목표시스템 구성도");
  s.addNotes("[7:00~9:30] 이 그림 한 장이 오늘의 지도입니다. CEO가 기억할 점: 모델(L3)은 매년 바뀝니다. 남는 것은 회사가 정리한 데이터와 지식(L1·L2)입니다. 그래서 투자의 중심도 L1·L2여야 합니다.");
}

// ───────────────────────── 6. Ⅱ-2 데이터 수집·표준화
{
  const s = base("Ⅱ. 제조 AX 서비스의 구성요소", "구성요소 ① 데이터 수집 · 표준화 (L1)", "흩어진 현장 데이터를 하나의 표준 언어로 모읍니다");
  const flow = ["수집", "룰 기반 파싱", "코드·단위 표준화", "품질 검증", "메타 생성 · 온톨로지 연결", "분기 적재"];
  flow.forEach((f, i) => {
    const x = 0.5 + i * 2.1;
    box(s, x, 1.85, 1.75, 0.75, f, { bold: true, size: 11.5, border: i === 2 || i === 4 ? C.accent : C.line });
    if (i < flow.length - 1) arrowR(s, x + 1.78, 2.07, 0.28, 0.3);
  });
  card(s, 0.5, 2.85, 4.0, 3.35, "어떤 데이터를 모으나", ["업무 시스템: ERP · MES · SCM · APS", "설비·공정: DCS 운전 시계열, 알람 이력", "품질: 검사 데이터, 고장탐구·시정조치 이력", "문서: 도면 · 매뉴얼 · 시험절차서 (OCR)", "사람: 명장 음성·판단 의견, 작업일지, 영상"], { size: 11.5, gap: 6 });
  card(s, 4.66, 2.85, 4.0, 3.35, "어떻게 연결하나 (어댑터)", ["REST API: 업무 시스템 표준 연동", "파일: Log · CSV · XML · PDF", "메시지: MQTT · Kafka 스트리밍", "설비 확장: OPC-UA 등 산업 프로토콜", "플러그인 방식으로 신규 설비 추가"], { size: 11.5, gap: 6 });
  card(s, 8.82, 2.85, 4.0, 3.35, "무엇으로 표준화하나", ["AAS(자산관리셸, IEC 63278-1) 서브모델 설계", "구성품 계층 4단계 코드 체계", "통합 코드 매핑 테이블 (시스템 간 코드 통일)", "제조 용어사전 (현장어 ↔ 표준어)", "→ RDB · 벡터DB · 그래프DB · 객체저장소로 분기"], { size: 11.5, gap: 6, barColor: C.accent2 });
  keyMsg(s, "같은 현상을 다른 말로 부르면 AI는 다른 사건으로 봅니다. 용어 통일이 정확도를 좌우합니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) 데이터 처리 체계 · 포항 AX실증산단 착수회의(2026.10)");
  s.addNotes("[9:30~12:00] AAS는 '설비와 자산의 디지털 명함'이라고 설명하면 이해가 쉽습니다. 국제표준(IEC 63278-1)이라 다른 회사·협력사와 데이터를 주고받을 때 공통 언어가 됩니다. 신규 센서를 무조건 늘리기보다 이미 있는 ERP·MES부터 연결하는 것이 순서입니다.");
}

// ───────────────────────── 7. Ⅱ-3 지식 계층: RAG vs Graph-RAG
{
  const s = base("Ⅱ. 제조 AX 서비스의 구성요소", "구성요소 ② 지식 계층 (L2·L3): 문서 검색과 관계 추론을 함께", "\"A 보드 출력 전압이 불안정하다. 원인은?\" 이라는 질문에 두 방식은 다르게 답합니다");
  card(s, 0.5, 1.85, 6.05, 2.6, "RAG (벡터 검색 기반)", ["문서를 의미 단위로 잘라 '비슷한 내용'을 찾음", "답: 관련 매뉴얼·절차서 페이지와 요약, 출처 표시", "강점: 매뉴얼·규정 Q&A, 보고서 초안", "한계: 원인→현상→조치처럼 여러 단계를 잇는 질문에 약함"], { size: 12, gap: 6 });
  card(s, 6.78, 1.85, 6.05, 2.6, "Graph-RAG (지식그래프 기반)", ["개체(설비·부품·증상)와 관계를 따라 추론", "답: 증상 → 회로영역 → 의심부품 → 측정 포인트 → 조치, 과거 성공률 함께 제시", "강점: 고장 원인 역추적, 다단계 판단, 근거 경로 제시", "전제: 온톨로지 설계와 검수가 필요"], { size: 12, gap: 6, barColor: C.accent2 });
  s.addText("질문 의도에 따라 자동 분기 (Query Router)", { x: 0.5, y: 4.6, w: 12, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: C.white });
  box(s, 0.5, 5.05, 2.3, 1.1, "사용자 질문", { bold: true, size: 13, border: C.accent });
  arrowR(s, 2.9, 5.45, 0.45, 0.3);
  box(s, 3.45, 5.05, 2.3, 1.1, "Query Router\n(의도 파악)", { bold: true, size: 12 });
  arrowR(s, 5.85, 5.45, 0.45, 0.3);
  box(s, 6.4, 5.05, 2.05, 1.1, "문서 질문\n→ RAG", { size: 11.5 });
  box(s, 8.6, 5.05, 2.05, 1.1, "원인·관계 질문\n→ Graph-RAG", { size: 11.5, border: C.accent2 });
  box(s, 10.8, 5.05, 2.03, 1.1, "실적·수치 질문\n→ 업무 DB 조회", { size: 11.5 });
  keyMsg(s, "문서는 RAG로, 판단은 지식그래프로. 두 가지를 함께 써야 '근거 있는 답'이 나옵니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) 범용 Chatbot 운영 구성 · RIST 의미추출 모듈 Kick-off(2026.08) Query Router");
  s.addNotes("[12:00~14:30] 실제 질문 하나로 차이를 보여줍니다. RAG는 '어느 문서에 있다'까지, Graph-RAG는 '어디를 먼저 재 보라'까지 갑니다. 회사의 판단 노하우는 문서보다 관계 속에 있기 때문에 지식그래프가 필요합니다.");
}

// ───────────────────────── 8. Ⅱ-4 AI 엔진 & Agent 서비스
{
  const s = base("Ⅱ. 제조 AX 서비스의 구성요소", "구성요소 ③ AI 엔진 · ④ Agent 서비스 (L3·L4·L5)", "업무 단위로 Agent를 만들고, 마지막 확정은 사람이 합니다");
  card(s, 0.5, 1.85, 3.9, 4.35, "AI 엔진", ["온프레미스 LLM (추론형 모델)", "검색엔진 기반 고정밀 RAG", "OCR · STT (도면·음성 디지털화)", "Vision · VLM (영상·이미지 분석)", "PHM (설비 예지정비)", "인과추론 모델 (공정↔설비 원인 역추적)"], { size: 12, gap: 7 });
  card(s, 4.56, 1.85, 4.1, 4.35, "업무 단위 Agent (예시)", ["고장탐구 AI 어시스턴트", "전원 비인가 / 인가 검사 Agent", "유사 고장사례 매칭 Agent", "보고서 자동 생성 (FA Report · 품질검사 · MRB)", "기술문서 검색 · 업무 Q&A 챗봇", "Agent Builder: 현업이 노코드로 Agent 제작"], { size: 12, gap: 7 });
  s.addText("HITL(사람 검증) 업무 흐름", { x: 8.85, y: 1.85, w: 4, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: C.white });
  const flow = [["고장 접수", C.panel], ["AI 분석: 추정 원인 · 의심 부품 · 확인 절차", C.panel], ["결과서 초안 자동 출력", C.panel], ["엔지니어 검증", C.panel2], ["최종 승인", C.panel2]];
  flow.forEach((f, i) => {
    const y = 2.35 + i * 0.78;
    box(s, 8.85, y, 3.98, 0.58, f[0], { fill: f[1], size: 11.5, border: i >= 3 ? C.accent2 : C.line, bold: i >= 3 });
    if (i < flow.length - 1) s.addShape(pptx.ShapeType.downArrow, { x: 10.7, y: y + 0.59, w: 0.25, h: 0.18, fill: { color: C.dim }, line: { color: C.dim, width: 0 } });
  });
  keyMsg(s, "AI는 초안과 근거를 만들고, 판단은 엔지니어가 확정합니다. 이것이 현장이 AI를 받아들이는 조건입니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) 고장탐구 업무 자동화 파이프라인 · 코난테크놀로지 회사소개서");
  s.addNotes("[14:30~17:00] Agent는 '업무 하나를 맡는 AI 직원'으로 설명합니다. 영풍전자는 접수→분석→결과서→검증→승인의 전 과정을 연결했습니다. 사람이 승인 단계를 쥐고 있으므로 책임 소재가 분명하고, 승인·수정 기록이 다시 지식으로 쌓입니다.");
}

// ───────────────────────── 9. Ⅱ-5 운영·인프라
{
  const s = base("Ⅱ. 제조 AX 서비스의 구성요소", "구성요소 ⑤ 운영 · 거버넌스와 인프라", "구축보다 운영이 어렵습니다. 6개월 뒤에도 쓰이는 시스템을 만드는 장치");
  const ops = [
    ["LLMOps", "모델 전 생애주기 관리", ["노코드로 모델 배포·교체", "GPU·응답 성능 대시보드", "추가학습·미세조정·평가"]],
    ["RAGOps", "검색 품질 관리", ["데이터 소스·저장소 설정", "문서 수집·전처리·청킹", "신규 문서 상시 반영"]],
    ["AgentOps", "Agent 실행 관리", ["입력·추론·도구 호출·응답 전 과정 관측", "성능·비용·위험도 평가", "실패 감지 · 가드레일"]],
  ];
  ops.forEach((o, i) => card(s, 0.5 + i * 2.75, 1.85, 2.6, 3.0, o[0], [{ text: o[1], options: { color: C.accent, bold: true, breakLine: true } }, ...o[2]], { size: 11, gap: 6, barColor: C.green }));
  card(s, 8.85, 1.85, 3.98, 4.35, "인프라 구성 예시 (영풍전자 실증)", ["주요 추론 서버 2대 (각 RTX 6000 Pro × 2)", "RAG 서버 · OCR 서버", "온톨로지(그래프DB) 서버", "LLMOps / AgentOps 관리 서버", "WEB/WAS · DB · 객체저장소", "Kubernetes 기반 통합 운영", { text: "→ 신규 HW·네트워크 11종 12식", options: { bold: true, color: C.accent2, breakLine: true } }], { size: 11.5, gap: 5, barColor: C.accent2 });
  card(s, 0.5, 5.0, 8.15, 1.2, "보안 원칙", ["온프레미스(사내 서버) 운영으로 도면·공정 데이터 외부 반출 없음 · 사용자별 권한 · 감사·승인 로그"], { size: 11.5, bullet: false, titleH: 0.4 });
  keyMsg(s, "견적을 받을 때 '구축비'만 보지 말고 '운영 체계(누가, 무엇으로 관리하나)'를 함께 확인하십시오.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) HW 구성도 · 코난테크놀로지 AX 플랫폼 구축 사례(2026.08) LLM Studio");
  s.addNotes("[17:00~19:00] 인프라 규모 감을 드리는 슬라이드입니다. 실제 규모는 업무 범위·동시 사용자 수에 따라 달라지므로 그대로 적용하면 안 된다고 꼭 말할 것. 핵심은 운영 도구(Ops)가 있어야 데이터가 늘고 모델이 바뀌어도 성능이 유지된다는 점입니다.");
}

// ───────────────────────── 10. Ⅲ-1 암묵지란
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "제조 암묵지란 무엇인가", "매뉴얼에 적힌 '무엇을'이 아니라, 명장 머릿속의 '왜 · 언제 · 얼마나'");
  card(s, 0.5, 1.85, 4.2, 2.6, "형식지 (Explicit)", ["문서로 적혀 있는 지식", "표준작업서 · 매뉴얼 · 도면 · 시방서", "\"이 순서대로 점검한다\""], { size: 12.5, gap: 8, barColor: C.dim });
  card(s, 0.5, 4.6, 4.2, 1.6, "암묵지 (Tacit)", ["경험 · 직관 · 판단 — 명장의 '감'", "\"이 소리가 나면 베어링부터 본다\""], { size: 12.5, gap: 6, barColor: C.accent2 });
  s.addText("현장의 말 → 회사의 언어로 (영풍전자 용어사전 예시)", { x: 5.0, y: 1.85, w: 7.8, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: C.white });
  const terms = [["\"뜬 납\"", "ColdSolder (냉납)"], ["\"터진 캡\"", "Capacitor : PhysicalDamage"], ["\"먹통\"", "FunctionalFailure (기능 불량)"]];
  terms.forEach((t, i) => {
    const y = 2.35 + i * 0.62;
    box(s, 5.0, y, 2.6, 0.5, t[0], { size: 13, bold: true, border: C.accent2 });
    arrowR(s, 7.75, y + 0.11, 0.4, 0.28);
    box(s, 8.3, y, 4.53, 0.5, t[1], { size: 12.5 });
  });
  s.addText("명장의 판단 → 재사용 가능한 규칙으로 (포항 철강 예시)", { x: 5.0, y: 4.35, w: 7.8, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: C.white });
  const rule = [["조건", "고속 가공 + 가혹 전단 부하"], ["현상", "툴 온도 급상승"], ["결과", "내부 미세 균열 위험 증가"]];
  rule.forEach((r, i) => {
    const x = 5.0 + i * 2.66;
    s.addShape(pptx.ShapeType.rect, { x, y: 4.85, w: 2.35, h: 1.3, fill: { color: C.panel2 }, line: { color: C.line, width: 1 } });
    s.addText(r[0], { x, y: 4.9, w: 2.35, h: 0.35, fontFace: F, fontSize: 11, bold: true, color: C.accent, align: "center" });
    s.addText(r[1], { x: x + 0.1, y: 5.25, w: 2.15, h: 0.85, fontFace: F, fontSize: 12, color: C.white, align: "center", valign: "middle" });
    if (i < 2) arrowR(s, x + 2.37, 5.35, 0.27, 0.28);
  });
  keyMsg(s, "암묵지 자산화 = 명장의 판단을 '조건 → 현상 → 결과'와 '근거'의 구조로 꺼내 기록하는 일입니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) PBA 전문 용어사전 · 포항 AX실증산단 착수회의(2026.10) 판단 체계 정형화");
  s.addNotes("[19:00~22:00] 객석 참여: '여러분 현장에도 뜬 납, 먹통 같은 말이 있으시죠?' 이 말들을 표준어와 연결해야 AI가 알아듣습니다. 두 번째 예시는 철강 명장의 판단을 규칙 한 줄로 바꾼 것입니다. 이 한 줄이 쌓인 것이 지식그래프입니다.");
}

// ───────────────────────── 11. Ⅲ-2 암묵지는 어디에 있나
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "암묵지는 어디에 있나: 새로 만들기보다 '이미 있는데 연결 안 된 것'부터", "네 곳에서 꺼냅니다");
  const src = [
    ["① 사람", ["명장 인터뷰 · 음성 · 판단 의견", "작업자–AI 대화 기록", "검수 과정의 수정 의견"], "인터뷰 대상자와 일정을 먼저 확보"],
    ["② 기록", ["작업일지 · 트러블 메모", "고장탐구 이력 · 시정조치", "알람 이력 · 클레임 대응 기록"], "형식이 제각각 → 용어 표준화 필요"],
    ["③ 시스템 데이터", ["DCS·MES 운전 시계열", "품질검사 데이터", "DB 스키마와 컬럼 설명"], "테이블 사이의 '관계'는 담당자 머릿속에"],
    ["④ 영상 · 문서", ["작업 영상 · 사진", "회로도 · 도면 · 데이터시트", "시험절차서 · 매뉴얼"], "OCR · Vision으로 디지털화"],
  ];
  src.forEach((c, i) => {
    const x = 0.5 + i * 3.12;
    card(s, x, 1.85, 2.95, 3.25, c[0], c[1], { size: 12, gap: 8 });
    s.addShape(pptx.ShapeType.rect, { x, y: 5.2, w: 2.95, h: 0.95, fill: { color: C.panel2 }, line: { color: C.line, width: 1 } });
    s.addText(c[2], { x: x + 0.12, y: 5.2, w: 2.71, h: 0.95, fontFace: F, fontSize: 11, color: C.accent2, valign: "middle" });
  });
  keyMsg(s, "영풍전자는 BOM 23,598개, 승인도면 27,795건이 이미 전산에 있었습니다. 출발점은 '보유 데이터 목록'입니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) AI 도입 환경 · 포항 AX실증산단 착수회의(2026.10) 6대 조업 데이터");
  s.addNotes("[22:00~24:00] 특히 ③번이 자주 간과됩니다. ERP·MES는 있는데 테이블끼리 어떻게 이어지는지는 담당자만 압니다. 이 부분을 다룬 것이 뒤에 나올 서연이화 사례입니다.");
}

// ───────────────────────── 12. Ⅲ-3 구현 3단계 + 파이프라인
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "암묵지 정형화 3단계 프로세스", "수집 → 정형화 → 추론 Agent");
  const st = [
    ["①", "멀티모달 수집", "숙련공 작업 로그 · 음성 · 영상 · 센서 데이터를 함께 수집"],
    ["②", "온톨로지 · 지식그래프 정형화", "조업 의사결정을 개체 · 관계 · 규칙으로 구조화"],
    ["③", "추론형 마스터 AI Agent", "그래프를 따라 원인을 추론하고 근거와 함께 설명"],
  ];
  st.forEach((t, i) => {
    const x = 0.5 + i * 4.2;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.85, w: 3.85, h: 1.75, fill: { color: C.panel }, line: { color: C.accent, width: 1 } });
    s.addText(t[0], { x: x + 0.15, y: 1.92, w: 0.6, h: 0.6, fontFace: F, fontSize: 26, bold: true, color: C.accent });
    s.addText(t[1], { x: x + 0.75, y: 1.95, w: 3.0, h: 0.55, fontFace: F, fontSize: 14.5, bold: true, color: C.white, valign: "middle" });
    s.addText(t[2], { x: x + 0.2, y: 2.6, w: 3.5, h: 0.9, fontFace: F, fontSize: 11.5, color: C.text, valign: "top" });
    if (i < 2) arrowR(s, x + 3.87, 2.58, 0.3, 0.3);
  });
  s.addText("세부 파이프라인 (LLM 자동 추출 + 전문가 검수)", { x: 0.5, y: 3.8, w: 12, h: 0.38, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  const p = ["자료 수집 ·\n전처리", "도메인\n용어사전", "의미 추출\n개체→관계→규칙", "그래프\n구조화", "전문가\n검수·승인", "AI 서비스\n적용"];
  p.forEach((t, i) => {
    const x = 0.5 + i * 2.1;
    box(s, x, 4.28, 1.8, 1.0, t, { size: 11.5, bold: i === 4, border: i === 4 ? C.accent2 : C.line });
    if (i < p.length - 1) arrowR(s, x + 1.82, 4.63, 0.26, 0.28);
  });
  s.addText("← 신규 용어 · 검수 결과 환류 (지식이 계속 쌓이는 구조)", { x: 4.7, y: 5.38, w: 8, h: 0.35, fontFace: F, fontSize: 11, color: C.accent2, align: "right" });
  s.addText("목표: 새로 참여하는 기업은 소량의 현장 데이터만으로 숙련공 수준의 판단 지원을 받을 수 있게 함 (포항 실증 목표)", { x: 0.5, y: 5.8, w: 12.3, h: 0.4, fontFace: F, fontSize: 11.5, color: C.muted });
  keyMsg(s, "사람이 처음부터 온톨로지를 손으로 짜지 않습니다. LLM이 초안을 뽑고, 명장은 검수에 집중합니다.");
  source(s, "포항 AX실증산단 사례(2026.08) 3단계 암묵지 정형화 프로세스 · RIST 의미추출 모듈 Kick-off(2026.08)");
  s.addNotes("[24:00~26:30] 과거 온톨로지 구축이 비싸고 오래 걸렸던 이유는 사람이 전부 손으로 했기 때문입니다. 지금은 LLM이 문서에서 개체·관계·규칙 초안을 뽑고, 명장은 맞다/틀리다만 판단합니다. 명장의 시간을 가장 적게 쓰는 방식입니다.");
}

// ───────────────────────── 13. Ⅲ-4 온톨로지 인과경로 (영풍)
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "온톨로지 설계: 고장탐구 노하우를 '인과 경로'로 (영풍전자 PBA)", "숙련 엔지니어의 머릿속 탐구 순서를 그래프로 옮깁니다");
  const chain = [["고장 현상", "Symptom"], ["회로 영역", "CircuitBlock"], ["의심 부품", "Component"], ["측정 포인트", "MeasurementPoint"], ["시정 조치", "CorrectiveAction"]];
  chain.forEach((c, i) => {
    const x = 0.5 + i * 2.53;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.95, w: 2.1, h: 1.15, fill: { color: i === 4 ? C.panel2 : C.panel }, line: { color: i === 4 ? C.accent2 : C.accent, width: 1.5 }, rectRadius: 0.1 });
    s.addText(c[0], { x, y: 2.0, w: 2.1, h: 0.6, fontFace: F, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addText(c[1], { x, y: 2.55, w: 2.1, h: 0.45, fontFace: F, fontSize: 10.5, color: C.muted, align: "center" });
    if (i < 4) arrowR(s, x + 2.13, 2.38, 0.36, 0.3);
  });
  card(s, 0.5, 3.35, 4.0, 2.85, "그래프 구성", ["노드 9종: 플랫폼 · 조립체 · PBA · 부품 · 회로블록 · 증상 · 고장모드 · 조치 · 측정포인트", "관계 10종: 증상을 보임 · 유발 · 해결됨 · 측정 필요 · 의심(추론) 등"], { size: 11.5, gap: 8 });
  card(s, 4.66, 3.35, 4.0, 2.85, "관계에 '경험'을 담는 법", ["해결 관계에 과거 성공률 (예: 0.95)", "평균 수리시간 (예: 30분)", "최근 발생일 · 신뢰도 (예: 0.8)", "→ \"이 조치가 과거 95% 성공\"을 근거로 제시"], { size: 11.5, gap: 8, barColor: C.accent2 });
  card(s, 8.82, 3.35, 4.0, 2.85, "활용", ["시나리오별 인과 경로 추적 → 의심부품과 조치 제안", "과거 효과적이었던 조치 식별", "복합 이상 진단: 증상과 회로영역 연결"], { size: 11.5, gap: 8 });
  keyMsg(s, "숙련자의 '탐구 순서'와 '과거 성공 경험'이 그래프의 길과 숫자가 됩니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) 고장탐구 인과관계 지식그래프(Neo4j) — 수치는 설계 예시값");
  s.addNotes("[26:30~29:00] 숙련자는 증상을 보면 어느 회로부터, 어느 부품을, 어디를 재볼지 순서가 머릿속에 있습니다. 그 순서를 그래프로 옮기고, 관계마다 과거 성공률·수리시간을 붙이면 신입도 같은 순서로 탐구할 수 있습니다. 숫자는 설계 문서의 예시값임을 언급.");
}

// ───────────────────────── 14. Ⅲ-5 문장에서 규칙 추출
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "문장 한 줄에서 지식을 꺼내는 과정 (철강 연주 공정 예시)", "LLM이 단계별로 개체 → 관계 → 규칙 → 근거를 추출합니다");
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.85, w: 12.33, h: 0.85, fill: { color: C.panel2 }, line: { color: C.accent2, width: 1 } });
  s.addText("원문: \"래들 잔여 용강이 200톤을 넘은 상태로 교환하면 필터가 턴디시로 배출되어 표면 결함을 야기할 수 있음\"", { x: 0.7, y: 1.85, w: 12, h: 0.85, fontFace: F, fontSize: 13.5, color: C.white, valign: "middle" });
  const steps = [
    ["P1 개체", ["래들 잔여 용강량 (변수)", "래들 교환 (이벤트)", "필터 배출 (결함 원인)"]],
    ["P2 관계", ["잔여 용강량 —[원인]→ 필터 배출", "조건: 200톤 초과", "인과·영향·포함 등 6유형"]],
    ["P3 규칙", ["'잔여 용강 과다' 인과 규칙", "우선순위: 높음", "원문 문장과 연결"]],
    ["P5 통합 · 근거", ["동의어를 대표어로 병합", "충돌 시 '검수 필요' 표시", "모든 노드·관계에 근거 ID 필수"]],
  ];
  steps.forEach((st, i) => {
    const x = 0.5 + i * 3.12;
    card(s, x, 2.9, 2.95, 2.55, st[0], st[1], { size: 11.5, gap: 6, barColor: i === 3 ? C.accent2 : C.accent });
    if (i < 3) arrowR(s, x + 2.96, 4.0, 0.15, 0.25);
  });
  s.addText("※ P4(질문 의도·필요 정보 항목)는 질의 응답 단계에서 사용", { x: 0.5, y: 5.55, w: 12, h: 0.3, fontFace: F, fontSize: 10.5, color: C.dim });
  s.addText("근거 연결 예: 출처 파일 · 문서 조각 번호 · 근거 문장을 별도 노드로 저장 → 답변마다 원문 추적 가능", { x: 0.5, y: 5.88, w: 12.3, h: 0.35, fontFace: F, fontSize: 11.5, color: C.muted });
  keyMsg(s, "AI가 만든 지식이라도 '어느 문서, 어느 문장에서 왔는지'가 항상 붙어 있어야 합니다.");
  source(s, "RIST 작업자-LLM 대화 기반 온톨로지 자동 구현 — 의미추출 및 그래프 구조화 모듈 Kick-off(2026.08)");
  s.addNotes("[29:00~31:00] 실제 연주 공정 문장 하나로 시연하듯 설명합니다. 핵심은 마지막 P5: 모든 지식에 근거가 붙습니다. 그래서 AI 답변을 보고 '이 말 어디서 나왔어?'라고 물으면 원문 문장까지 보여줄 수 있습니다.");
}

// ───────────────────────── 15. Ⅲ-6 용어사전·쿼리확장
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "현장 언어를 AI 언어로: 제조 용어사전과 검색 확장", "용어 정리는 지루하지만 정확도를 가장 크게 올리는 작업입니다");
  card(s, 0.5, 1.85, 4.0, 4.35, "한/영/약어 3중 용어사전", ["불량 유형 용어 50종", "부품 유형 용어 80종", "계측기·측정 용어 40종", "공정 단계 용어 30종", "군 규격 약어 200종 이상 (MIL-STD 등)", { text: "(영풍전자 방산 PBA 설계안)", options: { color: C.dim, breakLine: true } }], { size: 12, gap: 6 });
  card(s, 4.66, 1.85, 4.0, 4.35, "동의어 · 상위어 검색 확장", [{ text: "사용자 입력: '냉납'", options: { bold: true, color: C.accent2, breakLine: true } }, "→ ColdSolder · 솔더크랙 · 접촉불량 · Cold Joint 함께 검색", { text: "사용자 입력: '부품 불량'", options: { bold: true, color: C.accent2, breakLine: true } }, "→ 오삽 · 미삽 · 극성 오류 등 하위 유형 전체 검색"], { size: 12, gap: 8, bullet: false });
  card(s, 8.82, 1.85, 4.0, 4.35, "철강 연주 도메인 예시", ["침지노즐 = SEN = 노즐 = submerged entry nozzle", "몰드탕면레벨 = MdLv = 탕면 레벨 = mold level", "전문용어가 쪼개지지 않도록 사용자 사전 등록", "LLM 추출 → 전문가 검수 → 반영 → 신규 용어 환류"], { size: 12, gap: 8 });
  keyMsg(s, "용어사전은 한 번 만들고 끝이 아니라 계속 자라는 회사의 '제조 국어사전'입니다.");
  source(s, "영풍전자 신속상용화 발표평가(2026.05) PBA 전문 용어사전 · RIST 의미추출 모듈 Kick-off(2026.08) 도메인 사전");
  s.addNotes("[31:00~32:30] 같은 불량을 부서마다 다르게 부르면, 검색이 절반을 놓칩니다. 용어사전이 있으면 '냉납'을 검색해도 '솔더크랙'으로 적힌 과거 사례를 찾아줍니다.");
}

// ───────────────────────── 16. Ⅲ-7 신뢰 장치
{
  const s = base("Ⅲ. 제조 암묵지: 개념과 구현", "신뢰의 장치: 근거 · 검수 · 모르면 묻는 AI", "현장이 AI를 믿게 만드는 세 가지 설계 원칙");
  const items = [
    ["근거 제시 (Evidence-first)", ["모든 지식·답변에 출처 문서·문장·버전 연결", "지식그래프 경로와 유사 사례를 화면에 함께 표시"]],
    ["사람 검수 (HITL)", ["신뢰도가 낮거나 충돌하는 지식은 '검수 필요'로 분류", "검수 정합률 목표 80% 이상 (RIST 과제 기준)", "검수 확정본만 그래프DB에 적재"]],
    ["능동 질의 (모르면 묻기)", ["답에 필요한 정보가 없으면 빈칸으로 표시", "AI가 먼저 질문: \"압연기 출측온도의 정상범위 기준을 알려주세요\"", "답이 들어오면 빈칸을 지식으로 채움"]],
  ];
  items.forEach((it, i) => card(s, 0.5 + i * 4.16, 1.85, 4.0, 3.6, it[0], it[1], { size: 12.5, gap: 10, barColor: i === 2 ? C.accent2 : C.accent }));
  s.addText("\"모르는 것을 모른다고 말하는 AI\"가 현장에서 신뢰받습니다. 그럴듯한 오답이 가장 위험합니다.", { x: 0.5, y: 5.65, w: 12.3, h: 0.5, fontFace: F, fontSize: 14, color: C.white, italic: true });
  keyMsg(s, "AI 도입의 최대 리스크는 정확도 부족이 아니라 '근거 없는 확신'입니다.");
  source(s, "RIST 의미추출 모듈 Kick-off(2026.08) 자동 슬롯 채움·능동 질의 · 영풍전자 발표평가(2026.05)");
  s.addNotes("[32:30~34:00] Ⅲ 파트 마무리. 세 가지 원칙이 있어야 명장도 AI를 '내 노하우를 지키는 도구'로 받아들입니다. 이제 이 원칙이 실제 현장에서 어떻게 적용되고 있는지 세 사례로 넘어갑니다.");
}

// ───────────────────────── 17. Ⅳ-1 사례 개요
{
  const s = base("Ⅳ. 구축 사례", "세 가지 사례 한눈에 보기", "업종도, 암묵지의 형태도 다르지만 같은 구조로 풀었습니다");
  const rows = [
    [TH("구분"), TH("서연이화"), TH("영풍전자"), TH("포항 AX실증산단")],
    ["업종", "자동차부품 제조", "방산 전자회로기판(PBA)", "철강 (가공·조업)"],
    ["사업", "서연이화 AX 구축사업", "AI 응용제품 신속 상용화 지원사업", "경북포항 스마트그린 AX실증산단 구축사업"],
    ["기간", "2026.06 ~ 12", "2026.07 ~ 2027.06 (12개월)", "2026 ~ 2029 (4년)"],
    ["암묵지 대상", "DB 속 '관계' 지식\n(테이블·필드 간 의미)", "고장탐구 숙련자 노하우", "명장의 조업·정비 판단"],
    ["핵심 방법", "AI 관계 추정 + HITL 검수\n→ 온톨로지 자산화", "AAS 표준 + 온톨로지 +\nGraph-RAG + 업무 Agent", "AAS 표준 + 통합 온톨로지 +\n추론·설명 Agent"],
    ["현재 단계", "PoC 결과 확보", "구축 진행 중 (목표치 제시)", "착수 (목표치 제시)"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.8, w: 12.33, colW: [1.8, 3.4, 3.5, 3.63], fontFace: F, fontSize: 12, color: C.text, fill: { color: C.panel }, border: { type: "solid", color: C.line, pt: 1 }, rowH: [0.45, 0.5, 0.6, 0.5, 0.78, 0.78, 0.55], valign: "middle", align: "center" });
  keyMsg(s, "※ 영풍전자·포항은 진행 중인 사업으로, 수치는 '달성 실적'이 아닌 '목표치'입니다.", 6.45);
  source(s, "각 사업 발표·착수 자료 (2026.05~10)");
  s.addNotes("[34:00~35:00] 세 사례를 비교표로 먼저 보여드립니다. 정직하게 말씀드릴 점: 영풍전자와 포항은 진행 중이라 숫자는 목표입니다. 서연이화는 PoC 결과가 나와 있습니다.");
}

// ───────────────────────── 18. Ⅳ-2 서연이화
{
  const s = base("Ⅳ. 구축 사례  ①", "서연이화 (자동차부품): DB 스키마 속 '관계 지식'의 자산화", "ERP·MES는 있는데, 테이블끼리 어떻게 이어지는지는 담당자만 안다면?");
  card(s, 0.5, 1.85, 5.4, 1.75, "문제", ["금형 등 기준정보 DB는 있으나 테이블·필드 간 관계와 의미는 정의서 밖, 담당자 경험에 의존", "담당자 이동·퇴직 시 데이터 활용 단절"], { size: 12, gap: 6, barColor: C.accent2 });
  s.addText("접근 방법", { x: 0.5, y: 3.75, w: 5, h: 0.35, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  const fl = ["DB 스키마 +\n컬럼 설명", "AI 관계 추정\n(근거·득표)", "전문가 검수\n(HITL)", "YAML/OWL\n확정·그래프 적재"];
  fl.forEach((t, i) => {
    box(s, 0.5 + i * 1.37, 4.15, 1.2, 1.15, t, { size: 10, border: i === 2 ? C.accent2 : C.line, bold: i === 2 });
  });
  s.addText("PoC 대상: 금형마스터 7개 테이블 · 308개 필드 · FK 관계 16건", { x: 0.5, y: 5.45, w: 5.5, h: 0.5, fontFace: F, fontSize: 11.5, color: C.muted });
  s.addText("PoC 결과 (실측)", { x: 6.2, y: 1.85, w: 6, h: 0.35, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  const k = [
    ["F1 67~72%", "숨긴 관계 블라인드 복원 정확도"],
    ["12 / 16건", "전면 은닉한 관계 중 복원"],
    ["9건", "정의서에 없던 관계 신규 발굴\n(근거와 함께 추천)"],
    ["0건", "허위 관계 (잘못된 관계 확정)"],
    ["10 / 10회", "가짜 테이블 노이즈 방어"],
    ["HITL", "추정 → 검수 → 자산화"],
  ];
  k.forEach((v, i) => {
    const x = 6.2 + (i % 3) * 2.22;
    const y = 2.3 + Math.floor(i / 3) * 1.85;
    kpi(s, x, y, 2.08, 1.7, v[0], v[1], { vs: 20, ls: 10, vc: i === 3 ? C.green : C.accent });
  });
  keyMsg(s, "AI의 역할은 정답 확정이 아니라 '검토할 가치가 있는 후보'를 근거와 함께 올리는 것입니다.");
  source(s, "코난테크놀로지 AX 플랫폼 구축 사례(2026.08) · RIST 의미추출 모듈 Kick-off(2026.08) 온톨로지 레퍼런스②");
  s.addNotes("[35:00~39:00] 강조점 두 가지. (1) 허위 관계 0건: AI가 틀린 관계를 확정한 경우가 없었습니다. 애매하면 사람에게 넘깁니다. (2) 정의서에 없던 관계 9건: 사람이 문서화하지 않았던 지식을 AI가 찾아냈습니다. 바로 이것이 '데이터 속 암묵지'입니다. ※ 고객사 실명 공개 동의 확인 필요(내부 자료 일부는 비식별 처리됨).");
}

// ───────────────────────── 19. Ⅳ-3 영풍전자 AS-IS/TO-BE
{
  const s = base("Ⅳ. 구축 사례  ②", "영풍전자 (방산 PBA): 고장탐구 AI 어시스턴트", "다품종·소량, 고신뢰성 회로기판의 고장탐구 전 과정을 AI + 엔지니어 협업으로");
  card(s, 0.5, 1.85, 4.0, 2.25, "사업 개요", ["과제: AI 기반 제조 지식 자산화 및 전자회로기판 제조 공정 관리 자동화", "총 28.6억 원 (정부 20억 · 기업 8.6억)", "12개월 · 참여 43명 · 5개 기관"], { size: 11.5, gap: 5 });
  card(s, 0.5, 4.22, 4.0, 1.98, "이미 갖춘 IT 자산", ["SAP ERP · e-SCM · e-SFC(MES) · 도면관리", "BOM 조립체 23,598 · 품목 85,789", "승인도면 27,795 · 개발도면 11,353건"], { size: 11.5, gap: 5, barColor: C.green });
  card(s, 4.66, 1.85, 3.95, 4.35, "AS-IS: 전 과정 수작업", ["육안검사 → 전원 인가 → 회로도 분석 → 계측기 진단 → 부품 교환 → 보고서 수기 작성", "개인 재량·숙련도에 따라 수리시간 편차", "측정값 수기 비교로 판독 오류, SW 버전 체크 누락 위험", { text: "2025년 수리 공수 700MD · 7명 투입", options: { bold: true, color: C.accent2, breakLine: true } }, { text: "수리 불가 폐기 77개", options: { bold: true, color: C.accent2, breakLine: true } }], { size: 11.5, gap: 6, barColor: C.accent2 });
  card(s, 8.77, 1.85, 4.06, 4.35, "TO-BE: AI + 엔지니어 협업", ["고장 현상 입력 → 추정 원인 · 의심 부품 · 확인 절차 · 시정조치 단계별 제안", "전원 비인가/인가 검사 Agent: 기준값 자동 제시, SW 불일치 자동 감지", "유사 고장사례 자동 매칭", "FA Report · 품질검사 · MRB 보고서 초안 자동 생성", "엔지니어 검증 · 승인 (HITL)"], { size: 11.5, gap: 6 });
  keyMsg(s, "2019년 ERP·MES, 2024년 APS 등 스마트공장 투자가 쌓여 있었기에 AX로 바로 갈 수 있었습니다.");
  source(s, "영풍전자 「AI 응용제품 신속 상용화 지원사업」 발표평가(2026.05.19)");
  s.addNotes("[39:00~42:00] 방산 PBA는 다품종·소량이라 같은 고장이 반복되지 않는 것처럼 보이지만, 회로블록·부품 수준에서는 패턴이 있습니다. 그 패턴을 지식그래프로 만든 것이 핵심입니다. 기존 스마트공장 투자(ERP·MES)가 AX의 준비도로 이어졌다는 점을 강조. ※ 내부 수치(700MD·폐기 77개) 외부 공개 동의 확인 필요.");
}

// ───────────────────────── 20. Ⅳ-4 영풍전자 KPI·확산
{
  const s = base("Ⅳ. 구축 사례  ②", "영풍전자: 목표 성과지표와 단계적 확산 계획", "목표치는 현재 수준(기준선)이 측정되어 있기에 세울 수 있었습니다");
  s.addText("목표 KPI (사업 종료 시점 목표)", { x: 0.5, y: 1.8, w: 8, h: 0.35, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  const k = [
    ["50%↓", "고장 원인 분석 시간"], ["70%↓", "보고서 작성 시간"], ["80%↓", "기술문서 검색 시간"],
    ["85%↑", "유사 고장사례 매칭 정확도"], ["90%↑", "표준 코드 적용률"], ["95%↑", "데이터 정합성"],
  ];
  k.forEach((v, i) => {
    const x = 0.5 + (i % 3) * 2.75;
    const y = 2.25 + Math.floor(i / 3) * 1.55;
    kpi(s, x, y, 2.6, 1.4, v[0], v[1], { vs: 26, ls: 11 });
  });
  s.addText("단계적 확산", { x: 8.85, y: 1.8, w: 4, h: 0.35, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  const ph = [["1단계 · 사업 기간 내", "창원 본사 PBA 제조·검사 공정 시범 운영"], ["2단계 · 종료 후 1년", "대전·성남 사업장 확대, MES·ERP 통합 연동"], ["3단계 · 종료 후 2~3년", "국내 방산전자 동종업계 B2B 패키지 공급"]];
  ph.forEach((p, i) => {
    const y = 2.25 + i * 1.05;
    s.addShape(pptx.ShapeType.rect, { x: 8.85, y, w: 3.98, h: 0.9, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
    s.addText(p[0], { x: 9.0, y: y + 0.05, w: 3.7, h: 0.35, fontFace: F, fontSize: 11.5, bold: true, color: C.accent });
    s.addText(p[1], { x: 9.0, y: y + 0.4, w: 3.7, h: 0.45, fontFace: F, fontSize: 10.5, color: C.text });
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.45, w: 12.33, h: 0.8, fill: { color: C.panel2 }, line: { color: C.line, width: 1 } });
  s.addText([{ text: "확산 조건: ", options: { bold: true, color: C.accent2 } }, { text: "내부 KPI 달성 → 투자심의 → 확산 예산 확보.  KPI가 다음 투자의 근거가 되는 구조입니다." }], { x: 0.7, y: 5.45, w: 12, h: 0.8, fontFace: F, fontSize: 12.5, color: C.white, valign: "middle" });
  keyMsg(s, "\"작게 실증하고, 숫자로 증명하고, 그 숫자로 다음 투자를 결정한다.\"");
  source(s, "영풍전자 「AI 응용제품 신속 상용화 지원사업」 발표평가(2026.05.19) — 모든 수치는 목표치");
  s.addNotes("[42:00~44:00] 이 슬라이드는 CEO 의사결정 방식의 모범 사례로 소개합니다. 전사 도입을 한 번에 결정하지 않고, 본사 1개 공정에서 KPI를 확인한 뒤 투자심의로 확산합니다. Ⅴ 파트의 우선순위 결정 방법과 연결됩니다.");
}

// ───────────────────────── 21. Ⅳ-5 포항
{
  const s = base("Ⅳ. 구축 사례  ③", "포항 AX실증산단 (철강): 숙련공(명장) 조업 암묵지 자산화", "산단 단위로 만들고, 검증된 Agent를 '카세트 모듈'로 여러 기업에 확산");
  card(s, 0.5, 1.85, 4.1, 2.5, "사업 개요", ["경북포항 스마트그린 AX실증산단 구축사업", "RIST(총괄) · 동국제강 · 제일테크노스 · 코난테크놀로지", "4년 (2026~2029) · 업체별 실증 공정 1건부터"], { size: 11.5, gap: 6 });
  card(s, 0.5, 4.48, 4.1, 1.72, "AI 모델 4종", ["Vision·VLM 분석 · 공정-설비 통합 추론", "명장 판단 근거 설명 · 유사 사례 검색·설명"], { size: 11.5, gap: 5 });
  s.addText("현재 수준(Lv.2) → 목표 수준(Lv.4)", { x: 4.85, y: 1.85, w: 8, h: 0.35, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  const g = [["암묵지 디지털 자산화", "0%", "70% 이상"], ["신규 인력 숙련 도달 기간", "5년", "1년"], ["조업 판단 표준편차", "±30%", "±10% 이내"]];
  g.forEach((r, i) => {
    const y = 2.3 + i * 0.85;
    s.addShape(pptx.ShapeType.rect, { x: 4.85, y, w: 7.98, h: 0.72, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
    s.addText(r[0], { x: 5.0, y, w: 3.3, h: 0.72, fontFace: F, fontSize: 12.5, bold: true, color: C.white, valign: "middle" });
    s.addText(r[1], { x: 8.3, y, w: 1.5, h: 0.72, fontFace: F, fontSize: 18, color: C.muted, align: "center", valign: "middle" });
    arrowR(s, 9.9, y + 0.21, 0.5, 0.3);
    s.addText(r[2], { x: 10.5, y, w: 2.2, h: 0.72, fontFace: F, fontSize: 18, bold: true, color: C.accent, align: "center", valign: "middle" });
  });
  s.addText("연차별 계획", { x: 4.85, y: 4.95, w: 4, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: C.white });
  const yr = [["2026", "표준·데이터 체계 설계"], ["2027", "지식그래프 · 핵심 AI PoC"], ["2028", "통합추론 · 카세트형 SaaS"], ["2029", "현장 검증 · 전국 산단 전파"]];
  yr.forEach((y, i) => {
    const x = 4.85 + i * 2.0;
    box(s, x, 5.3, 1.88, 0.9, y[0] + "\n" + y[1], { size: 10.5, border: i === 0 ? C.accent : C.line });
  });
  keyMsg(s, "개별 기업이 혼자 하기 어려운 표준·온톨로지를 산단이 함께 만들고, 중소 협력사는 모듈을 가져다 씁니다.");
  source(s, "코난테크놀로지 회사소개서(2026.08) M.AX Core · 포항 AX실증산단 착수회의(2026.10) — 수치는 목표치");
  s.addNotes("[44:00~48:00] 대구·경북 CEO께 가장 가까운 사례. 중소 협력사가 혼자 AI를 도입하기엔 비용 장벽이 큽니다. 산단 차원에서 표준과 코어를 만들고, 검증된 Agent를 카세트처럼 끼워 쓰게 하는 것이 확산 전략입니다. 숙련 도달 5년→1년이 가장 강력한 메시지.");
}

// ───────────────────────── 22. Ⅳ-6 교훈
{
  const s = base("Ⅳ. 구축 사례", "세 사례가 주는 5가지 교훈", "업종과 규모가 달라도 성공 조건은 같았습니다");
  const L = [
    ["1", "작게 시작한다", "포항: 업체별 공정 1건 · 서연이화: 금형마스터 7개 테이블"],
    ["2", "있는 데이터부터 연결한다", "영풍전자: 기존 ERP·MES·도면 시스템 활용, 신규 수집은 그 다음"],
    ["3", "기준선을 먼저 잰다", "조치시간·분석시간·공수의 현재값이 없으면 성과를 증명할 수 없음"],
    ["4", "AI는 초안, 확정은 사람", "허위 관계 0건 · 엔지니어 검증·승인 단계 유지"],
    ["5", "처음부터 확산 구조로 설계한다", "카세트 모듈 · 사업장 → 동종업계 단계적 확산"],
  ];
  L.forEach((l, i) => {
    const y = 1.85 + i * 0.88;
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 12.33, h: 0.76, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
    s.addText(l[0], { x: 0.6, y, w: 0.7, h: 0.76, fontFace: F, fontSize: 24, bold: true, color: C.accent, align: "center", valign: "middle" });
    s.addText(l[1], { x: 1.4, y, w: 4.2, h: 0.76, fontFace: F, fontSize: 15, bold: true, color: C.white, valign: "middle" });
    s.addText(l[2], { x: 5.6, y, w: 7.1, h: 0.76, fontFace: F, fontSize: 12, color: C.text, valign: "middle" });
  });
  keyMsg(s, "그렇다면 우리 회사는 '어떤 업무 1건'으로 시작해야 할까요? → Ⅴ. 우선순위 결정");
  source(s, "서연이화 · 영풍전자 · 포항 AX실증산단 사업 자료 종합");
  s.addNotes("[48:00~49:00] 사례 파트 요약. 다음 질문으로 자연스럽게 넘어갑니다: 그래서 우리는 무엇부터 해야 하나.");
}

// ───────────────────────── 23. Ⅴ-1 원칙: 가치 × 실행가능성
{
  const s = base("Ⅴ. AX 도입 우선순위 결정", "우선순위의 원칙: 가치(Value) × 실행 가능성(Feasibility)", "\"가장 중요한 업무\"가 아니라 \"가치가 크고 지금 해낼 수 있는 업무\"부터");
  // 2x2
  const ox = 0.9, oy = 1.95, cw = 3.3, ch = 2.05;
  const q = [
    [0, 0, "전략 과제", "가치 高 · 실행 低", "데이터 정비·표준화를 먼저 하고 추진", C.panel],
    [1, 0, "1순위: Quick Win", "가치 高 · 실행 高", "PoC 대상 — 공정·업무 1건", C.panel2],
    [0, 1, "보류", "가치 低 · 실행 低", "현 단계에서 제외", C.panel],
    [1, 1, "개선 과제", "가치 低 · 실행 高", "LLM 사무자동화 등 범용 도구로 빠르게", C.panel],
  ];
  q.forEach(([cx, cy, t, a, d, fill]) => {
    const x = ox + cx * (cw + 0.1), y = oy + cy * (ch + 0.1);
    const hl = cx === 1 && cy === 0;
    s.addShape(pptx.ShapeType.rect, { x, y, w: cw, h: ch, fill: { color: fill }, line: { color: hl ? C.accent2 : C.line, width: hl ? 2 : 1 } });
    s.addText(t, { x: x + 0.15, y: y + 0.15, w: cw - 0.3, h: 0.45, fontFace: F, fontSize: 15, bold: true, color: hl ? C.accent2 : C.white });
    s.addText(a, { x: x + 0.15, y: y + 0.62, w: cw - 0.3, h: 0.35, fontFace: F, fontSize: 11, color: C.accent });
    s.addText(d, { x: x + 0.15, y: y + 1.0, w: cw - 0.3, h: 0.9, fontFace: F, fontSize: 11.5, color: C.text, valign: "top" });
  });
  s.addText("가치 ↑", { x: 0.2, y: 2.0, w: 0.65, h: 0.4, fontFace: F, fontSize: 11, color: C.muted });
  s.addText("실행 가능성 →", { x: 5.6, y: 6.2, w: 1.6, h: 0.3, fontFace: F, fontSize: 11, color: C.muted, align: "right" });
  card(s, 8.0, 1.95, 4.83, 4.2, "포항 실증 대상 선정 5대 기준", [
    { text: "가치 측", options: { bold: true, color: C.accent, breakLine: true } },
    "숙련공 의존도: 사람마다 판단 편차가 큰 공정",
    "확산성: 같은 공정을 가진 기업이 많음",
    { text: "실행 측", options: { bold: true, color: C.accent, breakLine: true } },
    "인터뷰 인력: 참여 가능한 숙련공 보유",
    "데이터 연계: 기존 예지보전·품질검사 데이터와 연결  ★최우선",
    "성과 측정: 조치시간 등 현재 수준 측정 가능  ★최우선",
  ], { size: 11.5, gap: 5, barColor: C.accent2 });
  keyMsg(s, "데이터 연계와 성과 측정이 안 되는 과제는, 아무리 중요해도 첫 과제로 고르지 않습니다.");
  source(s, "포항 AX실증산단 착수회의(2026.10) 실증 대상 선정 기준");
  s.addNotes("[49:00~51:00] 많은 기업이 '가장 골치 아픈 문제'를 첫 과제로 고르다가 실패합니다. 데이터가 없거나 성과를 잴 수 없으면 1년 뒤 '효과가 있었나?'라는 질문에 답할 수 없습니다. 포항 과제도 데이터 연계와 성과 측정을 최우선 기준으로 삼았습니다.");
}

// ───────────────────────── 24. Ⅴ-2 정량 스코어카드
{
  const s = base("Ⅴ. AX 도입 우선순위 결정", "정량 스코어카드: 후보 업무를 8개 지표로 점수화", "각 지표를 1~5점으로 매기고 가중치를 곱해 가치 50점 + 실행 가능성 50점으로 환산");
  const rows = [
    [TH("축"), TH("지표"), TH("측정 방법"), TH("1점"), TH("3점"), TH("5점"), TH("가중치")],
    ["가치", "V1 업무 부하", "해당 업무 연간 투입 공수(MD)", "50MD 미만", "약 200MD", "500MD 이상", "15"],
    ["가치", "V2 숙련 의존도", "담당자 간 처리시간·판단 편차", "10% 미만", "약 20%", "30% 이상", "15"],
    ["가치", "V3 품질·손실 비용", "불량·폐기·재작업·클레임 연간 비용", "미미", "중간", "매출 영향", "10"],
    ["가치", "V4 인력 리스크", "5년 내 은퇴·이직 예상 핵심인력 비율", "없음", "일부", "핵심 대부분", "10"],
    ["실행", "F1 데이터 연계성", "관련 데이터 전산화 수준", "종이·구두", "파일 산재", "ERP·MES 축적", "15"],
    ["실행", "F2 성과 측정성", "현재 기준선(시간·건수) 측정 가능 여부", "불가", "추정 가능", "이미 측정 중", "15"],
    ["실행", "F3 전문가 참여", "인터뷰·검수 가능한 숙련자 확보", "불가", "제한적", "전담 가능", "10"],
    ["실행", "F4 확산성", "같은 업무를 가진 사업장·협력사 수", "단독", "2~3곳", "다수", "10"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.75, w: 12.33, colW: [0.8, 2.0, 3.7, 1.35, 1.35, 1.6, 1.53], fontFace: F, fontSize: 11, color: C.text, fill: { color: C.panel }, border: { type: "solid", color: C.line, pt: 1 }, rowH: 0.44, valign: "middle", align: "center" });
  s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.85, w: 12.33, h: 0.5, fill: { color: C.panel2 }, line: { color: C.line, width: 1 } });
  s.addText([{ text: "산식  ", options: { bold: true, color: C.accent } }, { text: "가치 점수 = Σ(V 가중치 × 점수) ÷ 5   |   실행 점수 = Σ(F 가중치 × 점수) ÷ 5   |   각 축 30점 이상이면 '높음'" }], { x: 0.7, y: 5.85, w: 12, h: 0.5, fontFace: F, fontSize: 11.5, color: C.white, valign: "middle" });
  keyMsg(s, "※ 지표 구간·가중치·기준점(30점)은 코난테크놀로지 제안값입니다. 업종·규모에 맞게 조정해 쓰십시오.");
  source(s, "포항 AX실증산단 선정 기준(2026.10) · 영풍전자 성과지표 체계(2026.05)를 바탕으로 재구성한 제안 모델");
  s.addNotes("[51:00~53:30] 이 표는 그대로 가져가 사내 회의에서 쓰실 수 있습니다. 가치 4개, 실행 4개 지표. 가중치와 구간은 저희 제안값이므로 회사 사정에 맞게 바꾸시면 됩니다. 중요한 것은 '느낌'이 아니라 같은 잣대로 후보들을 비교하는 것입니다.");
}

// ───────────────────────── 25. Ⅴ-3 적용 예시 (가상)
{
  const s = base("Ⅴ. AX 도입 우선순위 결정", "스코어카드 적용 예시 (가상의 부품 제조기업)", "후보 업무 3개를 같은 잣대로 비교하면 '첫 과제'가 분명해집니다");
  const rows = [
    [TH("후보 업무"), TH("V1"), TH("V2"), TH("V3"), TH("V4"), TH("F1"), TH("F2"), TH("F3"), TH("F4"), TH("가치"), TH("실행"), TH("판정")],
    ["A. 고장탐구 지원", "5", "5", "4", "4", "4", "4", "4", "3", "46", "38", "1순위 PoC"],
    ["B. 품질보고서 자동 작성", "3", "2", "2", "2", "5", "5", "4", "4", "23", "46", "개선 과제"],
    ["C. 설비 예지정비", "3", "4", "5", "3", "2", "2", "3", "3", "37", "24", "전략 과제"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.8, w: 8.2, colW: [2.2, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45, 0.65, 0.65, 1.1], fontFace: F, fontSize: 11, color: C.text, fill: { color: C.panel }, border: { type: "solid", color: C.line, pt: 1 }, rowH: 0.55, valign: "middle", align: "center" });
  s.addText("계산 예 (A): 가치 = (15×5 + 15×5 + 10×4 + 10×4) ÷ 5 = 46점", { x: 0.5, y: 4.15, w: 8.2, h: 0.35, fontFace: F, fontSize: 11, color: C.muted });
  card(s, 0.5, 4.6, 8.2, 1.6, "해석", ["A: 가치·실행 모두 높음 → 숙련자 인터뷰와 기존 이력 데이터로 바로 PoC", "B: 가치는 낮지만 쉬움 → LLM 사무자동화 도구로 빠르게 적용", "C: 가치는 크지만 센서·이력 데이터 부족 → 데이터 수집 체계를 먼저 갖춘 뒤 추진"], { size: 11.5, gap: 4 });
  // 미니 매트릭스
  const mx = 9.0, my = 1.8, mw = 3.83, mh = 4.4;
  s.addShape(pptx.ShapeType.rect, { x: mx, y: my, w: mw, h: mh, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
  s.addShape(pptx.ShapeType.line, { x: mx + mw * 0.6, y: my, w: 0, h: mh, line: { color: C.dim, width: 1, dashType: "dash" } });
  s.addShape(pptx.ShapeType.line, { x: mx, y: my + mh * 0.4, w: mw, h: 0, line: { color: C.dim, width: 1, dashType: "dash" } });
  s.addText("가치 ↑", { x: mx + 0.05, y: my + 0.05, w: 1, h: 0.3, fontFace: F, fontSize: 9.5, color: C.muted });
  s.addText("실행 →", { x: mx + mw - 1.05, y: my + mh - 0.35, w: 1, h: 0.3, fontFace: F, fontSize: 9.5, color: C.muted, align: "right" });
  // 좌표: 실행 0~50 → x, 가치 0~50 → y(반전)
  const pt = (f, v) => [mx + (f / 50) * mw, my + mh - (v / 50) * mh];
  [["A", 38, 46, C.accent2], ["B", 46, 23, C.accent], ["C", 24, 37, C.accent]].forEach(([n, f, v, col]) => {
    const [px, py] = pt(f, v);
    s.addShape(pptx.ShapeType.ellipse, { x: px - 0.25, y: py - 0.25, w: 0.5, h: 0.5, fill: { color: col }, line: { color: C.white, width: 1 } });
    s.addText(n, { x: px - 0.25, y: py - 0.25, w: 0.5, h: 0.5, fontFace: F, fontSize: 13, bold: true, color: C.bg, align: "center", valign: "middle" });
  });
  s.addText("기준선 30점", { x: mx + mw * 0.6 + 0.05, y: my + mh * 0.4 + 0.05, w: 1.4, h: 0.3, fontFace: F, fontSize: 8.5, color: C.dim });
  keyMsg(s, "※ 위 점수는 설명을 위한 가상 예시이며, 특정 기업의 실제 데이터가 아닙니다.");
  source(s, "코난테크놀로지 제안 스코어카드 적용 예시 (가상 데이터)");
  s.addNotes("[53:30~55:00] 가상 예시임을 반드시 언급. 같은 회사 안에서도 '고장탐구'는 바로 할 수 있고, '예지정비'는 데이터부터 쌓아야 하고, '보고서 작성'은 범용 도구로 빨리 해결할 수 있습니다. 세 가지를 다르게 다뤄야 예산이 낭비되지 않습니다.");
}

// ───────────────────────── 26. Ⅴ-4 효과 산정
{
  const s = base("Ⅴ. AX 도입 우선순위 결정", "효과 산정과 KPI 설계: 투자 판단을 숫자로", "기준선 → 목표 → 효과 금액 → 회수 기간 순서로 계산합니다");
  card(s, 0.5, 1.85, 6.1, 2.55, "간이 효과 산정식", [
    { text: "연간 효과 =", options: { bold: true, color: C.accent, breakLine: true } },
    "① 절감 공수(MD) × MD당 인건비",
    "＋ ② 품질손실 감소액 (폐기·재작업·클레임)",
    "＋ ③ 숙련 기간 단축 효과 (교육·OJT 기간 × 인건비)",
    { text: "회수 기간 = (구축비 + 연간 운영비 − 정부지원금) ÷ 연간 효과", options: { bold: true, color: C.accent2, breakLine: true } },
  ], { size: 12, gap: 5, bullet: false });
  card(s, 6.78, 1.85, 6.05, 2.55, "기준선이 있어야 계산이 됩니다 (영풍전자)", [
    "기준선: 수리 공수 연 700MD · 수리 불가 폐기 77개 (2025)",
    "목표: 분석시간 50%↓ · 보고서 작성 70%↓",
    "재원: 총 28.6억 중 정부지원 20억 (약 70%), 기업 부담 8.6억",
  ], { size: 12, gap: 8, barColor: C.accent2 });
  s.addText("KPI는 선행지표와 결과지표를 함께 잡습니다", { x: 0.5, y: 4.55, w: 12, h: 0.35, fontFace: F, fontSize: 13.5, bold: true, color: C.white });
  card(s, 0.5, 4.95, 6.1, 1.3, "선행지표 (잘 쓰이고 있나)", ["AI 적용률 · 데이터 수집률 · 온톨로지 커버리지 · 검수 정합률"], { size: 11.5, bullet: false, titleH: 0.4 });
  card(s, 6.78, 4.95, 6.05, 1.3, "결과지표 (돈이 되었나)", ["분석·작성 시간 · 납기 단축률 · 불량 예측 정확도 · 폐기 수량"], { size: 11.5, bullet: false, titleH: 0.4, barColor: C.accent2 });
  keyMsg(s, "인건비 단가·구축비는 기업마다 다릅니다. 오늘 드린 산식에 자사 값을 넣어 보십시오.");
  source(s, "영풍전자 발표평가(2026.05.19) 성과지표·예산 · 산정식은 코난테크놀로지 제안 모델");
  s.addNotes("[55:00~56:30] 구체 금액은 회사마다 달라서 일반화하지 않습니다. 대신 산식과 기준선의 중요성을 강조합니다. 정부지원 비율(영풍전자 약 70%)은 다음 세션 조윤성 박사님의 정부지원 강의로 자연스럽게 연결합니다.");
}

// ───────────────────────── 27. Ⅴ-5 도입 로드맵
{
  const s = base("Ⅴ. AX 도입 우선순위 결정", "기업 AX 도입 4단계 로드맵", "진단 → PoC → 실증·확대 → 자립·확산");
  const st = [
    ["STEP 1", "진단", "1~2개월", ["후보 업무 목록화 · 스코어카드 평가", "기준선(시간·공수·불량) 측정", "보유 데이터 목록 작성", "숙련자 명단 · 인터뷰 대상 선정"]],
    ["STEP 2", "PoC", "3~6개월", ["업무 1건 · 공정 1건으로 한정", "용어사전 · 최소 온톨로지", "정보제공형 Agent (STEP 1 수준)", "HITL 검수 체계 구축"]],
    ["STEP 3", "실증 · 확대", "6~12개월", ["KPI 검증 → 투자심의", "판단지원형 Agent로 고도화", "타 공정 · 사업장 확대", "정부 지원사업 연계"]],
    ["STEP 4", "자립 · 확산", "이후", ["LLMOps · AgentOps 운영 내재화", "현업이 Agent Builder로 직접 제작", "협력사 · 동종업계 확산", "지식 자산 지속 축적"]],
  ];
  st.forEach((t, i) => {
    const x = 0.5 + i * 3.12;
    s.addShape(pptx.ShapeType.rect, { x, y: 1.85, w: 2.95, h: 1.05, fill: { color: i === 0 ? C.panel2 : C.panel }, line: { color: i === 0 ? C.accent2 : C.accent, width: 1.5 } });
    s.addText(t[0], { x, y: 1.88, w: 2.95, h: 0.35, fontFace: F, fontSize: 11, bold: true, color: C.accent, align: "center" });
    s.addText(t[1], { x, y: 2.2, w: 2.95, h: 0.4, fontFace: F, fontSize: 17, bold: true, color: C.white, align: "center" });
    s.addText(t[2], { x, y: 2.58, w: 2.95, h: 0.3, fontFace: F, fontSize: 10.5, color: C.accent2, align: "center" });
    card(s, x, 3.0, 2.95, 2.85, null, t[3], { size: 11.5, gap: 8, bar: false });
    if (i < 3) arrowR(s, x + 2.96, 2.22, 0.15, 0.3);
  });
  s.addText("※ 기간은 일반적인 제안이며, 업무 범위와 데이터 상태에 따라 달라집니다.", { x: 0.5, y: 5.95, w: 12, h: 0.3, fontFace: F, fontSize: 10.5, color: C.dim });
  keyMsg(s, "오늘 돌아가서 할 일은 STEP 1 하나입니다: 후보 업무 3개를 적고, 기준선을 재는 것.");
  source(s, "영풍전자 단계적 확산 계획 · 포항 AX실증산단 연차별 계획을 바탕으로 재구성");
  s.addNotes("[56:30~58:00] 로드맵의 첫 단계는 비용이 거의 들지 않습니다. 후보 업무 목록, 기준선 측정, 보유 데이터 목록, 숙련자 명단. 이 네 가지만 있으면 정부 지원사업 신청서의 절반이 준비됩니다.");
}

// ───────────────────────── 28. CEO 체크리스트
{
  const s = base("Ⅴ. AX 도입 우선순위 결정", "CEO 자가진단 체크리스트 10문항", "'예'의 개수로 지금 우리 회사의 출발점을 확인해 보십시오");
  const q = [
    "5년 내 은퇴 예정인 핵심 숙련자가 누구인지 명단이 있다",
    "같은 문제를 사람마다 다르게 판단·처리하는 업무가 있다",
    "그 업무의 현재 처리시간·공수를 숫자로 말할 수 있다",
    "고장·불량·조치 이력이 ERP·MES 등에 전산으로 쌓인다",
    "같은 현상을 부르는 사내 용어가 통일되어 있다",
    "도면·매뉴얼·절차서를 한 곳에서 검색할 수 있다",
    "숙련자가 AI 검수에 주 몇 시간이라도 참여할 수 있다",
    "데이터를 사내(온프레미스)에 둘지 외부를 쓸지 원칙이 있다",
    "AI 결과를 최종 승인할 책임자를 지정할 수 있다",
    "성공하면 확산할 사업장·협력사가 있다",
  ];
  q.forEach((t, i) => {
    const col = i < 5 ? 0 : 1;
    const row = i % 5;
    const x = 0.5 + col * 6.2, y = 1.85 + row * 0.72;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 6.0, h: 0.6, fill: { color: C.panel }, line: { color: C.line, width: 1 } });
    s.addShape(pptx.ShapeType.rect, { x: x + 0.15, y: y + 0.16, w: 0.28, h: 0.28, fill: { color: C.bg }, line: { color: C.accent, width: 1 } });
    s.addText(`${i + 1}. ${t}`, { x: x + 0.55, y, w: 5.4, h: 0.6, fontFace: F, fontSize: 11.5, color: C.text, valign: "middle" });
  });
  const r = [["7개 이상", "PoC 바로 착수 가능"], ["4~6개", "진단·데이터 정비부터"], ["3개 이하", "기초 전산화·용어 정리 선행"]];
  r.forEach((v, i) => kpi(s, 0.5 + i * 4.16, 5.5, 4.0, 0.85, v[0], v[1], { vs: 15, ls: 11 }));
  source(s, "코난테크놀로지 제안 (판정 구간은 제안값) — 세미나 사전 설문으로도 활용 가능");
  s.addNotes("[58:00~59:00] 현장에서 바로 체크해 보시도록 시간을 30초 드립니다. 이 체크리스트는 세미나 사전·사후 설문으로도 쓸 수 있고, 원하시는 기업은 1:1 진단으로 이어갈 수 있다고 안내합니다.");
}

// ───────────────────────── 29. 맺음
{
  const s = pptx.addSlide();
  pageNo++;
  s.background = { color: C.bg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText("맺음말", { x: 0.9, y: 0.8, w: 6, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: C.accent });
  s.addText("AI는 명장을 대체하지 않습니다.\n명장의 판단을 회사의 자산으로 남깁니다.", { x: 0.9, y: 1.3, w: 11.5, h: 1.5, fontFace: F, fontSize: 30, bold: true, color: C.white });
  const m = [
    ["1", "성패는 모델이 아니라 데이터와 지식(온톨로지)에서 갈립니다"],
    ["2", "AI는 근거 있는 초안을, 확정은 사람이 — 신뢰가 먼저입니다"],
    ["3", "업무 1건 · 기준선 측정 · 스코어카드로 작게 시작하십시오"],
  ];
  m.forEach((t, i) => {
    const y = 3.25 + i * 0.75;
    s.addText(t[0], { x: 0.9, y, w: 0.6, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: C.accent2, valign: "middle" });
    s.addText(t[1], { x: 1.5, y, w: 11, h: 0.6, fontFace: F, fontSize: 17, color: C.text, valign: "middle" });
  });
  s.addShape(pptx.ShapeType.line, { x: 0.9, y: 5.75, w: 11.5, h: 0, line: { color: C.line, width: 1 } });
  s.addText("㈜코난테크놀로지  |  www.konantech.com  |  konan@konantech.com  |  1533-1614", { x: 0.9, y: 5.95, w: 11.5, h: 0.4, fontFace: F, fontSize: 13, color: C.muted });
  s.addText("Q & A", { x: 10.4, y: 6.5, w: 2.4, h: 0.6, fontFace: F, fontSize: 24, bold: true, color: C.accent, align: "right" });
  s.addNotes("[59:00~60:00] 세 가지 메시지로 정리하고 Q&A. 다음 세션(AI 정부지원)으로 연결 멘트: '오늘 말씀드린 STEP 1을 정부 지원사업과 어떻게 연결하는지는 조윤성 박사님께서 이어서 설명해 주십니다.'");
}

// ───────────────────────── 30. 부록: 용어 풀이
{
  const s = base("부록", "용어 풀이 (CEO용)", "강의 중 나온 기술 용어를 한 줄로 정리했습니다");
  const rows = [
    [TH("용어"), TH("쉬운 설명")],
    ["AX (AI Transformation)", "AI로 업무 방식 자체를 바꾸는 것. 도구 도입이 아니라 프로세스 전환"],
    ["암묵지 / 형식지", "경험·직관으로만 아는 지식 / 문서로 적힌 지식"],
    ["온톨로지", "회사 업무의 개념과 관계를 정한 '설계도' (무엇이 무엇의 원인인가)"],
    ["지식그래프", "온톨로지에 따라 실제 데이터를 점(개체)과 선(관계)으로 연결한 지식 지도"],
    ["RAG / Graph-RAG", "문서를 찾아 근거로 답하는 AI / 지식그래프의 관계를 따라 추론해 답하는 AI"],
    ["AAS (자산관리셸)", "설비·자산 정보를 담는 국제표준(IEC 63278-1) 디지털 명함"],
    ["AI Agent", "업무 하나를 맡아 여러 단계를 스스로 처리하는 AI"],
    ["HITL (Human-in-the-Loop)", "AI 결과를 사람이 검토·승인하는 구조"],
    ["LLMOps · RAGOps · AgentOps", "모델 · 검색 · Agent를 도입 후에도 관리하고 개선하는 운영 체계"],
    ["온프레미스", "외부 클라우드가 아닌 사내 서버에서 운영 — 데이터 외부 반출 없음"],
  ];
  s.addTable(rows, { x: 0.5, y: 1.8, w: 12.33, colW: [3.3, 9.03], fontFace: F, fontSize: 12, color: C.text, fill: { color: C.panel }, border: { type: "solid", color: C.line, pt: 1 }, rowH: 0.44, valign: "middle" });
  s.addNotes("배포용 부록. 강의 중 넘기지 않고 질문이 나오면 참고 자료로 사용.");
}

pptx.writeFile({ fileName: process.argv[2] || "deck.pptx" }).then((f) => console.log("written", f));
