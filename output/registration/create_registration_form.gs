/**
 * 경남 제조AX 세미나 2026 — 참석 접수 폼 + 응답 구글시트 자동 생성
 * 사용법: script.google.com → 새 프로젝트 → 이 코드를 붙여넣기 → createSeminarForm 실행 → 실행 로그의 주소 확인
 */
function createSeminarForm() {
  var form = FormApp.create('경남 제조AX 세미나 2026 참석 접수');
  form.setDescription(
    '코난테크놀로지 경남(창원)사무소 개소 기념 「경남 제조AX 세미나」 참석 접수입니다.\n\n' +
    '· 일시: 2026년 12월 10일(목) 14:00 (등록 13:30부터)\n' +
    '· 장소: 창원컨벤션센터(CECO) 중회의실\n' +
    '· 만찬: 17:45, 같은 장소 뷔페\n' +
    '· 회신 마감: 11월 27일(금)\n\n' +
    '초청받으신 분만 접수하실 수 있습니다 (50명 한정).');
  form.setCollectEmail(false);
  form.setConfirmationMessage('접수되었습니다. 12월 9일(수)에 참석 안내 문자를 보내드리겠습니다. 감사합니다.');

  form.addTextItem().setTitle('성명').setRequired(true);
  form.addTextItem().setTitle('소속 (회사 · 기관명)').setRequired(true);
  form.addTextItem().setTitle('부서 · 직위').setRequired(true);
  form.addTextItem().setTitle('휴대전화').setHelpText('예: 010-0000-0000 (참석 안내 문자 발송용)').setRequired(true);
  form.addTextItem().setTitle('이메일').setRequired(true);
  form.addMultipleChoiceItem().setTitle('소재 지역')
      .setChoiceValues(['창원 · 마산', '사천', '거제', '기타']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('세미나 참석 여부')
      .setChoiceValues(['참석', '불참', '대리 참석 (아래 비고에 대리인 성명 기재)']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('만찬 참석 여부 (17:45, 뷔페)')
      .setChoiceValues(['참석', '불참']).setRequired(true);
  form.addTextItem().setTitle('차량번호').setHelpText('주차 등록이 필요한 경우에만 적어 주세요.');
  form.addCheckboxItem().setTitle('관심 있는 세션 (복수 선택)')
      .setChoiceValues([
        '세션 1. 제조 AX 도입 4단계 전략',
        '세션 2. 온톨로지 자동화 설계 기술 통한 제조 현장 AI 지식 구축',
        '세션 3. VLM 도면이해 기술 기반의 AX 서비스',
        '세션 4. OpenCode 하네스 기반 Agent 개발 플랫폼',
        '세션 5. 온디바이스 기반 전자교범 AI 기술 소개',
        '세션 6. 2027년 정부과제 및 지역 지원사업 소개']);
  form.addMultipleChoiceItem().setTitle('1:1 상담 희망 여부')
      .setChoiceValues(['희망', '희망하지 않음']).setRequired(true);
  form.addParagraphTextItem().setTitle('상담 희망 주제 · 비고');
  form.addCheckboxItem().setTitle('개인정보 수집 · 이용 동의')
      .setHelpText('수집 항목: 성명, 소속, 직위, 연락처, 이메일, 차량번호 / 목적: 행사 접수 · 안내 · 상담 연락 / 보유 기간: 행사 종료 후 6개월')
      .setChoiceValues(['동의합니다']).setRequired(true);

  // 응답을 받을 구글시트 생성 · 연결
  var ss = SpreadsheetApp.create('경남 제조AX 세미나 2026 접수 명단');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('접수 페이지(QR에 넣을 주소): ' + form.getPublishedUrl());
  Logger.log('폼 편집 주소: ' + form.getEditUrl());
  Logger.log('응답 구글시트: ' + ss.getUrl());
}
