/**
 * Google Apps Script — 수강 신청 폼 → 구글 시트 연동
 *
 * ── 설정 방법 ──────────────────────────────────────────
 * 1. Google Sheets 열기 (새 시트 만들기)
 * 2. 상단 메뉴 > 확장 프로그램 > Apps Script
 * 3. 아래 코드 전체를 붙여넣기 (기존 내용 덮어쓰기)
 * 4. 상단 [배포] > [새 배포] 클릭
 *    - 유형: 웹 앱
 *    - 설명: (자유롭게 입력)
 *    - 다음 사용자로 실행: 나 (내 계정)
 *    - 액세스 권한: 모든 사용자
 * 5. [배포] 클릭 → 웹 앱 URL 복사
 * 6. index.html 의 SCRIPT_URL 값에 복사한 URL 붙여넣기
 * ───────────────────────────────────────────────────────
 */

// 데이터를 쓸 시트 이름 (기본: 첫 번째 시트)
const SHEET_NAME = '';  // 비워두면 첫 번째 시트 사용

function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME
      ? ss.getSheetByName(SHEET_NAME)
      : ss.getSheets()[0];

    // 헤더 행이 없으면 자동 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['신청일시', '이름', '이메일', '연락처', '강의명']);
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.submittedAt || new Date().toLocaleString('ko-KR', {timeZone: 'Asia/Seoul'}),
      data.name  || '',
      data.email || '',
      data.phone || '',
      data.course || 'Claude Code 3시간 완성 클래스'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 배포 테스트용 — Apps Script 에디터에서 직접 실행해 확인
function testWrite() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['신청일시', '이름', '이메일', '연락처', '강의명']);
  }
  sheet.appendRow([
    new Date().toLocaleString('ko-KR', {timeZone:'Asia/Seoul'}),
    '테스트',
    'test@example.com',
    '010-0000-0000',
    'Claude Code 3시간 완성 클래스'
  ]);
  Logger.log('✅ 테스트 행 추가 완료');
}
