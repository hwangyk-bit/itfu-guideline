/**
 * 잇퓨 캠페인 신청 폼 → 구글시트 연동 스크립트
 *
 * [대상 시트]
 * https://docs.google.com/spreadsheets/d/16dX_zSESxgckgTNBdN26VfjqZYuAX5FCVS5kC1QNkfY/edit
 * 탭 이름: 시트1  (없으면 자동 생성)
 *
 * [배포 방법]
 * 1. 위 구글시트를 엽니다.
 * 2. 상단 메뉴 "확장 프로그램 > Apps Script" 클릭
 * 3. 기본 생성된 코드(myFunction)를 모두 지우고 이 파일 내용을 붙여넣기
 * 4. 저장 (Ctrl+S)
 * 5. 우측 상단 "배포 > 새 배포" 클릭
 *    - 유형 선택(톱니바퀴): "웹 앱"
 *    - 설명: 아무거나 (예: itfu-v1)
 *    - 실행 계정: 나(본인 계정)
 *    - 액세스 권한이 있는 사용자: "모든 사용자"   ← 반드시 이걸로
 *    - "배포" 클릭 → 권한 승인 (본인 계정 선택 → 고급 → 이동 → 허용)
 * 6. 배포 후 나오는 "웹 앱 URL" 복사
 *    (https://script.google.com/macros/s/AKfycb..../exec 형태)
 * 7. 잇퓨 사이트 3개 페이지(itfu/index.html, campaign-02/index.html,
 *    campaign-03/index.html)에서 아래 줄을 찾아 URL을 붙여넣기
 *      const GAS_WEB_APP_URL = '';
 *    →  const GAS_WEB_APP_URL = '복사한_웹앱_URL';
 * 8. 배포 (vercel deploy --prod)
 *
 * [동작 방식]
 * - 탭을 나누지 않고 "시트1" 한 곳에 모든 신청이 쌓입니다.
 * - 각 페이지가 보낸 등급(tier)·고료(price)가 행마다 컬럼으로 기록되므로,
 *   시트1에서 고료 컬럼으로 필터/정렬하면 고료별로 구분해 볼 수 있습니다.
 *     itfu        → 등급 A / 고료 5만원
 *     campaign-02 → 등급 B / 고료 10만원
 *     campaign-03 → 등급 C / 고료 20만원
 * - 컬럼 순서
 *     A 제출일시 | B 등급 | C 고료 | D 이름 | E 인스타그램 | F 연락처 |
 *     G 이메일 | H 우편번호 | I 주소 | J 요청사항 | K 유입경로
 * - 연락처·우편번호는 앞자리 0이 사라지지 않도록 텍스트로 저장됩니다.
 * - 같은 사람이 같은 페이지로 2번 제출하면 그대로 2줄 쌓입니다(중복 확인은 시트에서).
 *
 * [코드를 수정한 뒤 다시 배포하는 법]
 * "배포 > 배포 관리" → 연필(수정) → 버전 "새 버전" → 배포
 * (웹 앱 URL은 그대로 유지되므로 사이트 코드는 다시 안 바꿔도 됩니다)
 */

var SHEET_NAME = '시트1';
var HEADERS = ['제출일시', '등급', '고료', '이름', '인스타그램', '연락처', '이메일', '우편번호', '주소', '요청사항', '유입경로'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // 동시에 여러 명이 제출해도 같은 줄에 겹쳐 쓰이지 않도록 잠금
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('요청 본문이 비어 있습니다.');
    }
    var data = JSON.parse(e.postData.contents);

    var sheet = getSheet_();

    var row = [
      new Date(),
      toText_(data.tier, '기타'),
      toText_(data.price, '미지정'),
      toText_(data.name),
      toText_(data.instagram),
      "'" + toText_(data.phone),    // 앞자리 0 보존
      toText_(data.email),
      "'" + toText_(data.zipcode),  // 앞자리 0 보존
      toText_(data.address),
      toText_(data.note),
      toText_(data.page)
    ];
    sheet.appendRow(row);

    // 방금 추가된 행의 연락처(F)·우편번호(H)를 텍스트 서식으로 고정
    var r = sheet.getLastRow();
    sheet.getRange(r, 6).setNumberFormat('@');
    sheet.getRange(r, 8).setNumberFormat('@');

    return json_({ result: 'success', row: r });
  } catch (err) {
    // 실패해도 시트에 기록을 남겨 원인을 추적할 수 있게 함
    try {
      logError_(err, e);
    } catch (ignored) {}
    return json_({ result: 'error', message: String(err && err.message ? err.message : err) });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

function doGet() {
  return json_({ status: 'ok', message: '잇퓨 캠페인 신청 웹훅이 정상 동작 중입니다.' });
}

/* ---------- 내부 함수 ---------- */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // 완전히 빈 시트일 때만 헤더를 넣습니다.
  // (이미 데이터가 있는 시트라면 기존 컬럼 순서를 그대로 두고 아래에 이어 붙입니다)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150); // 제출일시
    sheet.setColumnWidth(5, 260); // 인스타그램
    sheet.setColumnWidth(9, 300); // 주소
  }

  return sheet;
}

function toText_(value, fallback) {
  if (value === null || value === undefined || value === '') {
    return fallback || '';
  }
  return String(value);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError_(err, e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('오류로그');
  if (!sheet) {
    sheet = ss.insertSheet('오류로그');
    sheet.appendRow(['시각', '메시지', '요청본문']);
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(),
    String(err && err.message ? err.message : err),
    e && e.postData ? String(e.postData.contents).slice(0, 5000) : ''
  ]);
}

/**
 * 배포 전 테스트용 — Apps Script 편집기에서 이 함수를 직접 실행하면
 * 시트1에 테스트 행이 1줄 추가됩니다. 확인 후 그 줄은 지우세요.
 */
function testAppend() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        campaign: '잇퓨 순율앰플 추석 이벤트',
        tier: 'A',
        price: '5만원',
        name: '테스트',
        instagram: 'https://instagram.com/itfu_test',
        phone: '010-1234-5678',
        email: 'test@example.com',
        zipcode: '06236',
        address: '서울 강남구 테헤란로 1 101동 1001호',
        note: '테스트 제출',
        page: '/itfu'
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}
