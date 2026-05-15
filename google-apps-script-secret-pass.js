/**
 * Accelerate 2026 — VIP 시크릿 패스 (MVP 한정) 트래킹 시스템
 *
 * 동작 흐름:
 *   1. 운영자가 '시크릿패스' 시트탭에 대상자 행을 수기로 추가 (B열 Name, C열 Email)
 *   2. onEditSecretPass 트리거가 A열에 고유 토큰(SP-XXXXXX) 자동 발급
 *   3. 운영자가 https://nc26.bni-korea.com/secret?t=토큰 형태로 카톡/메일 발송
 *   4. VIP가 페이지 접속 시 secret.html → verifySecretPass(token) 호출
 *      → Visited 컬럼에 첫 방문 시각 기록, 이후 방문 시 LastVisited 갱신
 *   5. 결제 완료 여부는 운영자가 linkpay 내역과 매칭하여 Paid 컬럼에 수기로 기록
 *
 * 설치:
 *   1. 기존 Apps Script 프로젝트(Code.gs 옆)에 이 파일을 새 스크립트로 추가
 *   2. onEditSecretPass 함수를 설치형 트리거(onEdit)로 등록
 *      - 함수: onEditSecretPass / 이벤트 소스: 스프레드시트 / 유형: 수정 시
 *   3. 기존 doGet에 verifySecretPass 분기 병합 (아래 doGet 예시 참조)
 *   4. 배포(웹 앱) — 기존 SCRIPT_URL 그대로 재사용
 */

var SECRET_SHEET_NAME = 'SecretPass';
var SECRET_HEADERS = ['Token', 'Name', 'Email', 'Created', 'Visited', 'LastVisited', 'VisitCount', 'Paid', 'Memo'];

/**
 * '시크릿패스' 시트 가져오거나 생성
 */
function getOrCreateSecretSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SECRET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SECRET_SHEET_NAME);
    sheet.appendRow(SECRET_HEADERS);
    var headerRange = sheet.getRange(1, 1, 1, SECRET_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#cf1f2e');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // 컬럼 너비
    sheet.setColumnWidth(1, 130); // Token
    sheet.setColumnWidth(2, 140); // Name
    sheet.setColumnWidth(3, 220); // Email
    sheet.setColumnWidth(4, 160); // Created
    sheet.setColumnWidth(5, 160); // Visited
    sheet.setColumnWidth(6, 160); // LastVisited
    sheet.setColumnWidth(7, 90);  // VisitCount
    sheet.setColumnWidth(8, 110); // Paid
    sheet.setColumnWidth(9, 240); // Memo
  }
  return sheet;
}

/**
 * 8자리 고유 토큰 생성 (혼동 문자 0/O/1/I 제외)
 *   예) SP-A3X7K2M9
 */
function generateSecretToken() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var token = '';
  for (var i = 0; i < 8; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return 'SP-' + token;
}

/**
 * onEdit 트리거 — 시크릿패스 시트에 대상자 행이 추가되면 토큰 자동 발급
 *
 * 트리거 등록: 함수 onEditSecretPass / 이벤트 소스: 스프레드시트 / 유형: 수정 시
 */
function onEditSecretPass(e) {
  if (!e || !e.range) return;
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SECRET_SHEET_NAME) return;

  var range = e.range;
  var row = range.getRow();
  if (row <= 1) return;

  // B열(Name) 또는 C열(Email)이 입력되면 토큰 발급 검토
  var col = range.getColumn();
  if (col !== 2 && col !== 3) return;

  var tokenCell = sheet.getRange(row, 1);
  var existing = (tokenCell.getValue() || '').toString().trim();
  if (existing) return; // 이미 토큰이 있으면 무시

  // 이름/이메일 둘 다 비어있으면 발급하지 않음
  var name = (sheet.getRange(row, 2).getValue() || '').toString().trim();
  var email = (sheet.getRange(row, 3).getValue() || '').toString().trim();
  if (!name && !email) return;

  // 중복되지 않는 토큰 생성
  var allTokens = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
  var tokenSet = {};
  for (var i = 0; i < allTokens.length; i++) {
    var t = (allTokens[i][0] || '').toString().trim();
    if (t) tokenSet[t] = true;
  }
  var token;
  var attempt = 0;
  do {
    token = generateSecretToken();
    attempt++;
    if (attempt > 50) break;
  } while (tokenSet[token]);

  tokenCell.setValue(token);
  var createdAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(row, 4).setValue(createdAt);
  Logger.log('Issued secret pass token: ' + token + ' for ' + email);
}

/**
 * 시크릿 패스 토큰 검증 + 방문 로깅
 *   GET ?action=verifySecretPass&token=SP-XXXXXX
 *   응답: { valid: boolean, name?: string }
 */
function verifySecretPass(token) {
  var result = { valid: false };
  if (!token) {
    return _jsonResponse(result);
  }
  token = token.toString().trim().toUpperCase();

  // 만료 시각 검증 (2026-05-16 00:00 KST = UTC 2026-05-15T15:00:00Z)
  var expiry = new Date('2026-05-15T15:00:00Z').getTime();
  if (Date.now() >= expiry) {
    result.message = 'expired';
    return _jsonResponse(result);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SECRET_SHEET_NAME);
  if (!sheet) return _jsonResponse(result);

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowToken = (data[i][0] || '').toString().trim().toUpperCase();
    if (rowToken === token) {
      var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      // Visited 컬럼이 비어있으면 첫 방문 시각 기록
      if (!data[i][4]) sheet.getRange(i + 1, 5).setValue(now);
      // LastVisited / VisitCount 갱신
      sheet.getRange(i + 1, 6).setValue(now);
      var count = parseInt(data[i][6], 10) || 0;
      sheet.getRange(i + 1, 7).setValue(count + 1);

      result.valid = true;
      result.name = data[i][1] || '';
      return _jsonResponse(result);
    }
  }
  return _jsonResponse(result);
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ─────────────────────────────────────────────────────────────────────
 * 기존 doGet 병합 예시 — 이미 다른 분기들이 있는 경우 if 블록만 추가
 * ─────────────────────────────────────────────────────────────────────
 *
 * function doGet(e) {
 *   if (e && e.parameter) {
 *     if (e.parameter.action === 'verifyPartyCode') {
 *       return verifyPartyCode(e.parameter.code);
 *     }
 *     if (e.parameter.action === 'verifySecretPass') {
 *       return verifySecretPass(e.parameter.token);
 *     }
 *   }
 *   return ContentService
 *     .createTextOutput(JSON.stringify({ status: 'ok' }))
 *     .setMimeType(ContentService.MimeType.JSON);
 * }
 */

/**
 * 초기 셋업 헬퍼 — 한 번만 수동 실행하여 시트탭을 미리 만들어 둘 수 있음
 */
function setupSecretPassSheet() {
  var sheet = getOrCreateSecretSheet();
  Logger.log('Secret pass sheet ready: ' + sheet.getName());
}

/**
 * 이미 입력된 행들에 토큰을 일괄 발급
 * — 트리거 등록 전에 데이터를 먼저 입력한 경우 사용
 */
function backfillSecretPassTokens() {
  var sheet = getOrCreateSecretSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows.');
    return;
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var tokenSet = {};
  for (var i = 0; i < data.length; i++) {
    var t = (data[i][0] || '').toString().trim();
    if (t) tokenSet[t] = true;
  }
  var issued = 0;
  var skipped = 0;
  for (var i = 0; i < data.length; i++) {
    var row = i + 2;
    var existingToken = (data[i][0] || '').toString().trim();
    var name = (data[i][1] || '').toString().trim();
    var email = (data[i][2] || '').toString().trim();
    if (existingToken) { skipped++; continue; }
    if (!name && !email) { skipped++; continue; }

    var newToken;
    var attempt = 0;
    do {
      newToken = generateSecretToken();
      attempt++;
      if (attempt > 50) break;
    } while (tokenSet[newToken]);
    tokenSet[newToken] = true;

    sheet.getRange(row, 1).setValue(newToken);
    var createdAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(row, 4).setValue(createdAt);
    issued++;
  }
  Logger.log('Backfill complete — issued: ' + issued + ', skipped: ' + skipped);
}
