/**
 * Accelerate 2026 — VIP 시크릿 패스 (MVP 한정) 트래킹 시스템
 *
 * 동작 흐름:
 *   1. 운영자가 '시크릿패스' 시트탭에 대상자 행을 수기로 추가 (B열 Name, C열 Email)
 *   2. onEditSecretPass 트리거가 A열에 고유 토큰(SP-XXXXXX) 자동 발급
 *   3. 운영자가 https://www.nc26-bnikorea.com/secret?t=토큰 형태로 카톡/메일 발송
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
 * ────────────────────────────────────────────────────────────────────────
 * Paid 자동 처리 — 결제 시트(Ticket & Booth_Kor.pay)와 SecretPass 매칭
 * ────────────────────────────────────────────────────────────────────────
 * 결제 시트 컬럼:
 *   A 상품명 / B applicantName / C applicantEmail / D applicantPhone
 *   E applicantRegion / F applyChapter / G formLinkpayID / H createdAt
 *   I modifiedAt / J statusSubmit / K orderID / L statusPayment / M isDelete
 *
 * 조건: A열 상품명에 '시크릿' 또는 'MVP' 포함 + L열 statusPayment = '결제 완료'
 * 매칭: SecretPass 시트의 Email(C) 우선, 미일치 시 Name+Phone fallback
 * 결과: SecretPass.Paid(H, 8번 열)에 '결제 완료 yyyy-MM-dd HH:mm:ss' 기록
 */
var PAY_SHEET_NAME = 'Ticket & Booth_Kor.pay';
var PAY_STATUS_COL = 12; // L열 statusPayment

function _isSecretPassProduct(productName) {
  if (!productName) return false;
  var p = productName.toString();
  return p.indexOf('시크릿') !== -1 || p.toUpperCase().indexOf('MVP') !== -1 || p.toUpperCase().indexOf('SECRET') !== -1;
}

function _normalizePhone(v) {
  if (v === null || v === undefined) return '';
  var digits = v.toString().replace(/\D/g, '');
  if (digits.length > 10) digits = digits.slice(-10);
  return digits;
}

function _normalizeEmail(v) {
  return (v || '').toString().trim().toLowerCase();
}

/**
 * SecretPass 행 매칭 → 일치 시 Paid 컬럼 갱신
 * 반환: 매칭 성공 여부 (true/false)
 */
function _markSecretPassPaid(name, email, phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SECRET_SHEET_NAME);
  if (!sheet) return false;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowName = (data[i][1] || '').toString().trim();
    var rowEmail = _normalizeEmail(data[i][2]);
    // Memo 컬럼(I, index 8)에 '010-XXXX-XXXX' 형태로 저장됨
    var rowPhone = _normalizePhone(data[i][8]);
    var rowPaid = (data[i][7] || '').toString().trim();

    var emailMatch = email && rowEmail && (email === rowEmail);
    var phoneTail = phone ? phone.slice(-8) : '';
    var rowPhoneTail = rowPhone ? rowPhone.slice(-8) : '';
    var phoneMatch = phoneTail && rowPhoneTail && (phoneTail === rowPhoneTail);
    var nameMatch = name && rowName && (name === rowName);

    if (emailMatch || (nameMatch && phoneMatch)) {
      if (rowPaid && rowPaid.indexOf('결제 완료') === 0) {
        Logger.log('SecretPass already paid (row ' + (i + 2) + '): ' + email);
        return true; // 중복 갱신 방지
      }
      var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(i + 2, 8).setValue('결제 완료 ' + stamp);
      Logger.log('SecretPass paid marked (row ' + (i + 2) + '): ' + email);
      return true;
    }
  }
  Logger.log('No SecretPass match for ' + email + ' / ' + name + ' / ' + phone);
  return false;
}

/**
 * onEdit 트리거 — 결제 시트 L열이 '결제 완료'로 변경되면 자동 매칭
 *
 * 트리거 등록: 함수 onEditPaymentToSecretPass / 이벤트 소스: 스프레드시트 / 유형: 수정 시
 */
function onEditPaymentToSecretPass(e) {
  if (!e || !e.range) return;
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== PAY_SHEET_NAME) return;

  var range = e.range;
  if (range.getColumn() !== PAY_STATUS_COL) return;
  var row = range.getRow();
  if (row <= 1) return;

  var status = (range.getValue() || '').toString().trim();
  if (status !== '결제 완료') return;

  var rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
  var product = (rowData[0] || '').toString().trim();
  if (!_isSecretPassProduct(product)) return;

  var name = (rowData[1] || '').toString().trim();
  var email = _normalizeEmail(rowData[2]);
  var phone = _normalizePhone(rowData[3]);

  _markSecretPassPaid(name, email, phone);
}

/**
 * 결제 시트의 기존 '결제 완료' 행을 일괄 스캔해 Paid 처리
 * — 트리거 등록 전에 이미 결제된 건이 있거나, 검증용으로 한 번 실행
 */
function backfillSecretPassPaid() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAY_SHEET_NAME);
  if (!sheet) {
    Logger.log('Pay sheet not found: ' + PAY_SHEET_NAME);
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No pay rows.');
    return;
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var marked = 0;
  var skippedProduct = 0;
  var skippedStatus = 0;
  for (var i = 0; i < data.length; i++) {
    var product = (data[i][0] || '').toString().trim();
    var status = (data[i][11] || '').toString().trim();
    if (status !== '결제 완료') { skippedStatus++; continue; }
    if (!_isSecretPassProduct(product)) { skippedProduct++; continue; }

    var name = (data[i][1] || '').toString().trim();
    var email = _normalizeEmail(data[i][2]);
    var phone = _normalizePhone(data[i][3]);
    if (_markSecretPassPaid(name, email, phone)) marked++;
  }
  Logger.log('Paid backfill — marked: ' + marked + ', skipped (status): ' + skippedStatus + ', skipped (product): ' + skippedProduct);
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
