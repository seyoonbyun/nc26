/**
 * Accelerate 2026 — VIP 시크릿 패스 (MVP 한정) 트래킹 시스템
 *
 * 동작 흐름:
 *   1. 운영자가 '시크릿패스' 시트탭에 대상자 행을 수기로 추가 (B열 Name, C열 Email)
 *   2. onEditSecretPass 트리거가 A열에 고유 토큰(SP-XXXXXX) 자동 발급
 *   3. 운영자가 https://www.nc26-bnikorea.com/secret?t=토큰 형태로 카톡/메일 발송
 *   4. VIP가 페이지 접속 시 secret.html → verifySecretPass(token) 호출
 *      → Visited 컬럼에 첫 방문 시각 기록, 이후 방문 시 LastVisited 갱신
 *   5. 결제 완료 = scanSecretPassPaid 1분 time-driven 트리거 (이 파일 내부,
 *      self-contained) 가 자동 매칭 — Email 정확 일치 기준으로
 *      SecretPass.Paid(H) 에 '결제 완료 ' + 타임스탬프 기록 (멱등)
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

// ────────────────────────────────────────────────────────────────────────
// Last Call 공유 페이지 — 토큰 없이 접근, 단순 방문 카운터
// ────────────────────────────────────────────────────────────────────────
var LASTCALL_SHEET_NAME = 'LastCallVisits';
var LASTCALL_HEADERS = ['VisitedAt', 'Referrer', 'UserAgent'];

function getOrCreateLastCallSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LASTCALL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LASTCALL_SHEET_NAME);
    sheet.appendRow(LASTCALL_HEADERS);
    var headerRange = sheet.getRange(1, 1, 1, LASTCALL_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#cf1f2e');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 400);
  }
  return sheet;
}

/**
 * Last Call 방문 로깅
 *   GET ?action=logLastCall&ref=...
 *   응답: { ok: true } (만료 시 expired:true)
 */
function logLastCall(referrer, userAgent) {
  var result = { ok: false };
  var expiry = new Date('2026-05-15T15:00:00Z').getTime();
  if (Date.now() >= expiry) {
    result.expired = true;
    return _jsonResponse(result);
  }
  try {
    var sheet = getOrCreateLastCallSheet();
    var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([stamp, referrer || '', userAgent || '']);
    result.ok = true;
  } catch (e) {
    result.error = String(e);
  }
  return _jsonResponse(result);
}

function setupLastCallSheet() {
  var sheet = getOrCreateLastCallSheet();
  Logger.log('Last call sheet ready: ' + sheet.getName());
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
 * Self-contained. full.js 같은 다른 스크립트에 의존하지 않음.
 *
 * 결제 시트 컬럼:
 *   A 상품명 / B applicantName / C applicantEmail / D applicantPhone
 *   E applicantRegion / F applyChapter / G formLinkpayID / H createdAt
 *   I modifiedAt / J statusSubmit / K orderID / L statusPayment / M isDelete
 *
 * 조건: A열 상품명에 '시크릿' / 'MVP' / 'SECRET' 포함 + L열 statusPayment 가 결제완료
 * 매칭: SecretPass 시트의 Email(C) 정확 일치 (소문자 normalize)
 * 결과: SecretPass.Paid(H, 8번 열) ← '결제 완료 yyyy-MM-dd HH:mm:ss' (멱등)
 *
 * 트리거 등록 (Apps Script 콘솔 시계 아이콘):
 *   - 함수: scanSecretPassPaid
 *   - 이벤트 소스: 시간 기반
 *   - 유형: 분 타이머 / 1분 (또는 5분)
 *
 * 수동 실행은 동일 함수 ▶ 한 번 — backfill 효과까지 같이 남.
 */

function _isSecretPassProduct(productName) {
  if (!productName) return false;
  var p = productName.toString();
  return p.indexOf('시크릿') !== -1 || p.toUpperCase().indexOf('MVP') !== -1 || p.toUpperCase().indexOf('SECRET') !== -1;
}

function _spIsPaymentComplete(v) {
  if (v === null || v === undefined) return false;
  var s = String(v).replace(/\s+/g, '').toLowerCase();
  return s === '결제완료' || s === 'paid' || s === '완료';
}

function _spNormalizeEmail(v) {
  return (v || '').toString().trim().toLowerCase();
}

/**
 * 1분 time-driven 트리거 — 결제 시트 전체 스캔 → 시크릿 상품 결제완료 행에
 * 대해 SecretPass.Paid 컬럼 자동 마킹 (멱등)
 *
 * 첫 1회 수동 실행 시 backfill 효과까지 같이 발생.
 */
function scanSecretPassPaid() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('scanSecretPassPaid 스킵: 다른 실행 진행 중');
    return;
  }
  try {
    SpreadsheetApp.flush();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paySheet = ss.getSheetByName('Ticket & Booth_Kor.pay');
    if (!paySheet) { Logger.log('Pay sheet not found'); return; }
    if (paySheet.getLastRow() < 2) { Logger.log('No pay rows.'); return; }

    var secretSheet = ss.getSheetByName(SECRET_SHEET_NAME);
    if (!secretSheet || secretSheet.getLastRow() < 2) {
      Logger.log('SecretPass sheet empty');
      return;
    }

    var secretData = secretSheet.getRange(2, 1, secretSheet.getLastRow() - 1, 9).getValues();
    var secretByEmail = {};
    for (var s = 0; s < secretData.length; s++) {
      var sEm = _spNormalizeEmail(secretData[s][2]);
      if (!sEm) continue;
      var sPaid = String(secretData[s][7] || '').trim();
      secretByEmail[sEm] = { rowIdx: s, paidAlready: sPaid.indexOf('결제 완료') === 0 };
    }

    var data = paySheet.getRange(2, 1, paySheet.getLastRow() - 1, 13).getDisplayValues();
    var marked = 0, alreadyPaid = 0, noMatch = 0, skippedProduct = 0, skippedStatus = 0;
    for (var i = 0; i < data.length; i++) {
      var product = String(data[i][0] || '').trim();
      var status = data[i][11];
      if (!_spIsPaymentComplete(status)) { skippedStatus++; continue; }
      if (!_isSecretPassProduct(product)) { skippedProduct++; continue; }

      var em = _spNormalizeEmail(data[i][2]);
      var hit = em && secretByEmail[em];
      if (!hit) {
        noMatch++;
        Logger.log('No SecretPass match: ' + em + ' (.pay 행 ' + (i + 2) + ')');
        continue;
      }
      if (hit.paidAlready) { alreadyPaid++; continue; }

      var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      secretSheet.getRange(hit.rowIdx + 2, 8).setValue('결제 완료 ' + stamp);
      hit.paidAlready = true;
      marked++;
      Logger.log('SecretPass Paid 기록: ' + em + ' @ ' + stamp);
    }
    Logger.log('scanSecretPassPaid — marked: ' + marked + ', already: ' + alreadyPaid + ', no match: ' + noMatch + ', skip status: ' + skippedStatus + ', skip product: ' + skippedProduct);
  } finally {
    lock.releaseLock();
  }
}

/** backwards-compat alias — 기존에 backfillSecretPassPaid 로 부르던 곳용 */
function backfillSecretPassPaid() {
  scanSecretPassPaid();
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
