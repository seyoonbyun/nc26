var SS = SpreadsheetApp.getActiveSpreadsheet();
var ADMIN_EMAILS = 'hq@joy-bnikorea.com,admin@bni-korea.com,ksoh7512@gmail.com';
var ADMIN_EMAILS_KOR = 'hq@joy-bnikorea.com,admin@bni-korea.com';

var TICKET_HEADERS = [
  'Timestamp', 'Name', 'Nationality', 'Email', 'Phone',
  'Position', 'Plan', 'Price', 'Language', 'Memo', 'Status'
];

var BOOTH_HEADERS = [
  'Timestamp', 'Booth No', 'Company', 'Display Name', 'Owner', 'Address',
  'Phone', 'Fax', 'Homepage', 'Email',
  'Applicant Name', 'Applicant Phone', 'Applicant Email',
  'Country', 'Chapter', 'License', 'Price', 'Logo File', 'Ad File', 'Status'
];

var BOOTH_KOR_HEADERS = [
  'Timestamp', 'Booth No', 'Company', 'Owner', 'Address',
  'Phone', 'Fax', 'Homepage', 'Email',
  'Applicant Name', 'Chapter', 'Applicant Phone', 'Applicant Email',
  'License', 'Price', 'Logo File', 'Ad File', 'Status'
];

var ADMIN_BOOTHS = { 'A19': 'BNI Korea' };

var PARTY_CODE_HEADERS = ['Code', 'Email', 'Name', 'Created', 'Used', 'Reserved'];

function getOrCreateSheet(name, headers) {
  var sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#cf1f2e');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function asText(val) {
  if (!val) return '';
  return "'" + val;
}

function formatTimestamp(ts) {
  try {
    var d = new Date(ts);
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  } catch (err) {
    return ts || '';
  }
}

// Google 일시 서버 에러에 대응하기 위한 지수 백오프 재시도 래퍼.
// Sheets/Gmail API 호출을 이걸로 감싸면 503/timeout 등 transient 에러가
// 대부분 자동 복구됨.
function retryable(label, fn) {
  var delays = [500, 1500, 4000, 8000];
  var lastErr;
  for (var i = 0; i <= delays.length; i++) {
    try { return fn(); }
    catch (err) {
      lastErr = err;
      Logger.log('retry[' + label + '] attempt ' + (i + 1) + ' failed: ' + (err && err.message));
      if (i < delays.length) Utilities.sleep(delays[i]);
    }
  }
  throw lastErr;
}

function setHyperlink(sheet, row, col, url, label) {
  if (!url) return;
  var cell = sheet.getRange(row, col);
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(label)
    .setLinkUrl(url)
    .build();
  cell.setRichTextValue(richText);
}

function doPost(e) {
  try {
    var p = e.parameter;

    if (e.postData && e.postData.type && e.postData.type.indexOf('text/plain') > -1) {
      try {
        p = JSON.parse(e.postData.contents);
      } catch (err) {
        Logger.log('JSON parse error: ' + err.message);
      }
    }

    if (p.type === 'booth-files') {
      var sheetName = (p.boothType === 'domestic') ? 'Booth_Kor' : 'Booth';
      var sheet = SS.getSheetByName(sheetName);
      if (sheet) {
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var logoCol = headers.indexOf('Logo File') + 1;
        var adCol = headers.indexOf('Ad File') + 1;
        var companyCol = headers.indexOf('Company') + 1;
        var lastRow = sheet.getLastRow();
        var targetRow = -1;
        if (lastRow > 1) {
          var data = sheet.getRange(2, companyCol, lastRow - 1, 1).getValues();
          for (var i = data.length - 1; i >= 0; i--) {
            if (data[i][0] === p.company) { targetRow = i + 2; break; }
          }
        }
        if (targetRow > 0) {
          if (p.logoFileBase64 && logoCol > 0) {
            var logoUrl = saveFileToDrive(p.logoFileName || 'logo', p.logoFileBase64, p.company);
            if (logoUrl) setHyperlink(sheet, targetRow, logoCol, logoUrl, p.logoFileName || 'logo');
          }
          if (p.adFileBase64 && adCol > 0) {
            var adUrl = saveFileToDrive(p.adFileName || 'ad', p.adFileBase64, p.company);
            if (adUrl) setHyperlink(sheet, targetRow, adCol, adUrl, p.adFileName || 'ad');
          }
        }
      }
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'ok' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (p.type === 'booth') {
      var logoUrl = '';
      var logoName = '';
      var adUrl = '';
      var adName = '';

      if (p.logoFileBase64) {
        logoName = p.logoFileName || 'logo';
        logoUrl = saveFileToDrive(logoName, p.logoFileBase64, p.company);
      }
      if (p.adFileBase64) {
        adName = p.adFileName || 'ad';
        adUrl = saveFileToDrive(adName, p.adFileBase64, p.company);
      }

      if (p.boothType === 'domestic') {
        var sheet = getOrCreateSheet('Booth_Kor', BOOTH_KOR_HEADERS);
        sheet.appendRow([
          formatTimestamp(p.timestamp),
          p.boothNo || '',
          p.company || '',
          p.owner || '',
          p.address || '',
          asText(p.phone),
          asText(p.fax),
          p.homepage || '',
          p.email || '',
          p.applicantName || '',
          p.chapter || '',
          asText(p.applicantPhone),
          p.applicantEmail || '',
          p.license || '',
          p.price || '',
          logoUrl ? logoName : '',
          adUrl ? adName : '',
          'Pending'
        ]);

        var newRow = sheet.getLastRow();
        if (p.homepage) setHyperlink(sheet, newRow, 8, p.homepage, p.homepage);
        if (p.email) setHyperlink(sheet, newRow, 9, 'mailto:' + p.email, p.email);
        if (p.applicantEmail) setHyperlink(sheet, newRow, 13, 'mailto:' + p.applicantEmail, p.applicantEmail);
        if (logoUrl) setHyperlink(sheet, newRow, 16, logoUrl, logoName);
        if (adUrl) setHyperlink(sheet, newRow, 17, adUrl, adName);

      } else {
        var sheet = getOrCreateSheet('Booth', BOOTH_HEADERS);
        sheet.appendRow([
          formatTimestamp(p.timestamp),
          p.boothNo || '',
          p.company || '',
          p.displayName || '',
          p.owner || '',
          p.address || '',
          asText(p.phone),
          asText(p.fax),
          p.homepage || '',
          p.email || '',
          p.applicantName || '',
          asText(p.applicantPhone),
          p.applicantEmail || '',
          p.country || '',
          p.chapter || '',
          p.license || '',
          p.price || '',
          logoUrl ? logoName : '',
          adUrl ? adName : '',
          'Pending'
        ]);

        var newRow = sheet.getLastRow();
        if (p.homepage) setHyperlink(sheet, newRow, 9, p.homepage, p.homepage);
        if (p.email) setHyperlink(sheet, newRow, 10, 'mailto:' + p.email, p.email);
        if (p.applicantEmail) setHyperlink(sheet, newRow, 13, 'mailto:' + p.applicantEmail, p.applicantEmail);
        if (logoUrl) setHyperlink(sheet, newRow, 18, logoUrl, logoName);
        if (adUrl) setHyperlink(sheet, newRow, 19, adUrl, adName);
      }

    } else {
      var sheet = getOrCreateSheet('Tickets', TICKET_HEADERS);
      sheet.appendRow([
        formatTimestamp(p.timestamp),
        p.name || '',
        p.nationality || '',
        p.email || '',
        asText(p.phone),
        p.position || '',
        p.plan || '',
        p.planPrice || '',
        p.lang || '',
        p.memo || '',
        'Pending'
      ]);

      var newRow = sheet.getLastRow();
      if (p.email) {
        setHyperlink(sheet, newRow, 4, 'mailto:' + p.email, p.email);
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'ok' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function saveFileToDrive(fileName, base64Data, companyName) {
  try {
    var folder = getOrCreateFolder('NC26_Booth_Files');
    var fullName = (companyName ? companyName + '_' : '') + fileName;
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'application/octet-stream',
      fullName
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    Logger.log('File save error: ' + err.message);
    return '';
  }
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

// =============================================================
// doGet — Party code verification API
// =============================================================
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'verifyPartyCode') {
    return verifyPartyCode(e.parameter.code);
  }

  if (action === 'getBoothStatus') {
    return getBoothStatus();
  }

  // ----- 관리자 API (PropertiesService에 저장된 ADMIN_KEY 필요) -----
  if (action === 'adminBackfill')    return adminGuard(e, function() { backfillPartyCodeUsed(); return { status: 'ok', action: 'adminBackfill' }; });
  if (action === 'adminReadSheet')   return adminGuard(e, function() { return adminReadSheet(e.parameter); });
  if (action === 'adminUpdateCell')  return adminGuard(e, function() { return adminUpdateCell(e.parameter); });
  if (action === 'adminDeleteRow')   return adminGuard(e, function() { return adminDeleteRow(e.parameter); });
  if (action === 'adminSetUsed')     return adminGuard(e, function() { return adminSetUsed(e.parameter); });
  if (action === 'adminListSheets')  return adminGuard(e, function() { return adminListSheets(); });
  if (action === 'adminRunScan')     return adminGuard(e, function() { scanAndSendPartyCodes(); return { status: 'ok', action: 'adminRunScan' }; });

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'NC26 Overseas API is running.' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// =============================================================
// Booth status API — 배치도 판매 현황 조회 (public, 인증 불필요)
//   Status == 'Paid' 인 행만 sold 로 반환
//   ADMIN_BOOTHS 에 정의된 부스는 항상 admin 으로 포함
// =============================================================
function getBoothStatus() {
  var result = { status: 'ok', sold: [], admin: [] };
  try {
    ['Booth_Kor', 'Booth'].forEach(function(sheetName) {
      var sheet = SS.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) return;
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var boothCol = headers.indexOf('Booth No') + 1;
      var companyCol = headers.indexOf('Company') + 1;
      var statusCol = headers.indexOf('Status') + 1;
      if (!boothCol || !companyCol || !statusCol) return;
      var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (var i = 0; i < data.length; i++) {
        var boothNo = String(data[i][boothCol - 1] || '').trim();
        var company = String(data[i][companyCol - 1] || '').trim();
        var status = String(data[i][statusCol - 1] || '').trim();
        if (!boothNo) continue;
        if (status.toLowerCase() === 'paid') {
          result.sold.push({ booth: boothNo, company: company });
        }
      }
    });
    Object.keys(ADMIN_BOOTHS).forEach(function(b) {
      result.admin.push({ booth: b, company: ADMIN_BOOTHS[b] });
    });
  } catch (err) {
    result.status = 'error';
    result.message = String(err && err.message || err);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================
// 관리자 API — Claude가 WebFetch로 시트를 직접 읽고 쓰기 위한 엔드포인트
//
// 사용 전 1회 세팅: Apps Script 편집기에서 setupAdminKey() 실행 → 로그로 키 확인
// =============================================================
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function adminGuard(e, fn) {
  try {
    var storedKey = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
    var providedKey = e && e.parameter && e.parameter.key;
    if (!storedKey) return jsonOut_({ status: 'error', message: 'ADMIN_KEY not set. Run setupAdminKey() first.' });
    if (providedKey !== storedKey) return jsonOut_({ status: 'error', message: 'unauthorized' });
    var result = fn();
    return jsonOut_(result && result.status ? result : { status: 'ok', result: result });
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err && err.message || err) });
  }
}

// Apps Script 편집기에서 1회 수동 실행 → 생성된 키를 로그에서 복사해 Claude에게 전달
function setupAdminKey() {
  var key = 'nc26-' + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', key);
  Logger.log('ADMIN_KEY 설정 완료.');
  Logger.log('이 키를 Claude에게 전달하세요 ↓');
  Logger.log(key);
  return key;
}

function rotateAdminKey() { return setupAdminKey(); }

function revokeAdminKey() {
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_KEY');
  Logger.log('ADMIN_KEY 삭제됨.');
}

function adminListSheets() {
  var sheets = SS.getSheets().map(function(s) {
    return { name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn() };
  });
  return { sheets: sheets };
}

function adminReadSheet(params) {
  var name = params.sheet;
  if (!name) return { status: 'error', message: 'sheet required' };
  var sheet = SS.getSheetByName(name);
  if (!sheet) return { status: 'error', message: 'sheet not found: ' + name };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0) return { status: 'ok', sheet: name, rows: 0, data: [] };

  var startRow = parseInt(params.startRow || '1', 10);
  var limit = parseInt(params.limit || String(lastRow), 10);
  var numRows = Math.min(lastRow - startRow + 1, limit);
  if (numRows <= 0) return { status: 'ok', sheet: name, rows: 0, data: [] };

  var data = sheet.getRange(startRow, 1, numRows, lastCol).getDisplayValues();
  return { status: 'ok', sheet: name, startRow: startRow, rows: numRows, cols: lastCol, data: data };
}

function adminUpdateCell(params) {
  var name = params.sheet;
  var row = parseInt(params.row, 10);
  var col = parseInt(params.col, 10);
  if (!name || !row || !col) return { status: 'error', message: 'sheet, row, col required' };
  var sheet = SS.getSheetByName(name);
  if (!sheet) return { status: 'error', message: 'sheet not found' };
  sheet.getRange(row, col).setValue(params.value != null ? params.value : '');
  return { status: 'ok', sheet: name, row: row, col: col, value: params.value };
}

function adminDeleteRow(params) {
  var name = params.sheet;
  var row = parseInt(params.row, 10);
  if (!name || !row || row < 2) return { status: 'error', message: 'sheet, row (>=2) required' };
  var sheet = SS.getSheetByName(name);
  if (!sheet) return { status: 'error', message: 'sheet not found' };
  if (row > sheet.getLastRow()) return { status: 'error', message: 'row out of range' };
  sheet.deleteRow(row);
  return { status: 'ok', sheet: name, deletedRow: row };
}

// PartyCodes 특정 코드 행의 Used / Reserved 값을 직접 기록
function adminSetUsed(params) {
  var code = String(params.code || '').trim().toUpperCase();
  if (!code) return { status: 'error', message: 'code required' };
  var sheet = getOrCreatePartyCodeSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === code) {
      var row = i + 1;
      if (params.used !== undefined) sheet.getRange(row, 5).setValue(params.used);
      if (params.reserved !== undefined) sheet.getRange(row, 6).setValue(params.reserved);
      return { status: 'ok', code: code, row: row };
    }
  }
  return { status: 'error', message: 'code not found: ' + code };
}

// =============================================================
// onEdit trigger — Overseas payment confirmation email (Tickets, Booth, Booth_Kor)
// =============================================================
function onEditInstallable(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var sheetName = sheet.getName();
    var col = range.getColumn();
    var row = range.getRow();

    if (sheetName !== 'Tickets' && sheetName !== 'Booth' && sheetName !== 'Booth_Kor') return;
    if (row <= 1) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('Status') + 1;

    if (statusCol === 0 || col !== statusCol || range.getValue() !== 'Paid') return;

    var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    if (sheetName === 'Tickets') {
      sendTicketConfirmationEmail(headers, rowData);
    } else if (sheetName === 'Booth' || sheetName === 'Booth_Kor') {
      sendBoothVoucherEmail(headers, rowData, sheetName);
    }
  } catch (err) {
    Logger.log('onEditInstallable error: ' + err.message);
  }
}

// =============================================================
// Time-based trigger — Scan "Ticket & Booth_Kor.pay" sheet
//
// 1) Issue code: A=BNI K. Member Pass + L=결제완료 -> generate code + send email
// 2) Mark Used: A=Networking Party Pass + L=결제완료 -> update PartyCodes Used
//
// Trigger setup: Editor > Triggers > Add trigger
//   - Function: scanAndSendPartyCodes
//   - Event source: Time-driven
//   - Type: Minutes timer (1 min recommended)
// =============================================================

// "결제 완료" / "결제완료" 모두 허용 (공백·대소문자 무시)
function isPaymentComplete(v) {
  if (v === null || v === undefined) return false;
  var s = String(v).replace(/\s+/g, '').toLowerCase();
  return s === '결제완료' || s === 'paid' || s === '완료';
}

function normalizeEmail(v) {
  if (!v) return '';
  return String(v).trim().toLowerCase();
}

// 핵심 프로세스 딱 두 가지:
//   Case 1. Ticket & Booth_Kor.pay 에 BNI K. Member Pass + 결제 완료 → PartyCodes 에 코드 발급 + 메일 발송
//   Case 2. Ticket & Booth_Kor.pay 에 Networking Party Pass + 결제 완료 → PartyCodes 의 해당 이메일 Used 에 시각 기록
// 나머지(중복 감지·태깅 등)는 전부 제거. 이 두 동작만 무조건 정확하게 돌린다.
function scanAndSendPartyCodes() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('scanAndSendPartyCodes 스킵: 다른 실행 진행 중');
    return;
  }

  try {
    SpreadsheetApp.flush();

    var paySheet = SS.getSheetByName('Ticket & Booth_Kor.pay');
    if (!paySheet) { Logger.log('시트 없음: Ticket & Booth_Kor.pay'); return; }
    if (paySheet.getLastRow() <= 1) { Logger.log('Ticket & Booth_Kor.pay 데이터 없음'); return; }

    var codeSheet = getOrCreatePartyCodeSheet();

    var payCols = Math.max(paySheet.getLastColumn(), 12);
    var payData = paySheet.getRange(2, 1, paySheet.getLastRow() - 1, payCols).getDisplayValues();
    var codeData = codeSheet.getDataRange().getValues();

    // PartyCodes: email → codeData 인덱스 (sheet row = idx + 1)
    var codeByEmail = {};
    for (var i = 1; i < codeData.length; i++) {
      var em = normalizeEmail(codeData[i][1]);
      if (em) codeByEmail[em] = i;
    }

    var issued = 0;
    var used = 0;

    for (var r = 0; r < payData.length; r++) {
      var title = String(payData[r][0] || '').trim();
      var name = payData[r][1];
      var email = payData[r][2];
      var status = payData[r][11];
      var normEmail = normalizeEmail(email);

      if (!normEmail || !isPaymentComplete(status)) continue;

      // ── Case 1: BNI K. Member Pass + 결제 완료 → 코드 발급 + 메일 발송
      if (title === 'BNI K. Member Pass') {
        if (codeByEmail[normEmail] !== undefined) continue; // 이미 발급됨 → 스킵

        var code = generatePartyCode();
        var issuedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        codeSheet.appendRow([code, email, name, issuedAt, '', '', '']);
        SpreadsheetApp.flush();

        var newIdx = codeSheet.getLastRow() - 1;
        codeByEmail[normEmail] = newIdx;
        codeData[newIdx] = [code, email, name, issuedAt, '', '', ''];

        try {
          sendPartyCodeEmail(email, name, code);
          Logger.log('코드 발급+메일 발송: ' + email + ' / ' + code);
          issued++;
        } catch (err) {
          Logger.log('메일 발송 실패(코드는 생성됨): ' + email + ' — ' + (err && err.message));
        }
        continue;
      }

      // ── Case 2: Networking Party Pass + 결제 완료 → Used 기록
      if (title === 'Networking Party Pass') {
        var idx = codeByEmail[normEmail];
        if (idx === undefined) {
          Logger.log('경고: 코드 없는 이메일이 Networking Party Pass 결제완료 → ' + email + ' (.pay 행 ' + (r + 2) + ')');
          continue;
        }
        if (codeData[idx][4]) continue; // Used 이미 있음 → 멱등 스킵

        var usedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        codeSheet.getRange(idx + 1, 5).setValue(usedAt);
        codeData[idx][4] = usedAt;
        Logger.log('Used 기록: ' + email + ' @ ' + usedAt);
        used++;
      }
    }

    Logger.log('scan 완료: 발급 ' + issued + '건, Used ' + used + '건');
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
// 수동 백필 (Used 컬럼만)
// - Ticket & Booth_Kor.pay의 Networking Party Pass 결제완료 이메일에 대해
//   PartyCodes Used가 비어있으면 현재 시각으로 채움
// - Apps Script 편집기에서 backfillPartyCodeUsed 선택 후 실행
// =============================================================
function backfillPartyCodeUsed() {
  Logger.log('backfillPartyCodeUsed 시작');
  var sheet = SS.getSheetByName('Ticket & Booth_Kor.pay');
  if (!sheet) { Logger.log('시트 없음'); return; }
  var codeSheet = getOrCreatePartyCodeSheet();

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('Ticket & Booth_Kor.pay 데이터 없음'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getDisplayValues();
  var codeData = codeSheet.getDataRange().getValues();

  var emailToRow = {};
  for (var i = 1; i < codeData.length; i++) {
    var em = normalizeEmail(codeData[i][1]);
    if (em) emailToRow[em] = i + 1;
  }

  var paidEmails = {};
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0] || '').trim() !== 'Networking Party Pass') continue;
    if (!isPaymentComplete(data[r][11])) continue;
    var normEmail = normalizeEmail(data[r][2]);
    if (normEmail) paidEmails[normEmail] = true;
  }

  var updated = 0;
  var ts = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  Object.keys(paidEmails).forEach(function(em) {
    var targetRow = emailToRow[em];
    if (!targetRow) { Logger.log('PartyCodes에 코드 없음: ' + em); return; }
    var usedCell = codeSheet.getRange(targetRow, 5);
    if (!usedCell.getValue()) {
      usedCell.setValue(ts);
      updated++;
      Logger.log('Used 백필: ' + em + ' @ ' + ts);
    }
  });

  Logger.log('백필 완료: Used ' + updated + '건');
}

// =============================================================
// PartyCodes 중복 코드 감사 (read-only)
// - 같은 이메일로 여러 코드가 발급된 케이스(17:44~20:11 race condition 피해) 전수 조사
// - 로그에만 출력, 시트는 건드리지 않음
// - 실행: Apps Script 편집기에서 auditDuplicateCodes 선택 후 ▶
// =============================================================
function auditDuplicateCodes() {
  Logger.log('auditDuplicateCodes 시작');
  var codeSheet = getOrCreatePartyCodeSheet();
  var data = codeSheet.getDataRange().getValues();

  var rowsByEmail = {};
  for (var i = 1; i < data.length; i++) {
    var em = normalizeEmail(data[i][1]);
    if (!em) continue;
    if (!rowsByEmail[em]) rowsByEmail[em] = [];
    rowsByEmail[em].push({
      sheetRow: i + 1,
      code: data[i][0],
      name: data[i][2],
      created: data[i][3],
      used: data[i][4],
      reserved: data[i][5]
    });
  }

  var affectedEmails = 0;
  var totalDupRows = 0;
  Object.keys(rowsByEmail).forEach(function(em) {
    var rows = rowsByEmail[em];
    if (rows.length <= 1) return;
    affectedEmails++;
    totalDupRows += (rows.length - 1);
    Logger.log('중복 발급 감지: ' + em + ' → ' + rows.length + '개 코드');
    rows.forEach(function(r) {
      Logger.log('  · 행 ' + r.sheetRow + ' | ' + r.code + ' | Created=' + r.created + ' | Used=' + (r.used || '-') + ' | Reserved=' + (r.reserved || '-'));
    });
  });

  Logger.log('auditDuplicateCodes 완료: 영향 이메일 ' + affectedEmails + '건, 삭제 대상 ' + totalDupRows + '개 행 예상');
  Logger.log('(실제 정리는 cleanupDuplicateCodes 함수 실행)');
}

// =============================================================
// PartyCodes 중복 코드 정리 (destructive — 반드시 auditDuplicateCodes 로그로 확인 후 실행)
//
// 동작 플로우 (이메일 기준 원자적 처리):
//   1) 같은 이메일이 여러 코드 행으로 등록된 경우, keeper 1개 선정
//      - Used 또는 Reserved가 차있는 행을 우선 (유저가 실제 사용/검증한 코드)
//      - 없으면 Created가 가장 오래된 행 (유저가 먼저 받은 코드)
//   2) keeper 행에 나머지 행의 비어있지 않은 state(Used/Reserved/Duplicate) 머지
//      → consolidate write 실패 시 이 이메일 전체 스킵 (삭제·메일 모두 안 함)
//   3) keeper의 유효 코드로 정정 안내 이메일 발송 (sendClarificationEmail)
//      → 메일 실패 시 삭제 스킵 (유저에게 알리지 못한 채 데이터만 지우는 상황 방지)
//   4) 이메일 성공 시에만 나머지 행 삭제 대상에 등록
//   5) 전체 스캔 끝난 뒤 행 번호 내림차순으로 deleteRow 일괄 실행 (번호 밀림 방지)
//
// 안전장치:
//   - LockService.tryLock(5000) — scanAndSendPartyCodes와 동시 실행 차단
//   - 각 단계 실패 시 조기 return으로 이 이메일의 뒷 단계 스킵 (부분 실패 시 데이터 보존)
//   - retryable로 모든 Sheets/Gmail 호출 보호
// =============================================================
function cleanupDuplicateCodes() {
  _cleanupDuplicateCodesCore(false);
}

// 이메일 발송 건너뛰고 DB만 정리 — Gmail 할당량 소진 시 또는 수동으로 이미 안내 완료한 경우 사용.
// 작동: 이메일 발송 단계만 스킵. 그 외 (lock, consolidate, delete) 동일.
function cleanupDuplicateCodesNoEmail() {
  _cleanupDuplicateCodesCore(true);
}

function _cleanupDuplicateCodesCore(skipEmail) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('cleanupDuplicateCodes 스킵: 다른 실행이 진행 중');
    return;
  }
  try {
    Logger.log('cleanupDuplicateCodes 시작 (skipEmail=' + skipEmail + ')');
    var codeSheet = getOrCreatePartyCodeSheet();
    var data = retryable('cleanupDup read PartyCodes', function() {
      return codeSheet.getDataRange().getValues();
    });

    var rowsByEmail = {};
    for (var i = 1; i < data.length; i++) {
      var em = normalizeEmail(data[i][1]);
      if (!em) continue;
      if (!rowsByEmail[em]) rowsByEmail[em] = [];
      rowsByEmail[em].push({ sheetRow: i + 1, data: data[i].slice() });
    }

    var toDeleteRows = [];
    var emailsProcessed = 0;
    var consolidationWrites = 0;
    var emailsNotified = 0;

    Object.keys(rowsByEmail).forEach(function(em) {
      var rows = rowsByEmail[em];
      if (rows.length <= 1) return;

      // 1) Keeper 선정: Used/Reserved 있는 행 우선, 없으면 Created 가장 오래된 행
      var keeperIdx = -1;
      for (var k = 0; k < rows.length; k++) {
        if (rows[k].data[4] || rows[k].data[5]) { keeperIdx = k; break; }
      }
      if (keeperIdx === -1) {
        keeperIdx = 0;
        for (var k = 1; k < rows.length; k++) {
          if (String(rows[k].data[3] || '') < String(rows[keeperIdx].data[3] || '')) keeperIdx = k;
        }
      }

      var keeper = rows[keeperIdx];
      var dups = [];
      for (var k = 0; k < rows.length; k++) { if (k !== keeperIdx) dups.push(rows[k]); }

      // 2) State consolidate — keeper에 빈 필드만 dup에서 보충 (Used/Reserved만)
      var merged = {
        used: keeper.data[4],
        reserved: keeper.data[5]
      };
      dups.forEach(function(d) {
        if (!merged.used && d.data[4]) merged.used = d.data[4];
        if (!merged.reserved && d.data[5]) merged.reserved = d.data[5];
      });

      var changed = (merged.used !== keeper.data[4]) || (merged.reserved !== keeper.data[5]);

      if (changed) {
        try {
          retryable('consolidate keeper', function() {
            codeSheet.getRange(keeper.sheetRow, 5, 1, 2).setValues([[
              merged.used || '',
              merged.reserved || ''
            ]]);
          });
          Logger.log('Consolidate: ' + em + ' → keeper 행 ' + keeper.sheetRow + ' (code ' + keeper.data[0] + ')');
          consolidationWrites++;
        } catch (err) {
          Logger.log('Consolidate 실패, 이 이메일 전체 스킵: ' + em + ' — ' + (err && err.message));
          return;
        }
      } else {
        Logger.log('Keep: ' + em + ' → 행 ' + keeper.sheetRow + ' (code ' + keeper.data[0] + ')');
      }

      // 3) 정정 안내 이메일 발송 — 실패 시 이 이메일의 삭제는 스킵
      var userEmail = keeper.data[1];
      var userName = String(keeper.data[2] || '').trim() || '고객';
      var validCode = keeper.data[0];
      if (skipEmail) {
        Logger.log('이메일 발송 스킵(skipEmail 모드): ' + userEmail);
      } else {
        try {
          retryable('send clarification email', function() {
            sendClarificationEmail(userEmail, userName, validCode);
          });
          Logger.log('Clarification 발송 완료: ' + userEmail + ' → code ' + validCode);
          emailsNotified++;
        } catch (err) {
          Logger.log('Clarification 발송 실패, 이 이메일 삭제 스킵 (데이터 보존): ' + em + ' — ' + (err && err.message));
          return;
        }
      }

      // 4) 이메일까지 성공한 경우에만 dup 행을 삭제 대상에 등록
      dups.forEach(function(d) {
        toDeleteRows.push(d.sheetRow);
        Logger.log('  삭제 예정: 행 ' + d.sheetRow + ' · code ' + d.data[0]);
      });
      emailsProcessed++;
    });

    // 5) 내림차순으로 deleteRow 일괄 실행 (번호 밀림 방지)
    toDeleteRows.sort(function(a, b) { return b - a; });
    var deleted = 0;
    toDeleteRows.forEach(function(rowNum) {
      try {
        retryable('deleteRow ' + rowNum, function() { codeSheet.deleteRow(rowNum); });
        deleted++;
      } catch (err) {
        Logger.log('deleteRow ' + rowNum + ' 실패: ' + (err && err.message));
      }
    });

    Logger.log('cleanupDuplicateCodes 완료: 영향 이메일 ' + emailsProcessed + '건, consolidate ' + consolidationWrites + '건, 안내메일 ' + emailsNotified + '건, 삭제 ' + deleted + '/' + toDeleteRows.length + '건');
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
// Networking Party Pass - Code generation & verification
// =============================================================

function getOrCreatePartyCodeSheet() {
  var sheet = SS.getSheetByName('PartyCodes');
  if (!sheet) {
    sheet = SS.insertSheet('PartyCodes');
    sheet.appendRow(PARTY_CODE_HEADERS);
    sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).setBackground('#cf1f2e');
    sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Reserved(F) 헤더가 없으면 보정
  var headerRow = sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).getValues()[0];
  var applyHeader = function(col, label) {
    var cell = sheet.getRange(1, col);
    cell.setValue(label);
    cell.setFontWeight('bold');
    cell.setBackground('#cf1f2e');
    cell.setFontColor('#ffffff');
  };
  if (!headerRow[5]) applyHeader(6, 'Reserved');

  // 레거시 Duplicate(G) / EmailSent(H) 컬럼 자동 제거 (신규 로직 미사용)
  // 끝 열부터 먼저 지워야 인덱스 밀림 없음
  while (sheet.getLastColumn() > 6) {
    var extraHeader = String(sheet.getRange(1, sheet.getLastColumn()).getValue() || '').trim();
    if (extraHeader === 'Duplicate' || extraHeader === 'EmailSent' || extraHeader === '') {
      sheet.deleteColumn(sheet.getLastColumn());
    } else {
      break;
    }
  }
  return sheet;
}

function generatePartyCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code = '';
  for (var i = 0; i < 8; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return 'NP-' + code;
}

/**
 * 인증코드 검증 API (ticket.html에서 호출)
 *
 * 동작:
 *   - 코드가 없으면 valid:false
 *   - Used(E) 또는 Reserved(F)에 값이 있으면 → 이미 사용된 코드 (message:'used')
 *   - 둘 다 비어있으면 → Reserved(F)에 현재 시각 즉시 기록(락) + valid:true 반환
 *
 * LockService로 동시 요청을 직렬화하여 중복 결제 차단.
 * 최종 결제완료 시점은 scanAndSendPartyCodes가 Used(E)에 덮어씀.
 */
// 코드 검증 — PartyCodes 시트의 Used(E) 컬럼만 체크
// - 코드가 존재하고 Used가 비어있으면 valid='true' + name 반환
// - Used에 값이 있으면 valid='false', message='used'
// - 코드를 찾지 못하면 valid='false'
// Reserved 컬럼은 건드리지 않음 (읽기 전용 검증)
// Typebot 호환: valid를 문자열 'true'/'false'로 반환
function verifyPartyCode(code) {
  var result = { valid: 'false' };
  var out = function(r) {
    return ContentService
      .createTextOutput(JSON.stringify(r))
      .setMimeType(ContentService.MimeType.JSON);
  };

  if (!code) return out(result);
  code = String(code).trim().toUpperCase();

  var sheet = getOrCreatePartyCodeSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== code) continue;

    if (data[i][4]) {
      result.valid = 'false';
      result.message = 'used';
      return out(result);
    }

    result.valid = 'true';
    result.name = data[i][2];
    return out(result);
  }
  return out(result);
}

// 오늘 남은 Gmail 발송 할당량 확인 (digits만 반환)
function checkEmailQuota() {
  var remaining = MailApp.getRemainingDailyQuota();
  Logger.log('오늘 남은 이메일 발송 가능 수: ' + remaining);
  return remaining;
}

// cleanupDuplicateCodes에서 강덕자에게 보낼 정정 안내 메일을 어드민에게만 먼저 발송.
// 실제 유저에겐 발송되지 않음. 본문 검수 후 문제 없으면 cleanupDuplicateCodes 실행.
function previewClarificationEmail() {
  var name = '강덕자';
  var code = 'NP-DFUF52A7';
  var adminTo = ADMIN_EMAILS_KOR.split(',')[0].trim(); // 첫 어드민 주소 (hq@joy-bnikorea.com)
  sendClarificationEmail(adminTo, name, code);
  Logger.log('Preview 발송 완료: ' + adminTo + ' (cc: ' + ADMIN_EMAILS_KOR + '). 유저에겐 발송 안 됨.');
}

// 중복 발송 정정 안내 이메일 — cleanupDuplicateCodes에서만 사용
// 유저에게 "시스템 오류로 여러 코드가 발송되었고, 유효한 코드는 하나뿐" 이라고 명시적으로 안내
function sendClarificationEmail(email, name, code) {
  var subject = '[Accelerate 2026] Networking Party Pass 인증코드 재안내 (중복 발송 정정)';

  var body = '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"/></head>'
    + '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica Neue,Arial,sans-serif;">'
    + '<div style="max-width:600px;margin:40px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'
    + '<div style="background:linear-gradient(135deg,#cf1f2e,#a31824);padding:32px 40px;text-align:center;">'
    + '<h1 style="color:#ffffff;margin:0;font-size:22px;font-weight:800;letter-spacing:1px;">ACCELERATE 2026</h1>'
    + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:13px;">BNI Korea National Conference</p>'
    + '</div>'
    + '<div style="padding:40px;">'
    + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">\uC778\uC99D\uCF54\uB4DC \uC7AC\uC548\uB0B4</h2>'
    + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 20px;">'
    + '<strong>' + name + '</strong>\uB2D8, \uC548\uB155\uD558\uC138\uC694.<br/><br/>'
    + '\uC2DC\uC2A4\uD15C \uC624\uB958\uB85C \uC778\uD574 \uB3D9\uC77C \uC774\uBA54\uC77C\uB85C <strong>\uB3D9\uC77C\uD55C \uC778\uC99D\uCF54\uB4DC\uAC00 \uC5EC\uB7EC \uBC88 \uBC1C\uC1A1\uB418\uC5C8\uC744 \uC218 \uC788\uC2B5\uB2C8\uB2E4</strong>. \uBD88\uD3B8\uC744 \uB4DC\uB824 \uC9C4\uC2EC\uC73C\uB85C \uC0AC\uACFC\uB4DC\uB9BD\uB2C8\uB2E4.<br/><br/>'
    + '\uC544\uB798 <strong style="color:#cf1f2e;">\uD55C \uAC1C\uC758 \uCF54\uB4DC\uB9CC \uC720\uD6A8</strong>\uD558\uBA70, \uC774\uC804\uC5D0 \uBC1B\uC73C\uC2E0 \uB2E4\uB978 \uCF54\uB4DC\uB294 \uC2DC\uC2A4\uD15C\uC5D0\uC11C \uBB34\uD6A8 \uCC98\uB9AC\uB418\uC5C8\uC2B5\uB2C8\uB2E4. \uD2F0\uCF13 \uAD6C\uB9E4 \uC2DC \uAC00\uC7A5 \uC544\uB798 \uCF54\uB4DC\uB9CC \uC0AC\uC6A9\uD574 \uC8FC\uC138\uC694.</p>'
    + '<div style="background:#fffbeb;border:2px dashed #f59e0b;border-radius:12px;padding:28px;margin-bottom:24px;text-align:center;">'
    + '<p style="color:#b45309;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 12px;">\uC720\uD6A8\uD55C \uCF54\uB4DC (Valid Code)</p>'
    + '<p style="color:#1a1a1a;font-size:28px;font-weight:900;margin:0 0 8px;letter-spacing:4px;font-family:monospace;">' + code + '</p>'
    + '<p style="color:#999;font-size:12px;margin:0;">\uD2F0\uCF13 \uAD6C\uB9E4 \uD398\uC774\uC9C0\uC5D0\uC11C \uC774 \uCF54\uB4DC\uB97C \uC785\uB825\uD574 \uC8FC\uC138\uC694</p>'
    + '</div>'
    + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:24px;">'
    + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
    + '<strong>\uC0AC\uC6A9 \uBC29\uBC95:</strong><br/>'
    + '1. <a href="https://www.nc26-bnikorea.com" style="color:#cf1f2e;">www.nc26-bnikorea.com</a> \uC811\uC18D<br/>'
    + '2. Networking Party Pass \uCE74\uB4DC\uC758 <strong>\u201C\uC778\uC99D\uCF54\uB4DC \uC785\uB825\u201D</strong> \uBC84\uD2BC \uD074\uB9AD<br/>'
    + '3. \uC704 \uC720\uD6A8\uD55C \uCF54\uB4DC \uC785\uB825 \uD6C4 \uACB0\uC81C \uC9C4\uD589</p>'
    + '</div>'
    + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
    + '\uBB38\uC758\uC0AC\uD56D: '
    + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
    + ' \uB610\uB294 <a href="http://pf.kakao.com/_xewxmrT/chat" style="color:#cf1f2e;">\uCE74\uCE74\uC624\uD1A1 \uCC44\uD305</a></p>'
    + '</div>'
    + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
    + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
    + '</div></div></body></html>';

  MailApp.sendEmail({ to: email, cc: ADMIN_EMAILS_KOR, subject: subject, htmlBody: body });
  Logger.log('Clarification email sent to: ' + email + ' with code: ' + code);
}

function sendPartyCodeEmail(email, name, code) {
  var subject = '[Accelerate 2026] Networking Party Pass 구매 인증코드';

  var body = '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"/></head>'
    + '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica Neue,Arial,sans-serif;">'
    + '<div style="max-width:600px;margin:40px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'
    + '<div style="background:linear-gradient(135deg,#cf1f2e,#a31824);padding:32px 40px;text-align:center;">'
    + '<h1 style="color:#ffffff;margin:0;font-size:22px;font-weight:800;letter-spacing:1px;">ACCELERATE 2026</h1>'
    + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:13px;">BNI Korea National Conference</p>'
    + '</div>'
    + '<div style="padding:40px;">'
    + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">\uD30C\uD2F0 \uD328\uC2A4 \uAD6C\uB9E4 \uC778\uC99D\uCF54\uB4DC</h2>'
    + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 28px;">'
    + '<strong>' + name + '</strong>\uB2D8, BNI K. Member Pass \uACB0\uC81C\uAC00 \uD655\uC778\uB418\uC5C8\uC2B5\uB2C8\uB2E4.<br/>'
    + '\uC544\uB798 \uC778\uC99D\uCF54\uB4DC\uB97C \uC0AC\uC6A9\uD558\uC5EC Networking Party Pass\uB97C \uAD6C\uB9E4\uD558\uC2E4 \uC218 \uC788\uC2B5\uB2C8\uB2E4.</p>'
    + '<div style="background:#fffbeb;border:2px dashed #f59e0b;border-radius:12px;padding:28px;margin-bottom:28px;text-align:center;">'
    + '<p style="color:#b45309;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 12px;">Networking Party Pass Code</p>'
    + '<p style="color:#1a1a1a;font-size:28px;font-weight:900;margin:0 0 8px;letter-spacing:4px;font-family:monospace;">' + code + '</p>'
    + '<p style="color:#999;font-size:12px;margin:0;">\uD2F0\uCF13 \uAD6C\uB9E4 \uD398\uC774\uC9C0\uC5D0\uC11C \uC774 \uCF54\uB4DC\uB97C \uC785\uB825\uD574 \uC8FC\uC138\uC694</p>'
    + '</div>'
    + '<div style="background:#fafafa;border:1px solid #eee;border-radius:12px;padding:24px;margin-bottom:28px;">'
    + '<table style="width:100%;border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#999;padding:6px 0;width:140px;">\uD328\uC2A4 \uC885\uB958</td><td style="color:#cf1f2e;font-weight:700;padding:6px 0;">Networking Party Pass</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">\uC5BC\uB9AC\uBC84\uB4DC</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">\u20A977,000</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">\uC815\uC0C1\uAC00</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">\u20A988,000</td></tr>'
    + '</table></div>'
    + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:28px;">'
    + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
    + '<strong>\uC0AC\uC6A9 \uBC29\uBC95:</strong><br/>'
    + '1. <a href="https://www.nc26-bnikorea.com" style="color:#cf1f2e;">www.nc26-bnikorea.com</a> \uC811\uC18D<br/>'
    + '2. Networking Party Pass \uCE74\uB4DC\uC758 <strong>\u201C\uC778\uC99D\uCF54\uB4DC \uC785\uB825\u201D</strong> \uBC84\uD2BC \uD074\uB9AD<br/>'
    + '3. \uC704 \uCF54\uB4DC \uC785\uB825 \uD6C4 \uACB0\uC81C \uC9C4\uD589</p>'
    + '</div>'
    + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
    + '\uBB38\uC758\uC0AC\uD56D: '
    + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
    + ' \uB610\uB294 <a href="http://pf.kakao.com/_xewxmrT/chat" style="color:#cf1f2e;">\uCE74\uCE74\uC624\uD1A1 \uCC44\uD305</a></p>'
    + '</div>'
    + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
    + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
    + '</div></div></body></html>';

  var adminCc = ADMIN_EMAILS_KOR;
  MailApp.sendEmail({ to: email, cc: adminCc, subject: subject, htmlBody: body });
  Logger.log('Party code email sent to: ' + email);
}

// =============================================================
// Email functions (Overseas ticket / Booth)
// =============================================================

function sendTicketConfirmationEmail(headers, row) {
  var get = function(key) { return row[headers.indexOf(key)] || ''; };

  var name = get('Name');
  var email = get('Email');
  var plan = get('Plan');
  var price = get('Price');
  var nationality = get('Nationality');
  var lang = get('Language');

  if (!email) return;

  var subject = lang === 'ja'
    ? '\u3010Accelerate 2026\u3011\u304A\u652F\u6255\u3044\u78BA\u8A8D\u5B8C\u4E86'
    : lang === 'zh'
    ? '\u3010Accelerate 2026\u3011\u4ED8\u6B3E\u786E\u8BA4\u5B8C\u6210'
    : '\u3010Accelerate 2026\u3011Payment Confirmed';

  var body = '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"/></head>'
    + '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica Neue,Arial,sans-serif;">'
    + '<div style="max-width:600px;margin:40px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'
    + '<div style="background:#cf1f2e;padding:32px 40px;text-align:center;">'
    + '<h1 style="color:#ffffff;margin:0;font-size:22px;font-weight:800;letter-spacing:1px;">ACCELERATE 2026</h1>'
    + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:13px;">BNI Korea National Conference</p>'
    + '</div>'
    + '<div style="padding:40px;">'
    + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">Payment Confirmed &#10003;</h2>'
    + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 28px;">'
    + 'Dear <strong>' + name + '</strong>, your registration has been confirmed.</p>'
    + '<div style="background:#fafafa;border:1px solid #eee;border-radius:12px;padding:24px;margin-bottom:28px;">'
    + '<table style="width:100%;border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#999;padding:6px 0;width:120px;">Name</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + name + '</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">Nationality</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + nationality + '</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">Ticket</td><td style="color:#cf1f2e;font-weight:700;padding:6px 0;">' + plan + '</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">Amount Paid</td><td style="color:#1a1a1a;font-weight:700;padding:6px 0;">' + price + '</td></tr>'
    + '</table></div>'
    + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:28px;">'
    + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
    + '&#128197; <strong>June 22-23, 2026</strong><br/>'
    + '&#128205; <strong>Swiss Grand Hotel, Seoul</strong><br/>'
    + '&#127760; <a href="https://nc26.bni-korea.com" style="color:#cf1f2e;">nc26.bni-korea.com</a></p>'
    + '</div>'
    + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
    + 'If you have any questions, please contact us at '
    + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
    + ' or via <a href="https://wa.me/821023778835" style="color:#cf1f2e;">WhatsApp</a>.</p>'
    + '</div>'
    + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
    + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
    + '</div></div></body></html>';

  MailApp.sendEmail({ to: email, cc: ADMIN_EMAILS, subject: subject, htmlBody: body });
  Logger.log('Ticket confirmation email sent to: ' + email + ' cc: ' + ADMIN_EMAILS);
}

function sendBoothVoucherEmail(headers, row, sheetName) {
  var get = function(key) { return row[headers.indexOf(key)] || ''; };

  var company = get('Company');
  var displayName = get('Display Name');
  var email = get('Email');
  var applicantName = get('Applicant Name');
  var applicantEmail = get('Applicant Email');
  var country = get('Country');
  var price = get('Price');

  Logger.log('Booth email debug - Email: [' + email + '] ApplicantEmail: [' + applicantEmail + ']');
  var recipients = [];
  if (email) recipients.push(email);
  if (applicantEmail && applicantEmail !== email) recipients.push(applicantEmail);
  if (recipients.length === 0) return;
  var recipientEmail = recipients.join(',');
  Logger.log('Booth sending to: ' + recipientEmail);

  var isKor = (sheetName === 'Booth_Kor');
  var voucherNo = 'BNI-BOOTH-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd') + '-' + Math.random().toString(36).substring(2, 7).toUpperCase();
  var recipientName = applicantName || company;

  var subject = isKor
    ? '\u3010Accelerate 2026\u3011\uBD80\uC2A4 \uACB0\uC81C \uD655\uC778 \uC644\uB8CC'
    : '\u3010Accelerate 2026\u3011Booth Payment Confirmed - Voucher';

  var body;
  if (isKor) {
    body = '<!DOCTYPE html>'
      + '<html><head><meta charset="utf-8"/></head>'
      + '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica Neue,Arial,sans-serif;">'
      + '<div style="max-width:600px;margin:40px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'
      + '<div style="background:linear-gradient(135deg,#cf1f2e,#a31824);padding:32px 40px;text-align:center;">'
      + '<h1 style="color:#ffffff;margin:0;font-size:22px;font-weight:800;letter-spacing:1px;">ACCELERATE 2026</h1>'
      + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:13px;">BNI Korea National Conference</p>'
      + '</div>'
      + '<div style="padding:40px;">'
      + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">&#10003; &#xBD80;&#xC2A4; &#xACB0;&#xC81C;&#xAC00; &#xD655;&#xC778;&#xB418;&#xC5C8;&#xC2B5;&#xB2C8;&#xB2E4;</h2>'
      + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 28px;">'
      + '<strong>' + recipientName + '</strong>&#xB2D8;, &#xBD80;&#xC2A4; &#xCC38;&#xAC00; &#xC2E0;&#xCCAD;&#xC774; &#xD655;&#xC778;&#xB418;&#xC5C8;&#xC2B5;&#xB2C8;&#xB2E4;.</p>'
      + '<div style="background:#fffbeb;border:2px dashed #f59e0b;border-radius:12px;padding:24px;margin-bottom:28px;text-align:center;">'
      + '<p style="color:#b45309;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 8px;">Exhibition Booth Voucher</p>'
      + '<p style="color:#1a1a1a;font-size:22px;font-weight:900;margin:0 0 4px;letter-spacing:2px;">' + voucherNo + '</p>'
      + '<p style="color:#999;font-size:11px;margin:0;">&#xD589;&#xC0AC;&#xC7A5;&#xC5D0;&#xC11C; &#xC774; &#xBC14;&#xC6B0;&#xCC98;&#xB97C; &#xC81C;&#xC2DC;&#xD574; &#xC8FC;&#xC138;&#xC694;</p>'
      + '</div>'
      + '<div style="background:#fafafa;border:1px solid #eee;border-radius:12px;padding:24px;margin-bottom:28px;">'
      + '<table style="width:100%;border-collapse:collapse;font-size:14px;">'
      + '<tr><td style="color:#999;padding:6px 0;width:140px;">&#xC5C5;&#xCCB4;&#xBA85;</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + company + '</td></tr>'
      + '<tr><td style="color:#999;padding:6px 0;">&#xBD80;&#xC2A4; &#xD06C;&#xAE30;</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">2m x 2m x 2.5m</td></tr>'
      + '<tr><td style="color:#999;padding:6px 0;">&#xACB0;&#xC81C; &#xAE08;&#xC561;</td><td style="color:#cf1f2e;font-weight:700;padding:6px 0;">' + price + '</td></tr>'
      + '</table></div>'
      + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:28px;">'
      + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
      + '&#128197; <strong>2026&#xB144; 6&#xC6D4; 22&#xC77C;(&#xC6D4;) ~ 23&#xC77C;(&#xD654;)</strong><br/>'
      + '&#128205; <strong>&#xC2A4;&#xC704;&#xC2A4;&#xADF8;&#xB79C;&#xB4DC;&#xD638;&#xD154; &#xCEE8;&#xBCA4;&#xC158;&#xC13C;&#xD130; 4F</strong><br/>'
      + '&#128679; <strong>&#xBD80;&#xC2A4; &#xC14B;&#xD305;:</strong> 6&#xC6D4; 22&#xC77C;(&#xC6D4;) &#xC0C8;&#xBCBD; 2&#xC2DC;~<br/>'
      + '&#127760; <a href="https://nc26.bni-korea.com" style="color:#cf1f2e;">nc26.bni-korea.com</a></p>'
      + '</div>'
      + '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:20px;margin-bottom:28px;">'
      + '<p style="margin:0 0 12px;font-size:13px;font-weight:700;color:#166534;">&#128203; &#xD589;&#xC0AC; &#xC804; &#xC900;&#xBE44;&#xC0AC;&#xD56D;</p>'
      + '<p style="margin:0;font-size:12px;color:#15803d;line-height:2;">'
      + '&#9989; &#xB85C;&#xACE0; &#xD30C;&#xC77C; &#xC81C;&#xCD9C; (AI/PNG) - &#xBBF8;&#xC81C;&#xCD9C; &#xC2DC;<br/>'
      + '&#9989; &#xD504;&#xB85C;&#xADF8;&#xB7A8;&#xBD81; &#xD64D;&#xBCF4;&#xC790;&#xB8CC; &#xC81C;&#xCD9C; (A5 &#xAC00;&#xB85C; 210x148mm)<br/>'
      + '&#9989; &#xBD80;&#xC2A4; &#xC804;&#xC2DC; &#xBB3C;&#xD488; &#xC900;&#xBE44;</p>'
      + '</div>'
      + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
      + '&#xBD80;&#xC2A4; &#xAD00;&#xB828; &#xBB38;&#xC758;&#xB294; '
      + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
      + ' &#xB610;&#xB294; <a href="http://pf.kakao.com/_xewxmrT/chat" style="color:#cf1f2e;">&#xCE74;&#xCE74;&#xC624;&#xD1A1; &#xCC44;&#xD305;</a>&#xC73C;&#xB85C; &#xC5F0;&#xB77D;&#xD574; &#xC8FC;&#xC138;&#xC694;.</p>'
      + '</div>'
      + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
      + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
      + '</div></div></body></html>';
  } else {
    var displayNameRow = displayName ? '<tr><td style="color:#999;padding:6px 0;">Display Name</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + displayName + '</td></tr>' : '';
    var countryRow = country ? '<tr><td style="color:#999;padding:6px 0;">Country</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + country + '</td></tr>' : '';

    body = '<!DOCTYPE html>'
      + '<html><head><meta charset="utf-8"/></head>'
      + '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica Neue,Arial,sans-serif;">'
      + '<div style="max-width:600px;margin:40px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'
      + '<div style="background:linear-gradient(135deg,#cf1f2e,#a31824);padding:32px 40px;text-align:center;">'
      + '<h1 style="color:#ffffff;margin:0;font-size:22px;font-weight:800;letter-spacing:1px;">ACCELERATE 2026</h1>'
      + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:13px;">Exhibition Booth Voucher</p>'
      + '</div>'
      + '<div style="padding:40px;">'
      + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">Booth Payment Confirmed &#10003;</h2>'
      + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 28px;">'
      + 'Dear <strong>' + recipientName + '</strong>, your booth registration has been confirmed.</p>'
      + '<div style="background:#fffbeb;border:2px dashed #f59e0b;border-radius:12px;padding:24px;margin-bottom:28px;text-align:center;">'
      + '<p style="color:#b45309;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 8px;">Exhibition Booth Voucher</p>'
      + '<p style="color:#1a1a1a;font-size:22px;font-weight:900;margin:0 0 4px;letter-spacing:2px;">' + voucherNo + '</p>'
      + '<p style="color:#999;font-size:11px;margin:0;">Please present this voucher at the venue</p>'
      + '</div>'
      + '<div style="background:#fafafa;border:1px solid #eee;border-radius:12px;padding:24px;margin-bottom:28px;">'
      + '<table style="width:100%;border-collapse:collapse;font-size:14px;">'
      + '<tr><td style="color:#999;padding:6px 0;width:140px;">Company</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + company + '</td></tr>'
      + displayNameRow
      + countryRow
      + '<tr><td style="color:#999;padding:6px 0;">Booth Size</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">2m x 2m x 2.5m</td></tr>'
      + '<tr><td style="color:#999;padding:6px 0;">Amount Paid</td><td style="color:#cf1f2e;font-weight:700;padding:6px 0;">' + price + '</td></tr>'
      + '</table></div>'
      + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:28px;">'
      + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
      + '&#128197; <strong>June 22-23, 2026</strong><br/>'
      + '&#128205; <strong>Swiss Grand Hotel Convention Center 4F, Seoul</strong><br/>'
      + '&#128679; <strong>Booth Setup:</strong> June 22 (Mon) AM<br/>'
      + '&#127760; <a href="https://nc26.bni-korea.com" style="color:#cf1f2e;">nc26.bni-korea.com</a></p>'
      + '</div>'
      + '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:20px;margin-bottom:28px;">'
      + '<p style="margin:0 0 12px;font-size:13px;font-weight:700;color:#166534;">&#128203; Before the Event</p>'
      + '<p style="margin:0;font-size:12px;color:#15803d;line-height:2;">'
      + '&#9989; Submit logo file (AI/PNG) - if not already uploaded<br/>'
      + '&#9989; Submit program book ad material (A5 Landscape 210x148mm)<br/>'
      + '&#9989; Prepare booth display items</p>'
      + '</div>'
      + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
      + 'For booth-related inquiries, please contact '
      + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
      + ' or via <a href="https://wa.me/821023778835" style="color:#cf1f2e;">WhatsApp</a>.</p>'
      + '</div>'
      + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
      + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
      + '</div></div></body></html>';
  }

  var adminCc = (sheetName === 'Booth_Kor') ? ADMIN_EMAILS_KOR : ADMIN_EMAILS;
  MailApp.sendEmail({ to: recipientEmail, cc: adminCc, subject: subject, htmlBody: body });
  Logger.log('Booth voucher email sent to: ' + recipientEmail + ' cc: ' + adminCc);
}
