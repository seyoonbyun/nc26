var SS = SpreadsheetApp.getActiveSpreadsheet();
var ADMIN_EMAILS = 'hq@joy-bnikorea.com,admin@bni-korea.com,ksoh7512@gmail.com';
var ADMIN_EMAILS_KOR = 'hq@joy-bnikorea.com,admin@bni-korea.com';

var TICKET_HEADERS = [
  'Timestamp', 'Name', 'Nationality', 'Email', 'Phone',
  'Position', 'Plan', 'Price', 'Language', 'Memo', 'Status'
];

var BOOTH_HEADERS = [
  'Timestamp', 'Company', 'Display Name', 'Owner', 'Address',
  'Phone', 'Fax', 'Homepage', 'Email',
  'Applicant Name', 'Applicant Phone', 'Applicant Email',
  'Country', 'Chapter', 'License', 'Price', 'Logo File', 'Ad File', 'Status'
];

var BOOTH_KOR_HEADERS = [
  'Timestamp', 'Company', 'Owner', 'Address',
  'Phone', 'Fax', 'Homepage', 'Email',
  'Applicant Name', 'Chapter', 'Applicant Phone', 'Applicant Email',
  'License', 'Price', 'Logo File', 'Ad File', 'Status'
];

var PARTY_CODE_HEADERS = ['Code', 'Email', 'Name', 'Created', 'Used', 'Reserved', 'Duplicate'];

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
        if (p.homepage) setHyperlink(sheet, newRow, 7, p.homepage, p.homepage);
        if (p.email) setHyperlink(sheet, newRow, 8, 'mailto:' + p.email, p.email);
        if (p.applicantEmail) setHyperlink(sheet, newRow, 12, 'mailto:' + p.applicantEmail, p.applicantEmail);
        if (logoUrl) setHyperlink(sheet, newRow, 15, logoUrl, logoName);
        if (adUrl) setHyperlink(sheet, newRow, 16, adUrl, adName);

      } else {
        var sheet = getOrCreateSheet('Booth', BOOTH_HEADERS);
        sheet.appendRow([
          formatTimestamp(p.timestamp),
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
        if (p.homepage) setHyperlink(sheet, newRow, 8, p.homepage, p.homepage);
        if (p.email) setHyperlink(sheet, newRow, 9, 'mailto:' + p.email, p.email);
        if (p.applicantEmail) setHyperlink(sheet, newRow, 12, 'mailto:' + p.applicantEmail, p.applicantEmail);
        if (logoUrl) setHyperlink(sheet, newRow, 17, logoUrl, logoName);
        if (adUrl) setHyperlink(sheet, newRow, 18, adUrl, adName);
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

  // ----- 관리자 API (PropertiesService에 저장된 ADMIN_KEY 필요) -----
  if (action === 'adminBackfill')    return adminGuard(e, function() { backfillPartyCodeUsed(); return { status: 'ok', action: 'adminBackfill' }; });
  if (action === 'adminReadSheet')   return adminGuard(e, function() { return adminReadSheet(e.parameter); });
  if (action === 'adminUpdateCell')  return adminGuard(e, function() { return adminUpdateCell(e.parameter); });
  if (action === 'adminSetUsed')     return adminGuard(e, function() { return adminSetUsed(e.parameter); });
  if (action === 'adminListSheets')  return adminGuard(e, function() { return adminListSheets(); });
  if (action === 'adminRunScan')     return adminGuard(e, function() { scanAndSendPartyCodes(); return { status: 'ok', action: 'adminRunScan' }; });

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'NC26 Overseas API is running.' })
  ).setMimeType(ContentService.MimeType.JSON);
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

// PartyCodes 특정 코드 행의 Used / Duplicate 값을 직접 기록
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
      if (params.duplicate !== undefined) sheet.getRange(row, 7).setValue(params.duplicate);
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

function scanAndSendPartyCodes() {
  Logger.log('scanAndSendPartyCodes 시작');
  var sheet = SS.getSheetByName('Ticket & Booth_Kor.pay');
  if (!sheet) { Logger.log('시트 없음: Ticket & Booth_Kor.pay'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('데이터 없음'); return; }

  var lastCol = Math.max(sheet.getLastColumn(), 12);
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  var codeSheet = getOrCreatePartyCodeSheet();
  var codeData = codeSheet.getDataRange().getValues();

  // PartyCodes의 Email(B열) 기준으로 인덱스 매핑 (정규화하여 저장)
  var issuedEmails = {};
  for (var i = 1; i < codeData.length; i++) {
    var em = normalizeEmail(codeData[i][1]);
    if (em) issuedEmails[em] = i;
  }

  var sentCount = 0;
  var usedCount = 0;

  for (var r = 0; r < data.length; r++) {
    var title = String(data[r][0] || '').trim();
    var name = data[r][1];
    var email = data[r][2];
    var statusPayment = data[r][11];
    var normEmail = normalizeEmail(email);

    if (!normEmail || !isPaymentComplete(statusPayment)) continue;

    if (title === 'BNI K. Member Pass') {
      if (issuedEmails[normEmail] !== undefined) continue;

      var code = generatePartyCode();
      var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      Logger.log('발급: ' + email + ' 코드: ' + code);
      codeSheet.appendRow([code, email, name, timestamp, '', '', '']);
      issuedEmails[normEmail] = codeSheet.getLastRow() - 1;

      sendPartyCodeEmail(email, name, code);
      Logger.log('이메일 발송 완료: ' + email);
      sentCount++;
    }

    if (title === 'Networking Party Pass') {
      var codeRowIdx = issuedEmails[normEmail];
      if (codeRowIdx === undefined) {
        Logger.log('Used 스킵 (해당 이메일의 발급 코드 없음): ' + email);
        continue;
      }

      // E열 Used에 첫 결제완료 시각 기록. 이미 있으면 중복이므로 G열(Duplicate)에 기록.
      // (F열 Reserved는 verifyPartyCode가 락 용도로 사용하며 건드리지 않음)
      var payRow = r + 2; // .pay 시트 실제 행번호
      var usedCell = codeSheet.getRange(codeRowIdx + 1, 5);
      var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

      if (!usedCell.getValue()) {
        usedCell.setValue(timestamp);
        Logger.log('Used 처리: ' + email + ' @ ' + timestamp + ' (.pay 행 ' + payRow + ')');
        usedCount++;
      } else {
        // 중복 결제 발생 — Duplicate(G)에 행번호 누적
        var dupCell = codeSheet.getRange(codeRowIdx + 1, 7);
        var existing = String(dupCell.getValue() || '');
        var tag = '.pay 행 ' + payRow;
        if (existing.indexOf(tag) === -1) {
          dupCell.setValue(existing ? (existing + ' / ' + tag + ' @ ' + timestamp)
                                    : ('중복 구매 / ' + tag + ' @ ' + timestamp));
          Logger.log('Duplicate 감지: ' + email + ' (.pay 행 ' + payRow + ')');
        }
      }
    }
  }

  Logger.log('완료: ' + sentCount + '건 발급, ' + usedCount + '건 Used');
}

// =============================================================
// 수동 백필 / 진단 도구
// - Apps Script 편집기에서 함수 선택 후 "실행" 버튼으로 호출
// - Ticket & Booth_Kor.pay 시트의 Networking Party Pass 결제완료 행을 전부 훑어서
//   PartyCodes의 Used(E) / Duplicate(G)에 채워 넣음
// - 같은 이메일로 2회 이상 구매된 경우 첫 건은 Used, 나머지는 Duplicate(G)에 시트 행번호로 기록
// =============================================================
function backfillPartyCodeUsed() {
  Logger.log('backfillPartyCodeUsed 시작');
  var sheet = SS.getSheetByName('Ticket & Booth_Kor.pay');
  if (!sheet) { Logger.log('시트 없음'); return; }
  var codeSheet = getOrCreatePartyCodeSheet();

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('.pay 데이터 없음'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 12)).getDisplayValues();
  var codeData = codeSheet.getDataRange().getValues();

  // PartyCodes: email → 시트 행번호
  var emailToRow = {};
  for (var i = 1; i < codeData.length; i++) {
    var em = normalizeEmail(codeData[i][1]);
    if (em) emailToRow[em] = i + 1;
  }

  // .pay 시트를 훑어 이메일별 Networking Party Pass 결제완료 행번호 수집
  var hitsByEmail = {};
  for (var r = 0; r < data.length; r++) {
    var title = String(data[r][0] || '').trim();
    if (title !== 'Networking Party Pass') continue;
    var normEmail = normalizeEmail(data[r][2]);
    if (!normEmail) continue;
    if (!isPaymentComplete(data[r][11])) continue;

    if (!hitsByEmail[normEmail]) hitsByEmail[normEmail] = [];
    hitsByEmail[normEmail].push(r + 2); // .pay 시트의 실제 행번호
  }

  var updated = 0;
  var dupMarked = 0;
  var ts = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  Object.keys(hitsByEmail).forEach(function(em) {
    var rows = hitsByEmail[em]; // [rowA, rowB, ...]
    var targetRow = emailToRow[em];
    if (!targetRow) { Logger.log('PartyCodes에 코드 없음: ' + em); return; }

    var usedCell = codeSheet.getRange(targetRow, 5);
    var dupCell = codeSheet.getRange(targetRow, 7);

    if (!usedCell.getValue()) {
      usedCell.setValue(ts);
      updated++;
      Logger.log('Used 백필: ' + em + ' @ ' + ts + ' (.pay 행 ' + rows[0] + ')');
    }

    if (rows.length > 1) {
      var existing = String(dupCell.getValue() || '');
      var note = rows.length + '회 구매 (중복) / .pay 행: ' + rows.join(', ');
      if (existing !== note) {
        dupCell.setValue(note);
        dupMarked++;
        Logger.log('Duplicate 표시: ' + em + ' → ' + note);
      }
    }
  });

  Logger.log('백필 완료: Used ' + updated + '건, Duplicate ' + dupMarked + '건');
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
  // 기존 시트에 Reserved(F) / Duplicate(G) 헤더가 없으면 보정
  var headerRow = sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).getValues()[0];
  var applyHeader = function(col, label) {
    var cell = sheet.getRange(1, col);
    cell.setValue(label);
    cell.setFontWeight('bold');
    cell.setBackground('#cf1f2e');
    cell.setFontColor('#ffffff');
  };
  if (!headerRow[5]) applyHeader(6, 'Reserved');
  if (!headerRow[6]) applyHeader(7, 'Duplicate');
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
function verifyPartyCode(code) {
  var result = { valid: false };
  var out = function(r) {
    return ContentService
      .createTextOutput(JSON.stringify(r))
      .setMimeType(ContentService.MimeType.JSON);
  };

  if (!code) return out(result);
  code = String(code).trim().toUpperCase();

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return out(result);
  }

  try {
    var sheet = getOrCreatePartyCodeSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== code) continue;

      var used = data[i][4];
      var reserved = data[i][5];
      if (used || reserved) {
        result.valid = false;
        result.message = 'used';
        return out(result);
      }

      var ts = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(i + 1, 6).setValue(ts); // F열 = Reserved
      SpreadsheetApp.flush();
      result.valid = true;
      result.name = data[i][2];
      return out(result);
    }
    return out(result);
  } finally {
    lock.releaseLock();
  }
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
