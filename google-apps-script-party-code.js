/**
 * Networking Party Pass 인증코드 시스템
 *
 * BNI K. Member Pass 구매자(L열 = "결제 완료")에게
 * Networking Party Pass 구매용 인증코드를 이메일로 발송합니다.
 *
 * 설정 방법:
 * 1. 이 코드를 기존 Google Apps Script 프로젝트에 추가하거나 새 프로젝트로 생성
 * 2. onEditPartyCode 함수를 설치형 트리거(onEdit)로 등록
 * 3. doGet 함수가 이미 있으면 기존 doGet에 verifyPartyCode 분기를 병합
 */

var PARTY_CODE_HEADERS = ['Code', 'Email', 'Name', 'Created', 'Used'];

/**
 * "PartyCodes" 시트를 가져오거나 새로 생성
 */
function getOrCreatePartyCodeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PartyCodes');
  if (!sheet) {
    sheet = ss.insertSheet('PartyCodes');
    sheet.appendRow(PARTY_CODE_HEADERS);
    sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).setBackground('#cf1f2e');
    sheet.getRange(1, 1, 1, PARTY_CODE_HEADERS.length).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 8자리 고유 인증코드 생성 (예: NP-A3X7K2M9)
 */
function generatePartyCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 혼동 방지: 0,O,1,I 제외
  var code = '';
  for (var i = 0; i < 8; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return 'NP-' + code;
}

/**
 * 설치형 onEdit 트리거 — L열(12번째 열)이 "결제 완료"로 변경될 때 실행
 *
 * BNI K. Member Pass 구매자에게만 인증코드를 발송합니다.
 *
 * 트리거 등록: 편집기 > 트리거 > 트리거 추가
 *   - 함수: onEditPartyCode
 *   - 이벤트 소스: 스프레드시트에서
 *   - 이벤트 유형: 수정 시
 */
function onEditPartyCode(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var sheetName = sheet.getName();

  // Tickets_Kor 시트 또는 국내 티켓을 관리하는 시트명으로 변경 가능
  // 해당 시트에서만 동작하도록 제한
  if (sheetName !== 'Tickets_Kor' && sheetName !== 'Tickets') return;

  // L열(12번째 열) 확인
  if (range.getColumn() !== 12) return;
  if (range.getValue() !== '결제 완료') return;

  var row = range.getRow();
  if (row <= 1) return;

  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Plan 열에서 BNI K. Member Pass인지 확인
  var planCol = headers.indexOf('Plan');
  if (planCol === -1) planCol = headers.indexOf('plan');
  if (planCol === -1) planCol = headers.indexOf('플랜');

  var plan = planCol >= 0 ? rowData[planCol] : '';

  // BNI K. Member Pass가 아니면 무시
  if (plan.indexOf('Member') === -1 && plan.indexOf('member') === -1 && plan.indexOf('멤버') === -1) return;

  // 이메일, 이름 추출
  var emailCol = headers.indexOf('Email');
  if (emailCol === -1) emailCol = headers.indexOf('email');
  if (emailCol === -1) emailCol = headers.indexOf('이메일');
  var email = emailCol >= 0 ? rowData[emailCol] : '';
  if (!email) return;

  var nameCol = headers.indexOf('Name');
  if (nameCol === -1) nameCol = headers.indexOf('name');
  if (nameCol === -1) nameCol = headers.indexOf('이름');
  var name = nameCol >= 0 ? rowData[nameCol] : '';

  // 이미 코드가 발급되었는지 확인 (중복 방지)
  var codeSheet = getOrCreatePartyCodeSheet();
  var codeData = codeSheet.getDataRange().getValues();
  for (var i = 1; i < codeData.length; i++) {
    if (codeData[i][1] === email) {
      Logger.log('Party code already issued for: ' + email);
      return;
    }
  }

  // 인증코드 생성 및 저장
  var code = generatePartyCode();
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  codeSheet.appendRow([code, email, name, timestamp, '']);

  // 인증코드 이메일 발송
  sendPartyCodeEmail(email, name, code);
  Logger.log('Party code sent to: ' + email + ' code: ' + code);
}

/**
 * 인증코드 이메일 발송
 */
function sendPartyCodeEmail(email, name, code) {
  var subject = '【Accelerate 2026】Networking Party Pass 구매 인증코드';

  var body = '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"/></head>'
    + '<body style="margin:0;padding:0;background:#f5f5f5;font-family:Helvetica Neue,Arial,sans-serif;">'
    + '<div style="max-width:600px;margin:40px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'

    // Header
    + '<div style="background:linear-gradient(135deg,#cf1f2e,#a31824);padding:32px 40px;text-align:center;">'
    + '<h1 style="color:#ffffff;margin:0;font-size:22px;font-weight:800;letter-spacing:1px;">ACCELERATE 2026</h1>'
    + '<p style="color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:13px;">BNI Korea National Conference</p>'
    + '</div>'

    // Body
    + '<div style="padding:40px;">'
    + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">\uD30C\uD2F0 \uD328\uC2A4 \uAD6C\uB9E4 \uC778\uC99D\uCF54\uB4DC</h2>'
    + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 28px;">'
    + '<strong>' + name + '</strong>\uB2D8, BNI K. Member Pass \uACB0\uC81C\uAC00 \uD655\uC778\uB418\uC5C8\uC2B5\uB2C8\uB2E4.<br/>'
    + '\uC544\uB798 \uC778\uC99D\uCF54\uB4DC\uB97C \uC0AC\uC6A9\uD558\uC5EC Networking Party Pass\uB97C \uAD6C\uB9E4\uD558\uC2E4 \uC218 \uC788\uC2B5\uB2C8\uB2E4.</p>'

    // Code box
    + '<div style="background:#fffbeb;border:2px dashed #f59e0b;border-radius:12px;padding:28px;margin-bottom:28px;text-align:center;">'
    + '<p style="color:#b45309;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 12px;">Networking Party Pass Code</p>'
    + '<p style="color:#1a1a1a;font-size:28px;font-weight:900;margin:0 0 8px;letter-spacing:4px;font-family:monospace;">' + code + '</p>'
    + '<p style="color:#999;font-size:12px;margin:0;">\uD2F0\uCF13 \uAD6C\uB9E4 \uD398\uC774\uC9C0\uC5D0\uC11C \uC774 \uCF54\uB4DC\uB97C \uC785\uB825\uD574 \uC8FC\uC138\uC694</p>'
    + '</div>'

    // Info
    + '<div style="background:#fafafa;border:1px solid #eee;border-radius:12px;padding:24px;margin-bottom:28px;">'
    + '<table style="width:100%;border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#999;padding:6px 0;width:140px;">\uD328\uC2A4 \uC885\uB958</td><td style="color:#cf1f2e;font-weight:700;padding:6px 0;">Networking Party Pass</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">\uC5BC\uB9AC\uBC84\uB4DC</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">\u20A977,000</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">\uC815\uC0C1\uAC00</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">\u20A988,000</td></tr>'
    + '</table></div>'

    // How to use
    + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:28px;">'
    + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
    + '<strong>\uC0AC\uC6A9 \uBC29\uBC95:</strong><br/>'
    + '1. <a href="https://nc26.bni-korea.com/ticket.html" style="color:#cf1f2e;">nc26.bni-korea.com/ticket.html</a> \uC811\uC18D<br/>'
    + '2. Networking Party Pass \uCE74\uB4DC\uC758 <strong>\u201C\uC778\uC99D\uCF54\uB4DC \uC785\uB825\u201D</strong> \uBC84\uD2BC \uD074\uB9AD<br/>'
    + '3. \uC704 \uCF54\uB4DC \uC785\uB825 \uD6C4 \uACB0\uC81C \uC9C4\uD589</p>'
    + '</div>'

    + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
    + '\uBB38\uC758\uC0AC\uD56D: '
    + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
    + ' \uB610\uB294 <a href="http://pf.kakao.com/_xewxmrT/chat" style="color:#cf1f2e;">\uCE74\uCE74\uC624\uD1A1 \uCC44\uD305</a></p>'
    + '</div>'

    // Footer
    + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
    + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
    + '</div></div></body></html>';

  var adminCc = 'hq@joy-bnikorea.com,admin@bni-korea.com';
  MailApp.sendEmail({ to: email, cc: adminCc, subject: subject, htmlBody: body });
}

/**
 * 인증코드 검증 API (doGet 요청)
 *
 * 호출 예: GET ?action=verifyPartyCode&code=NP-A3X7K2M9
 *
 * 기존 doGet이 있는 경우 아래 분기를 기존 doGet에 병합하세요:
 *   if (e.parameter.action === 'verifyPartyCode') {
 *     return verifyPartyCode(e.parameter.code);
 *   }
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'verifyPartyCode') {
    return verifyPartyCode(e.parameter.code);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'NC26 API is running.' })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 인증코드 유효성 검증
 * - 코드가 존재하고 아직 사용되지 않았으면 valid
 * - 사용 시 Used 열에 타임스탬프 기록 (1회용)
 */
function verifyPartyCode(code) {
  var result = { valid: false };

  if (!code) {
    return ContentService.createTextOutput(
      JSON.stringify(result)
    ).setMimeType(ContentService.MimeType.JSON);
  }

  code = code.trim().toUpperCase();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PartyCodes');
  if (!sheet) {
    return ContentService.createTextOutput(
      JSON.stringify(result)
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === code) {
      if (data[i][4]) {
        // 이미 사용된 코드
        result.valid = false;
        result.message = 'used';
        break;
      }
      // 코드 사용 처리
      var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(i + 1, 5).setValue(timestamp);
      result.valid = true;
      result.name = data[i][2];
      break;
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify(result)
  ).setMimeType(ContentService.MimeType.JSON);
}
