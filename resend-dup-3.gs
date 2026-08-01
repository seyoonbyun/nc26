// =====================================================================
// PartyCodes 누락분 3건 추가 발송 — admin@bni-korea.com 계정 전용
//
// 케이스: 같은 이메일로 다른 사람 명의 멤버패스 결제 → 코드 미발급분
//
// [실행]
// 1. https://script.google.com 에 admin@bni-korea.com 로그인 상태로 접속
// 2. 좀 전에 만든 91건 재발송 프로젝트의 Code.gs 내용을 전부 지우고
//    이 코드 통째로 붙여넣기
// 3. 함수 드롭다운에서 resendDupEmails 선택 → ▶ 실행
// 4. 권한은 이미 동의된 상태라 바로 발송됨
// =====================================================================

var RESEND_LIST = [
  ['guswns0255@hanmail.net', '임현준', 'NP-TTUAPNTG'],
  ['zeo3352@naver.com', '윤영진', 'NP-2DJPKG6U'],
  ['zeo3352@naver.com', '구미경', 'NP-8MPERT5B']
];

function resendDupEmails() {
  var sent = 0;
  var failed = 0;
  var total = RESEND_LIST.length;
  Logger.log('=== 추가 발송 시작 (총 ' + total + '건) ===');
  for (var i = 0; i < total; i++) {
    var email = RESEND_LIST[i][0];
    var name  = RESEND_LIST[i][1];
    var code  = RESEND_LIST[i][2];
    try {
      sendPartyCodeEmail(email, name, code);
      sent++;
      Logger.log((i+1) + '/' + total + ' OK  ' + name + ' / ' + email + ' / ' + code);
    } catch (err) {
      failed++;
      Logger.log((i+1) + '/' + total + ' FAIL ' + email + ' — ' + (err && err.message));
    }
    Utilities.sleep(250);
  }
  Logger.log('=== 완료: 성공 ' + sent + ' / 실패 ' + failed + ' ===');
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
    + '<h2 style="color:#1a1a1a;font-size:20px;margin:0 0 8px;">파티 패스 구매 인증코드</h2>'
    + '<p style="color:#666;font-size:14px;line-height:1.6;margin:0 0 28px;">'
    + '<strong>' + name + '</strong>님, BNI K. Member Pass 결제가 확인되었습니다.<br/>'
    + '아래 인증코드를 사용하여 Networking Party Pass를 구매하실 수 있습니다.</p>'
    + '<div style="background:#fffbeb;border:2px dashed #f59e0b;border-radius:12px;padding:28px;margin-bottom:28px;text-align:center;">'
    + '<p style="color:#b45309;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:0 0 12px;">Networking Party Pass Code</p>'
    + '<p style="color:#1a1a1a;font-size:28px;font-weight:900;margin:0 0 8px;letter-spacing:4px;font-family:monospace;">' + code + '</p>'
    + '<p style="color:#999;font-size:12px;margin:0;">티켓 구매 페이지에서 이 코드를 입력해 주세요</p>'
    + '</div>'
    + '<div style="background:#fafafa;border:1px solid #eee;border-radius:12px;padding:24px;margin-bottom:28px;">'
    + '<table style="width:100%;border-collapse:collapse;font-size:14px;">'
    + '<tr><td style="color:#999;padding:6px 0;width:140px;">패스 종류</td><td style="color:#cf1f2e;font-weight:700;padding:6px 0;">Networking Party Pass</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">얼리버드</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">₩77,000</td></tr>'
    + '<tr><td style="color:#999;padding:6px 0;">정상가</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">₩88,000</td></tr>'
    + '</table></div>'
    + '<div style="background:#fff8f8;border-left:4px solid #cf1f2e;padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:28px;">'
    + '<p style="margin:0;font-size:13px;color:#333;line-height:1.8;">'
    + '<strong>사용 방법:</strong><br/>'
    + '1. <a href="https://www.nc26-bnikorea.com" style="color:#cf1f2e;">www.nc26-bnikorea.com</a> 접속<br/>'
    + '2. Networking Party Pass 카드의 <strong>"인증코드 입력"</strong> 버튼 클릭<br/>'
    + '3. 위 코드 입력 후 결제 진행</p>'
    + '</div>'
    + '<p style="color:#999;font-size:12px;line-height:1.6;margin:0;">'
    + '문의사항: '
    + '<a href="mailto:admin@bni-korea.com" style="color:#cf1f2e;">admin@bni-korea.com</a>'
    + ' 또는 <a href="http://pf.kakao.com/_xewxmrT/chat" style="color:#cf1f2e;">카카오톡 채팅</a></p>'
    + '</div>'
    + '<div style="background:#f9f9f9;border-top:1px solid #eee;padding:20px 40px;text-align:center;">'
    + '<p style="margin:0;font-size:11px;color:#bbb;">&copy; 2026 BNI Korea. All rights reserved.</p>'
    + '</div></div></body></html>';

  MailApp.sendEmail({
    to: email,
    cc: 'hq@joy-bnikorea.com',
    subject: subject,
    htmlBody: body
  });
}
