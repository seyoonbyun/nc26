// =====================================================================
// 누락 인증코드 일괄 재발송 — admin@bni-korea.com 계정 전용
//
// [실행 방법]
// 1. admin@bni-korea.com 으로 Google 로그인 (다른 계정 아닐 것!)
// 2. https://script.google.com 접속 → 새 프로젝트
// 3. 이 파일 내용 전체 복사해서 Code.gs 에 붙여넣기
// 4. 상단 함수 드롭다운에서 "resendMissingEmails" 선택 → ▶ 실행
// 5. 권한 동의창 뜨면 Gmail 발송 권한 허용
// 6. 하단 실행 로그에 진행 상황 / 성공·실패 카운트 표시됨
//
// 안전:
//  - 각 메일 사이 250ms 대기 (Gmail rate limit 회피)
//  - 실패 건도 로그에 찍힘 — 따로 골라 재시도 가능
//  - admin@ 일일 쿼터 (Workspace 1500) 의 91건만 사용 → 여유
//
// 발송 후:
//  - 보낸편지함에서 91건 확인 가능
//  - CC: hq@joy-bnikorea.com (자기 자신은 CC 제외)
// =====================================================================

var RESEND_LIST = [
  ['lozy211@gmail.com', '전홍귝', 'NP-4L3HHQEX'],
  ['ccoret1215@naver.com', '이종빈', 'NP-SP5C3B74'],
  ['naynot07@naver.com', '강보경', 'NP-SM2NNM3Q'],
  ['lee_shiwu@naver.com', '이지언', 'NP-FGBGNZNY'],
  ['africa2424@naver.com', '김필오', 'NP-3XCLXURY'],
  ['legendplus11@naver.com', '조형목', 'NP-Q4BDQQ4W'],
  ['ocm789@naver.com', '김선근', 'NP-DXDZXDD2'],
  ['gillhang@naver.com', '박선주', 'NP-BYS6B277'],
  ['datasave@data-save.net', '김태원', 'NP-PQR8YLY6'],
  ['sbsohn@kifac.biz', '손상백', 'NP-2FZY9GGU'],
  ['jwook11.lee@gmail.com', '이재욱', 'NP-JVDP45Y7'],
  ['posejeong@naver.com', '정용진', 'NP-FZ23S2WL'],
  ['loidmtp@naver.com', '오재협', 'NP-ZSLVQQXW'],
  ['dooho3323@gmail.com', '신두호', 'NP-S5DN2BRQ'],
  ['meccasalt-1@naver.com', '허승희', 'NP-7KVYK4YQ'],
  ['ghdcjs7@naver.com', '김건우', 'NP-PDNS66WY'],
  ['sunjeongkim@canarykorea.com', '김선정', 'NP-MBG775VW'],
  ['ryouseungyeol@gmail.com', '류승열', 'NP-A94SLR42'],
  ['d024729@naver.com', '김은규', 'NP-H6JAX5ET'],
  ['gayaiangel@gmail.com', '송현지', 'NP-CUHAMEPW'],
  ['ceo@mvmfilm.com', '차일웅', 'NP-BZDRERYQ'],
  ['dalors_vin@naver.com', '유광렬', 'NP-RVVBV8YS'],
  ['metro120@naver.com', '이재우', 'NP-3S4WQWG4'],
  ['synachang@gmail.com', '나성연', 'NP-P2U9DHF3'],
  ['Mankiu@hanmail.net', '유만기', 'NP-YKZWNTST'],
  ['lym0612@okfood.co.kr', '이영미', 'NP-Z5C4HC38'],
  ['huzodang1671@naver.com', '진성훈', 'NP-TE42AKR6'],
  ['btmtg@naver.com', '문종환', 'NP-99QTS7MG'],
  ['jasonjob@naver.com', '성민석', 'NP-5UUZM5UT'],
  ['cms@cmyj.kr', '최민수', 'NP-BMU89JEF'],
  ['papa2130@daum.net', '이원상', 'NP-H67EFCD4'],
  ['0082yyy@naver.com', '윤삼자', 'NP-S8QZPEZR'],
  ['jinoplus200818@gmail.com', '홍정현', 'NP-C63JG9U9'],
  ['yangik77@hanmail.net', '양일광', 'NP-R78HB23B'],
  ['bymaju@naver.com', '김은경', 'NP-J3WW3FXP'],
  ['ifoxcomet@naver.com', '윤혜성', 'NP-XZDLW3W4'],
  ['etshayyim@naver.com', '오형덕', 'NP-C9A2RPEE'],
  ['wootja1004@hanmail.net', '허진아', 'NP-WPKJ8FFD'],
  ['lawyer.shin@daum.net', '신동철', 'NP-8CAW4ZYA'],
  ['tt1194@acromobility.com', '김기환', 'NP-QQAWA3QW'],
  ['hdkim@acromobility.com', '김현동', 'NP-6NN5M7JE'],
  ['hwangh6293@naver.com', '황현', 'NP-TCRVWC58'],
  ['bhj2036@naver.com', '방혜정', 'NP-G9K4BKAB'],
  ['gljh00@hanmail.net', '이정한', 'NP-6B92MSUY'],
  ['guguent@naver.com', '구자민', 'NP-SFYRW7LY'],
  ['neoldk04@gmail.com', '이동관', 'NP-NA7ZMTTG'],
  ['lisse2015@naver.com', '최미진', 'NP-2GF5EB53'],
  ['80abba@gmail.com', '김시현', 'NP-LU9EBSCS'],
  ['166herjae@naver.com', '허형재', 'NP-2ASVV63J'],
  ['hippocra0419@gmail.com', '이정민', 'NP-ZZN6W9SX'],
  ['hapyoon@naver.com', '윤숙현', 'NP-E3QE8LR3'],
  ['delpicorp@naver.com', '이도겸', 'NP-CEHJRGJA'],
  ['jihoonl51@naver.com', '이지훈', 'NP-DTCGPGWR'],
  ['lcwid89@gmail.com', '이충원', 'NP-FGJBW4LD'],
  ['anyway80@hanmail.net', '오영훈', 'NP-9SELCLLC'],
  ['tidotttt13@naver.com', '정재영', 'NP-M2XD7LGU'],
  ['Justkim@kakao.com', 'JamesKim', 'NP-CLXZCBPN'],
  ['el@flowerhd.co.kr', '김수헌', 'NP-5M7PKMYA'],
  ['loveik78@naver.com', '조주훈', 'NP-FJR45NTE'],
  ['leeyoon0697@gmail.coml', '이윤', 'NP-YHSQGYLR'],
  ['fcmbyfreedom@gmail.com', '이지혜', 'NP-3PUE2Z4J'],
  ['jukj2000@hanmail.net', '장석흥', 'NP-XCMPU58A'],
  ['lawhojun@gmail.com', '이호준', 'NP-CVA3KCFS'],
  ['crochok@hanmail.net', '박원기', 'NP-RWA6XT9W'],
  ['Kimjaesung59@gmail.com', '김재성', 'NP-SCCKL8M4'],
  ['bmtax32@gmail.com', '임석인', 'NP-7T35VJ4L'],
  ['sgrida@naver.com', '노승준', 'NP-4LHCVFNL'],
  ['jdg2715@js-on.co.kr', '정대길', 'NP-7XD9WEXS'],
  ['yoonki@gjgs.kr', '김윤기', 'NP-5GZUCEPZ'],
  ['hongtax1001@naver.com', '이홍재', 'NP-2S8J744D'],
  ['bethe.jjj@gmail.com', '이기주', 'NP-TDE5AG6Z'],
  ['jinsuk-han@hanmail.net', '한진석', 'NP-BAPDRTX7'],
  ['support@reloalabs.com', '김도영', 'NP-N3B8UFZT'],
  ['viki0208n@naver.com', '곽여진', 'NP-ACH8AX78'],
  ['lmsruth@naver.com', '이명성', 'NP-33JPBXLC'],
  ['a01063121337@gmail.com', '박종민', 'NP-TA7C9EXT'],
  ['1559jungjin@naver.com', '박정짓', 'NP-6QW6BWTD'],
  ['dhhong97@naver.com', '홍동호', 'NP-WMYLNPZX'],
  ['aocceo0414@gmail.com', '정종태', 'NP-UFB2EMVL'],
  ['titedios2071@gmail.com', '이영기', 'NP-XPTUQV9T'],
  ['injichem.ai@gmail.com', '강근보', 'NP-9HS3W3E4'],
  ['vora1757@naver.com', '오정화', 'NP-BPGH6CBG'],
  ['hcn0815@naver.com', '한차남', 'NP-65ZJCB6P'],
  ['poohpjw77@naver.com', '박진우', 'NP-62MWNHQG'],
  ['digiry1718@hanmail.net', '곽대운', 'NP-HWLXSU7V'],
  ['zeo3352@naver.com', '이재호', 'NP-76RQCQ7Q'],
  ['bong@chaesimdang.com', '김봉석', 'NP-7F9VDDWL'],
  ['j2iman@naver.com', '주종일', 'NP-3ZFGXDNT'],
  ['ilmare517@naver.com', '윤미경', 'NP-AFYNYUW8'],
  ['hong2vv@naver.com', '정재홍', 'NP-53AG2AMT'],
  ['blueswane@hanmail.net', '한정아', 'NP-DCAM29QC']
];

function resendMissingEmails() {
  var sent = 0;
  var failed = 0;
  var total = RESEND_LIST.length;
  Logger.log('=== 재발송 시작 (총 ' + total + '건) ===');

  for (var i = 0; i < total; i++) {
    var email = RESEND_LIST[i][0];
    var name  = RESEND_LIST[i][1];
    var code  = RESEND_LIST[i][2];
    try {
      sendPartyCodeEmail(email, name, code);
      sent++;
      Logger.log((i+1) + '/' + total + ' OK  ' + email + ' / ' + code);
    } catch (err) {
      failed++;
      Logger.log((i+1) + '/' + total + ' FAIL ' + email + ' — ' + (err && err.message));
    }
    Utilities.sleep(250);
  }

  Logger.log('=== 완료: 성공 ' + sent + ' / 실패 ' + failed + ' / 총 ' + total + ' ===');
  Logger.log('남은 일일 쿼터: ' + MailApp.getRemainingDailyQuota());
}

function checkAdminQuota() {
  Logger.log('admin 계정 남은 일일 쿼터: ' + MailApp.getRemainingDailyQuota());
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
