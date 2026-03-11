/**
 * Accelerate 2026 - 등록 폼 → Google Sheets 저장 + 결제완료 시 자동 이메일 발송
 *
 * 설정 방법:
 * 1. Google Sheets에서 확장 프로그램 > Apps Script
 * 2. 이 코드를 Code.gs에 붙여넣기
 * 3. 배포 > 새 배포 > 웹 앱 (액세스: 모든 사용자)
 * 4. 생성된 URL을 pricing.html의 localStorage에 설정
 * 5. onEdit 트리거 설정: Apps Script 편집기 > 트리거(시계 아이콘) >
 *    + 트리거 추가 > 함수: onSheetEdit / 이벤트: 스프레드시트에서 / 유형: 수정 시
 */

var ADMIN_EMAILS = "hq@joy-bnikorea.com, admin@bni-korea.com, ksoh7512@gmail.com";
var SHEET_NAME = "Registrations";
var HEADERS = [
  "Timestamp", "Name", "Nationality", "Email", "Phone",
  "Plan", "Price", "Language", "Position", "Memo", "Status", "Mail Result"
];
var STATUS_COL = 11; // "Status" 열 번호 (K열)

// ─────────────────────────────────────────────
// 1. 폼 데이터 수신 → 시트 저장
// ─────────────────────────────────────────────

function doPost(e) {
  try {
    // form submit (e.parameter) 또는 JSON fetch (e.postData) 모두 지원
    var data;
    if (e.parameter && e.parameter.name) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = {};
    }

    var sheet = getOrCreateSheet();

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name || "",
      data.nationality || "",
      data.email || "",
      "'" + (data.phone || ""),
      data.plan || "",
      data.planPrice || "",
      data.lang || "en",
      data.position || "",
      data.memo || "",
      "Pending",
      ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Accelerate 2026 Registration API is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
// 2. 시트에서 "결제완료" 입력 감지 → 이메일 자동 발송
// ─────────────────────────────────────────────

function onSheetEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  // Status 열(J)이 수정되었고, 헤더 행이 아닌 경우
  if (col !== STATUS_COL || row < 2) return;

  var newValue = (range.getValue() || "").toString().trim();
  if (newValue !== "paid") return;

  // 해당 행의 데이터 읽기
  var rowData = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  var registration = {};
  for (var i = 0; i < HEADERS.length; i++) {
    registration[HEADERS[i]] = rowData[i];
  }
  registration._row = row;

  // 이메일 발송
  try {
    var ticketNumber = sendConfirmationEmail(registration);
    sendAdminNotification(registration, ticketNumber);
    // 발송 완료 표시: "결제완료" → "결제완료 ✓ 메일발송"
    sheet.getRange(row, STATUS_COL + 1).setValue("✓ 메일발송");
  } catch (err) {
    // 발송 실패 시 표시
    sheet.getRange(row, STATUS_COL + 1).setValue("✗ 발송실패: " + err.message);
  }
}

// ─────────────────────────────────────────────
// 3. 확인 이메일 발송
// ─────────────────────────────────────────────

function sendConfirmationEmail(reg) {
  var lang = (reg.Language || "en").toString().toLowerCase();
  var ticketNumber = generateTicketNumber(reg.Name, reg.Phone);
  var s = getLocalizedStrings(lang);

  var htmlBody = buildEmailHtml(reg, s, ticketNumber);

  MailApp.sendEmail({
    to: reg.Email,
    subject: s.subject + " — " + ticketNumber,
    htmlBody: htmlBody
  });

  return ticketNumber;
}

function sendAdminNotification(reg, ticketNumber) {
  var name = reg.Name || "";
  var email = reg.Email || "";
  var plan = reg.Plan || "";
  var price = reg.Price || "";
  var lang = reg.Language || "";
  var nationality = reg.Nationality || "";
  var phone = reg.Phone || "";
  var memo = reg.Memo || "";

  var subject = "[Accelerate 2026] 컨펌 메일 발송 완료 — " + name + " (" + ticketNumber + ")";

  var html = '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"></head>'
    + '<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 20px;">'
    + '<tr><td align="center">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);">'

    + '<tr><td style="background:#1a1a2e;padding:24px 32px;">'
    + '<h2 style="margin:0;color:#fff;font-size:18px;">Accelerate 2026 — Admin Notification</h2>'
    + '</td></tr>'

    + '<tr><td style="padding:24px 32px;">'
    + '<p style="font-size:15px;color:#333;margin:0 0 16px;">아래 참가자에게 컨펌 이메일이 성공적으로 발송되었습니다.</p>'

    + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-bottom:20px;">'
    + adminRow("티켓 번호", ticketNumber)
    + adminRow("이름", name)
    + adminRow("이메일", email)
    + adminRow("국적", nationality)
    + adminRow("연락처", phone)
    + adminRow("플랜", plan)
    + adminRow("금액", price)
    + adminRow("언어", lang)
    + adminRow("메모", memo || "(없음)")
    + '</table>'

    + '<p style="font-size:12px;color:#999;margin:0;">이 메일은 구글 시트에서 "paid" 입력 시 자동 발송됩니다.</p>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';

  MailApp.sendEmail({
    to: ADMIN_EMAILS,
    subject: subject,
    htmlBody: html
  });
}

function adminRow(label, value) {
  return '<tr>'
    + '<td style="padding:8px 0;border-bottom:1px solid #f0f0f0;font-size:13px;color:#888;width:30%;">' + label + '</td>'
    + '<td style="padding:8px 0;border-bottom:1px solid #f0f0f0;font-size:14px;color:#222;font-weight:600;">' + value + '</td>'
    + '</tr>';
}

function generateTicketNumber(name, phone) {
  var n = (name || "USER").toString().trim();
  var digits = (phone || "").toString().replace(/[^0-9]/g, "");
  var last4 = digits.length >= 4 ? digits.slice(-4) : ("0000" + digits).slice(-4);
  return "ACC26-" + n + "-" + last4;
}

// ─────────────────────────────────────────────
// 4. 다국어 문자열
// ─────────────────────────────────────────────

function getLocalizedStrings(lang) {
  var map = {
    en: {
      subject: "Accelerate 2026 — Registration Confirmed",
      greeting: "Dear",
      greetingSuffix: ",",
      confirmed: "CONFIRMED",
      ticketLabel: "Ticket Number",
      planLabel: "Selected Plan",
      priceLabel: "Price",
      nameLabel: "Attendee",
      dateLabel: "Date",
      venueLabel: "Venue",
      dateValue: "June 1 (Mon) – June 2 (Tue), 2026",
      venueValue: "Swiss Grand Hotel Seoul, South Korea",
      qrNote: "Please present this ticket at the venue entrance for check-in.",
      closingLine: "We look forward to seeing you at Accelerate 2026!",
      teamSign: "BNI KOREA NATIONAL SUPPORT TEAM 2026",
      footer: "If you have any questions, please contact admin@bni-korea.com"
    },
    ja: {
      subject: "Accelerate 2026 — ご登録確認",
      greeting: "",
      greetingSuffix: " 様",
      confirmed: "登録確認済み",
      ticketLabel: "チケット番号",
      planLabel: "選択プラン",
      priceLabel: "料金",
      nameLabel: "参加者名",
      dateLabel: "開催日程",
      venueLabel: "会場",
      dateValue: "2026年6月1日（月）〜 6月2日（火）",
      venueValue: "スイスグランドホテルソウル、韓国",
      qrNote: "会場入口でこのチケットをご提示ください。",
      closingLine: "Accelerate 2026でお会いできることを楽しみにしております！",
      teamSign: "BNI KOREA NATIONAL SUPPORT TEAM 2026",
      footer: "ご不明な点がございましたら admin@bni-korea.com までお問い合わせください"
    },
    zh: {
      subject: "Accelerate 2026 — 注册确认",
      greeting: "尊敬的 ",
      greetingSuffix: "",
      confirmed: "注册已确认",
      ticketLabel: "票号",
      planLabel: "所选方案",
      priceLabel: "价格",
      nameLabel: "参会者",
      dateLabel: "日期",
      venueLabel: "地点",
      dateValue: "2026年6月1日（星期一）～ 6月2日（星期二）",
      venueValue: "瑞士大酒店首尔，韩国",
      qrNote: "请在会场入口出示此票办理签到。",
      closingLine: "我们期待在 Accelerate 2026 与您相见！",
      teamSign: "BNI KOREA NATIONAL SUPPORT TEAM 2026",
      footer: "如有任何问题，请联系 admin@bni-korea.com"
    },
    ko: {
      subject: "Accelerate 2026 — 등록 확인",
      greeting: "",
      greetingSuffix: " 님",
      confirmed: "등록 확인 완료",
      ticketLabel: "티켓 번호",
      planLabel: "선택 플랜",
      priceLabel: "결제 금액",
      nameLabel: "참가자",
      dateLabel: "행사 일정",
      venueLabel: "장소",
      dateValue: "2026년 6월 1일 (월) ~ 6월 2일 (화)",
      venueValue: "스위스그랜드호텔 서울",
      qrNote: "행사장 입구에서 이 티켓을 제시해 주세요.",
      closingLine: "Accelerate 2026에서 만나 뵙겠습니다!",
      teamSign: "BNI KOREA NATIONAL SUPPORT TEAM 2026",
      footer: "문의사항이 있으시면 admin@bni-korea.com으로 연락해 주세요"
    }
  };
  return map[lang] || map["en"];
}

// ─────────────────────────────────────────────
// 5. HTML 이메일 템플릿
// ─────────────────────────────────────────────

function buildEmailHtml(reg, s, ticketNumber) {
  var name = reg.Name || "";
  var greetingLine = s.greeting + name + s.greetingSuffix;

  return '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>'
    + '<body style="margin:0;padding:0;background:#f0f0f0;font-family:Georgia,\'Times New Roman\',serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f0f0;padding:40px 20px;">'
    + '<tr><td align="center">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">'

    // ── Header ──
    + '<tr><td style="background:#cf1f2e;padding:36px 40px;text-align:center;">'
    + '<h1 style="margin:0;color:#fff;font-size:26px;letter-spacing:3px;font-weight:800;">ACCELERATE 2026</h1>'
    + '<p style="margin:8px 0 0;color:rgba(255,255,255,0.8);font-size:13px;letter-spacing:1px;">BNI KOREA NATIONAL CONFERENCE</p>'
    + '</td></tr>'

    // ── Confirmed Badge ──
    + '<tr><td style="padding:32px 40px 16px;text-align:center;">'
    + '<div style="display:inline-block;background:#cf1f2e;color:#fff;padding:10px 36px;border-radius:40px;font-size:16px;font-weight:800;letter-spacing:3px;">'
    + '&#10003; ' + s.confirmed
    + '</div>'
    + '</td></tr>'

    // ── Greeting ──
    + '<tr><td style="padding:12px 40px 8px;">'
    + '<p style="font-size:16px;color:#333;margin:0;">' + greetingLine + '</p>'
    + '</td></tr>'

    // ── Ticket Number Box ──
    + '<tr><td style="padding:12px 40px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fef2f2;border:2px solid #fca5a5;border-radius:10px;">'
    + '<tr><td style="padding:20px;text-align:center;">'
    + '<div style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:2px;margin-bottom:6px;">' + s.ticketLabel + '</div>'
    + '<div style="font-size:24px;font-weight:800;color:#cf1f2e;letter-spacing:2px;font-family:monospace;">' + ticketNumber + '</div>'
    + '</td></tr></table>'
    + '</td></tr>'

    // ── Details Table ──
    + '<tr><td style="padding:20px 40px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">'
    + detailRow(s.nameLabel, name)
    + detailRow(s.planLabel, reg.Plan || "")
    + detailRow(s.priceLabel, reg.Price || "")
    + detailRow(s.dateLabel, s.dateValue)
    + detailRow(s.venueLabel, '<a href="https://www.swissgrand.co.kr/" style="color:#cf1f2e;text-decoration:underline;">' + s.venueValue + '</a>')
    + '</table>'
    + '</td></tr>'

    // ── Divider ──
    + '<tr><td style="padding:0 40px;"><div style="border-top:1px solid #eee;"></div></td></tr>'

    // ── Closing ──
    + '<tr><td style="padding:24px 40px 8px;">'
    + '<p style="font-size:17px;color:#222;margin:0;line-height:1.6;font-weight:700;text-align:center;">' + s.closingLine + '</p>'
    + '</td></tr>'
    + '<tr><td style="padding:4px 40px 32px;">'
    + '<p style="font-size:12px;color:#888;margin:0;font-weight:400;text-align:center;">' + s.teamSign + '</p>'
    + '</td></tr>'

    // ── Footer ──
    + '<tr><td style="background:#1a1a2e;padding:24px 40px;text-align:center;">'
    + '<p style="margin:0 0 8px;color:#fff;font-size:13px;font-weight:700;">Accelerate Your Success with BNI Korea</p>'
    + '<p style="margin:0 0 8px;color:rgba(255,255,255,0.6);font-size:11px;">Swiss Grand Hotel Seoul &bull; Jun 1–2, 2026</p>'
    + '<p style="margin:0;color:rgba(255,255,255,0.4);font-size:11px;">' + s.footer + '</p>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}

function detailRow(label, value) {
  return '<tr>'
    + '<td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-size:13px;color:#888;width:35%;">' + label + '</td>'
    + '<td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-size:14px;color:#111;font-weight:600;">' + value + '</td>'
    + '</tr>';
}

// ─────────────────────────────────────────────
// 유틸: 시트 생성/조회
// ─────────────────────────────────────────────

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#cf1f2e");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(STATUS_COL, 180);
  }

  return sheet;
}

// ─────────────────────────────────────────────
// 디버그용 (테스트 후 삭제 가능)
// ─────────────────────────────────────────────

function debugCheck() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registrations");
  var val = sheet.getRange(2, 10).getValue();
  sheet.getRange(2, 11).setValue("Value:[" + val + "] Type:" + typeof val + " Match:" + (val.toString().trim() === "paid"));
}
