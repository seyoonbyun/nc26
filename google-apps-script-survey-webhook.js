// ============================================================
// NC26 Survey Webhook — Typebot → Airtable + Thank-you Email
// ============================================================
// Deploy: Apps Script Web App (Execute as me, Anyone can access)
// Typebot webhook POSTs survey data here.
// This script:
//   1. Extracts numeric ratings from Typebot choice text ("3 보통" → 3)
//   2. Inserts a row into Airtable
//   3. Sends a thank-you email via Gmail

var CONFIG = {
  AIRTABLE_TOKEN: PropertiesService.getScriptProperties().getProperty("AIRTABLE_TOKEN"),
  AIRTABLE_BASE:  "app3rkaUQZX4JE0x7",
  AIRTABLE_TABLE: "tblL98VCc1ieSsUfi",
  SENDER_NAME:    "NC26 BNI Korea",
  SENDER_EMAIL:   "syoon850@gmail.com"
};

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var record = buildRecord(body);
    insertAirtable(record);
    sendThankYouEmail(body.email, body.language || "한국어");
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log("doPost error: " + err);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- Airtable row builder ---
function buildRecord(b) {
  return {
    fields: {
      language:              b.language || "",
      email:                 b.email || "",
      attendee_type:         b.attendee_type || "",
      schedule:              b.schedule || "",
      overall_satisfaction:  extractNumber(b.overall_satisfaction),
      expectations:          extractNumber(b.expectations),
      nps:                   extractNumber(b.nps),
      reattend_intent:       b.reattend_intent || "",
      reattend_reason:       b.reattend_reason || "",
      reattend_detail:       b.reattend_detail || "",
      day1_overall:          extractNumber(b.day1_overall),
      day1_best_session:     b.day1_best_session || "",
      day1_best_reason:      b.day1_best_reason || "",
      day1_worst_session:    b.day1_worst_session || "",
      day1_worst_reason:     b.day1_worst_reason || "",
      day1_comment:          b.day1_comment || "",
      day2_overall:          extractNumber(b.day2_overall),
      day2_best_session:     b.day2_best_session || "",
      day2_best_reason:      b.day2_best_reason || "",
      day2_worst_session:    b.day2_worst_session || "",
      day2_worst_reason:     b.day2_worst_reason || "",
      day2_comment:          b.day2_comment || "",
      ops_checkin:           extractNumber(b.ops_checkin),
      ops_guidance:          extractNumber(b.ops_guidance),
      ops_facilities:        extractNumber(b.ops_facilities),
      ops_food:              extractNumber(b.ops_food),
      ops_booth:             extractNumber(b.ops_booth),
      ops_networking:        extractNumber(b.ops_networking),
      next_programs:         b.next_programs || "",
      next_topics:           b.next_topics || "",
      best_thing:            b.best_thing || "",
      improvement:           b.improvement || "",
      additional:            b.additional || "",
      submitted_at:          new Date().toISOString()
    }
  };
}

// "3 보통" → 3,  "7" → 7,  "" → null
function extractNumber(val) {
  if (val == null || val === "") return null;
  var s = String(val).trim();
  var m = s.match(/^(\d+)/);
  return m ? parseInt(m[1], 10) : null;
}

// --- Airtable insert ---
function insertAirtable(record) {
  var url = "https://api.airtable.com/v0/" + CONFIG.AIRTABLE_BASE + "/" + CONFIG.AIRTABLE_TABLE;
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + CONFIG.AIRTABLE_TOKEN },
    payload: JSON.stringify(record),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error("Airtable " + res.getResponseCode() + ": " + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

// --- Thank-you email ---
function sendThankYouEmail(email, lang) {
  if (!email) return;

  var subject, htmlBody;

  if (lang === "English") {
    subject = "Thank you for completing the NC26 Survey!";
    htmlBody = buildEmailHtml(
      "Thank You!",
      "Thank you for taking the time to share your feedback on NC26.",
      "Your responses will help us create an even better National Conference next time.",
      "We look forward to seeing you again!",
      "NC26 BNI Korea Team"
    );
  } else if (lang === "日本語") {
    subject = "NC26 アンケートにご回答いただきありがとうございます！";
    htmlBody = buildEmailHtml(
      "ありがとうございます！",
      "NC26に関する貴重なご意見をお寄せいただき、誠にありがとうございます。",
      "いただいたご回答は、次回のナショナルカンファレンスをより良いものにするために活用させていただきます。",
      "またお会いできることを楽しみにしております！",
      "NC26 BNI Korea チーム"
    );
  } else if (lang === "中文") {
    subject = "感谢您完成NC26问卷调查！";
    htmlBody = buildEmailHtml(
      "感谢您！",
      "感谢您抽出宝贵时间分享您对NC26的反馈。",
      "您的回答将帮助我们打造更好的下一届全国大会。",
      "期待再次与您相见！",
      "NC26 BNI Korea 团队"
    );
  } else {
    subject = "NC26 설문에 응답해주셔서 감사합니다!";
    htmlBody = buildEmailHtml(
      "감사합니다!",
      "NC26에 대한 소중한 의견을 나눠주셔서 진심으로 감사드립니다.",
      "보내주신 응답은 다음 내셔널 컨퍼런스를 더욱 발전시키는 데 소중하게 활용하겠습니다.",
      "다음 행사에서 다시 뵙기를 기대합니다!",
      "NC26 BNI Korea 팀"
    );
  }

  GmailApp.sendEmail(email, subject, "", {
    htmlBody: htmlBody,
    name: CONFIG.SENDER_NAME,
    from: CONFIG.SENDER_EMAIL
  });
}

function buildEmailHtml(heading, line1, line2, line3, signature) {
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:\'Apple SD Gothic Neo\',\'Malgun Gothic\',sans-serif;background:#f5f5f5">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:40px 20px">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.08)">'
    + '<tr><td style="background:#db0000;padding:32px 40px;text-align:center">'
    + '<h1 style="color:#fff;margin:0;font-size:28px">' + heading + '</h1>'
    + '</td></tr>'
    + '<tr><td style="padding:40px">'
    + '<p style="font-size:16px;line-height:1.7;color:#333;margin:0 0 16px">' + line1 + '</p>'
    + '<p style="font-size:16px;line-height:1.7;color:#333;margin:0 0 16px">' + line2 + '</p>'
    + '<p style="font-size:16px;line-height:1.7;color:#333;margin:0 0 32px">' + line3 + '</p>'
    + '<p style="font-size:14px;color:#888;margin:0;border-top:1px solid #eee;padding-top:20px">' + signature + '</p>'
    + '</td></tr>'
    + '</table></td></tr></table></body></html>';
}
