var SS = SpreadsheetApp.getActiveSpreadsheet();

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

    // text/plain JSON (booth with files)
    if (e.postData && e.postData.type && e.postData.type.indexOf('text/plain') > -1) {
      try {
        p = JSON.parse(e.postData.contents);
      } catch (err) {
        Logger.log('JSON parse error: ' + err.message);
      }
    }

    // 파일 별도 수신 (booth-files)
    if (p.type === 'booth-files') {
      var sheet = SS.getSheetByName('Booth');
      if (sheet) {
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var logoCol = headers.indexOf('Logo File') + 1;
        var adCol = headers.indexOf('Ad File') + 1;
        // 해당 회사의 마지막 행 찾기
        var companyCol = headers.indexOf('Company') + 1;
        var data = sheet.getRange(2, companyCol, sheet.getLastRow() - 1, 1).getValues();
        var targetRow = -1;
        for (var i = data.length - 1; i >= 0; i--) {
          if (data[i][0] === p.company) { targetRow = i + 2; break; }
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

      var boothSheetName = (p.boothType === 'domestic') ? 'Booth_Kor' : 'Booth';
      var sheet = getOrCreateSheet(boothSheetName, BOOTH_HEADERS);
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

      // Homepage -> clickable link
      if (p.homepage) {
        setHyperlink(sheet, newRow, 8, p.homepage, p.homepage);
      }
      // Email -> mailto link
      if (p.email) {
        setHyperlink(sheet, newRow, 9, 'mailto:' + p.email, p.email);
      }
      // Applicant Email -> mailto link
      if (p.applicantEmail) {
        setHyperlink(sheet, newRow, 12, 'mailto:' + p.applicantEmail, p.applicantEmail);
      }
      // Logo File -> Drive link
      if (logoUrl) {
        setHyperlink(sheet, newRow, 17, logoUrl, logoName);
      }
      // Ad File -> Drive link
      if (adUrl) {
        setHyperlink(sheet, newRow, 18, adUrl, adName);
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

      // Email -> mailto link
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

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'NC26 Overseas API is running.' })
  ).setMimeType(ContentService.MimeType.JSON);
}

function onEditInstallable(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var sheetName = sheet.getName();

  if (sheetName !== 'Tickets' && sheetName !== 'Booth' && sheetName !== 'Booth_Kor') return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusCol = headers.indexOf('Status') + 1;
  if (statusCol === 0) return;

  if (range.getColumn() !== statusCol) return;
  if (range.getValue() !== 'Paid') return;

  var row = range.getRow();
  if (row <= 1) return;

  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

  if (sheetName === 'Tickets') {
    sendTicketConfirmationEmail(headers, rowData);
  } else if (sheetName === 'Booth' || sheetName === 'Booth_Kor') {
    sendBoothVoucherEmail(headers, rowData);
  }
}

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

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
  Logger.log('Ticket confirmation email sent to: ' + email);
}

function sendBoothVoucherEmail(headers, row) {
  var get = function(key) { return row[headers.indexOf(key)] || ''; };

  var company = get('Company');
  var displayName = get('Display Name');
  var email = get('Email');
  var applicantName = get('Applicant Name');
  var applicantEmail = get('Applicant Email');
  var country = get('Country');
  var price = get('Price');

  var recipients = [];
  if (email) recipients.push(email);
  if (applicantEmail && applicantEmail !== email) recipients.push(applicantEmail);
  if (recipients.length === 0) return;
  var recipientEmail = recipients.join(',');

  var voucherNo = 'BNI-BOOTH-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd') + '-' + Math.random().toString(36).substring(2, 7).toUpperCase();

  var subject = '\u3010Accelerate 2026\u3011Booth Payment Confirmed - Voucher';

  var displayNameRow = displayName ? '<tr><td style="color:#999;padding:6px 0;">Display Name</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + displayName + '</td></tr>' : '';
  var countryRow = country ? '<tr><td style="color:#999;padding:6px 0;">Country</td><td style="color:#1a1a1a;font-weight:600;padding:6px 0;">' + country + '</td></tr>' : '';
  var recipientName = applicantName || company;

  var body = '<!DOCTYPE html>'
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

  MailApp.sendEmail({ to: recipientEmail, subject: subject, htmlBody: body });
  Logger.log('Booth voucher email sent to: ' + recipientEmail);
}
