/**
 * nc26 굿즈 부스 재고 관리 (Code_deploy.gs에 추가)
 *
 * 시트 자동 생성:
 *   - "굿즈_재고"    : 마스터 + 초기재고/판매수량/남은재고
 *   - "굿즈_판매기록" : 모든 판매·취소 로그 (append-only)
 *
 * doGet(e) 에 다음 분기 추가 ── (Code_deploy.gs)의 doGet 함수 안:
 *   if (action === "goodsList")   return goodsList();
 *   if (action === "goodsSale")   return goodsSale(e.parameter);
 *   if (action === "goodsCancel") return goodsCancel(e.parameter);
 *   if (action === "goodsLog")    return goodsLog(e.parameter);
 *
 * goods.html 의 SCRIPT_URL 은 기존 Code_deploy.gs 의 웹앱 URL 그대로 사용.
 * (배포 > 새 버전 만들기 한 번 눌러야 변경사항 반영됨)
 */

// ─────────────────────────────────────────────
// 시트 정의
// ─────────────────────────────────────────────
var GOODS_STOCK_SHEET = "굿즈_재고";
var GOODS_STOCK_HEADERS = ["ID", "Category", "Name", "Variant", "InitStock", "SoldQty", "Remain", "PriceSpc"];

var GOODS_LOG_SHEET = "굿즈_판매기록";
var GOODS_LOG_HEADERS = ["Timestamp", "Ref", "Action", "ID", "Name", "Variant", "Qty", "PaymentMethod", "Staff", "Memo", "Status"];

// 초기 재고 시드 (goods.html 의 GOODS_MASTER 와 1:1 매칭)
var GOODS_SEED = [
  // id, cat, name, variant, init, priceSpc
  ["tee-w-s",    "의류",    "화이트 카라티",    "S",   9,   35000],
  ["tee-w-m",    "의류",    "화이트 카라티",    "M",   17,  35000],
  ["tee-w-l",    "의류",    "화이트 카라티",    "L",   23,  35000],
  ["tee-w-xl",   "의류",    "화이트 카라티",    "XL",  23,  35000],
  ["tee-w-2xl",  "의류",    "화이트 카라티",    "2XL", 9,   35000],
  ["tee-b-s",    "의류",    "블랙 카라티",      "S",   10,  35000],
  ["tee-b-m",    "의류",    "블랙 카라티",      "M",   15,  35000],
  ["tee-b-l",    "의류",    "블랙 카라티",      "L",   14,  35000],
  ["tee-b-xl",   "의류",    "블랙 카라티",      "XL",  22,  35000],
  ["tee-b-2xl",  "의류",    "블랙 카라티",      "2XL", 7,   35000],
  ["cap-blk",    "액세서리", "BNI 볼캡",         "블랙", 20,  22000],
  ["cap-red",    "액세서리", "BNI 볼캡",         "레드", 20,  22000],
  ["scarf",      "액세서리", "BNI 실크스카프",   "",    30,  39000],
  ["tie-std",    "액세서리", "BNI 넥타이",       "일반", 100, 38000],
  ["tie-auto",   "액세서리", "BNI 넥타이",       "자동", 100, 38000],
  ["pouch",      "액세서리", "nc26 랩탑 파우치",  "",    100, 29000],
  ["mug",        "굿즈",    "BNI 머그컵",       "",    83,  15000],
  ["book",       "굿즈",    "멘토링 도서",       "",    32,  20000],
  ["patch",      "굿즈",    "와펜 (런칭)",      "",    8,   27500],
  ["badge-fire", "뱃지",    "횃불 뱃지",         "",    150, 10000],
  ["badge-conf", "뱃지",    "2026 컨퍼런스 뱃지","",    1000,10000],
  ["tablecov",   "대형",    "BNI 테이블보",     "",    2,   170000]
];

// ─────────────────────────────────────────────
// 시트 헬퍼
// ─────────────────────────────────────────────
function getOrCreateGoodsStockSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GOODS_STOCK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GOODS_STOCK_SHEET);
    sheet.getRange(1, 1, 1, GOODS_STOCK_HEADERS.length).setValues([GOODS_STOCK_HEADERS]).setFontWeight("bold").setBackground("#f0f0f0");
    sheet.setFrozenRows(1);
    // seed
    var rows = GOODS_SEED.map(function (s) {
      // [ID, Cat, Name, Variant, Init, Sold=0, Remain=Init, PriceSpc]
      return [s[0], s[1], s[2], s[3], s[4], 0, s[4], s[5]];
    });
    sheet.getRange(2, 1, rows.length, GOODS_STOCK_HEADERS.length).setValues(rows);
    // Remain = InitStock - SoldQty (formula)
    for (var i = 0; i < rows.length; i++) {
      sheet.getRange(i + 2, 7).setFormula("=E" + (i + 2) + "-F" + (i + 2));
    }
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 180); sheet.setColumnWidth(4, 70);
    sheet.autoResizeColumns(5, 4);
  }
  return sheet;
}

function getOrCreateGoodsLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GOODS_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GOODS_LOG_SHEET);
    sheet.getRange(1, 1, 1, GOODS_LOG_HEADERS.length).setValues([GOODS_LOG_HEADERS]).setFontWeight("bold").setBackground("#f0f0f0");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 70);  sheet.setColumnWidth(4, 110);
    sheet.setColumnWidth(5, 160); sheet.setColumnWidth(6, 70);
    sheet.setColumnWidth(7, 60);  sheet.setColumnWidth(8, 90);
    sheet.setColumnWidth(9, 80);  sheet.setColumnWidth(10, 200);
    sheet.setColumnWidth(11, 80);
  }
  return sheet;
}

function _goodsFindStockRow(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function _goodsFindLogRow(sheet, ref) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var refs = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < refs.length; i++) {
    if (String(refs[i][0]) === String(ref)) return i + 2;
  }
  return -1;
}

function _goodsJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _goodsGenRef() {
  var d = new Date();
  var ts = Utilities.formatDate(d, "Asia/Seoul", "yyMMddHHmmss");
  var rnd = Math.floor(Math.random() * 1000).toString().padStart(3, "0");
  return "G" + ts + rnd;
}

// ─────────────────────────────────────────────
// API: 재고 목록 조회
// ─────────────────────────────────────────────
function goodsList() {
  try {
    var sheet = getOrCreateGoodsStockSheet();
    var last = sheet.getLastRow();
    if (last < 2) return _goodsJson({ ok: true, items: [] });
    var data = sheet.getRange(2, 1, last - 1, GOODS_STOCK_HEADERS.length).getValues();
    var items = data.map(function (r) {
      return {
        id: r[0], cat: r[1], name: r[2], variant: r[3],
        init: Number(r[4] || 0), sold: Number(r[5] || 0), remain: Number(r[6] || 0),
        priceSpc: Number(r[7] || 0)
      };
    });
    return _goodsJson({ ok: true, items: items });
  } catch (err) {
    return _goodsJson({ ok: false, error: String(err && err.message || err) });
  }
}

// ─────────────────────────────────────────────
// API: 판매 기록 (atomic)
// ─────────────────────────────────────────────
function goodsSale(params) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return _goodsJson({ ok: false, error: "lock timeout" }); }
  try {
    var id   = String(params.id || "").trim();
    var qty  = Number(params.qty || 0);
    var pay  = String(params.pay || "").trim() || "카드";
    var staff = String(params.staff || "").trim() || "Joy";
    var memo = String(params.memo || "").trim();
    if (!id || qty <= 0) return _goodsJson({ ok: false, error: "invalid params" });

    var stockSheet = getOrCreateGoodsStockSheet();
    var row = _goodsFindStockRow(stockSheet, id);
    if (row < 0) return _goodsJson({ ok: false, error: "item not found: " + id });

    var name    = stockSheet.getRange(row, 3).getValue();
    var variant = stockSheet.getRange(row, 4).getValue();
    var sold    = Number(stockSheet.getRange(row, 6).getValue() || 0);
    stockSheet.getRange(row, 6).setValue(sold + qty);
    SpreadsheetApp.flush();
    var remain = Number(stockSheet.getRange(row, 7).getValue() || 0);

    var ref = _goodsGenRef();
    var logSheet = getOrCreateGoodsLogSheet();
    logSheet.appendRow([
      new Date(), ref, "SALE", id, name, variant, qty, pay, staff, memo, "ACTIVE"
    ]);

    return _goodsJson({ ok: true, ref: ref, id: id, remain: remain });
  } catch (err) {
    return _goodsJson({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// API: 판매 취소 (atomic) — 로그 Status 변경 + 재고 복원
// ─────────────────────────────────────────────
function goodsCancel(params) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return _goodsJson({ ok: false, error: "lock timeout" }); }
  try {
    var ref = String(params.ref || "").trim();
    if (!ref) return _goodsJson({ ok: false, error: "missing ref" });

    var logSheet = getOrCreateGoodsLogSheet();
    var logRow = _goodsFindLogRow(logSheet, ref);
    if (logRow < 0) return _goodsJson({ ok: false, error: "log not found: " + ref });

    var status = String(logSheet.getRange(logRow, 11).getValue() || "");
    if (status === "CANCELLED") return _goodsJson({ ok: false, error: "already cancelled" });

    var id   = String(logSheet.getRange(logRow, 4).getValue() || "");
    var name = logSheet.getRange(logRow, 5).getValue();
    var variant = logSheet.getRange(logRow, 6).getValue();
    var qty  = Number(logSheet.getRange(logRow, 7).getValue() || 0);

    // 재고 복원
    var stockSheet = getOrCreateGoodsStockSheet();
    var stRow = _goodsFindStockRow(stockSheet, id);
    if (stRow < 0) return _goodsJson({ ok: false, error: "stock row not found" });
    var sold = Number(stockSheet.getRange(stRow, 6).getValue() || 0);
    stockSheet.getRange(stRow, 6).setValue(Math.max(0, sold - qty));
    SpreadsheetApp.flush();
    var remain = Number(stockSheet.getRange(stRow, 7).getValue() || 0);

    // 원본 로그 Status 변경 + 취소 로그 1행 추가
    logSheet.getRange(logRow, 11).setValue("CANCELLED");
    var cancelRef = _goodsGenRef();
    logSheet.appendRow([
      new Date(), cancelRef, "CANCEL", id, name, variant, qty, "-", "", "취소: " + ref, "ACTIVE"
    ]);

    return _goodsJson({ ok: true, ref: cancelRef, origRef: ref, id: id, remain: remain });
  } catch (err) {
    return _goodsJson({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────
// API: 최근 판매 내역
// ─────────────────────────────────────────────
function goodsLog(params) {
  try {
    var limit = Math.min(200, Math.max(1, Number((params && params.limit) || 50)));
    var sheet = getOrCreateGoodsLogSheet();
    var last = sheet.getLastRow();
    if (last < 2) return _goodsJson({ ok: true, logs: [] });
    var start = Math.max(2, last - limit + 1);
    var data = sheet.getRange(start, 1, last - start + 1, GOODS_LOG_HEADERS.length).getValues();
    // 최신순
    data.reverse();
    var logs = data.map(function (r) {
      return {
        ts: r[0] ? new Date(r[0]).toISOString() : "",
        ref: r[1], action: r[2], id: r[3],
        name: (r[5] ? r[4] + " · " + r[5] : r[4]),
        qty: r[6], pay: r[7], staff: r[8], memo: r[9], status: r[10]
      };
    });
    return _goodsJson({ ok: true, logs: logs });
  } catch (err) {
    return _goodsJson({ ok: false, error: String(err && err.message || err) });
  }
}

// ─────────────────────────────────────────────
// 운영: 시드 강제 재실행 (수동 호출 전용)
//   - 초기재고만 다시 채움. 판매기록·SoldQty는 건드리지 않음
// ─────────────────────────────────────────────
function goodsReseedInit() {
  var sheet = getOrCreateGoodsStockSheet();
  var last = sheet.getLastRow();
  if (last < 2) return;
  GOODS_SEED.forEach(function (s) {
    var row = _goodsFindStockRow(sheet, s[0]);
    if (row > 0) {
      sheet.getRange(row, 5).setValue(s[4]);     // InitStock
      sheet.getRange(row, 8).setValue(s[5]);     // PriceSpc
    } else {
      sheet.appendRow([s[0], s[1], s[2], s[3], s[4], 0, s[4], s[5]]);
      var newRow = sheet.getLastRow();
      sheet.getRange(newRow, 7).setFormula("=E" + newRow + "-F" + newRow);
    }
  });
}
