/**
 * nc26 Apps Script 프로젝트 — Self Health Check
 *
 * 5/11 발송 사고 + 5/17 트리거 fail 사고 같은 "함수가 코드에 없는데 트리거는
 * 살아있어서 매분 폭주" 패턴을 자동 감지.
 *
 * 점검 항목:
 *   1. 트리거 vs 함수 정의 일치성 (등록된 trigger handler가 실제 globalThis에
 *      존재하는지) — 이번 사고의 직접 원인
 *   2. 핵심 시트 존재 + 최소 컬럼 수
 *   3. Gmail 발송 일일 잔량
 *
 * 설치:
 *   1. nc26 Apps Script 콘솔에 새 파일 'HealthCheck' 추가
 *   2. 이 파일 전체 복붙
 *   3. 함수 nc26HealthCheck ▶ 1회 수동 실행 (권한 승인 — MailApp 추가됨)
 *   4. 시간 기반 트리거 등록: nc26HealthCheck / 일일 타이머 / 오전 8-9시
 *
 * 결과 출력:
 *   - 시트 '__HealthLog' 에 매회 1행 append (Timestamp / Status / Issues)
 *   - FAIL 시 NC26_HEALTH_ALERT_RECIPIENT 로 알림 메일
 */

// 알림 수신자 — 셋업 시 사용자 입력으로 치환
var NC26_HEALTH_ALERT_RECIPIENT = 'hq@joy-bnikorea.com';

// 예상 시트 구조 (시트명 → 최소 컬럼 수)
var NC26_EXPECTED_SHEETS = {
  'Ticket & Booth_Kor.pay': 13,
  'SecretPass': 9,
  'PartyCodes': 6
};

// Gmail 잔량 경고 임계치
var NC26_MAIL_QUOTA_THRESHOLD = 50;

function nc26HealthCheck() {
  var issues = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. 트리거 vs 함수 정의 일치성
  try {
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t) {
      var fn = t.getHandlerFunction();
      var exists = false;
      try {
        // Apps Script 글로벌 함수는 this[fn] 로 접근 가능
        exists = (typeof this[fn] === 'function');
      } catch (e) {
        exists = false;
      }
      if (!exists) {
        var src = t.getEventType ? t.getEventType() : 'unknown';
        issues.push('TRIGGER_FN_MISSING: handler=' + fn + ' (event=' + src + ')');
      }
    });
  } catch (e) {
    issues.push('TRIGGER_SCAN_ERROR: ' + e);
  }

  // ── 2. 핵심 시트 존재 + 최소 컬럼
  Object.keys(NC26_EXPECTED_SHEETS).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      issues.push('SHEET_MISSING: ' + name);
      return;
    }
    var cols = sh.getLastColumn();
    var minCols = NC26_EXPECTED_SHEETS[name];
    if (cols < minCols) {
      issues.push('SHEET_COL_SHORT: ' + name + ' has ' + cols + ' cols (need ' + minCols + ')');
    }
  });

  // ── 3. Gmail 일일 잔량
  try {
    var remaining = MailApp.getRemainingDailyQuota();
    if (remaining < NC26_MAIL_QUOTA_THRESHOLD) {
      issues.push('LOW_MAIL_QUOTA: ' + remaining + ' remaining (threshold ' + NC26_MAIL_QUOTA_THRESHOLD + ')');
    }
  } catch (e) {
    issues.push('MAIL_QUOTA_ERROR: ' + e);
  }

  // ── 로그 시트에 기록
  var logSheet = ss.getSheetByName('__HealthLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('__HealthLog');
    logSheet.appendRow(['Timestamp', 'Status', 'Issues']);
    var hdr = logSheet.getRange(1, 1, 1, 3);
    hdr.setFontWeight('bold');
    hdr.setBackground('#cf1f2e');
    hdr.setFontColor('#ffffff');
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 170);
    logSheet.setColumnWidth(2, 70);
    logSheet.setColumnWidth(3, 800);
    logSheet.hideSheet(); // 운영자 시야에서 숨김 (수동으로 다시 보기 가능)
  }
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  var status = issues.length === 0 ? 'OK' : 'FAIL';
  logSheet.appendRow([stamp, status, issues.join(' | ')]);

  // ── FAIL 시 알림 메일 (멱등: 같은 issue 셋이 연속이면 1일 1회만 발송)
  if (issues.length > 0) {
    var lastAlertProp = PropertiesService.getScriptProperties().getProperty('NC26_HEALTH_LAST_ALERT');
    var todayKey = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    var issueKey = issues.join('||');
    var thisKey = todayKey + ':' + Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, issueKey
    ).map(function(b){return ('0'+(b&0xff).toString(16)).slice(-2);}).join('');

    if (lastAlertProp !== thisKey) {
      try {
        MailApp.sendEmail({
          to: NC26_HEALTH_ALERT_RECIPIENT,
          subject: '[nc26 헬스체크] ' + issues.length + '건 이상 — ' + stamp,
          body: '시각: ' + stamp + '\n프로젝트: nc26 Apps Script\n\n이상 항목:\n' +
                issues.map(function(i){return '  - ' + i;}).join('\n') +
                '\n\n조치:\n' +
                '  • TRIGGER_FN_MISSING — 트리거 페이지에서 해당 트리거 삭제 또는 함수 본체 복구\n' +
                '  • SHEET_MISSING / SHEET_COL_SHORT — 시트 복구 또는 코드의 기대 구조 갱신\n' +
                '  • LOW_MAIL_QUOTA — 24시간 이내 대량 발송 작업 보류\n\n' +
                '__HealthLog 시트에서 전체 이력 확인 가능 (숨김 처리되어 있으니 시트 탭 우클릭 > 숨기기 해제).\n\n' +
                '— nc26HealthCheck (자동 발송, 같은 이슈 셋은 하루 1회만)'
        });
        PropertiesService.getScriptProperties().setProperty('NC26_HEALTH_LAST_ALERT', thisKey);
      } catch (e) {
        Logger.log('Alert mail failed: ' + e);
      }
    } else {
      Logger.log('Same issue set as last alert today — skip mail (idempotent)');
    }
  }

  Logger.log('Health check: ' + status + (issues.length ? ' — ' + issues.join(' | ') : ''));
  return { status: status, issues: issues, timestamp: stamp };
}

/**
 * 수동 점검 — 결과를 Logger 에만 출력하고 로그 시트/알림 메일은 건드리지 않음
 * (이미 등록된 시간 기반 트리거를 방해하지 않고 조용히 한 번 보고 싶을 때)
 */
function nc26HealthCheckDryRun() {
  var issues = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    var fn = t.getHandlerFunction();
    var exists = false;
    try { exists = (typeof this[fn] === 'function'); } catch (e) {}
    if (!exists) issues.push('TRIGGER_FN_MISSING: ' + fn);
  });

  Object.keys(NC26_EXPECTED_SHEETS).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { issues.push('SHEET_MISSING: ' + name); return; }
    if (sh.getLastColumn() < NC26_EXPECTED_SHEETS[name]) {
      issues.push('SHEET_COL_SHORT: ' + name);
    }
  });

  try {
    var remaining = MailApp.getRemainingDailyQuota();
    if (remaining < NC26_MAIL_QUOTA_THRESHOLD) {
      issues.push('LOW_MAIL_QUOTA: ' + remaining);
    }
  } catch (e) {}

  Logger.log('DryRun: ' + (issues.length === 0 ? 'OK' : 'FAIL — ' + issues.join(' | ')));
  return issues;
}
