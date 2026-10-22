/**
 * SheetData.gs — 웹앱(Index.html)과 구글 시트(CCPC 예산 앱)를 연결하는 함수들.
 * 웹앱 화면에서 google.script.run 으로 호출됩니다.
 *
 * 필요한 것: 이 스크립트를 실행하는 계정이 아래 스프레드시트에 접근할 수 있어야 합니다.
 * 처음 실행할 때 구글 권한 승인 창이 뜨면 승인해 주세요.
 * (Apps Script 편집기에서 getSheetPayload를 한 번 실행해 보면 승인 창이 바로 뜹니다.)
 */

var BUDGET_SPREADSHEET_ID = '1jxv76QJ_Q8AZ_zPhUXIPPPJ0gYHQUgXRfyP334JaMxc';

function budgetSs_() {
  return SpreadsheetApp.openById(BUDGET_SPREADSHEET_ID);
}

function csvCell_(v) {
  if (v == null) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'America/Chicago', 'yyyy-MM-dd');
  }
  var s = String(v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/** 웹앱이 시트 데이터를 한 번에 받아 가는 함수 */
function getSheetPayload() {
  var ss = budgetSs_();

  // 1) 거래내역 → CSV 문자열 (헤더 제외)
  var tx = ss.getSheetByName('거래내역');
  var data = tx.getDataRange().getValues();
  var lines = [];
  for (var i = 1; i < data.length; i++) {
    lines.push(data[i].map(csvCell_).join(','));
  }

  // 2) 예산안 탭의 입력값(성격/조정률/일회성 예정액) — 웹앱에 그대로 반영
  var overrides = {};
  var bud = ss.getSheetByName('예산안');
  if (bud && bud.getLastRow() > 1) {
    var bd = bud.getRange(2, 1, bud.getLastRow() - 1, 11).getValues();
    for (var j = 0; j < bd.length; j++) {
      var key = String(bd[j][0]) + '|' + String(bd[j][3]);
      overrides[key] = {
        kind: bd[j][4] || '',
        rate: bd[j][8] || 0,
        once: (bd[j][10] === '' || bd[j][10] == null) ? '' : bd[j][10],
      };
    }
  }

  // 3) 설정
  var cfg = { baseYear: 2026, budgetYear: 2027, months: 9 };
  var conf = ss.getSheetByName('설정');
  if (conf) {
    var cv = conf.getRange('A2:B5').getValues();
    for (var k = 0; k < cv.length; k++) {
      if (cv[k][0] === '기준연도') cfg.baseYear = Number(cv[k][1]) || cfg.baseYear;
      if (cv[k][0] === '예산연도') cfg.budgetYear = Number(cv[k][1]) || cfg.budgetYear;
      if (cv[k][0] === '실적 개월 수') cfg.months = Number(cv[k][1]) || cfg.months;
    }
  }

  return { csv: lines.join('\n'), overrides: overrides, config: cfg, count: lines.length };
}

/** 웹앱에서 조정한 성격/조정률/일회성 예정액을 시트 예산안 탭에 저장 */
function saveBudgetSettings(rows) {
  var ss = budgetSs_();
  var bud = ss.getSheetByName('예산안');
  if (!bud || bud.getLastRow() < 2) return { ok: false, message: '예산안 탭을 찾을 수 없습니다.' };

  var n = bud.getLastRow() - 1;
  var keys = bud.getRange(2, 1, n, 4).getValues();
  var idx = {};
  for (var i = 0; i < n; i++) idx[String(keys[i][0]) + '|' + String(keys[i][3])] = i;

  var kinds = bud.getRange(2, 5, n, 1).getValues();
  var rates = bud.getRange(2, 9, n, 1).getValues();
  var onces = bud.getRange(2, 11, n, 1).getValues();

  var changed = 0;
  (rows || []).forEach(function (r) {
    var p = idx[r.acct + '|' + r.cat];
    if (p == null) return;
    kinds[p][0] = r.kind;
    rates[p][0] = Number(r.rate) || 0;
    onces[p][0] = (r.once === '' || r.once == null) ? '' : Number(r.once);
    changed++;
  });

  bud.getRange(2, 5, n, 1).setValues(kinds);
  bud.getRange(2, 9, n, 1).setValues(rates);
  bud.getRange(2, 11, n, 1).setValues(onces);
  return { ok: true, changed: changed };
}
