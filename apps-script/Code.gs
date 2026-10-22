/**
 * CCPC 예산 앱 (Google Apps Script)
 * 시카고 언약 장로교회 — 매년 재사용하는 예산 편성 앱
 *
 * 계산 규칙 (재정부 지침):
 *  - 고정 지출: 월 평균 × 12
 *  - 변동 지출: (실적 개월 합계 ÷ 개월 수) × 12  →  예: 9개월 실적 × 1.333
 *  - 연 1회/일회성 지출: 비례 환산하지 않고 실제 예정 금액만 반영
 *  - Refunds & Credits(환불/조정) 항목은 예산 산출에서 제외
 *  - 4분기 계절 요인은 항목별 '조정률'로 반영
 *
 * 탭 구성: 사용법 / 설정 / 계정목록 / 분류기준 / 거래내역 / 예산안 / 비교표 / 인사관리
 */

// ---------- 설정 읽기 ----------
function getConfig_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설정');
  var cfg = { baseYear: 2026, budgetYear: 2027, months: 9 };
  if (sh) {
    var vals = sh.getRange('A1:B10').getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] === '기준연도') cfg.baseYear = Number(vals[i][1]) || cfg.baseYear;
      if (vals[i][0] === '예산연도') cfg.budgetYear = Number(vals[i][1]) || cfg.budgetYear;
      if (vals[i][0] === '실적 개월 수') cfg.months = Number(vals[i][1]) || cfg.months;
    }
  }
  return cfg;
}

// ---------- 메뉴 ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 예산 앱')
    .addItem('거래내역 → 예산안 자동 생성', 'buildBudgetFromTransactions')
    .addItem('분류기준 기본값 채우기', 'seedClassificationRules')
    .addSeparator()
    .addItem('새 연도 워크북 만들기', 'createNewYearWorkbook')
    .addToUi();
}

// ---------- 분류 기본 규칙 (키워드) ----------
function classifyDefault_(category) {
  var s = String(category || '').toLowerCase();
  if (s.indexOf('refund') >= 0 || s.indexOf('credit') >= 0) return '제외';
  var fixed = ['payroll', 'salary', 'stipend', 'fica', 'medicare',
               'subscription', 'software', 'internet', 'telephone', 'phone',
               'loan interest'];
  for (var i = 0; i < fixed.length; i++) if (s.indexOf(fixed[i]) >= 0) return '고정';
  var once = ['insurance', 'membership', 'scholarship'];
  for (var j = 0; j < once.length; j++) if (s.indexOf(once[j]) >= 0) return '일회성';
  return '변동';
}

function loadRules_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('분류기준');
  var rules = {};
  if (sh && sh.getLastRow() > 1) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) rules[String(data[i][0]).trim()] = String(data[i][1] || '변동').trim();
    }
  }
  return rules;
}

function seedClassificationRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tx = ss.getSheetByName('거래내역');
  var sh = ss.getSheetByName('분류기준');
  if (!tx || !sh) { SpreadsheetApp.getUi().alert('거래내역/분류기준 탭이 필요합니다.'); return; }
  var seen = {};
  var data = tx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var cat = String(data[i][2] || '').trim();
    if (cat) seen[cat] = true;
  }
  var existing = loadRules_();
  var rows = [];
  Object.keys(seen).sort().forEach(function (cat) {
    rows.push([cat, existing[cat] || classifyDefault_(cat), '']);
  });
  sh.getRange(2, 1, Math.max(rows.length, 1), 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  SpreadsheetApp.getUi().alert('분류기준 ' + rows.length + '개 항목을 채웠습니다. 성격을 확인/수정해 주세요.');
}

// ---------- 예산안 자동 생성 ----------
// 거래내역을 (계정, 카테고리)로 집계해 예산안 탭을 채운다.
// 기존 행의 성격/조정률/일회성 예정액/비고는 같은 (계정, 항목) 키로 보존한다.
function buildBudgetFromTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig_();
  var tx = ss.getSheetByName('거래내역');
  var bud = ss.getSheetByName('예산안');
  var acctSheet = ss.getSheetByName('계정목록');
  if (!tx || !bud) { SpreadsheetApp.getUi().alert('거래내역/예산안 탭이 필요합니다.'); return; }

  // 계정 → (계정명, 담당팀)
  var acctInfo = {};
  if (acctSheet && acctSheet.getLastRow() > 1) {
    var ad = acctSheet.getRange(2, 1, acctSheet.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < ad.length; i++) {
      acctInfo[String(ad[i][1])] = { name: ad[i][2], team: ad[i][3] };
    }
  }

  // 기존 예산안 입력값 보존 (E성격, I조정률, K일회성금액, N비고)
  var saved = {};
  if (bud.getLastRow() > 1) {
    var bd = bud.getRange(2, 1, bud.getLastRow() - 1, 14).getValues();
    for (var j = 0; j < bd.length; j++) {
      var key0 = String(bd[j][0]) + '|' + String(bd[j][3]);
      saved[key0] = { kind: bd[j][4], rate: bd[j][8], once: bd[j][10], note: bd[j][13] };
    }
  }

  var rules = loadRules_();
  var data = tx.getDataRange().getValues();
  var agg = {};   // key -> {acct, name, team, cat, total, lines}
  for (var r = 1; r < data.length; r++) {
    var acct = String(data[r][0] || '').trim();
    var name = String(data[r][1] || '').trim();
    var cat = String(data[r][2] || '').trim();
    var amt = Number(data[r][8]) || 0;
    if (!cat) continue;
    var key = acct + '|' + cat;
    if (!agg[key]) {
      var info = acctInfo[acct] || {};
      agg[key] = { acct: acct, name: name || info.name || '', team: info.team || '', cat: cat, total: 0, lines: 0 };
    }
    agg[key].total += amt;
    agg[key].lines += 1;
  }

  var keys = Object.keys(agg).sort();
  var values = [], formulasG = [], formulasH = [], formulasJ = [], formulasL = [], formulasM = [];
  keys.forEach(function (key, idx) {
    var row = idx + 2;
    var a = agg[key];
    var sv = saved[key] || {};
    var kind = sv.kind || rules[a.cat] || classifyDefault_(a.cat);
    values.push([a.acct, a.name, a.team, a.cat, kind, Math.round(a.total * 100) / 100,
                 null, null, (sv.rate === '' || sv.rate == null) ? 0 : sv.rate, null,
                 sv.once || '', null, '', sv.note || '']);
    formulasG.push(['=IF(F' + row + '="","",F' + row + '/설정!$B$4)']);
    formulasH.push(['=IF(OR(E' + row + '="일회성",E' + row + '="제외"),0,G' + row + '*12)']);
    formulasJ.push(['=H' + row + '*(1+I' + row + ')']);
    formulasL.push(['=IF(E' + row + '="일회성",K' + row + ',J' + row + ')']);
    formulasM.push(['=IF(E' + row + '="고정","월 평균 × 12",IF(E' + row + '="변동","실적 ÷ ' +
                    '"&설정!$B$4&" × 12",IF(E' + row + '="일회성","연 1회성 — 실제 예정액","예산 제외 (Refunds & Credits)")))']);
  });

  if (bud.getLastRow() > 1) bud.getRange(2, 1, bud.getLastRow() - 1, 14).clearContent();
  if (values.length) {
    bud.getRange(2, 1, values.length, 14).setValues(values);
    bud.getRange(2, 7, values.length, 1).setFormulas(formulasG);
    bud.getRange(2, 8, values.length, 1).setFormulas(formulasH);
    bud.getRange(2, 10, values.length, 1).setFormulas(formulasJ);
    bud.getRange(2, 12, values.length, 1).setFormulas(formulasL);
    bud.getRange(2, 13, values.length, 1).setFormulas(formulasM);
  }
  SpreadsheetApp.getUi().alert('예산안 ' + values.length + '개 항목을 생성했습니다. 성격·조정률·일회성 금액을 검토해 주세요.');
}

// ---------- 새 연도 워크북 ----------
function createNewYearWorkbook() {
  var ui = SpreadsheetApp.getUi();
  var cfg = getConfig_();
  var res = ui.prompt('새 연도 워크북', '만들 예산 연도를 입력하세요 (예: ' + (cfg.budgetYear + 1) + ')', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var year = Number(res.getResponseText());
  if (!year) { ui.alert('연도를 숫자로 입력해 주세요.'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var copy = ss.copy('CCPC 예산 앱 (' + year + ')');
  var conf = copy.getSheetByName('설정');
  if (conf) {
    conf.getRange('B2').setValue(year - 1);  // 기준연도
    conf.getRange('B3').setValue(year);      // 예산연도
  }
  // 거래내역·인사관리 비우기 (헤더만 유지), 예산안 비우기
  ['거래내역', '예산안'].forEach(function (name) {
    var sh = copy.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
  var per = copy.getSheetByName('인사관리');
  if (per && per.getLastRow() > 1) {
    // 인사관리는 구조(구분/성명/항목/비고)만 남기고 금액 열 비우기
    var rows = per.getRange(2, 1, per.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < rows.length; i++) { rows[i][3] = ''; rows[i][4] = ''; rows[i][5] = ''; }
    per.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  ui.alert('새 워크북이 만들어졌습니다: ' + copy.getName() + '\n' + copy.getUrl());
}
