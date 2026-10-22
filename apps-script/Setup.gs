/**
 * Setup.gs — 빈 구글 시트에서 실행하면 CCPC 예산 앱의 전체 탭 구조를 만듭니다.
 * 사용법: Apps Script 편집기에서 setupWorkbook()을 한 번 실행하세요.
 * 이미 있는 탭/데이터는 건드리지 않습니다 (없는 탭만 생성).
 */

var TAB_DEFS = [
  { name: '사용법', headers: [] },
  { name: '설정', headers: ['항목', '값'] },
  { name: '계정목록', headers: ['그룹', '계정번호', '계정명', '담당팀', '부서', '거래건수', '항목수', '기준연도 실적'] },
  { name: '분류기준', headers: ['카테고리', '지출 성격', '메모'] },
  { name: '거래내역', headers: ['계정번호', '계정명', '카테고리', '날짜', '거래유형', 'Num', '거래처', '내역', '금액'] },
  { name: '예산안', headers: ['계정번호', '계정명', '담당팀', '세부항목', '지출성격', '실적 합계', '월 평균', '12개월 환산', '조정률', '조정 후 환산', '일회성 예정액', '최종 예산안', '산출 근거', '비고'] },
  { name: '비교표', headers: ['구분', '계정명(한글)', '계정명(영문)', '계정번호', '기준연도 예산', '실적', '실적/예산 %', '12개월 예상', '예산안 합계'] },
  { name: '인사관리', headers: ['구분', '성명', '항목', '기준연도 예산', '실적', '예산안', '비고'] },
];

function setupWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  TAB_DEFS.forEach(function (def) {
    var sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (def.headers.length && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sh.getRange(1, 1, 1, def.headers.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });

  var conf = ss.getSheetByName('설정');
  if (conf.getLastRow() === 1) {
    var year = new Date().getFullYear();
    conf.getRange('A2:B5').setValues([
      ['기준연도', year],
      ['예산연도', year + 1],
      ['실적 개월 수', 9],
      ['기본 조정률', 0],
    ]);
  }

  var help = ss.getSheetByName('사용법');
  if (help.getLastRow() === 0) {
    help.getRange('A1:A12').setValues([
      ['CCPC 예산 앱 — 사용법'],
      [''],
      ['1) 거래내역 탭에 기준연도 지출 내역(QuickBooks P&L Detail 형식)을 붙여넣습니다.'],
      ['2) 메뉴 [📊 예산 앱] → "분류기준 기본값 채우기"로 카테고리별 지출 성격(고정/변동/일회성/제외)을 정합니다.'],
      ['3) 메뉴 [📊 예산 앱] → "거래내역 → 예산안 자동 생성"을 실행하면 예산안 초안이 만들어집니다.'],
      ['4) 예산안 탭에서 조정률(4분기 계절 요인 등)과 일회성 예정액을 입력하면 최종 예산안이 완성됩니다.'],
      ['5) 비교표 탭에서 기준연도 예산·실적과 새 예산안을 한눈에 비교합니다.'],
      ['6) 다음 해에는 메뉴의 "새 연도 워크북 만들기"로 빈 워크북을 복사해 같은 방식으로 반복합니다.'],
      [''],
      ['계산 규칙: 고정 = 월 평균 × 12 · 변동 = 실적 ÷ 개월 수 × 12 · 일회성 = 실제 예정액 · Refunds & Credits = 제외'],
      ['주의: 인사관리 탭의 금액은 확정 전 내부 검토용입니다. 공유 범위를 확인하고 사용하세요.'],
    ]);
    help.getRange('A1').setFontWeight('bold').setFontSize(13);
    help.setColumnWidth(1, 900);
  }

  // 성격 드롭다운
  var rules = ss.getSheetByName('분류기준');
  var kindRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['고정', '변동', '일회성', '제외'], true).build();
  rules.getRange('B2:B500').setDataValidation(kindRule);
  var bud = ss.getSheetByName('예산안');
  bud.getRange('E2:E1000').setDataValidation(kindRule);

  SpreadsheetApp.getUi().alert('CCPC 예산 앱 구조가 준비됐습니다.');
}
