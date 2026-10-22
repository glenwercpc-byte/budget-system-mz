/**
 * WebApp.gs — 웹앱 배포용 진입점
 *
 * 이 파일이 있어야 배포 주소(/exec)가 오류 없이 열립니다.
 * Apps Script 프로젝트에 이 파일과 Index.html 파일을 함께 넣고
 * [배포 → 배포 관리 → 새 버전]으로 다시 배포하면 같은 주소로 앱이 열립니다.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CCPC 예산 앱')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
