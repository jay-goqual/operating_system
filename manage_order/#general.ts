//스프레드시트 열릴시
function onOpen() {
  //ui 생성
  SpreadsheetApp.getUi()
  .createMenu('출고관리')
  //.addItem('주문 양식 변환', functionName)
  .addSeparator()
  .addToUi();
}

/*
//업로드, 또는 생성된 파일이 있는지 체크 << 진짜 필요할까? 보류
function check_Update() {
  //출고요청 업로드 확인
  const order_upload = DriveApp.getFolderById(ref['출고요청/업로드']).getFiles();
  
  //송장 업로드 확인
  const invoice_upload = DriveApp.getFolderById(ref['송장업로드']).getFiles();
  
  //송장 다운로드 확인
  const invoice_download = DriveApp.getFolderById(ref['송장전달']).getFiles();
  
  //각 폴더 내 파일이 있는지 확인
  let stat = new Array();
  stat[0].push(order_upload.hasNext());
  stat[0].push(invoice_upload.hasNext());
  stat[0].push(invoice_download.hasNext());  
  
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('현황판');
}
*/