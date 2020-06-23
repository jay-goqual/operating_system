function initiative() {
  SpreadsheetApp.getUi()
  .createMenu('운영관리')
  .addItem('업데이트', 'update_Status')
  .addSubMenu(
    SpreadsheetApp.getUi().createMenu('주문 변환')
    .addItem('엑셀 변환', 'convert_Excel')
    .addItem('양식 변환', 'convert_Orderform'))
  .addItem('에러 확인', 'check_total_Error')
  .addSubMenu(
    SpreadsheetApp.getUi().createMenu('굿스코아 출고처리')
    .addItem('굿스코아 양식 추출', 'convert_eSCM')
    .addItem('엑셀파일 추출', 'convert_to_xlsx'))
  .addItem('송장 정보 입력', 'handle_Invoice')
  .addSubMenu(
    SpreadsheetApp.getUi().createMenu('송장 등록')
    .addItem('송장 추출', 'exportInvoice')
    .addItem('메일 전송', 'send_eMail'))
  .addToUi();
  update_Status();
  return;
}

function handle_eSCM() {
  convert_eSCM();
  convert_to_xlsx();
}

function convert_Excel() {
  handle_excelFiles(true);
  SpreadsheetApp.getUi().alert('변환 완료');
  handle_origOrder(false);
  return;
}

function convert_eSCM() {
  convert_to_eSCM_form();
  SpreadsheetApp.getUi().alert('변환 완료\n엑셀파일 추출을 실행해주세요');
  //check_eSCM();
}

function get_Ref() {
  var reft = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().getValues();
  var ref = new Map();
  
  for (i in reft) {
    ref[reft[i][0]] = reft[i][2];
  }
  
  return ref;
}

function update_Status() {
  handle_excelFiles(false);
  handle_origOrder(false);
  check_total_Error();
  //check_ConvertedOrder();
  //check_eSCM();
  SpreadsheetApp.getUi().alert('업데이트 완료');
  return;
}

function convert_Orderform() {
  var deleteFiles = handle_origOrder(true);
  var ref = get_Ref();
  
  deleteFiles.forEach(function(deleteFile) {
    var f = DriveApp.getFileById(deleteFile);
    f.getParents().next().removeFile(f);
    DriveApp.getFolderById(ref['출고요청/아카이브']).addFile(f);
  });
    
  SpreadsheetApp.getUi().alert('변환 완료\n주문현황을 확인해주세요');
  
  check_total_Error();
  //check_ConvertedOrder();
  
}

function handle_Invoice() {
  convert_invoice_to_sheet();
  connect_invoice();
}
