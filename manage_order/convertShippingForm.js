function convert_to_eSCM_form() {
  //var productInfo = get_product_info();
  var ref = get_Ref();
  
  var q = 'select D, E, F, G, H, I, J, K, L, M, N, B, V where Z is null AND AA = TRUE AND W = "굿스코아"';
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['주문현황'] + '/gviz/tq?tqx=out:csv&sheet=주문접수&tq=' + encodeURIComponent(q);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  //var orderData = String();
  var orderData = Utilities.parseCsv(csv);
  
  orderData.splice(0, 1);
  
  /*
  orderData.map(function(data){
    data.push(productInfo[data[2]]);
    return data;
  });
  */
  
  var targetSheet = SpreadsheetApp.openById(ref['출고대기']).getSheetByName('출고대기');
  
  targetSheet.getRange(2, 1, orderData.length, 13).setNumberFormat('@STRING@');
  targetSheet.getRange(2, 1, orderData.length, 13).setValues(orderData);
  
  var origSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('주문접수');
  
  var t = true;
  origSheet.getRange(2, 26, origSheet.getLastRow() - 1, 1).setValue(t).insertCheckboxes();
  
  //SpreadsheetApp.getUi().alert('5초 후 확인을 눌러주세요');
  
  return;
}

function convert_to_xlsx() {
  var ref = get_Ref();
  var origsheet = SpreadsheetApp.openById(ref['출고대기']);
  
  //var formattedDate = Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd' 'HH:mm");
  //var name = "Backup Copy " + formattedDate;
  var destination = DriveApp.getFolderById('1VGpr_rN7Yg7tSQyQWSRKasb3Hj-4fmiR');

  var today = Utilities.formatDate(new Date(), "GMT+9", "yy/MM/dd");
  var time = Utilities.formatDate(new Date(), "GMT+9", "MMddHHmm");
  // Added
  var url = 'https://docs.google.com/spreadsheets/d/' + ref['출고대기'] + '/export?format=xlsx&access_token=' + ScriptApp.getOAuthToken();
  var blob = UrlFetchApp.fetch(url).getBlob().setName(time + '_출고지시.xlsx'); // Modified
  destination.createFile(blob);
  
  origsheet.getSheetByName('출고대기').deleteRows(2, origsheet.getSheetByName('출고대기').getLastRow() - 1);
  
  //SpreadsheetApp.getUi().alert('엑셀 추출 완료\n출고관리 폴더 내 [eSCM 등록.xlsx]파일을 eSCM에 등록해주세요');
  
  //check_eSCM();
  
  return;
}

/*
function check_eSCM() {
  var ref = get_Ref();
  
  var query = 'select L, count(A) group by L';
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['출고대기'] + '/gviz/tq?tqx=out:csv&tq=' + encodeURIComponent(query);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  var eSCMlist = Utilities.parseCsv(csv);
  eSCMlist.splice(0, 1);
  
  var statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('현황판');
  var targetList = statusSheet.getRange(9, 10, statusSheet.getLastRow() - 8, 1);
  var channelName = statusSheet.getRange(9, 9, statusSheet.getLastRow() - 8, 1).getValues();
  var eSCMmap = new Map();
  for (i = 0; i < statusSheet.getLastRow() - 8; i++) {
    eSCMmap[channelName[i][0]] = 0;
  }
  
  for (i = 0; i < eSCMlist.length; i++) {
    eSCMmap[eSCMlist[i][0]] = eSCMlist[i][1];
  }
  
  var eSCMlistInput = new Array();
  for (i = 0; i < statusSheet.getLastRow() - 8; i++) {
    eSCMlistInput[i] = [eSCMmap[channelName[i][0]]];
  }
  
  eSCMlistInput[0][0] = '=J10 + J15';
  eSCMlistInput[1][0] = '=SUM(J11:J14)';
  eSCMlistInput[6][0] = '=SUM(J16:J27)';
  
  targetList.setValues(eSCMlistInput);
}
*/