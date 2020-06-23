/*function debug() {
  handle_origOrder(true);
}*/

//소량발주, CS발주 관리
function handle_ownOrder(convert) {
}

function handle_origOrder(convert) {
  var statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('현황판');
  var ref = get_Ref();
  
  var orderStat = statusSheet.getRange(10, 3, statusSheet.getLastRow() - 9, 1);
  var orderName = statusSheet.getRange(10, 2, statusSheet.getLastRow() - 9, 1).getValues();
  var orderStatmap = new Map();
  
  var orderS = orderStat.getValues();
  
  for (i = 0; i < statusSheet.getLastRow() - 9; i++) {
    orderStatmap[orderName[i][0]] = 0;
  }
  
  
  var folder = DriveApp.getFolderById(ref['출고요청/업로드']);
  var sheetFiles = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  var deleteFiles = new Array();
  
  while (sheetFiles.hasNext()) {
    var sheetFile = sheetFiles.next();
    var identifier = sheetFile.getName().split('_');
    var t = 0;
    
    if (convert) {
      //시트파일 양식 변환 후 출고요청으로 복사
      convert_origOrder(sheetFile, identifier[0], identifier[1], identifier[2]);
      deleteFiles.push(sheetFile.getId());
    } 
    
    else {
      if (orderStatmap[identifier[1]] == null) {
        SpreadsheetApp.getUi().alert('잘못된 파일명: ' + sheetFile.getName() + '\n해당 파일은 삭제됩니다.엑셀 파일 업로드부터 다시 실행해주세요');
        //deleteFiles.push(sheetFile.getId());
        continue;
      } else {
        t = (SpreadsheetApp.openById(sheetFile.getId()).getSheets()[0].getLastRow() - 1);
      }
      orderStatmap[identifier[1]] += t;
      
    }
    
  }
  
  
  var orderStatdata = new Array();
  for (i = 0; i < statusSheet.getLastRow() - 9; i++) {
    orderStatdata[i] = new Array();
    orderStatdata[i][0] = orderStatmap[orderName[i][0]];
  }
  orderStatdata[0][0] = '=C11 + C16';
  orderStatdata[1][0] = '=SUM(C12:C15)';
  orderStatdata[6][0] = '=SUM(C17:C30)';
  orderStat.setValues(orderStatdata);
  
  /*
  var targetSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('에러확인');
  
  targetSheet.getDataRange().getFilter().remove();
  targetSheet.getDataRange().createFilter();
  */
    
  return deleteFiles;
}

function convert_origOrder(sheetFile, Rdate, contractor, channel) {
  var sheets = SpreadsheetApp.openById(sheetFile.getId()).getSheets();
  
  //양식정보 가져오기
  var ref = get_Ref();
  
  var q = 'select * where A = "' + channel + '"';
  
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['출고요청/요청양식'] + '/gviz/tq?tqx=out:csv&sheet=양식정보' + '&tq=' + encodeURIComponent(q);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  var formData = Utilities.parseCsv(csv);
  
  //채널이 스마트스토어면 모든 데이터 받아서 미리 처리할것 [A, B] 이면 A열, B열 받아서 A열이 비어있으면 B열 가져와서 A열에 채울것 + [A, B]를 [A]로 변환
  if (channel == '스마트스토어' || channel == '헤이홈양식') {
    formData[0] = formData[0].map(function(column) {
      var s = column.split(',');
      if (s.length > 1) {
        return s[0];
      }
      return column;
    });
  }

  //기반으로 쿼리만들기
  //var query = new Array();
  
  //var query = new String();
  //query = query.concat('select');
  
  var query = 'select';
  
  formData[0].forEach(function(column) {
    var text = new String();
    if (column == channel || column == 'none') {
      return;
    }
    query = query.concat(' ' + column + ',');
  });
  
  query = query.substring(0, query.length - 1);
  
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + sheetFile.getId() + '/gviz/tq?tqx=out:csv&tq=' + encodeURIComponent(query);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  
  var orderData = String();
  var orderData = Utilities.parseCsv(csv);
  
  var targetSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('에러확인');
  var startAt = targetSheet.getLastRow() + 1;
  
  /*
  var spl = String(orderData[0]).split('upper()');
  if (spl.length > 1) {
    orderData.splice(0, 1);
  }
  */
  if (String(orderData[0]).indexOf('upper()') != -1) {
    orderData.splice(0, 1);
  }
  orderData.splice(0, 1);
  
  //get Product
  orderData = get_Product(orderData);
  
  targetSheet.getRange(startAt, 4, orderData.length, orderData[0].length).setNumberFormat('@STRING@');
  targetSheet.getRange(startAt, 4, orderData.length, orderData[0].length).setValues(orderData);
  
  var today = Utilities.formatDate(new Date(), "GMT+9", "yy/MM/dd");
  
  targetSheet.getRange(startAt, 1, orderData.length, 1).setValue(today);
  targetSheet.getRange(startAt, 2, orderData.length, 1).setValue(contractor);
  targetSheet.getRange(startAt, 3, orderData.length, 1).setValue(channel);
  
  if (channel == '헤이홈양식') {
    var q = 'select N';
    var qvizURL = 'https://docs.google.com/spreadsheets/d/' + sheetFile.getId() + '/gviz/tq?tqx=out:csv&tq=' + encodeURIComponent(q);
    var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
    var csv = UrlFetchApp.fetch(qvizURL, options);
    var channel_data = Utilities.parseCsv(csv);
    channel_data.splice(0, 1);
    
    for (var i in channel_data) {
      if (channel_data[i][0] == '') {
        channel_data[i][0] = channel;
      }
    }
    
    targetSheet.getRange(startAt, 3, channel_data.length, 1).setValues(channel_data);
  }
  
  //targetSheet.getRange(startAt, targetSheet.getLastColumn(), orderData.length, 1).setFormulaR1C1('=ROW(R[0]C[0]) - 2');
  
  return;
}

function get_Product(orderData) {
  var ref = get_Ref();
  var productInfo = get_product_info();
  
  orderData.map(function(data){
    if(productInfo[data[2]]) {
      data[18] = productInfo[data[2]][0];
      data[19] = productInfo[data[2]][1];
    } else {
      data[18] = '';
      data[19] = '';
    }
    return data;
  });
  
  return orderData;
}

function get_product_info() {
  var ref = get_Ref();
  
  var q = 'select A, B, C';
  
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['상품DB'] + '/gviz/tq?tqx=out:csv&tq=' + encodeURIComponent(q);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  var origData = Utilities.parseCsv(csv);
  
  var productInfo = new Map();
  
  origData.forEach(function(data){
    productInfo[data[0]] = new Array();
    productInfo[data[0]][0] = data[1];
    productInfo[data[0]][1] = data[2];
  });
  
  return productInfo;
}