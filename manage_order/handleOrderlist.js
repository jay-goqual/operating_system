function check_ConvertedOrder() {
  var ref = get_Ref();
  
  var statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('현황판');
  var targetList = statusSheet.getRange(9, 6, statusSheet.getLastRow() - 8, 2);
  //var errorStat = statusSheet.getRange(9, 7, statusSheet.getLastRow() - 8, 1);
  var channelName = statusSheet.getRange(9, 5, statusSheet.getLastRow() - 8, 1).getValues();
  var orderListmap = new Array();
  orderListmap[0] = new Map();
  orderListmap[1] = new Map();
  for (i = 0; i < statusSheet.getLastRow() - 8; i++) {
    orderListmap[0][channelName[i][0]] = 0;
    orderListmap[1][channelName[i][0]] = 0;
  }
  
  
  var query = 'select B, count(A) where AA = TRUE AND W = "굿스코아" group by B';
  var query_e = 'select B, count(A) where AA = FALSE AND W = "굿스코아" group by B';
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['주문현황'] + '/gviz/tq?tqx=out:csv&sheet=주문접수' + '&tq=' + encodeURIComponent(query);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  var orderList = Utilities.parseCsv(csv);
  orderList.splice(0, 1);
  
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['주문현황'] + '/gviz/tq?tqx=out:csv&sheet=주문접수' + '&tq=' + encodeURIComponent(query_e);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  var errorList = Utilities.parseCsv(csv);
  errorList.splice(0, 1);
  
  for (i = 0; i < orderList.length; i++) {
    orderListmap[0][orderList[i][0]] = orderList[i][1];
  }
  for (i = 0; i < errorList.length; i++) {
    orderListmap[1][errorList[i][0]] = errorList[i][1];
  }
  
  var orderListdata = new Array();
  for (i = 0; i < statusSheet.getLastRow() - 8; i++) {
    orderListdata[i] = new Array();
    orderListdata[i][0] = orderListmap[0][channelName[i][0]];
    orderListdata[i][1] = orderListmap[1][channelName[i][0]];
  }
  orderListdata[0][0] = '=F10 + F15';
  orderListdata[0][1] = '=G10 + G15';
  orderListdata[1][0] = '=SUM(F11:F14)';
  orderListdata[1][1] = '=SUM(G11:G14)';
  orderListdata[6][0] = '=SUM(F16:F27)';
  orderListdata[6][1] = '=SUM(G16:G27)';
  targetList.setValues(orderListdata);
}