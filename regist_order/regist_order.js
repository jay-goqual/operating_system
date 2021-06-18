// 이전 직접발주 스프레드시트의 코드이며, 현재는 폐기되었습니다.

function handle_order() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var orderData = sheets[0].getRange(3, 2, 1, 11).getValues();
  var product = sheets[0].getRange(3, 14, 15, 1).getValues();
  var count = sheets[0].getRange(3, 16, 15, 1).getValues();
  var price = sheets[0].getRange(3, 15, 15, 1).getValues();
  
  for (i = 0; i < 15; i++) {
    if ((product[i][0] && count[i][0] && String(price[i][0])) || (!product[i][0] && !count[i][0] && !String(price[i][0]))) {
      continue;
    }
    SpreadsheetApp.getUi().alert('상품 정보를 입력해주세요');
    return;
  }
    
  for (i in orderData[0]) {
    if (i == 9 || i == 2 || i == 3 || i == 4) {
      continue;
    }
    if (!orderData[0][i]) {
      SpreadsheetApp.getUi().alert(sheets[0].getRange(2, parseInt(i) + 2, 1, 1).getValue() + '를 입력해주세요');
      return;
    }
  }
    
  sheets[1].getRange(4, 13, 1, 1).setValue('');
  sheets[1].getRange(5, 13, 1, 1).setValue('');
  sheets[1].getRange(6, 13, 1, 1).setValue('');
  
  sheets[1].getRange(8, 2, 16, 1).setValue('');
  sheets[1].getRange(8, 12, 16, 1).setValue('');
  sheets[1].getRange(8, 14, 16, 1).setValue('');
  
  sheets[1].getRange(17, 2, 1, 1).setValue('');
  sheets[1].getRange(17, 12, 1, 1).setValue('');
  sheets[1].getRange(17, 14, 1, 1).setValue('');
  
  
  sheets[1].getRange(4, 13, 1, 1).setValue(orderData[0][0]);
  sheets[1].getRange(5, 13, 1, 1).setValue(orderData[0][1]);
  sheets[1].getRange(6, 13, 1, 1).setValue(orderData[0][2]);
  
  sheets[1].getRange(8, 2, 15, 1).setValues(product);
  sheets[1].getRange(8, 12, 15, 1).setValues(count);
  sheets[1].getRange(8, 14, 15, 1).setValues(price);
  
  if (orderData[0][10] == '기본 (수량당 3000원)') {
    sheets[1].getRange(23, 2, 1, 1).setValue('배송비');
    sheets[1].getRange(23, 14, 1, 1).setValue(3000);
  }
  
  var checkData = sheets[3].getDataRange().getValues();
  checkData.join(sheets[2].getDataRange().getValues());
  var index;
  var c = false;
  var orderCount = 1;
  while (true) {
    index = String(Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')) + orderData[0][0].split('-')[2] + Utilities.formatString('%02d', orderCount);
    for (i in checkData) {
      if (checkData[i][0] == index) {
        c = true;
        break;
      }
    }
    if (c == true) {
      orderCount++;
      c = false;
    } else {
      break;
    }
  }
  
  var money_wait = new Array();
  var out_wait = new Array();
  money_wait[0] = new Array();
  sheets[2].insertRowAfter(1);
  money_wait[0].push(index);
  money_wait[0].push(orderData[0][3]);
  money_wait[0].push(orderData[0][4]);
  
  sheets[2].getRange(2, 1, 1, 3).setValues(money_wait).setNumberFormat('@STRING@');
  sheets[2].getRange(2, 5, 1, 1).setValue(false).insertCheckboxes();
  
  for (i in product) {
    if (product[i][0]) {
      out_wait[i] = new Array();
      if (sheets[1].getRange(26, 17).getValue() == 0) {
        out_wait[i].push(orderData[0][1]);
        out_wait[i].push('무상');
        sheets[2].getRange(2, 5).setValue(true).insertCheckboxes();
        sheets[2].getRange(2, 6).setValue('무상').setNumberFormat('@STRING@');
      } else {
        out_wait[i].push('소량발주');
        out_wait[i].push(orderData[0][1]);
      }
      out_wait[i].push(index);
      out_wait[i].push(index + '-' + Utilities.formatString('%02d', parseInt(i) + 1));
      out_wait[i].push(orderData[0][1]);
      out_wait[i].push(orderData[0][6]);
      out_wait[i].push(orderData[0][5]);
      out_wait[i].push(orderData[0][6]);
      out_wait[i].push(orderData[0][7]);
      out_wait[i].push(orderData[0][8]);
      out_wait[i].push(orderData[0][9]);
      out_wait[i].push(product[i][0]);
      out_wait[i].push(count[i][0]);
    }
  }
  sheets[3].insertRowsAfter(1, out_wait.length);
  sheets[3].getRange(2, 1, out_wait.length, 13).setValues(out_wait).setNumberFormat('@STRING@');
  if (orderData[0][10] == '직접수령') {
    sheets[3].getRange(2, 14, out_wait.length, 1).setValue(true).insertCheckboxes();
  } else {
    sheets[3].getRange(2, 14, out_wait.length, 1).setValue(false).insertCheckboxes();
  }
  sheets[3].getRange(2, 16, out_wait.length, 1).setValue(false).insertCheckboxes();
  
  SpreadsheetApp.getUi().alert('등록 완료\n거래명세표를 확인하세요.');
  
  
  var orderData = sheets[0].getRange(3, 2, 1, 11).setValue('');
  var product = sheets[0].getRange(3, 14, 15, 1).setValue('');
  var count = sheets[0].getRange(3, 16, 15, 1).setValue('');
  var price = sheets[0].getRange(3, 15, 15, 1).setValue('');
}

function print_bill() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  if (!sheets[1].getRange(23, 12, 1, 1).getValue() && sheets[1].getRange(23, 2, 1, 1).getValue()) {
    SpreadsheetApp.getUi().alert('배송비 수량을 입력해주세요');
    return;
  }
  
  var url = 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActiveSpreadsheet().getId() + '/export?&gid=' + sheets[1].getSheetId() + '&exportFormat=pdf&size=A4&portrait=true&scale=1&gridlines=false&horizontal_alignment=CENTER&top_margin=0.5&bottom_margin=2.0&left_margin=0.0&right_margin=0.0';
  var html = '<p><a href="' + url + '" target="blank">다운로드</a><p>';
  var htmlOutput = HtmlService.createHtmlOutput(html)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .setWidth(100)
  .setHeight(80);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '거래명세서');
  
  var pdf_name = Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd') + '_' + sheets[1].getRange(5, 13, 1, 1).getValue() + '_거래명세서_' + sheets[2].getRange(2, 1, 1, 1).getValue();
  var blob = UrlFetchApp.fetch(url + '&access_token=' + ScriptApp.getOAuthToken()).getBlob().setName(pdf_name);
  var dest = DriveApp.getFolderById('1uCPascZr4U0mAyPX5raqGGnMtGi7IQAA');
  dest.createFile(blob);
  
  sheets[2].getRange(2, 4, 1, 1).setValue(sheets[1].getRange(26, 17, 1, 1).getValue());
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('소량발주')
  .addItem('명세표 생성', 'handle_order')
  .addItem('명세표 출력', 'print_bill')
  .addToUi();
}