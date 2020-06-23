function convert_invoice_to_sheet() {
  var ref = get_Ref();
  
  var folder = DriveApp.getFolderById(ref['송장업로드']);
  var excelFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while(excelFiles.hasNext()) {
    var excelFile = excelFiles.next();
    var blob = excelFile.getBlob();
    var resource = {
      title: '송장정보',
      parents : [{id: folder.getId()}],
      mimeType: MimeType.GOOGLE_SHEETS,
    };
    change_Format(Drive.Files.insert(resource, blob, {convert: true}).id);
    Drive.Files.remove(excelFile.getId());
  }
  
  return;
}

function connect_invoice() {
  var ref = get_Ref();
  
  var folder = DriveApp.getFolderById(ref['송장업로드']);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  var targetSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('주문접수');
  
  var tmap = new Map();
  var deleteFiles = new Array();
  
  while(files.hasNext()) {
    var file = files.next();
    var invoiceData = SpreadsheetApp.openById(file.getId()).getDataRange().getValues();
    
    deleteFiles.push(file.getId());
    if (invoiceData[0][5] == '송장번호') {
      invoiceData.splice(0, 1);
      for (var i in invoiceData) {
        tmap[invoiceData[i][13]] = invoiceData[i][5];
      }
    } else if (invoiceData[0][10] == '송장번호') {
      invoiceData.splice(0, 1);
      for (var i in invoiceData) {
        tmap[invoiceData[i][12]] = invoiceData[i][10];
      }
    }
  }
  
  var tdata = targetSheet.getDataRange().getValues();
  tdata.splice(0, 1);
  var inv = new Array();
  
  for (var i in tdata) {
    inv[i] = [tdata[i][23]];
    if (!tdata[i][23]) {
      inv[i] = [tmap[tdata[i][4]]];
    }
  }
  
  targetSheet.getRange(2, 24, inv.length, 1).setValues(inv);
  
  //targetSheet.getRange(2, 1, inv.lenght, 27).sort([{column: 24}, {column: 25, ascending: false}]);
  //targetSheet.getDataRange().sort([{column: 24}, {column: 25, ascending: true}]);
    
  for (var i in deleteFiles) {
    Drive.Files.trash(deleteFiles[i]);
  }
  
}

function exportInvoice() {
  var ref = get_Ref();
  
  var orderSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('주문접수');
  var order = orderSheet.getDataRange().getValues();
  
  var today = Utilities.formatDate(new Date(), "GMT+9", "yy/MM/dd");
  
  var exportData = new Map();
  var channelName = new Array();
  
  order.map(function(o) {
    if (o[23] && !o[24]) {
      o[24] = today;
      if (!exportData[o[1]]) {
        exportData[o[1]] = new Array();
        channelName.push(o[1]);
        if (o[1] == '공식몰') {
          exportData[o[1]][0] = ['상품주문번호', '택배사', '송장번호', '주문 상태 변경'];
          exportData[o[1]][1] = ['-', '-', '-', '-'];
        } else if (o[1] == '엔분의일') {
          exportData[o[1]][0] = ['상품주문번호', '배송방법', '택배사', '송장번호'];
        } else {
          exportData[o[1]][0] = ['상품주문번호', '주문자', '수령인', '송장번호'];
        }
      }
      if (o[1] == '공식몰') {
        exportData[o[1]].push([o[4], '우체국택배', o[23], '배송중']);
      } else if (o[1] == '엔분의일') {
        exportData[o[1]].push([o[4], '택배발송', '우체국택배', o[23]]);
      } else {
        exportData[o[1]].push([o[4], o[7], o[9], o[23]]);
      }
    }
    return o;
  });
  
  orderSheet.getDataRange().setValues(order);
  
  var exportFolder = DriveApp.getFolderById(ref['송장전달']);
  today = Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd');
  
  for (i in channelName) {
    var t = String(today) + '_' + channelName[i] + '_출고완료';
    var res = {
      title: t,
      parents : [{id: exportFolder.getId()}],
      mimeType: MimeType.GOOGLE_SHEETS,
    }
    var file = Drive.Files.insert(res);
    var ss = SpreadsheetApp.openById(file.id);
    var sheet = ss.getSheets();
    sheet[0].setName('발송처리');
    sheet[0].getRange(1, 1, exportData[channelName[i]].length, 4).setNumberFormat('@STRING@');
    sheet[0].getRange(1, 1, exportData[channelName[i]].length, 4).setValues(exportData[channelName[i]]);
    
    //send_eMail(file);    
    /*//var url = 'https://docs.google.com/spreadsheets/d/' + file.id + '/export?' + 'format=xlsx' +  '&gid=' + sheetgId+ "&portrait=true" + "&exportFormat=pdf";
    var url = 'https://docs.google.com/spreadsheets/d/' + file.id + '/export?' + 'format=xlsx' + '&exportFormat=xlsx';
    var content = UrlFetchApp.fetch(
      url,
      {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}},
    ).getContent();
    MailApp.sendEmail({
      to: 'jaysin@goqual.com',
      subject: '[헤이홈] ' + today + ' 운송장 정보',
      htmlBody: '테스트',
      attachments: [{fileName: t+'.xlsx', content: content, mimeType: MimeType.MICROSOFT_EXCEL}]
    });*/
  }
}

function send_eMail() {
  var ref = get_Ref();
  var exportFolder = DriveApp.getFolderById(ref['송장전달']);
  var address = new Map();
  
  var addressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var addressData = addressSheet[2].getDataRange().getValues();
  addressData.forEach(function(addr) {
    address[addr[0]] = addr[1];
  });
  
  //var file = DriveApp.getFileById('1qY67dnZ6SFIAizFBQ198gOQphZjkHTgWqzVRaQcgijU');
  //var url = 'https://docs.google.com/spreadsheets/d/' + file.id + '/export?' + 'format=xlsx' +  '&gid=' + sheetgId+ "&portrait=true" + "&exportFormat=pdf";
  
  var today = Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd');
  
  exportFiles = exportFolder.getFiles()
  while (exportFiles.hasNext()) {
    var file = exportFiles.next();
    var spl = file.getName().split('_');
    
    if (address[spl[1]]) {
      var url = 'https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?' + 'format=xlsx' + '&exportFormat=xlsx';
      var content = UrlFetchApp.fetch(
        url,
        {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}},
      ).getContent();
      MailApp.sendEmail({
        to: address[spl[1]],
        cc: 'service@goqual.com',
        //to: 'jaysin@goqual.com',
        subject: '[헤이홈] ' + today + '일자 운송장 정보',
        htmlBody:         
        '<div dir="ltr"><div dir="ltr"><div dir="ltr">안녕하세요.<br>주식회사 고퀄의 커머스팀 매니저 신재유입니다.<br><br>금일 주문 건에 대한 운송장 정보 전달드립니다.<br><br>감사합니다 :)<br>신재유 드림.</div></div></div><div><br></div>-- <br><div dir="ltr" class="gmail_signature" data-smartmail="gmail_signature"><div dir="ltr"><div style="color:rgb(34,34,34)"><p class="MsoNormal"><b><span style="color:rgb(102,102,102)">신재유&nbsp;</span></b><span style="color:rgb(102,102,102)">매니저<b><span lang="EN-US">&nbsp;</span></b><span lang="EN-US">/ 커머스</span>팀</span><span lang="EN-US"><u></u><u></u></span></p></div><div style="color:rgb(34,34,34)"><p class="MsoNormal"><b><span lang="EN-US" style="font-size:10pt;color:rgb(102,102,102)">MOBILE</span></b><span lang="EN-US" style="font-size:10pt;color:rgb(102,102,102)">&nbsp;010-3725-2198</span><span lang="EN-US"><u></u><u></u></span></p></div><div style="color:rgb(34,34,34)"><p class="MsoNormal"><b><span lang="EN-US" style="font-size:7.5pt;color:rgb(102,102,102)">TEL</span></b><span lang="EN-US" style="font-size:7.5pt;color:rgb(102,102,102)">&nbsp;070-4738-3000&nbsp;&nbsp;<b>FAX</b>&nbsp;0303-3444-2230</span><span lang="EN-US"><u></u><u></u></span></p></div><div style="color:rgb(34,34,34)"><p class="MsoNormal"><b><span lang="EN-US" style="font-size:7.5pt;color:rgb(102,102,102)">EMAIL</span></b><span lang="EN-US" style="font-size:7.5pt;color:rgb(102,102,102)">&nbsp;<a href="mailto:jaysin@goqual.com" style="color:rgb(17,85,204)" target="_blank">jaysin@goqual.com</a></span><span lang="EN-US"><u></u><u></u></span></p></div><div style="color:rgb(34,34,34)"><p class="MsoNormal"><b><span lang="EN-US" style="font-size:7.5pt;color:rgb(102,102,102)">ADDRESS</span></b><span lang="EN-US" style="font-size:7.5pt;color:rgb(102,102,102)">&nbsp;</span><span style="font-size:7.5pt;color:rgb(102,102,102)">서울특별시 금천구 가산디지털<span lang="EN-US">2</span>로<span lang="EN-US">&nbsp;184&nbsp;</span>벽산 경인디지털밸리<span lang="EN-US">2 1413</span>호</span></p></div></div></div>',
        attachments: [{fileName: file.getName()+'.xlsx', content: content, mimeType: MimeType.MICROSOFT_EXCEL}],
                        });
    }
  }
}



















