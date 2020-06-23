function debug2() {
  handle_excelFiles(true);
}

function handle_excelFiles(change) {
  var statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('현황판');
  var ref = get_Ref();
  
  var folder = DriveApp.getFolderById(ref['출고요청/업로드']);
  var excelFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL_LEGACY);
  
  var count = 0;
  
  while (excelFiles.hasNext()) {
    if(change) {
      change_Excelfile(excelFiles.next(), folder);
    } else {
      count++;
      excelFiles.next();
    }
  }
  
  var excelFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while (excelFiles.hasNext()) {
    if(change) {
      change_Excelfile(excelFiles.next(), folder);
    } else {
      count++;
      excelFiles.next();
    }
  }
  
  statusSheet.getRange(5, 3, 1, 1).setValue(count);
  
  return;
}

function change_Excelfile(excelFile, folder) {
  var blob = excelFile.getBlob();
  var resource = {
    title: excelFile.getName(),
    parents : [{id: folder.getId()}],
    mimeType: MimeType.GOOGLE_SHEETS,
  };
  change_Format(Drive.Files.insert(resource, blob, {convert: true}).id);
  Drive.Files.remove(excelFile.getId());
  
  return;
}

function change_Format(sheetID) {
  var ss = SpreadsheetApp.openById(sheetID);
  var spl = ss.getName().split('_');
  var channel = spl[2];
  var sheets = ss.getSheets();
  
  var ref = get_Ref();
  
  var q = 'select * where A = "' + channel + '"';
  
  var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ref['출고요청/요청양식'] + '/gviz/tq?tqx=out:csv&sheet=양식정보' + '&tq=' + encodeURIComponent(q);
  var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  var csv = UrlFetchApp.fetch(qvizURL, options);
  var formData = Utilities.parseCsv(csv);
  
  //채널이 공식몰이면 A열 B열 " " 지워줄 것 효율적으로 하는 방법 없을지 고민해봅시다
  //A,B,L,AS,AU 스트링 변환
  if (channel == '아임웹') {
    var fix = sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).getValues();
    
    /*
    if (fix[1][0].indexOf(' ') != -1) {
      for (i = 1; i < sheets[0].getLastRow(); i++) {
        var spl = fix[i][0].split(' ');
        if (spl[1]) {
          fix[i][0] = spl[1];
          //fix[i][0] = String(spl[1]);
          //fix[i][0] = String(fix[i][0]);
        }
        var spl = fix[i][1].split(' ');
        if (spl[1]) {
          fix[i][1] = spl[1];
          //fix[i][1] = String(spl[1]);
          //fix[i][1] = String(fix[i][1]);
        }
      }
      sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).setValues(fix).setNumberFormat('@STRING@');
    }*/
    
    fix = fix.map(function(f) {
      f[0] = f[0].split(' ').join('');
      f[1] = f[1].split(' ').join('');
      return f;
    });
    
    sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).setValues(fix).setNumberFormat('@STRING@');
    
    //배송메세지 \n 지우기
    
    var temp = sheets[0].getRange(1, 41, sheets[0].getLastRow(), 1).getValues();
 
    /*
    temp = temp.map(function(t) {
      var spl = t[0].split('\n');
      if (spl.length > 1) {
        var t2 = new Array();
        t2[0] = new String();
        for (i = 0; i < spl.length; i++) {
          t2[0] = t2[0].concat(spl[i]);
        }
        return t2;
      }
      return t;
    });
    */
    
    temp = temp.map(function(t) {
      t[0] = t[0].split('\n').join(' ');
      return t;
    });
    
    sheets[0].getRange(1, 41, sheets[0].getLastRow(), 1).setValues(temp);
    
    
    sheets[0].getRange(1, 12, sheets[0].getLastRow(), 1).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 45, sheets[0].getLastRow(), 1).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 47, sheets[0].getLastRow(), 1).setNumberFormat('@STRING@');
  }
  
  //채널이 스마트스토어면 A, B, AN, AQ, AR 스트링으로 변환할것
  //U~Y, AH~AJ, AX~AY, BA
  if (channel == '스마트스토어') {
    if (sheets[0].getRange(1, 1, 1, 1).getValue() != '상품주문번호') {
        sheets[0].deleteRow(1);
    }
    
    sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 41, sheets[0].getLastRow(), 1).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 44, sheets[0].getLastRow(), 2).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 22, sheets[0].getLastRow(), 5).setNumberFormat('0');
    sheets[0].getRange(1, 35, sheets[0].getLastRow(), 3).setNumberFormat('0');
    sheets[0].getRange(1, 51, sheets[0].getLastRow(), 2).setNumberFormat('0');
    sheets[0].getRange(1, 54, sheets[0].getLastRow(), 1).setNumberFormat('0');
    
    formData[0] = formData[0].map(function(column) {
      var s = column.split(',');
      if (s.length > 1) {
        var q = 'select ' + s[0] + ', ' + s[1];
        var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/gviz/tq?tqx=out:csv&tq=' + encodeURIComponent(q);
        var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
        var csv = UrlFetchApp.fetch(qvizURL, options);
        var temp = Utilities.parseCsv(csv);
        
        for (i = 1; i < temp.length; i++) {
          if (!temp[i][0]) {
            sheets[0].getRange(s[0] + String(i + 1)).setValue(sheets[0].getRange(s[1] + String(i + 1)).getValue());
          }
        }
        return s[0];
      }
      return column;
    });
    
    var temp = sheets[0].getRange(1, 45, sheets[0].getLastRow(), 1).getValues();
    
    /*temp = temp.map(function(t) {
      var spl = t[0].split('\n');
      if (spl.length > 1) {
        var t2 = new Array();
        t2[0] = new String();
        for (i = 0; i < spl.length; i++) {
          t2[0] = t2[0].concat(spl[i]) + ' ';
        }
        return t2;
      }
      return t;
    });*/
    temp = temp.map(function(t) {
      t[0] = t[0].split('\n').join(' ');
      return t;
    });
    
    
    sheets[0].getRange(1, 45, sheets[0].getLastRow(), 1).setValues(temp);
  }
  
  //채널이 헤이홈양식일 경우 상품주문번호 처리
  if (channel == '헤이홈양식') {
    sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 4, sheets[0].getLastRow(), 1).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 6, sheets[0].getLastRow(), 2).setNumberFormat('@STRING@');
    sheets[0].getRange(1, 12, sheets[0].getLastRow(), 1).setNumberFormat('0');
    
    //if (fix[1][0].indexOf(' ') != -1) {
      
    //}
    
    formData[0] = formData[0].forEach(function(column) {
      var s = column.split(',');
      if (s.length > 1) {
        var q = 'select ' + s[0] + ', ' + s[1];
        var qvizURL = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/gviz/tq?tqx=out:csv&tq=' + encodeURIComponent(q);
        var options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
        var csv = UrlFetchApp.fetch(qvizURL, options);
        var temp = Utilities.parseCsv(csv);
        
        var count = 1;
        var str = '0';
        for (i = 1; i < temp.length; i++) {
          if (!temp[i][0] && !temp[i][1]) {
            continue;
          }
          if (!temp[i][0] || temp[i][0] == '-') {
            if (i > 1 && temp[i][1] == temp[i - 1][1]) {
              count++;
              if (count > 9) {
                str = '1';
                count = 0;
              }
            } else {
              count = 1;
              str = '0';
            }
            /*var spl = temp[i][1].split(' ');
            if (spl.length > 1) {
              temp[i][1] = spl[0];
            }*/
            
            // var inp = String(temp[i][1]) + '-' + str + String(count);
            sheets[0].getRange(s[0] + String(i + 1)).setValue(String(temp[i][1]) + '-' + str + String(count));
          }
        }
      }
    });
    
    var fix = sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).getValues();
    
    /*for (i = 1; i < sheets[0].getLastRow(); i++) {
      var spl = fix[i][0].split(' ');
      if (spl[0]) {
        var c = '';
        for (j in spl) {
          c = c.concat(spl[j]);
        }
        fix[i][0] = c;
        //fix[i][0] = String(spl[1]);
        //fix[i][0] = String(fix[i][0]);
      }
      var spl = fix[i][1].split(' ');
      if (spl[0]) {
        var c = '';
        for (j in spl) {
          c = c.concat(spl[j]);
        }
        fix[i][1] = c;
        //fix[i][1] = String(spl[1]);
        //fix[i][1] = String(fix[i][1]);
      }
    }*/
    
    fix = fix.map(function(f) {
      f[0] = f[0].split(' ').join('');
      f[1] = f[1].split(' ').join('');
      return f;
    });
    
    sheets[0].getRange(1, 1, sheets[0].getLastRow(), 2).setValues(fix).setNumberFormat('@STRING@');
    
    /*
    var temp = sheets[0].getRange(1, 9, sheets[0].getLastRow(), 1).getValues();
    
    temp = temp.map(function(t) {
      var spl = t[0].split('\n');
      if (spl.length > 1) {
        var t2 = new Array();
        t2[0] = new String();
        for (i = 0; i < spl.length; i++) {
          t2[0] = t2[0].concat(spl[i]) + ' ';
        }
        return t2;
      }
      return t;
    });
        
    sheets[0].getRange(1, 9, sheets[0].getLastRow(), 1).setValues(temp);
    
    var temp = sheets[0].getRange(1, 8, sheets[0].getLastRow(), 1).getValues();
    
    temp = temp.map(function(t) {
      var spl = t[0].split('\n');
      if (spl.length > 1) {
        var t2 = new Array();
        t2[0] = new String();
        for (i = 0; i < spl.length; i++) {
          t2[0] = t2[0].concat(spl[i]) + ' ';
        }
        return t2;
      }
      return t;
    });
    
    sheets[0].getRange(1, 8, sheets[0].getLastRow(), 1).setValues(temp);
    */
    
    var temp = sheets[0].getRange(1, 8, sheets[0].getLastRow(), 2).getValues();
    
    temp = temp.map(function(t) {
      t[0] = t[0].split('\n').join(' ');
      t[1] = t[1].split('\n').join(' ');
      return t;
    });
    
    sheets[0].getRange(1, 8, sheets[0].getLastRow(), 2).setValues(temp);
    
    sheets[0].getRange(1, 11, sheets[0].getLastRow(), 1).setNumberFormat('@STRING@');
  }
  
  if (channel == '카카오스토어') {
    var temp = sheets[0].getRange(1, 19, sheets[0].getLastRow(), 1).getValues();
    
    /*
    temp = temp.map(function(t) {
      var spl = t[0].split('\n');
      if (spl.length > 1) {
        var t2 = new Array();
        t2[0] = new String();
        for (i = 0; i < spl.length; i++) {
          t2[0] = t2[0].concat(spl[i]);
        }
        return t2;
      }
      return t;
    });
    */
    
    temp = temp.map(function(t) {
      t[0] = t[0].split('\n').join(' ');
      return t;
    });
    
    sheets[0].getRange(1, 19, sheets[0].getLastRow(), 1).setValues(temp);
  }
  
  if (channel == '원룸만들기') {
    var temp = sheets[0].getRange(1, 20, sheets[0].getLastRow(), 1).getValues();
    
    temp = temp.map(function(t) {
      t[0] = t[0].split('\n').join(' ');
      return t;
    });
    
    sheets[0].getRange(1, 20, sheets[0].getLastRow(), 1).setValues(temp);
    
    var fix = sheets[0].getRange(1, 16, sheets[0].getLastRow(), 1).getValues();
    
    fix = fix.map(function(f) {
      if (f[0].length < 5) {
        f[0] = Utilities.formatString("%05d", f[0]);
      }
      return f;
    });
  }
  
  return;
}