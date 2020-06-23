
// 주문접수DB 내 에러 확인
// W 가 false 인 경우만 체크
// A~T 까지 빈칸 있을 시 에러
// A,B 스트링 아닐 시 에러
// M 5자 미만일 시 에러
// I, K ?
function check_Error(checkSheet, targetSheet) {
  var checkRange = checkSheet.getDataRange();
  //var check = String();
  var check = checkRange.getValues();
  var move = new Array();
  /*
  if (check[0][0] == '접수일자') {
    //check.splice(0, 1);
  }
  */
  var del = 0;
  var checkbox = new Array;
  for (i = 1; i < check.length; i++) {
    var c = true;
    for (j = 0; j < 13; j++) {
      if (!check[i][j]) {
        checkRange.getCell(i + 1, j + 1).setBackground("#f4cccc");
        c = false;
      }
    }
    for (j = 21; j < 23; j++) {
      if (!check[i][j]) {
        checkRange.getCell(i + 1, j + 1).setBackground("#f4cccc");
        c = false;
      }
    }
    if (!check[i][21]) {
      checkRange.getCell(i + 1, 22).setBackground("#f4cccc");
      c = false;
    }
    if (c == false) {
      continue;
    }
    if (typeof check[i][3] != 'string' || check[i][3].indexOf(' ') != -1) {
      checkRange.getCell(i + 1, 4).setBackground("#f4cccc");
      continue;
    }
    if (typeof check[i][4] != 'string' || check[i][4].indexOf(' ') != -1) {
      checkRange.getCell(i + 1, 5).setBackground("#f4cccc");
      continue;
    }
    if (typeof check[i][12] != 'string' || check[i][12].length < 5) {
      checkRange.getCell(i + 1, 13).setBackground("#f4cccc");
      continue;
    }
    if (check[i][6].indexOf(' ') != -1) {
      checkRange.getCell(i + 1, 7).setBackground("#f4cccc");
      continue;
    }
    check[i][26] = c;
    checkbox.push([c]);
    move.push(check[i]);
    //checkSheet.deleteRow(i + 2 - del);
    del++;
  }
  
  if (move[0]) {
    var startAt = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startAt, 2, move.length, move[0].length - 2).setNumberFormat('@STRING@').setBackground(null);
    targetSheet.getRange(startAt, 1, move.length, 1).setNumberFormat('yy/mm/dd').setBackground(null);
    //targetSheet.getRange(targetSheet.getLastRow() + 1, 1, move.length, move[0].length - 1).setBackground(null);
    
    //Logger.log(targetSheet.getLastRow() - move.length + 1);
    
    targetSheet.getRange(startAt, 1, move.length, move[0].length).setValues(move);
    //targetSheet.getRange(startAt, move[0].length, move.length, 1).setValues(checkbox).insertCheckboxes();
    targetSheet.getRange(startAt, move[0].length, move.length, 1).insertCheckboxes();
    //targetSheet.getRange(targetSheet.getLastRow() + 1, 1, move.length, move[0].length).setValues(move);
  }
  checkSheet.getRange(2, checkSheet.getLastColumn(), checkSheet.getLastRow() - 1, 1).insertCheckboxes();
  checkSheet.getDataRange().setValues(check);
  checkSheet.getDataRange().sort({column: 27});
  if (del != 0) {
    checkSheet.deleteRows(2, del);
  }
  
  if (checkSheet.getLastRow() > 1) {
    SpreadsheetApp.getUi().alert('에러가 있습니다.\n에러확인 시트를 확인해주세요.');
  }
}

function check_total_Error() {
  var ref = get_Ref();
  var checkSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('에러확인');
  if (checkSheet.getLastRow() == 1) {
    return;
  }
  //var checkRange = checkSheet.getDataRange();
  //var targetRange = checkSheet.getRange(2, checkSheet.getLastColumn() - 1, checkSheet.getLastRow(), 1);
  var targetSheet = SpreadsheetApp.openById(ref['주문현황']).getSheetByName('주문접수');
  
  check_Error(checkSheet, targetSheet);
}
