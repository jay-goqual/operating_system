//ref 전역선언
var ref = get_Ref();

//[전역]레퍼런스(id) 가져오기
function get_Ref() {
  const table: Array<Array<string>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().setNumberFormat('@').getValues();
  let ref: Map<string, string> = new Map();
  
  table.forEach((t) => {
    ref.set(t[0], t[2]);
  });
  
  return ref;
}

//Ref 찾기 되려나
function find_Ref(key: string) {
  return ref.get(key);
}

//[전역]이메일 가져오기
//셀러관리 이후 셀러관리에서 셀러정보 가져오는 것으로 변경
function get_Client() {
  const table: Array<Array<string>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이메일').getDataRange().setNumberFormat('@').getValues();
  let email_address: Map<string, string> = new Map();
  
  table.forEach((t) => {
    email_address.set(t[0], t[1]);
  });
  
  return email_address;
}

//[전역]변환양식 가져오기
//요청양식에서 변환양식 가져옴
function get_Form() {
  //const table = SpreadsheetApp.openById(ref['출고요청/요청양식']).getDataRange().getValues();
  const table: Array<Array<string>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('요청양식').getDataRange().setNumberFormat('@').getValues();
  table.splice(0, 1);
  
  let fetch_form: Map<string, Array<string>> = new Map();
  
  table.forEach((t) => {
    fetch_form.set(t[0], t.splice(1, t.length));
  });

  return fetch_form;
}