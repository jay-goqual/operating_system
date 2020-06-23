const form_data: Map<string, Array<string>> = get_Form();

//[출고요청/업로드]폴더 >> 주문현황/에러확인으로 데이터 옮김
async function fetch_Order() {
  //const form_data: Map<string, Array<string>> = get_Form();
  //@ts-ignore
  const files: any = DriveApp.getFolderById(find_Ref('업로드')).getFilesByType(MimeType.GOOGLE_SHEETS);
  
  const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
  
  while (files.hasNext()){
    let file = files.next();
    let separator = file.getName().split('_');
    if (form_data.get(separator[2])) {
      let input_data: Array<Array<string>>;
      if (separator[2] == '스마트스토어') {
        if (SpreadsheetApp.openById(file.getId()).getSheets()[0].getRange(1, 1).getValue().length > 10) {
          SpreadsheetApp.openById(file.getId()).getSheets()[0].deleteRow(1);
        }
      }
      await set_Format(file.getId(), '@');
      input_data = await fetch_Data(file.getId(), form_data.get(separator[2]));
      target_sheet.insertRowsAfter(1, input_data.length);
      target_sheet.getRange(2, 4, input_data.length, input_data[0].length).setNumberFormat('@').setValues(input_data);
    }
  }

  return;
}

async function set_Format(id: string, format: string) {
  SpreadsheetApp.openById(id).getDataRange().setNumberFormat(format);
}

//각 시트에서 주문현황/에러확인으로 데이터 옮기기
async function fetch_Data(file_id: string, form: Array<string>) {
  //데이터에 , 가 있을 경우 처리
  /*
  if (form.indexOf(',')) {
    form = form.map((f) => {
      if (f.indexOf(',') != -1) {
        let s: Array<string> = f.split(',');
        let temp = SpreadsheetApp.openById(file_id).getSheets()[0];
        let col: Array<any> = new Array();
        
        col.push(convert_Column(s[0]));
        col.push(convert_Column(s[1]));
        col.push(temp.getLastColumn() + 1);
        
        temp.insertColumnAfter(temp.getLastColumn());
        temp.getRange(1, col[2], temp.getLastRow(), 1).setFormulaR1C1(`=if(R[0]C[${(col[0] - col[2])}] = "", R[0]C[${(col[1] - col[2])}], R[0]C[${(col[0] - col[2])}])`);
       
        return convert_Column(temp.getLastColumn()) as string;
      }
      return f;
    });
  }

  const query: string = 'select ' + await get_Query_string(form);
  const qvizURL = `https://docs.google.com/spreadsheets/d/${file_id}/gviz/tq?&tqx=out:csv&headers=1&tq=${encodeURIComponent(query)}`;
  const options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
  const csv = UrlFetchApp.fetch(qvizURL, options);
  //@ts-ignore
  let input_data: Array<Array<string>> = Utilities.parseCsv(csv);
  if (input_data[0].indexOf('upper') != -1) {
    input_data.splice(0, 1);
  }
  input_data.splice(0, 1);
  */

  //데이터 가져오기
  const orig_data = SpreadsheetApp.openById(file_id).getDataRange().getValues();
  orig_data.splice(0, 1);

  //input_data 만들기
  let input_data: Array<Array<string>> = new Array();
  orig_data.forEach((o: Array<string>, index: number) => {
    input_data[index] = new Array();
    form.forEach((f: string) => {
      console.log(f);
      if (f = 'none') {
        input_data[index].push('');
        return;
      }
      let spl = f.split(',');
      let temp: string;
      for (let s of spl) {
        console.log(s);
        let ss = convert_Column(s);
        console.log(ss);
        if (o[ss as number - 1]) {
          temp = o[ss as number - 1];
          break;
        }
      }
      console.log(temp);
      input_data[index].push(temp);
    });
  });

  console.log(input_data);

  return input_data;
}

function convert_Column(col: string | number) {
  if (typeof col === 'string') {
    let num: number = 0;
    if (col.length > 1) {
      num += (col.charCodeAt(0) - 64) * 26 + (col.charCodeAt(1) - 64);
    } else {
      num += (col.charCodeAt(0) - 64);
    }
    return num;
  }
  if (typeof col === 'number') {
    let str: string;
    if (col > 26) {
      str = String.fromCharCode((col / 26) + 64) + String.fromCharCode((col % 26) + 64);
    } else {
      str = String.fromCharCode(col + 64);
    }
    return str;
  }
}

//쿼리 스트링 만들기
async function get_Query_string(form: Array<string>) {

  let query: string = form.filter(f => f != 'none').join(', ') as string;
  return query;
}

//빈 데이터 채우기