var fetch_form: Map<string, Array<string>> = get_Fetch_form();
var ref = get_Ref();
var order_form = get_Order_form();

//[출고요청/업로드]폴더 >> 주문현황/에러확인으로 데이터 옮김
async function fetch_Order() {
  //const fetch_form: Map<string, Array<string>> = get_Form();
  //@ts-ignore
  const files: any = DriveApp.getFolderById(ref.get('업로드')).getFilesByType(MimeType.GOOGLE_SHEETS);
  const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
  const trash_files = new Array();
  
  while (files.hasNext()){
    let file = files.next();
    let separator = file.getName().split('_');
    if (fetch_form.get(separator[2])) {
      let input_data: Array<Array<string>>;
      if (separator[2] == '스마트스토어') {
        if (SpreadsheetApp.openById(file.getId()).getSheets()[0].getRange(1, 1).getValue().length > 10) {
          SpreadsheetApp.openById(file.getId()).getSheets()[0].deleteRow(1);
        }
      }
      await set_Format(file.getId(), '@');
      input_data = await fetch_Data(file.getId(), fetch_form.get(separator[2]), separator);
      target_sheet.insertRowsAfter(1, input_data.length);
      target_sheet.getRange(2, 1, input_data.length, input_data[0].length).setNumberFormat('@').setValues(input_data);
    }
    trash_files.push(file);
  }

  return trash_files;
}

async function set_Format(id: string, format: string) {
  SpreadsheetApp.openById(id).getDataRange().setNumberFormat(format);
}

//각 시트에서 주문현황/에러확인으로 데이터 옮기기
async function fetch_Data(file_id: string, fetch_form: Array<string>, identifier: Array<string>) {
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
  const temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('요청양식').getDataRange().getValues();
  
  let form_index = new Array();
  temp[0].forEach((t, i) => {
    form_index[i] = t;
  });

  orig_data.forEach((o: Array<string>, index: number) => {
    let i = convert_Column(fetch_form[0]);
    if (!o[i as number - 1]) {
      return;
    }
    input_data[index] = new Array();
    input_data[index].push(Utilities.formatDate(new Date(), 'GMT+9', 'yy/MM/dd HH:mm'));
    input_data[index].push(identifier[1]);
    input_data[index].push(identifier[2]);

    fetch_form.forEach((f: string, id: number) => {
      let column = form_index[id + 1];
      let j = order_form.get(column);
      if (f === 'none') {
        input_data[index][j] = '';
        return;
      }
      
      //배송메세지 \n 삭제
      if (column === '배송메세지') {
        let i = convert_Column(f);
        input_data[index][j]= o[i as number - 1].split('\n').join('');
        return;
      }

      //우편번호 양식 변경
      if (column === '우편번호') {
        let i = convert_Column(f);
        /*
        if (o[i as number - 1].indexOf('-') != -1) {
          input_data[index][j] = o[i as number - 1];
        } else {
          input_data[index][j] = Utilities.formatString('%05d', o[i as number - 1]);
        }
        */
        input_data[index][j] = Utilities.formatString('%05d', o[i as number - 1].split('-').join(''));
        return;
      }
      
      //전화번호 양식변경
      if (column === '주문자연락처' || column === '수령인연락처') {
        let i = convert_Column(f);
        if (o[i as number - 1].indexOf('-') == -1) {
          input_data[index][j] = o[i as number - 1];
        } else {
          input_data[index][j] = o[i as number - 1].split('-').join('');
        }
        return;
      }
      
      let comma = f.split(',');
      let temp: any;
      for (let c of comma) {
        let i = convert_Column(c);
        if (o[i as number - 1]) {
          temp = o[i as number - 1];
          break;
        }
      }
      input_data[index][j] = temp as string;
    });
  });

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