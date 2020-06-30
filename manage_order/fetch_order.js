var fetch_form = get_Fetch_form();
var ref = get_Ref();
var order_form = get_Order_form();
var client = get_Client();

async function fetch_All_order() {
    change_type();
}

//[출고요청/업로드]폴더 >> 주문현황/에러확인으로 데이터 옮김
async function fetch_Order(file) {
    //const fetch_form: Map<string, Array<string>> = get_Form();
    //const files: any = DriveApp.getFolderById(ref.get('업로드')).getFilesByType(MimeType.GOOGLE_SHEETS);

    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');

    /* while (files.hasNext()){
        let file = files.next(); */
    let separator = file.getName().split('_');

    let client_info = client.get(separator[1]);
    let client_form = client_info.get('업로드양식');
    if (fetch_form.get(client_form)) {
        //let input_data: Array<Array<string>>;
        if (client_form == '스마트스토어') {
            if (SpreadsheetApp.openById(file.getId()).getSheets()[0].getRange(1, 1).getValue().length > 10) {
                SpreadsheetApp.openById(file.getId()).getSheets()[0].deleteRow(1);
            }
        }
        if (client_form == '천삼백케이') {
            if (SpreadsheetApp.openById(file.getId()).getSheets()[0].getRange(1, 1).getValue() == '') {
                SpreadsheetApp.openById(file.getId()).getSheets()[0].deleteRow(1);
            }
        }
        await set_Format(file.getId(), '@');
        
        let input_data = await fetch_Data(file.getId(), fetch_form.get(client_form), client_info);
        target_sheet.insertRowsAfter(1, input_data.length);
        target_sheet.getRange(2, 1, input_data.length, input_data[0].length).setNumberFormat('@').setValues(input_data);

        //아카이브로 보내버리기
        DriveApp.getFolderById(ref.get('업로드/아카이브')).addFile(file);
        file.getParents().next().removeFile(file);
    }
    //}
}

async function set_Format(id, format) {
    SpreadsheetApp.openById(id).getDataRange().setNumberFormat(format);
}

//각 시트에서 주문현황/에러확인으로 데이터 옮기기
async function fetch_Data(file_id, fetch_form, client_info) {

    //데이터 가져오기
    const orig_data = SpreadsheetApp.openById(file_id).getDataRange().getValues();
    orig_data.splice(0, 1);

    //input_data 만들기
    let input_data = new Array();
    const temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('요청양식').getDataRange().getValues();

    let form_index = new Array();
    temp[0].forEach((t, i) => {
        form_index[i] = t;
    });

    orig_data.forEach((o, index) => {
        let i = convert_Column(fetch_form[0]);
        if (!o[i - 1] || o[i - 1] == ' ') {
            return;
        }
        input_data[index] = new Array();
        // input_data[index].push(Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm'));
        input_data[index][order_form.get('셀러명')] = client_info.get('셀러명');
        // input_data[index].push(client_info.get('셀러코드'));

        fetch_form.forEach((f, id) => {
            let column = form_index[id + 1];
            let j = order_form.get(column);
            if (f === 'none') {
                input_data[index][j] = '';
                return;
            }

            if (column !== '주소' && column !== '배송메세지' && column !== '옵션정보' && column !== '결제일') {
                let i = convert_Column(f);
                if (o[i - 1]) {
                    o[i - 1] = o[i - 1].split(' ').join('');
                }
            }

            //배송메세지 \n 삭제
            if (column === '배송메세지') {
                let i = convert_Column(f);
                input_data[index][j] = o[i - 1].split('\n').join('');
                return;
            }

            //우편번호 양식 변경
            if (column === '우편번호') {
                let i = convert_Column(f);
                /*
                if (o[i - 1].indexOf('-') != -1) {
                input_data[index][j] = o[i - 1];
                } else {
                input_data[index][j] = Utilities.formatString('%05d', o[i - 1]);
                }
                */
                input_data[index][j] = Utilities.formatString('%05d', o[i - 1].split('-').join(''));
                return;
            }

            //전화번호 양식변경
            if (column === '주문자연락처' || column === '수령인연락처') {
                let i = convert_Column(f);
                if (o[i - 1].indexOf('-') == -1) {
                    input_data[index][j] = o[i - 1];
                } else {
                    input_data[index][j] = o[i - 1].split('-').join('');
                }
                return;
            }

            let comma = f.split(',');
            for (let c of comma) {
                let plus = c.split('+');
                let temp = '';
                plus.forEach((p, k) => {
                    let i = convert_Column(p);
                    if (k == 0) {
                        temp += o[i - 1];
                    } else {
                        temp += '-' + o[i - 1];
                    }
                });
                if (temp != '' && temp != '-') {
                    input_data[index][j] = temp;
                    break;
                }
            }
        });
    });

    return input_data;
}

function convert_Column(col) {
    if (typeof col === 'string') {
        let num = 0;
        if (col.length > 1) {
            num += (col.charCodeAt(0) - 64) * 26 + (col.charCodeAt(1) - 64);
        } else {
            num += (col.charCodeAt(0) - 64);
        }
        return num;
    }
    if (typeof col === 'number') {
        let str;
        if (col > 26) {
            str = String.fromCharCode((col / 26) + 64) + String.fromCharCode((col % 26) + 64);
        } else {
            str = String.fromCharCode(col + 64);
        }
        return str;
    }
}
/*
//쿼리 스트링 만들기
async function get_Query_string(form: Array<string>) {

let query: string = form.filter(f => f != 'none').join(', ') as string;
return query;
}
*/