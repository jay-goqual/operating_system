// CX_operating (cx 대시보드)로 대체되어 사용되지 않을 파일입니다.

var order_form = get_Order_form();

async function find_order() {
    const result_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('커튼주문검색');
    const result_table = result_sheet.getDataRange().getValues();
    const file_name = `${result_table[2][1]} ${result_table[2][2]}`;
    const channel = result_table[2][3];
    const key_name = result_table[2][4];
    const result = new Array();

    if (key_name == null || key_name == '') {
        return;
    }

    let month_file = DriveApp.getFolderById(ref.get('출고DB')).getFilesByName(file_name);
    const db_file = SpreadsheetApp.openById(month_file.next().getId());
    db_file.getSheets().forEach((sheet) => {
        const db_table = sheet.getDataRange().getValues();
        let find = db_table.filter(t => (t[order_form.get('송장번호')] != null && t[order_form.get('송장번호')] != '' && t[order_form.get('출고일시')] != '') && t[order_form.get('출고채널')] == channel && (t[order_form.get('주문자')] == key_name || t[order_form.get('수령인')] == key_name))
        find.forEach((f) => {
            result.push(f);
        })
    })

    result_sheet.getRange(6, 1, result_sheet.getLastRow() - 4, result_sheet.getLastColumn()).clear();
    if (result.length > 0) {
        result_sheet.getRange(6, 1, result.length, result_table[4].length).setValues(result);
    }
}

async function find_order2() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName().indexOf('검색') == -1) {
        return;
    }

    const data = sheet.getDataRange().getValues();
    const name = data[2][2];
    const phone = data[2][3];
    
    if (name == '' && phone == '') {
        return;
    }
    
    const find_sheet = SpreadsheetApp.openById('1LzKdF7futwfIw_bw1tfko36TRQ86Yf-9jdjNPZQCdac').getSheetByName(sheet.getName());

    sheet.getRange(6, 1, 95, 14).clear().setNumberFormat('@');

    find_sheet.getRange(1, 1).setValue(name);
    find_sheet.getRange(1, 2).setValue(phone);

    const result = find_sheet.getDataRange().getValues();

    result.splice(0, 2);

    if (result.length > 0) {
        sheet.getRange(6, 1, result.length, 14).setValues(result);
    }
}

async function find_Order3() {
    const file = DriveApp.getFileById('16AzZFrMNIQS8R_H2Vn7loH3DhE0lAjTz');
    const file2 = DriveApp.getFileById('1w4PIC2lQprb5jRjykdi9xWKC62J7gmWs');
    const file3 = DriveApp.getFileById('188A8Tjl3UzgBWn6W8P-4dFgajTUK4N6p');
    const content = file.getBlob().getDataAsString();
    const content2 = file2.getBlob().getDataAsString();
    const content3 = file3.getBlob().getDataAsString();
    const json = [JSON.parse(content), JSON.parse(content2), JSON.parse(content3)];

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName().indexOf('이전주문검색') == -1) {
        return;
    }
    const data = sheet.getDataRange().getValues();

    sheet.getRange(6, 1, 95, 20).clear().setNumberFormat('@');

    const name = data[2][2];
    const phone = data[2][3];

    if (name == '' && phone == '') {
        return;
    }

    const result = new Array();
    const result2 = new Array();

    let find = new Array();

    for (j of json) {
    for (t in j) {
        for (i in j[t]) {
            if (phone == '' || name == '') {
                let temp = new Array();
                if(((j[t][i]['order_phone'] == phone || j[t][i]['customer_phone'] == phone) && phone != '') || ((j[t][i]['order_name'] == name || j[t][i]['customer_name'] == name) && name != '')) {
                    temp.push(j[t][i]['date_receipt'], j[t][i]['seller_name'], j[t][i]['order_id'], j[t][i]['order_uid'], j[t][i]['order_name'], j[t][i]['order_phone'], j[t][i]['customer_name'], j[t][i]['customer_phone'], j[t][i]['customer_address'], j[t][i]['customer_zipcode'], j[t][i]['product_code'], j[t][i]['product_name'], j[t][i]['product_num']);
                    result.push(temp);
                }
            } else {
                if ((j[t][i]['order_phone'] == phone && j[t][i]['order_name'] == name) || (j[t][i]['customer_phone'] == phone && j[t][i]['customer_name'] == name)) {
                    let temp = new Array();
                    temp.push(j[t][i]['date_receipt'], j[t][i]['seller_name'], j[t][i]['order_id'], j[t][i]['order_uid'], j[t][i]['order_name'], j[t][i]['order_phone'], j[t][i]['customer_name'], j[t][i]['customer_phone'], j[t][i]['customer_address'], j[t][i]['customer_zipcode'], j[t][i]['product_code'], j[t][i]['product_name'], j[t][i]['product_num']);
                    result.push(temp);
                }
            }
        }
    }
    }

    if (result.length > 0) {
        sheet.getRange(6, 1, result.length, 13).setValues(result);
    }
}