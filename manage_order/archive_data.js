var ref = get_Ref();
var order_form = get_Order_form();

async function archive_Data() {
    const data_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const data = data_sheet.getDataRange().getValues();
    data.splice(0, 1);

    if (data.length == 0) {
        return;
    }

    let month_file = DriveApp.getFolderById(ref.get('출고DB')).getFilesByName(Utilities.formatDate(new Date(), 'GMT+9', 'yy년 MM월'));
    if (!month_file.hasNext()) {
        const source = {
            title: Utilities.formatDate(new Date(), 'GMT+9', 'yy년 MM월'),
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [{id: ref.get('출고DB')}]
        }
        month_file = DriveApp.getFileById(Drive.Files.insert(source).id);
    } else {
        month_file = month_file.next();
    }

    const month_target = SpreadsheetApp.openById(month_file.getId());
    let new_sheet = data_sheet.copyTo(month_target);
    new_sheet.setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'd')}일`);

    let check_sheet = month_target.getSheets()[0];
    if (check_sheet.getName().indexOf('일') == -1) {
        month_target.deleteSheet(check_sheet);
    }

    const manage_sale = SpreadsheetApp.openById(ref.get('이번달DB')).getSheetByName('주문');
    const total = SpreadsheetApp.openById('1VM1iKCp9RkiktD_4CXfVkENmA1GLyM66de6OAt9-sg0').getSheetByName('6개월주문DB');
    const sale_form = ['접수일', '셀러명', '셀러코드', '주문번호', '상품주문번호', '상품코드', '수량', '결제일', '판매액', '배송비', '수수료', '송장번호', '출고일시', '주문자', '수령인', '수령인연락처'];

    let push_table = new Array();
    let count = 0;
    data.forEach((d, i) => {
        if (d[order_form.get('출고일시')]) {
            push_table[count] = new Array();
            sale_form.forEach((f) => {
                push_table[count].push(d[order_form.get(f)]);
            });
            count++;
        }
    });

    if (push_table.length > 0) {
        manage_sale.insertRowsAfter(1, push_table.length);
        manage_sale.getRange(2, 1, push_table.length, sale_form.length).setValues(push_table);

        total.insertRowsAfter(1, push_table.length);
        total.getRange(2, 1, push_table.length, sale_form.length).setValues(push_table);

        data_sheet.sort(24, false);
        data_sheet.deleteRows(2, push_table.length);
    }

    const c_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    c_sheet.getRange(2, 2, c_sheet.getLastRow() - 1, 2).setValue(0);
    c_sheet.getRange(2, 6, 4, 1).setValue(0);
}