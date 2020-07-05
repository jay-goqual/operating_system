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
    const sale_form = ['접수일', '셀러명', '셀러코드', '주문번호', '상품주문번호', '상품코드', '수량', '결제일', '판매액', '배송비', '수수료', '송장번호'];

    let push_table = new Array();
    data.forEach((d, i) => {
        push_table[i] = new Array();
        sale_form.forEach((f) => {
            push_table[i].push(d[order_form.get(f)]);
        })
    });
    
    manage_sale.insertRowsAfter(1, data.length);
    manage_sale.getRange(2, 1, data.length, sale_form.length).setValues(push_table);

    data_sheet.deleteRows(2, data_sheet.getLastRow() - 1);
}