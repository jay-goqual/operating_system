var ref = get_Ref();
var order_form = get_Order_form();

async function fetch_Invoice(file) {

    const invoice_data = SpreadsheetApp.openById(file.getId()).getDataRange().getValues();
    const order_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();

    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    const check_data = check_sheet.getDataRange().getValues();

    let invoice_form = new Map();
    invoice_data[0].forEach((d, i) => {
        invoice_form.set(d, i);
    });

    let check = 0;

    invoice_data.splice(0, 1);

    order_data.map((o) => {
        if (o[order_form.get('송장번호')]) {
            return o;
        }
        let find;
        if (invoice_form.get('출고상태')) {
            //find = invoice_data.filter(d => (d[invoice_form.get('주문번호')] == o[order_form.get('상품주문번호')]) && d[invoice_form.get('출고상태')] == '확정');
            find = invoice_data.filter(d => (d[invoice_form.get('주문번호')] == o[order_form.get('상품주문번호')]));
            check = 1;
        } else {
            find = invoice_data.filter(d => (d[invoice_form.get('주문번호')] == o[order_form.get('주문번호')]));
            check = 2;
        }
        if (find.length > 0) {
            o[order_form.get('송장번호')] = find[0][invoice_form.get('송장번호')];
            if (o[order_form.get('택배사')] == '2 - 롯데택배') {
                o[order_form.get('택배사')] = '롯데택배'
            }
        }
        return o;
    });

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황')
    sheet.getDataRange().setValues(order_data);

    if (check == 1) {
        check_data[1][5]++;
    } else if (check == 2) {
        check_data[2][5]++;
    }
    check_sheet.getDataRange().setValues(check_data);

    //sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn() - 1).sort({column: order_form.get('송장번호') + 1, ascending: true});

    DriveApp.getFolderById(ref.get('업로드/아카이브')).addFile(file);
    file.getParents().next().removeFile(file);
}

async function fetch_Invoice_curtain() {
    const dashboard = SpreadsheetApp.openById(ref.get('제이에스비즈'));
    const invoice_data = dashboard.getSheetByName('송장번호 전달').getDataRange().getValues();
    const order_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();

    invoice_data.splice(0, 1);

    order_data.map((o) => {
        if (o[order_form.get('송장번호')]) {
            return o;
        }
        let find = invoice_data.filter(d => d[0] == o[order_form.get('주문번호')]);
        if (find.length > 0 && o[order_form.get('출고채널')] == '제이에스비즈') {
            o[order_form.get('송장번호')] = find[0][5];
        }
        return o;
    });

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황')
    sheet.getDataRange().setValues(order_data);

    dashboard.getSheetByName('송장번호 전달').getRange(2, 6, dashboard.getSheetByName('송장번호 전달').getLastRow() - 1, 1).clear();
    dashboard.getSheetByName('송장번호 전달').getRange(2, 6, dashboard.getSheetByName('송장번호 전달').getLastRow() - 1, 1).clearFormat();

    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    const check_data = check_sheet.getDataRange().getValues();

    check_data[3][5]++;
    check_sheet.getDataRange().setValues(check_data);
}