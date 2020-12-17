var ref = get_Ref();
var order_form = get_Order_form();

async function fetch_Invoice(file) {

    const invoice_data = SpreadsheetApp.openById(file.getId()).getDataRange().getValues();
    const order_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();

    let invoice_form = new Map();
    invoice_data[0].forEach((d, i) => {
        invoice_form.set(d, i);
    });

    invoice_data.splice(0, 1);

    order_data.map((o) => {
        if (o[order_form.get('송장번호')]) {
            return o;
        }
        let find = invoice_data.filter(d => d[invoice_form.get('주문번호')] == o[order_form.get('상품주문번호')]);
        if (find.length > 0) {
            o[order_form.get('송장번호')] = find[0][invoice_form.get('송장번호')];
        }
        return o;
    });

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황')
    sheet.getDataRange().setValues(order_data);
    //sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn() - 1).sort({column: order_form.get('송장번호') + 1, ascending: true});

    DriveApp.getFolderById(ref.get('업로드/아카이브')).addFile(file);
    file.getParents().next().removeFile(file);
}