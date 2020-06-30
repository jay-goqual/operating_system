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
    order_data.splice(0, 1);

    order_data.map((o) => {
        let find = invoice_data.filter(d => 
            d[invoice_form.get('주문번호')] == o[order_form.get('상품주문번호')]
        );
        o[order_form.get('송장번호')] = find[invoice_form('송장번호')];
        return o;
    });
}