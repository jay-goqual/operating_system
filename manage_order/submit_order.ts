var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();

async function get_Error() {

}

async function modify_Error() {

}

async function fetch_Product_info() {
    const productInfo = get_Product();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const order = sheet.getDataRange().getValues();
    order.splice(0, 1);

    let total: Map<string, number> = new Map();

    order.map((o) => {
        let code = order_form.get('상품코드');
        let p = productInfo.get(o[code]);
        if (p) {
            o[order_form.get('상품명')] = p[0];
            o[order_form.get('출고채널')] = p[1];
            o[order_form.get('택배사')] = delivery.get(p[1]);
            o[order_form.get('판매액')] = Number(p[2]) * o[order_form.get('수량')];
            if (total.has(o[order_form.get('주문번호')])) {
                total.set(o[order_form.get('주문번호')], total.get(o[order_form.get('주문번호')]) + o[order_form.get('판매액')]);
            } else {
                total.set(o[order_form.get('주문번호')], o[order_form.get('판매액')]);
            }
        } else {
            o[order_form.get('상품명')] = `can't find product`;
            o[order_form.get('출고채널')] = `can't find product`;
            o[order_form.get('택배사')] = `can't find product`;
            o[order_form.get('판매액')] = `can't find product`;
        }
        return o;
    });

    //배송비 구하기
    order.map((o) => {
        let orderId = o[order_form.get('주문번호')];
        let code = o[order_form.get('상품코드')];
        let t = total.get(orderId);
        let p = productInfo.get(code);
        if (p) {
            if (Number(t) > Number(p[3]) || t == -1) {
                o[order_form.get('배송비')] = 0;
            } else {
                o[order_form.get('배송비')] = p[4];
                total.set(orderId, -1);
            }
        } else {
            o[order_form.get('배송비')] = 'undefined';
        }
        return o;
    });

    sheet.getRange(2, 1, order.length, order[0].length).setValues(order);
}