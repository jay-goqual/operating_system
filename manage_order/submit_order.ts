var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();

async function submit_Order() {
    const error_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const table = error_sheet.getDataRange().getValues();
    table.splice(0, 1);
    let submit_table = new Array();
    let error_table = new Array();
    table.forEach((t) => {
        if (t[order_form.get('에러확인')] == true) {
            submit_table.push(t);
        } else {
            error_table.push(t);
        }
    });

    if (submit_table.length > 0) {
        target_sheet.insertRowsAfter(1, submit_table.length);
        target_sheet.getRange(2, 1, submit_table.length, submit_table[0].length).setNumberFormat('@').setValues(submit_table);
    }
    if (error_table.length > 0) {
        error_sheet.deleteRows(2, error_sheet.getLastRow() - 1);
        error_sheet.insertRowsAfter(1, error_table.length);
        error_sheet.getRange(2, 1, error_table.length, error_table[0].length).setNumberFormat('@').setValues(error_table);
    }
}

async function modify_Error() {
}

async function fetch_Additional_info() {

    //상품정보 가져오기
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
            o[order_form.get('에러확인')] = true;
        } else {
            o[order_form.get('상품명')] = `error`;
            o[order_form.get('출고채널')] = `error`;
            o[order_form.get('택배사')] = `error`;
            o[order_form.get('판매액')] = `error`;
            o[order_form.get('에러확인')] = false;
        }
        return o;
    });

    order.map((o) => {

        //배송비 구하기
        let orderId: string = o[order_form.get('주문번호')];
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
            o[order_form.get('배송비')] = 'error';
        }

        //결제일 양식 통일 및 변경
        let date = order_form.get('결제일');
        if (!o[date]) {
            let assume: Date;
            if (orderId[0] == 'N') {
                assume = new Date(orderId.substring(1, 5) + '-' + orderId.substring(5, 7) + '-' + orderId.substring(7, 9) + ' 00:00');
            } else {
                assume = new Date(orderId.substring(0, 4) + '-' + orderId.substring(4, 6) + '-' + orderId.substring(6, 8) + ' 00:00');
            }
            let today = new Date(o[order_form.get('접수일')]);

            //@ts-ignore
            if (Math.abs(assume.getFullYear() - today.getFullYear()) > 3 || isNaN(assume.getFullYear())) {
                o[date] = today;
            } else {
                o[date] = assume;
            }
        } else {
            o[date] = new Date(o[date]);
        }
        o[date] = Utilities.formatDate(o[date], 'GMT+9', 'yyyy/MM/dd HH:mm');

        return o;
    });

    sheet.getRange(2, 1, order.length, order[0].length).setValues(order);
}