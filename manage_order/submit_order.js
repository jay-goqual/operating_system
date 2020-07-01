var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();
var client = get_Client();

async function submit_Order() {
    const error_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const table = error_sheet.getDataRange().getValues();
    table.splice(0, 1);
    let submit_table = new Map();
    let total_table = new Array();
    let error_table = new Array();

    let count = 0;
    table.forEach((t) => {
        if (t[order_form.get('에러확인')] == true) {
            if (!submit_table.get(t[order_form.get('출고채널')])) {
                submit_table.set(t[order_form.get('출고채널')], new Array());
            }
            let temp = submit_table.get(t[order_form.get('출고채널')]);
            temp.push(t);
            submit_table.set(t[order_form.get('출고채널')], temp);
            total_table.push(t);
        } else {
            count++;
            error_table.push(t);
        }
    });

    submit_table.forEach((t, k) => {
        let target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(k);
        target.insertRowsAfter(1, t.length);
        target.getRange(2, 1, t.length, t[0].length).setValues(t);
    });

    if (total_table.length > 0) {
        target_sheet.insertRowsAfter(1, total_table.length);
        target_sheet.getRange(2, 1, total_table.length, total_table[0].length).setValues(total_table);
        error_sheet.sort(order_form.get('에러확인') + 1);
        error_sheet.deleteRows(count + 2, total_table.length);
    }
}

async function catch_Error(index, order, order_list, num) {
    order[order_form.get('에러확인')] = true;

    if (order_list.filter(x => x[order_form.get('상품주문번호')].indexOf(order[order_form.get('상품주문번호')]) != -1).length > 1) {
        order[order_form.get('상품주문번호')] = `${order[order_form.get('상품주문번호')]}-${Utilities.formatString('%02d', num)}`;
        //SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인').getRange(index + 2, order_form.get('상품주문번호') + 1).setBackground('#f4cccc');
        //order[order_form.get('에러확인')] = false;
    }

    const check = ['셀러명', '주문번호', '상품주문번호', '상품코드', '수량', '주문자', '주문자연락처', '수령인', '수령인연락처', '주소', '우편번호', '상품명', '출고채널', '택배사', '판매액', '배송비', '수수료'];
    check.forEach((c) => {
        // if (!order[order_form.get(c)] && order[order_form.get(c)] != 0) {
        if (order[order_form.get(c)] === '') {
            order[order_form.get('에러확인')] = false;
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인').getRange(index + 2, order_form.get(c) + 1).setBackground('#f4cccc');
        }
    });

    return order;
}

async function fetch_Additional_info() {

    //상품정보 가져오기
    const productInfo = get_Product();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const order = sheet.getDataRange().getValues();
    order.splice(0, 1);

    let total = new Map();

    order.map((o) => {
        o[order_form.get('접수일')] = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm');
        o[order_form.get('셀러코드')] = client.get(o[order_form.get('셀러명')]).get('셀러코드');

        let code = order_form.get('상품코드');
        let p = productInfo.get(o[code]);
        if (p) {
            o[order_form.get('상품명')] = p.get('상품명');
            o[order_form.get('출고채널')] = p.get('출고채널');
            o[order_form.get('택배사')] = delivery.get(p.get('출고채널'));
            o[order_form.get('판매액')] = Number(p.get('판매가')) * o[order_form.get('수량')];

            //수수료구하기
            let rate;
            let client_info = client.get(o[order_form.get('셀러명')]);
            if (client_info.get('공급방식') == '고정수수료') {
                rate = Number(client_info.get('고정수수료율'));
            } else if (client_info.get('공급방식') == '가산수수료') {
                rate = Number(p.get('상품수수료율')) + Number(client_info.get('가산수수료율'));
            } else {
                rate = Number(client_info.get('고정수수료율'));
            }

            if (p.get(o[order_form.get('셀러코드')])) {
                o[order_form.get('수수료')] = o[order_form.get('판매액')] - (Number(p.get(o[order_form.get('셀러코드')])) * o[order_form.get('수량')]);
            } else {
                o[order_form.get('수수료')] = Math.floor((Number(p.get('판매가')) * rate) / 10) * 10 * o[order_form.get('수량')];
            }

            if (total.has(o[order_form.get('주문번호')])) {
                total.set(o[order_form.get('주문번호')], total.get(o[order_form.get('주문번호')]) + o[order_form.get('판매액')]);
            } else {
                total.set(o[order_form.get('주문번호')], o[order_form.get('판매액')]);
            }
        }
        return o;
    });

    let num = 1;

    order.map(async (o, i) => {

        //배송비 구하기
        let orderId = o[order_form.get('주문번호')];
        let code = o[order_form.get('상품코드')];
        let t = total.get(orderId);
        let p = productInfo.get(code);
        if (p) {
            if (Number(t) > Number(p.get('무료배송기준')) || t == -1) {
                o[order_form.get('배송비')] = 0;
            } else {
                o[order_form.get('배송비')] = p.get('상품배송비');
                total.set(orderId, -1);
            }
        }

        //결제일 양식 통일 및 변경
        let date = order_form.get('결제일');
        if (!o[date]) {
            let assume;
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

        if (i > 1) {
            if (o[order_form.get('주문번호')] == order[i - 1][order_form.get('주문번호')]) {
                num++;
            } else {
                num = 1;
            }
        }
        
        o = await catch_Error(i, o, order, num)
        
        return o;
    });

    sheet.getRange(2, 1, order.length, order[0].length).setValues(order);
}