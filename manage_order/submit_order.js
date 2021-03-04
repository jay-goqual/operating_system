var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();
var client = get_Client();
var postpone = new Array();

async function submit_Order() {
    const error_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('확인요청');
    const table = error_sheet.getDataRange().getValues();
    const target_table =target_sheet.getDataRange().getValues();

    if (table.length == 1) {
        return;
    };
    
    table.splice(0, 1);
    // let submit_table = new Map();
    let total_table = new Array();
    let error_table = new Array();
    let check_table = new Array();

    let count = 0;
    let check = false;
    table.forEach((t, i) => {
        if (target_table.filter(x =>
            (x[order_form.get('주문번호')] == t[order_form.get('주문번호')] && 
            x[order_form.get('상품주문번호')].split('-')[0] == t[order_form.get('상품주문번호')].split('-')[0])).length > 0) {
                check = true;
                error_sheet.getRange(i + 2, 1, 1, table[0].length).setBackground('#f4cccc');
                error_sheet.getRange(i + 2, order_form.get('에러확인') + 1).setValue(false);
                t[order_form.get('에러확인')] = false;
            }
        if (table.filter(x =>
            (x[order_form.get('주문번호')] == t[order_form.get('주문번호')] && 
            x[order_form.get('상품주문번호')] == t[order_form.get('상품주문번호')])).length > 1) {
                check = true;
                error_sheet.getRange(i + 2, 1, 1, table[0].length).setBackground('#f4cccc');
                error_sheet.getRange(i + 2, order_form.get('에러확인') + 1).setValue(false);
                t[order_form.get('에러확인')] = false;
            }
        /* if (postpone.filter(x => (x == t[order_form.get('주문번호')])).length > 0) {
            check = true;
            error_sheet.getRange(i + 2, 1, 1, table[0].length).setBackground('#f4cccc');
            error_sheet.getRange(i + 2, order_form.get('에러확인') + 1).setValue(false);
            t[order_form.get('에러확인')] = false;
        } */
        /* if (total_table.filter(x =>
            (x[order_form.get('주문자')] == t[order_form.get('주문자')] &&
            x[order_form.get('수령인')] == t[order_form.get('수령인')] &&
            x[order_form.get('주소')] == t[order_form.get('주소')] && 
            x[order_form.get('상품코드')] == t[order_form.get('상품코드')] && 
            x[order_form.get('셀러명')] == t[order_form.get('셀러명')] &&
            x[order_form.get('수량')] == t[order_form.get('수량')]) &&
            (x[order_form.get('상품주문번호')].indexOf(t[order_form.get('상품주문번호')].split('-')[0]) != -1)).length > 0) {
                check = true;
                error_sheet.getRange(i + 2, 1, 1, table[0].length).setBackground('#f4cccc');
                error_sheet.getRange(i + 2, order_form.get('에러확인') + 1).setValue(false);
                t[order_form.get('에러확인')] = false;
            } */
        if (t[order_form.get('에러확인')] == true) {
            /* if (!submit_table.get(t[order_form.get('출고채널')])) {
                submit_table.set(t[order_form.get('출고채널')], new Array());
            }
            let temp = submit_table.get(t[order_form.get('출고채널')]);
            temp.push(t);
            submit_table.set(t[order_form.get('출고채널')], temp); */
            if (t[order_form.get('출고채널')].indexOf('대기') != -1) {
                check_table.push(t);
            }
            total_table.push(t);
        } else {
            count++;
            error_table.push(t);
        }
    });

    /* submit_table.forEach((t, k) => {
        if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(k)) {
            SpreadsheetApp.getActiveSpreadsheet().insertSheet().setName(k);
        }
        let target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(k);
        target.insertRowsAfter(1, t.length);
        target.getRange(2, 1, t.length, t[0].length).setValues(t);
    }); */

    if (total_table.length > 0) {
        target_sheet.insertRowsAfter(1, total_table.length);
        target_sheet.getRange(2, 1, total_table.length, total_table[0].length).setValues(total_table);
        error_sheet.sort(order_form.get('에러확인') + 1);
        error_sheet.deleteRows(count + 2, total_table.length);
        if (check_table.length > 0) {
            check_sheet.insertRowsAfter(1, check_table.length);
            check_sheet.getRange(2, 1, check_table.length, check_table[0].length).setValues(check_table);
            // check_sheet.getRange(2, check_table[0].length + 1, check_table.length, 1).setValue(false);
            check_sheet.getRange(2, check_table[0].length + 1, check_table.length, 1).insertCheckboxes();
        }
    };

    if (check) {
        SpreadsheetApp.getUi().alert('중복 주문이 감지되었습니다. 에러확인 시트를 확인해 주세요.');
    }
}

// async function catch_Error(index, order, order_list, num) {
async function catch_Error(index, order) {
    order[order_form.get('에러확인')] = true;

    /* if (order_list.filter(x => x[order_form.get('상품주문번호')].indexOf(order[order_form.get('상품주문번호')]) != -1).length > 1) {
        order[order_form.get('상품주문번호')] = `${order[order_form.get('상품주문번호')]}-${Utilities.formatString('%02d', num)}`;
        //SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인').getRange(index + 2, order_form.get('상품주문번호') + 1).setBackground('#f4cccc');
        //order[order_form.get('에러확인')] = false;
    } */

    const check = ['셀러명', '주문번호', '상품주문번호', '상품코드', '수량', '주문자', '주문자연락처', '수령인', '수령인연락처', '주소', '상품명', '출고채널', '택배사', '판매액', '배송비', '수수료'];
    check.forEach((c) => {
        // if (!order[order_form.get(c)] && order[order_form.get(c)] != 0) {
        if (order[order_form.get(c)] === '' || order[order_form.get(c)] === null) {
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

    if (order.length == 1) {
        return;
    }

    order.splice(0, 1);

    let total = new Map();

    order.map((o) => {
        /* if (o[order_form.get('상품코드')] == 'GQ-PM-A000-2101') {
            postpone.push(o[order_form.get('주문번호')]);
        } */

        o[order_form.get('접수일')] = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm');
        if (o[order_form.get('셀러명')] != '직접발주') {
            o[order_form.get('셀러코드')] = client.get(o[order_form.get('셀러명')]).get('셀러코드');
        }
        // o[order_form.get('셀러코드')] = client.get(o[order_form.get('셀러명')]).get('셀러코드');

        let code = order_form.get('상품코드');
        let p = productInfo.get(o[code]);
        if (p) {
            o[order_form.get('상품명')] = p.get('상품명');
            if (!o[order_form.get('출고채널')]) {
                o[order_form.get('출고채널')] = p.get('출고채널');
                if ((o[order_form.get('셀러코드')][0] == '2' || o[order_form.get('셀러코드')] == '30007') && o[order_form.get('출고채널')] == '대기_커튼') {
                    o[order_form.get('출고채널')] = '제이에스비즈';
                } else if ((o[order_form.get('셀러코드')][0] == '2' || o[order_form.get('셀러코드')] == '30007') && o[order_form.get('출고채널')] == '대기_커튼천') {
                    o[order_form.get('출고채널')] = '건인디앤씨';
                }
                o[order_form.get('택배사')] = delivery.get(p.get('출고채널'));
            }
            if (o[order_form.get('셀러명')] != '직접발주') {
                o[order_form.get('판매액')] = Number(p.get('판매가')) * Number(o[order_form.get('수량')]);
            }

            if (total.has(o[order_form.get('주문번호')])) {
                total.set(o[order_form.get('주문번호')], total.get(o[order_form.get('주문번호')]) + o[order_form.get('판매액')]);
            } else {
                total.set(o[order_form.get('주문번호')], o[order_form.get('판매액')]);
            }
            //수수료구하기

            if (o[order_form.get('셀러명')] != '직접발주') {

            
            let rate;
            let client_info = client.get(o[order_form.get('셀러명')]);
            if (client_info.get('공급방식') == '고정수수료') {
                rate = Number(client_info.get('고정수수료율'));
            } else if (client_info.get('공급방식') == '가산수수료') {
                if (p.get('가산수수료') == 'Y') {
                    rate = Number(p.get('상품수수료율')) + Number(client_info.get('가산수수료율'));
                } else {
                    rate = Number(p.get('상품수수료율'));
                }
            } else {
                rate = Number(client_info.get('고정수수료율'));
            }

            if (p.get(String(o[order_form.get('셀러코드')]))) {
                // o[order_form.get('수수료')] = o[order_form.get('판매액')] - (Number(p.get(String(o[order_form.get('셀러코드')]))) * o[order_form.get('수량')]);
                o[order_form.get('판매액')] = Number(p.get(String(o[order_form.get('셀러코드')]))) * o[order_form.get('수량')];
                o[order_form.get('수수료')] = 0;
            } else {
                o[order_form.get('수수료')] = Math.ceil((Number(p.get('판매가')) * rate) / 10) * 10 * o[order_form.get('수량')];
            }

            }

            /* if (total.has(o[order_form.get('주문번호')])) {
                if (client_info.get('공급방식') == '공급가') {
                    total.set(o[order_form.get('주문번호')], total.get(o[order_form.get('주문번호')]) + Number(p.get('판매가')) * Number(o[order_form.get('수량')]));
                } else {
                    total.set(o[order_form.get('주문번호')], total.get(o[order_form.get('주문번호')]) + o[order_form.get('판매액')]);
                }
            } else {
                if (client_info.get('공급방식') == '공급가') {
                    total.set(o[order_form.get('주문번호')], Number(p.get('판매가')) * Number(o[order_form.get('수량')]));
                } else {
                    total.set(o[order_form.get('주문번호')], o[order_form.get('판매액')]);
                }
            } */
        }
        return o;
    });

    let num = 0;

    order.map(async (o, i) => {

        //배송비 구하기
        let orderId = o[order_form.get('주문번호')];
        let code = o[order_form.get('상품코드')];
        let t = total.get(orderId);
        let p = productInfo.get(code);
        if (o[order_form.get('셀러명')] != '직접발주') {
            if (p) {
                if (Number(t) > Number(p.get('무료배송기준')) || t == -1) {
                    o[order_form.get('배송비')] = 0;
                } else {
                    o[order_form.get('배송비')] = p.get('상품배송비');
                    total.set(orderId, -1);
                }
            }   
        }

        //결제일 양식 통일 및 변경
        let date = order_form.get('결제일');
        if (!o[date] || isNaN(new Date(o[date]).getTime())) {
            let n;
            if (orderId[0] == 'N') {
                n = 1;
            } else {
                n = 0;
            }
            let assume = new Date(`${orderId.substring(n, n + 4)}-${orderId.substring(n + 4, n + 6)}-${orderId.substring(n + 6, n + 8)} 00:00`);
            let today = new Date(o[order_form.get('접수일')]);
            o[date] = today;
            let oneDay = 24 * 60 * 60 * 1000;
            if (Math.round(Math.abs(assume - today) / oneDay) < 180) {
                o[date] = assume;
            }
        } else {
            o[date] = new Date(o[date]);
        }
        o[date] = Utilities.formatDate(o[date], 'GMT+9', 'yyyy/MM/dd HH:mm');


        /* if (order.filter(x => x[order_form.get('상품주문번호')].indexOf(o[order_form.get('상품주문번호')]) != -1).length > 1) {
            num++;
        } else {
            num = 0;
        } */

        if (order.filter(x => x[order_form.get('상품주문번호')].split('-')[0] == o[order_form.get('상품주문번호')]).length > 1) {
            num++;
        } else {
            num = 0;
        }

        if (num > 0 && (o[order_form.get('셀러명')] == '직접발주' || o[order_form.get('셀러명')] == '씨씨티비프렌즈' || o[order_form.get('셀러명')] == '나혼자살림' || o[order_form.get('셀러명')] == '도치퀸' || o[order_form.get('셀러명')] == '오늘의집')) {
            o[order_form.get('상품주문번호')] = `${o[order_form.get('상품주문번호')]}-${Utilities.formatString('%02d', num)}`;
        }

        /* if (i > 1) {
            if (o[order_form.get('주문번호')] == order[i - 1][order_form.get('주문번호')]) {
                num++;
            } else {
                num = 1;
            }
        } */
        
        // o = await catch_Error(i, o, order, num)
        
        o = await catch_Error(i, o);
        
        return o;
    });

    sheet.getRange(2, 1, order.length, order[0].length).setValues(order);
}