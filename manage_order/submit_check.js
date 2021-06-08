var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();
var client = get_Client();

async function submit_Check() {
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const target_table = target_sheet.getDataRange().getValues();
    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('확인요청');
    const check_table = check_sheet.getDataRange().getValues();

    check_table.splice(0, 1);

    var count = 0;
    check_table.forEach((c) => {
        if (c[26] == false) { 
            return;
        }
        let find = target_table.findIndex(t => t[order_form.get('상품주문번호')] == c[order_form.get('상품주문번호')]);
        if (find) {
            target_table[find] = c.splice(0, 26);
            if(target_table[find][order_form.get('상품코드')].indexOf('CF201') > 0) {
                target_table[find][order_form.get('출고채널')] = '건인디앤씨'; 
                target_table[find][order_form.get('택배사')] = 'CJ대한통운';
            } else if (target_table[find][order_form.get('상품코드')].indexOf('CF211') > 0) {
                target_table[find][order_form.get('출고채널')] = '드림캐쳐'; 
                target_table[find][order_form.get('택배사')] = '로젠택배';
            } else {
                target_table[find][order_form.get('출고채널')] = '제이에스비즈';
            }
            count++;
        }
    })

    target_sheet.getDataRange().setValues(target_table);

    check_sheet.sort(order_form.get('에러확인') + 2);
    check_sheet.deleteRows(check_table.length - count + 2, count);
}