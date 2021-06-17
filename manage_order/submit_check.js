// 해피콜 과정이 완료된 주문제작 상품(커튼) 주문을 병합하는 파일입니다.

// 전역함수 호출
var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();
var client = get_Client();

// 확인이 완료된 주문과 기존 [주문현황]에 작성된 주문을 병합하는 함수입니다.
async function submit_Check() {
    // [주문현황] 시트 불러오기
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const target_table = target_sheet.getDataRange().getValues();
    // [확인요청] 시트 불러오기
    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('확인요청');
    const check_table = check_sheet.getDataRange().getValues();

    // [확인요청] 시트 첫번째 제목행 삭제
    check_table.splice(0, 1);

    // count = 확인 완료된 주문의 개수
    var count = 0;

    // [확인요청] 시트의 모든 행을 탐색
    check_table.forEach((c) => {

        // 확인이 완료되지 않은 행에서는 리턴
        if (c[26] == false) { 
            return;
        }

        // 확인이 완료되었을 경우, [주문현황] 시트에서 상품주문번호가 같은 행의 index를 찾음
        let find = target_table.findIndex(t => t[order_form.get('상품주문번호')] == c[order_form.get('상품주문번호')]);
        if (find) {
            // [확인요청] 시트의 최신화된 데이터를 [주문현황] 시트의 열 개수에 맞게 자른 후 동기화
            target_table[find] = c.splice(0, 26);

            // 각 출고처와 택배사를 출고채널 열에 입력
            if (target_table[find][order_form.get('상품코드')].indexOf('CF201') > 0) {
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

    // 데이터 업데이트
    target_sheet.getDataRange().setValues(target_table);

    // 확인이 완료된 순서로 sorting 한 후, 확인 완료된 행은 삭제
    check_sheet.sort(order_form.get('에러확인') + 2);
    check_sheet.deleteRows(check_table.length - count + 2, count);
}