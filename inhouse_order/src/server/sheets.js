// https://github.com/enuchi/React-Google-Apps-Script
// 위 라이브러리를 활용한 apps script 입니다.

// 전역함수 선언
const getSheets = () => SpreadsheetApp.getActive().getSheets();
const getActiveSheetName = () => SpreadsheetApp.getActive().getSheetName();
const getInputsheet = () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주하기');
const getTargetsheet = () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName('접수내역');

// [발주하기] 시트의 데이터를 검수하고 [출고관리] 스프레드시트로 옮기기 위한 더미 시트에 옮겨놓는 함수입니다.
export const pushOrder = () => {

    // [발주하기] 시트 데이터 로드
    const data = getInputsheet().getDataRange().getValues();

    const channel = data[2][1];

    if (channel == '') {
        throw new Error('요청자를 입력해주세요.');
    }

    // 주문번호 자동생성을 위한 count 값 불러오기
    let last_num = data[5][1];

    const date = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd");
    // 최종 주문번호가 오늘자가 아닐경우, 오늘자 첫번째 주문번호로 대체
    // 주문번호 생성규칙 = yyyyMMdd + 5자리 셀러코드 + 2자리 count
    if (date != String(data[5][1]).substr(0, 8)) {
        last_num = `${date}1000001`;
    }

    const customer = new Array();
    const product = new Array();
    const product2 = new Array();
    const order = new Array();

    // 각 위치의 데이터를 배열화하기
    for (var i = 0; i < 25; i++) {
        if (!(data[2 + i][3] == '' || data[2 + i][4] == '' || data[2 + i][5] == '' || data[2 + i][6] == '' || data[2 + i][7] == '' || data[2 + i][8] == '')) {
            const c_length = customer.length;
            customer.push([]);
            for (var j = 0; j < 7; j++) {
                customer[c_length].push(data[2 + i][3 + j]);
            }
        }
        if (!(data[2 + i][11] == '' || data[2 + i][14] === '' || data[2 + i][15] === '')) {
            const p_length = product.length;
            product.push([]);
            product2.push([]);
            product[p_length].push(data[2 + i][12]);
            product[p_length].push(data[2 + i][15]);
            product2[p_length].push(data[2 + i][17]);
            product2[p_length].push('');
            product2[p_length].push(data[2 + i][11]);
            product2[p_length].push('');
            product2[p_length].push('');
            product2[p_length].push(data[2 + i][16]);
            product2[p_length].push(0);
            product2[p_length].push(0);
        }
    }

    // 주문자정보가 없을 경우 throw error
    if (customer.length < 1) {
        throw new Error('발주 정보를 모두 입력해주세요.');
    }
    // 주문상품 정보가 없을 경우 throw error
    if (product.length < 1 || product2.length < 1) {
        throw new Error('발주 정보를 모두 입력해주세요.');
    }

    // 모든 주문자에 대해서 발주상품정보 연결
    for (var i of customer) {
        for (var j in product) {
            const o_length = order.length;
            order.push([]);
            order[o_length].push(channel);
            order[o_length].push(last_num);
            order[o_length].push(last_num);
            for (var k in product[j]) {
                order[o_length].push(product[j][k]);
            }
            for (var k in i) {
                order[o_length].push(i[k]);
            }
            for (var k in product2[j]) {
                order[o_length].push(product2[j][k]);
            }
        }
        last_num++;
    }

    // [접수내역] 시트로 이동 및 [발주하기] 시트 데이터 초기화
    getInputsheet().getRange(6, 2).setNumberFormat('@').setValue(String(last_num));

    getTargetsheet().insertRowsAfter(1, order.length);
    getTargetsheet().getRange(2, 1, order.length, order[0].length).setNumberFormat('@').setValues(order);
    getInputsheet().getRange(3, 2).setNumberFormat('@').setValue('');
    getInputsheet().getRange(3, 4, 1, 7).setNumberFormat('@').setValue('');
    getInputsheet().getRange(3, 12, 25, 1).setNumberFormat('@').setValue('');
    getInputsheet().getRange(3, 14, 25, 3).setValue('');
    getInputsheet().getRange(3, 18, 25, 1).setNumberFormat('@').setValue('');
}