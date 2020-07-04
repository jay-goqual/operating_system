function Init() {
    SpreadsheetApp.getUi()
        .createMenu('헤이홈')
            .addItem('주문제출', 'push_Order')
    .addToUi();
}

async function push_Order() {
    const order_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문접수');
    const data = order_sheet.getDataRange().getValues();
    let error = [false, ''];
    data.splice(0, 1);

    if (data.length == 0) {
        SpreadsheetApp.getUi().alert('주문이 없습니다.');
        return;
    }

    data.forEach((row, i) => {
        row.forEach((d, j) => {
            if (j == 10 || j == 11) {
                return;
            }
            if (d == '') {
                order_sheet.getRange(i + 2, j + 1).setBackground('f4cccc');
                error[0] = true;
                error[1] += '빈 데이터가 있습니다.\\n'
            }

            //상품주문번호 중복 확인
            if (j == 1) {
                if (data.filter(x => x[1] == d).length > 1) {
                    order_sheet.getRange(i + 2, j + 1).setBackground('f4cccc');
                    error[0] = true;
                    error[1] += '중복된 상품주문번호가 있습니다.\\n'
                }
            }

            //상품코드 확인
            if (j == 2) {
                
            }
        });
    });

    if (error[0]) {
        SpreadsheetApp.getUi().alert(`${error[1]}위 에러와 빨간 셀을 참고해 주문을 수정한 후 다시 실행해주세요.`);
        return;
    }
}