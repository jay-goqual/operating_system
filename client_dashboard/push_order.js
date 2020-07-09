function Init() {
    SpreadsheetApp.getUi()
        .createMenu('헤이홈')
            .addItem('주문제출', 'push_Order')
    .addToUi();
}

async function push_Order() {
    const order_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문접수');
    order_sheet.getRange(2, 1, order_sheet.getLastRow() - 1, order_sheet.getLastColumn()).setBackground(null);
    let data = order_sheet.getDataRange().getValues();
    let error = [false, ''];
    data.splice(0, 1);

    if (data.length == 0) {
        SpreadsheetApp.getUi().alert('주문이 없습니다.');
        return;
    }

    data = data.map((row, i) => {
        return row.map((d, j) => {

            //배송메세지 \n 삭제
            if (j == 10) {
                return d.split('\n').join(' ');
            }

            if (j == 11 || j == 12) {
                return d;
            }

            if (d == '' || d == '-') {
                order_sheet.getRange(i + 2, j + 1).setBackground('#f4cccc');
                error[0] = true;
                error[1] += `[${convert_Column(j + 1)}${i + 2}] 빈 데이터가 있습니다.\n`;
            }

            //상품주문번호 중복 확인
            if (j == 1) {
                if (data.filter(x => x[1] == d).length > 1) {
                    order_sheet.getRange(i + 2, j + 1).setBackground('#f4cccc');
                    error[0] = true;
                    error[1] += `[${convert_Column(j + 1)}${i + 2}] 중복된 상품주문번호가 있습니다.\n`;
                }
            }

            const product = get_Product();
            //상품코드 확인
            if (j == 2) {
                if (!product.get(d)) {
                    order_sheet.getRange(i + 2, j + 1).setBackground('#f4cccc');
                    error[0] = true;
                    error[1] += `[${convert_Column(j + 1)}${i + 2}] 상품코드가 존재하지 않거나 판매종료된 상품입니다.\n`;
                }
            }

            //우편번호 양식 변경
            if (j == 9) {
                let temp = d.split('-').join('');
                if (temp.lenth > 6) {
                    order_sheet.getRange(i + 2, j + 1).setBackground('#f4cccc');
                    error[0] = true;
                    error[1] += `[${convert_Column(j + 1)}${i + 2}] 잘못된 우편번호입니다.\n`;
                }
                return Utilities.formatString('%05d', temp);
            }

            //전화번호 양식변경
            if (j == 5 || j == 7) {
                let temp = d.split('-').join('');
                if (Number(temp) != temp) {
                    order_sheet.getRange(i + 2, j + 1).setBackground('#f4cccc');
                    error[0] = true;
                    error[1] += `[${convert_Column(j + 1)}${i + 2}] 잘못된 전화번호입니다.\n`;
                }
                return temp;
            }

            return d;
        });
    });


    if (error[0]) {
        SpreadsheetApp.getUi().alert(`${error[1]}\n위 에러와 셀 위치를 참고해 주문을 수정한 후 다시 실행해주세요.`);
        return;
    }

    // const target = SpreadsheetApp.openById('1CXXuEY-lf00i8xCMWJT5keCFfHrPsE0MILIIeueTYMA').getSheetByName('에러확인');
    const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문데이터');

    target.insertRowsAfter(1, data.length);
    // target.getRange(2, 2, data.length, 1).setValue(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(2, 3).getValue());
    target.getRange(2, 1, data.length, data[0].length).setNumberFormat('@').setValues(data);

    SpreadsheetApp.getUi().alert(`주문 제출이 완료되었습니다.\n`);

    order_sheet.getRange(2, 1, order_sheet.getLastRow() - 1, order_sheet.getLastColumn()).clear();

    return;
}

function get_Product() {
    const table = SpreadsheetApp.openById('13STuUesnhhhAoy27t1dzCDDyx6ImvZNEG8adf7JqXIc').getSheetByName('상품DB').getDataRange().getValues();
    let product = new Map();

    table.forEach((t) => {
        let temp = new Map();
        t.forEach((x, i) => {
            temp.set(table[0][i], x);
        })
        product.set(t[0], temp);
    });

    return product;
}

function convert_Column(col) {
    if (typeof col === 'string') {
        let num = 0;
        if (col.length > 1) {
            num += (col.charCodeAt(0) - 64) * 26 + (col.charCodeAt(1) - 64);
        } else {
            num += (col.charCodeAt(0) - 64);
        }
        return num;
    }
    if (typeof col === 'number') {
        let str;
        if (col > 26) {
            str = String.fromCharCode((col / 26) + 64) + String.fromCharCode((col % 26) + 64);
        } else {
            str = String.fromCharCode(col + 64);
        }
        return str;
    }
}