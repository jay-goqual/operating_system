const getSheets = () => SpreadsheetApp.getActive().getSheets();
const getActiveSheetName = () => SpreadsheetApp.getActive().getSheetName();
const getInputsheet = () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주하기');
const getTargetsheet = () => SpreadsheetApp.getActiveSpreadsheet().getSheetByName('접수내역');

/* export const getSheetsData = () => {
  const activeSheetName = getActiveSheetName();
  return getSheets().map((sheet, index) => {
    const name = sheet.getName();
    return {
      name,
      index,
      isActive: name === activeSheetName,
    };
  });
};

export const addSheet = sheetTitle => {
  SpreadsheetApp.getActive().insertSheet(sheetTitle);
  return getSheetsData();
};

export const deleteSheet = sheetIndex => {
  const sheets = getSheets();
  SpreadsheetApp.getActive().deleteSheet(sheets[sheetIndex]);
  return getSheetsData();
};

export const setActiveSheet = sheetName => {
  SpreadsheetApp.getActive()
    .getSheetByName(sheetName)
    .activate();
  return getSheetsData();
}; */

export const pushOrder = () => {
    const data = getInputsheet().getDataRange().getValues();
    const channel = data[2][1];

    if (channel == '') {
        throw new Error('요청자를 입력해주세요.');
    }

    let last_num = data[5][1];

    const date = Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd");
    if (date != String(data[5][1]).substr(0, 8)) {
        last_num = `${date}1000001`;
    }

    const customer = new Array();
    const product = new Array();
    const product2 = new Array();
    const order = new Array();

    for (var i = 0; i < 25; i++) {
        if (!(data[2 + i][3] == '' || data[2 + i][4] == '' || data[2 + i][5] == '' || data[2 + i][6] == '' || data[2 + i][7] == '' || data[2 + i][8] == '')) {
            const c_length = customer.length;
            customer.push([]);
            for (var j = 0; j < 7; j++) {
                customer[c_length].push(data[2 + i][3 + j]);
            }
        }
        if (!(data[2 + i][11] == '' || data[2 + i][14] == '' || data[2 + i][15] == '')) {
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

    if (customer.length < 1) {
        throw new Error('발주 정보를 모두 입력해주세요.');
    }
    if (product.length < 1 || product2.length < 1) {
        throw new Error('발주 정보를 모두 입력해주세요.');
    }

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

    getInputsheet().getRange(6, 2).setNumberFormat('@').setValue(String(last_num));

    getTargetsheet().insertRowsAfter(1, order.length);
    getTargetsheet().getRange(2, 1, order.length, order[0].length).setNumberFormat('@').setValues(order);
    getInputsheet().getRange(3, 2).setNumberFormat('@').setValue('');
    getInputsheet().getRange(3, 4, 25, 7).setNumberFormat('@').setValue('');
    getInputsheet().getRange(3, 12, 25, 1).setNumberFormat('@').setValue('');
    getInputsheet().getRange(3, 14, 25, 3).setValue('');
    getInputsheet().getRange(3, 18, 25, 1).setNumberFormat('@').setValue('');
}

/* export const findZipcode = () => {
    const data = getInputsheet.getDataRange().getValues();
    const address = data[2][5];
    Logger.log(`https://search.naver.com/search.naver?where=nexearch&ie=utf8&X_CSA=address_search&query=${address} 우편번호`);
    return `https://search.naver.com/search.naver?where=nexearch&ie=utf8&X_CSA=address_search&query=${address} 우편번호`;
} */