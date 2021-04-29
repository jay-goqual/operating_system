// const getSheets = () => SpreadsheetApp.getActive().getSheets();

// const getActiveSheetName = () => SpreadsheetApp.getActive().getSheetName();

// export const getSheetsData = () => {
//   const activeSheetName = getActiveSheetName();
//   return getSheets().map((sheet, index) => {
//     const name = sheet.getName();
//     return {
//       name,
//       index,
//       isActive: name === activeSheetName,
//     };
//   });
// };

// export const addSheet = sheetTitle => {
//   SpreadsheetApp.getActive().insertSheet(sheetTitle);
//   return getSheetsData();
// };

// export const deleteSheet = sheetIndex => {
//   const sheets = getSheets();
//   SpreadsheetApp.getActive().deleteSheet(sheets[sheetIndex]);
//   return getSheetsData();
// };

// export const setActiveSheet = sheetName => {
//   SpreadsheetApp.getActive()
//     .getSheetByName(sheetName)
//     .activate();
//   return getSheetsData();
// };

export const findOrder = input => {
    const url = 'https://docs.google.com/spreadsheets/d/1LzKdF7futwfIw_bw1tfko36TRQ86Yf-9jdjNPZQCdac/gviz/tq?gid=0&tq=';
    const query = `select A, B, D, E, H, I, J, K, L, M, F, Q, G, O where H contains '${input}' or J contains '${input}'`;

    const response = UrlFetchApp.fetch(url + query, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const clean = response.getContentText();

    const t = clean.substring(47, clean.length - 2);
    const temp = JSON.parse(t);

    const r = new Array();
    temp.table.rows.forEach((k, i) => {
        r.push({});
        k.c.forEach((key, j) => {
            let a = temp.table.cols[j].label;
            if (key.f) {
                // Object.assign(r[i], {a: key.f});
                r[i][a] = key.f
            } else {
                r[i][a] = key.v
                // Object.assign(r[i], {a: key.v});
            }
        });
    });

    return r;
}

export const getProducts = () => {
    const url = 'https://docs.google.com/spreadsheets/d/13STuUesnhhhAoy27t1dzCDDyx6ImvZNEG8adf7JqXIc/gviz/tq?gid=0&tq=';
    const query = `select A, B where N = 'all' or N = '10001'`;

    const response = UrlFetchApp.fetch(url + query, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const clean = response.getContentText();

    const t = clean.substring(47, clean.length - 2);
    const temp = JSON.parse(t);

    const r = new Array();
    temp.table.rows.forEach((k, i) => {
        r.push({});
        r[i] = {value: k.c[0].v, label: k.c[1].v}
    });

    return r;
}

export const getData = (cs, back, send) => {
    const back_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('회수필요');
    const send_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('출고필요');
    const cs_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('접수');

    let uid = cs_sheet.getRange(2, 1).getValue();
    let today = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy. MM. dd');
    if (uid == '' || !uid) {
        uid = `${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}001`;
    } else {
        uid = String(Number(uid) + 1);
    }

    if (cs && cs.length > 0) {
        cs_sheet.insertRowsAfter(1, cs.length);
        cs_sheet.getRange(2, 3, cs.length, cs[0].length).setValues(cs).setNumberFormat('@');
        cs_sheet.getRange(2, 1, cs.length).setValue(uid).setNumberFormat('@');
        cs_sheet.getRange(2, 2, cs.length).setValue(today).setNumberFormat('yyyy. M. d');
    }

    if (back && back.length > 0) {
        back_sheet.insertRowsAfter(1, back.length);
        back_sheet.getRange(2, 2, back.length, back[0].length).setValues(back).setNumberFormat('@');
        back_sheet.getRange(2, 1, back.length).setValue(uid).setNumberFormat('@');
    }

    if (send && send.length > 0) {
        send_sheet.insertRowsAfter(1, send.length);
        send_sheet.getRange(2, 2, send.length, send[0].length).setValues(send).setNumberFormat('@');
        send_sheet.getRange(2, 1, send.length).setValue(uid).setNumberFormat('@');
    }
}