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
    const query = `select A, B, D where N = 'all' or N = '10001'`;

    const response = UrlFetchApp.fetch(url + query, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const clean = response.getContentText();

    const t = clean.substring(47, clean.length - 2);
    const temp = JSON.parse(t);

    const r = new Array();
    temp.table.rows.forEach((k, i) => {
        r.push({});
        r[i] = {value: k.c[0].v, label: k.c[1].v, channel: k.c[2].v}
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

export const pushData = () => {
    const cs_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('접수');
    const data = cs_sheet.getDataRange().getValues();
    data.splice(0, 1);

    if (data.length < 1) {
        return;
    }

    let wait = new Array(), complete = new Array();

    data.map((d) => {
        if (d[2] == '교환' || d[2] == '단순반품' || d[2] == '보상반품' || d[2] == '검수필요') {
            wait.push(d);
        } else {
            complete.push(d);
        }
    });

    if (wait.length > 0) {
        const match_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('매칭대기');
        match_sheet.insertRowsAfter(1, wait.length);
        match_sheet.getRange(2, 1, wait.length, wait[0].length).setValues(wait);
    }

    if (complete.length > 0) {
        const complete_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('최종확인');
        complete_sheet.insertRowsAfter(1, complete.length);
        complete_sheet.getRange(2, 1, complete.length, complete[0].length).setValues(complete);
    }

    cs_sheet.deleteRows(2, data.length);
}

export const getInspection = () => {
    const inspection_sheet = SpreadsheetApp.openById('12tU0D6wku0XBgH3y-8cypuv6Gs45V06WvUP6c_kBbjI').getSheets()[0];
    const push_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('검수완료');

    const data = inspection_sheet.getDataRange().getValues();
    data.splice(0, 1);

    if (data.length > 0) {
        push_sheet.insertRowsAfter(1, data.length);
        push_sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
}

export const matchData = () => {
    const match_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('매칭대기');
    const inspection_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('검수완료');

    const match_data = match_sheet.getDataRange().getValues();
    const inspection_data = inspection_sheet.getDataRange().getValues();

    match_data.splice(0, 1);
    inspection_data.splice(0, 1);

    let complete = new Array();
    let inspection = inspection_data;
    let match = match_data;

    // match_data.map((md) => {
    //     let temp = inspection_data.filter(x => x[1] == md[3] && x[3] == md[10] && x[5] == md[12]);
    //     console.log(temp);
    //     if (temp.length == 1) {
    //         complete.push([...md.slice(0, -1), ...temp[0]]);
    //         inspection = inspection_data.filter(x => x != temp[0]);
    //         match = match_data.filter(x => x != md);
    //         return;
    //     }
    //     if (temp.length < 1) {
    //         md[15] = '검수대기';
    //     }
    //     if (temp.length > 1) {
    //         md[15] = '검수확인필요';
    //     }
    //     return md;
    // });

    inspection.map((i, index) => {
        let temp = match.filter(x => x[3] == i[1] && x[10] == i[3] && x[12] == i[5]);

        if (temp.length < 1) {
            i[14] = '반품접수필요';
            return i;
        }
        if (temp.length > 1) {
            i[14] = '수동매칭필요';
            return i;
        }

        if (temp.length == 1) {
            complete.push([...temp[0].slice(0, -1), ...i]);
            inspection = inspection.filter(x => x != i);
            match = match.filter(x => x != temp[0]);
        }
    });

    match_sheet.deleteRows(2, match_sheet.getLastRow() - 1);
    inspection_sheet.deleteRows(2, inspection_sheet.getLastRow() - 1);

    if (match.length > 0) {
        match_sheet.insertRowsAfter(1, match.length);
        match_sheet.getRange(2, 1, match.length, match[0].length).setValues(match);
    }

    if (inspection.length > 0) {
        inspection_sheet.insertRowsAfter(1, inspection.length);
        inspection_sheet.getRange(2, 1, inspection.length, inspection[0].length).setValues(inspection);
    }

    if (complete.length > 0) {
        const complete_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('최종확인');
        complete_sheet.insertRowsAfter(1, complete.length);
        complete_sheet.getRange(2, 1, complete.length, complete[0].length).setValues(complete);
    }
}

export const pushArchive = () => {
    const last = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('최종확인');
    const archive = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('아카이브대기');
    
    const today = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy. MM. dd');
    const data = last.getDataRange().getValues();

    let count = 0;

    if (data.length > 1) {
        data.forEach((d, i) => {
            if (i == 0) return;
            if (d[29] != '' && d[29] != null) {
                data[i][30] = today;
                count++;
            }
        });
    
        last.getDataRange().setValues(data).sort({column: 31, ascending: false});

        const move = last.getRange(2, 1, count, data[0].length).getValues();
        last.deleteRows(2, count);
        archive.insertRowsAfter(1, count);
        archive.getRange(2, 1, count, data[0].length).setValues(move);
    }
}

export const archiveData = () => {
    const this_year = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy') + '년';
    const folder = DriveApp.getFolderById('19uccVeoDg81X3MA2dgEWuKRT4-9obd8d');
    const files = folder.getFilesByName(this_year);
    let file;

    if (files.hasNext()) {
        file = files.next();
    } else {
        const source = {
            title: this_year,
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [{id: folder.getId()}]
        }
        file = DriveApp.getFileById(Drive.Files.insert(source).id);
        let ss2 = SpreadsheetApp.openById(file.getId()).getSheets()[0];
        const firstrow = [['UID', '접수일', '구분', '수령인명', '전화번호', '주소', '우편번호', '주문번호', '상품주문번호', '셀러명', '상품코드', '상품명', '수량', '문의메모', '반품배송비', '입고일', '발송인명', '발송주소지', '입고상품코드', '입고상품명', '입고상품수량', '사용여부', '시리얼넘버', 'LOT', '검수내용', '결과구분', '검수결과', '검수완료일', '검수메모', '처리결과', '최종처리일자']];
        ss2.getRange(1, 1, 1, firstrow[0].length).setValues(firstrow);
        ss2.deleteRows(2, 999);
    }

    const ss = SpreadsheetApp.openById(file.getId());
    const ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('아카이브대기');
    const data = ss1.getDataRange().getValues();
    data.splice(0, 1);
    if (data.length > 0) {
        ss1.deleteColumns(2, data.length);
        ss.insertRowsAfter(1, data.length);
        ss.getRange(2, 1, data.length, data[0].length).setValues(data);
    };
}