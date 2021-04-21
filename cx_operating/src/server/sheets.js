const getSheets = () => SpreadsheetApp.getActive().getSheets();

const getActiveSheetName = () => SpreadsheetApp.getActive().getSheetName();

export const getSheetsData = () => {
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
};

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

    Logger.log(r);
    return r;
}