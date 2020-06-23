function fetch_monthly_data() {
    const file_id = '1UPt4EUjTlmm6DFt7LF5H4ytnT5GkiUaps5BaUTJooJA';
    const month_sheet = SpreadsheetApp.openById(file_id).getSheets();
    const target_sheet = SpreadsheetApp.openById('13pPLRFACeHQD9fuMD0TmTh2z5uABQtV8jq49AQxu-9A').getSheets()[0];
    const date_check = target_sheet.getRange(1, target_sheet.getLastColumn(), target_sheet.getLastRow(), 1).getValues();

    month_sheet.forEach(async (sheet) => {
        //await fetch_daily_data(new Date().getDate());
        const check = [sheet.getName()];
        if (date_check.find(d => JSON.stringify(d) == JSON.stringify(check) ? false : true)) {
            await fetch_daily_data(sheet.getName().split('일')[0], file_id);
        }
    });
}

function fetch_daily_data(date, file_id) {
    const query = "select A, B, C, D, E, F, G, X, Y";
    const gvizURL = `https://docs.google.com/spreadsheets/d/${file_id}/gviz/tq?&tqx=out:csv&headers=1&sheet=${date}일&tq=${encodeURIComponent(query)}`;
    const options = {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}};
    const csv = UrlFetchApp.fetch(gvizURL, options);

    //const sheet = SpreadsheetApp.openById('1UPt4EUjTlmm6DFt7LF5H4ytnT5GkiUaps5BaUTJooJA').getSheetByName(date + '일');
    //const data = sheet.getDataRange().getValues();

    const data = Utilities.parseCsv(csv);
    data.splice(0, 1);

    const target_sheet = SpreadsheetApp.openById('13pPLRFACeHQD9fuMD0TmTh2z5uABQtV8jq49AQxu-9A').getSheets()[0];

    target_sheet.insertRowsAfter(1, data.length);
    target_sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    target_sheet.getRange(2, data[0].length + 1, data.length, 1).setValue(date + '일');
}

function fecth_data() {
    fetch_daily_data(new Date().getDate - 1, '1UPt4EUjTlmm6DFt7LF5H4ytnT5GkiUaps5BaUTJooJA')
}