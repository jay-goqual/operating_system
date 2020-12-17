function myFunction() {
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data').getDataRange().getValues();
    let out = new Array();
    let sht = new Array();

    out.push(data[0])
    data.forEach((d) => {
        if (out.filter(order => order[0] == d[0]).length == 0) {
            out.push(d);
        }
    })

    out.forEach((o, i) => {
        out[i][57] = false;
        if (out.filter(order => (order[11] == o[11]) && (order[48] == o[48]) && (o[13] != '입금대기')).length > 1) {
            out[i][57] = true;
        }
    })

    out.forEach((o, i) => {
        sht[i] = Array();
        sht[i][0] = o[0];
        sht[i][1] = o[14];
        sht[i][2] = o[11];
        sht[i][3] = o[48];
        sht[i][4] = o[35];
        sht[i][5] = o[57];
    })

    SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getRange(1, 1, sht.length, 6).setValues(sht);
}

function my2() {
    const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시트1').getDataRange().getValues();
    let out = new Array();
    let sht = new Array();

    for (i = 0; i < data.length; i++) {
        data[i][6] = 1;
        data[i][7] = 0;
        if (data[i][5] == 'true') {
            for (j = i - 1; j > 0; j--) {
                if (data[i][2] == data[j][2] && data[i][3] == data[j][3]) {
                    data[i][6] = data[j][6] + 1;
                    data[i][7] = parseInt((new Date(data[j][1]) - new Date(data[i][1])) / (1000 * 60 * 60 * 24), 10);
                    break;
                }
            }
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시트1').getRange(1, 1, data.length, 8).setValues(data);
}