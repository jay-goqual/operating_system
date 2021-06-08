var order_form = get_Order_form();
var ref = get_Ref();

async function download_Order_new(channel) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    channels.forEach((c) => {
        if (ss.getSheetByName(c).getLastRow <= 1) {
            return;
        }
        
    })
}

async function download_Order(channels) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // const channels = [ss.getSheetByName(channel)];

    const target = SpreadsheetApp.openById(ref.get('출고요청'));
    const former = target.getSheets();
    former.forEach((f, i) => {
        f.setName(i);
    });

    const time = ss.getSheetByName('주문현황').getDataRange().getValues();

    channels.forEach((c) => {
        if (c.getLastRow() <= 1) {
            return;
        }

        let table = c.getDataRange().getValues();

        target.insertSheet().setName(c.getName()).getRange(1, 1, table.length, table[0].length).setNumberFormat('@').setValues(table);

        table.forEach((t) => {
            /* time.forEach((x, i) => {
                if (x[order_form.get('상품주문번호')] == t[order_form.get('상품주문번호')]) {
                    return i;
                }
            }); */

            let temp = time.findIndex((x) => {
                if (c.getName() == '제이에스비즈' || c.getName() == '건인디앤씨' || c.getName() == '드림캐쳐') {
                    return x[order_form.get('상품주문번호')] == t[1];
                } else if (c.getName() == '박스풀') {
                    return x[order_form.get('상품주문번호')] == t[4];
                }else {
                    return x[order_form.get('상품주문번호')] == t[order_form.get('상품주문번호')];
                }
            });

            if (temp && temp != -1) {
                //if ((t.length > 17 && t[17] != '굿스코아/제이에스비즈') || t.length <= 17) {
                    time[temp][order_form.get('출고요청')] = Utilities.formatDate(new Date(), 'GMT+9', 'yy/MM/dd HH:mm');
                //}
            }
        });
    });

    if (former.length == target.getSheets().length) {
        SpreadsheetApp.getUi().alert(`주문이 없습니다.`);
        return;
    }

    former.forEach((f) => {
        target.deleteSheet(f);
    });

    const ex = target.getSheets();
    target.rename(`${Utilities.formatDate(new Date(), 'GMT+9', 'MMddHHmm_출고요청')}`);

    let source = '';
    ex.forEach((x) => {
        const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${target.getId()}\/export?gid=${x.getSheetId()}`;
        const response = UrlFetchApp.fetch(url, {headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`}});
        // const convert_url = DriveApp.getFolderById(ref.get('다운로드/아카이브')).createFile(response.getBlob().setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'MMddHHmm')}_${x.getSheetName()}_출고지시.xlsx`)).getDownloadUrl();
        // source += `<a href="${convert_url}" target="_blank">${x.getSheetName()}<\/a><\/br>`;
        if (x.getSheetName() == '제이에스비즈') {
            MailApp.sendEmail({
                to: 'jj-smart@naver.com',
                cc: 'service@goqual.com',
                subject: `[헤이홈] ${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}일자 주문 내역`,
                htmlBody: '<div dir="ltr">안녕하세요.<br>주식회사 고퀄의 커머스팀입니다.<br><br>금일 주문 접수 건 공유드립니다.<br><br>감사합니다 :)<br>헤이홈 드림.</div>',
                //attachments: [{fileName: x.getName(), content: response.getContent()}]
                attachments: [response.getBlob().setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}_${x.getName()}.xlsx`)]
            });
        } else {
            source += `<a href="${url}" target="_blank">${x.getSheetName()}<\/a><\/br>`;
        }
    });

    const html = HtmlService.createHtmlOutput(source);
    SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');

    ss.getSheetByName('주문현황').getDataRange().setValues(time);

    ss.getSheetByName('주문현황').getRange(2, 1, ss.getSheetByName('주문현황').getLastRow() - 1, ss.getSheetByName('주문현황').getLastColumn() - 1).sort({column: order_form.get('출고요청') + 1, ascending: false});
    
    // const target = SpreadsheetApp.openById(`1xXUAVL3S0NCytOegk52zgLAeiVU89skUyMh_hMElBwg`);
    
    // let now = Utilities.formatDate(new Date(), 'GMT+9', 'MMddHHmm');
    /* target.rename(`${now}_출고지시`);
    target.getSheets().forEach((t, i) => {
        if (i > 0) {
            target.deleteSheet(t);
        } else {
            t.setName('d');
        }
    })

    channel.forEach((c) => {
        c.copyTo(target);
        target.getSheets()[target.getSheets().length - 1].setName(c.getName());
    });

    target.deleteSheet(target.getSheets()[0]); */

    /* const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${ss.getId()}\/export?format=xlsx&contentsonly=true&gid=${channel_1.getSheetId()}&access_token=${ScriptApp.getOAuthToken()}`;
    const html = HtmlService.createHtmlOutput(`<input type="button" value="Download" onClick="location.href='${url}'" >`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'test'); */

    // const res = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}});
    //const sheet = target.getSheets();
    /* let source = '';
    channel.forEach((c) => {
        if (c.getLastRow() > 1) {
            const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${ss.getId()}\/export?gid=${c.getSheetId()}`;
            const response = UrlFetchApp.fetch(url, {headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`}});
            const convert_url = DriveApp.getFolderById(ref.get('다운로드/아카이브')).createFile(response.getBlob().setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'MMddHHmm')}_${c.getSheetName()}_출고지시.xlsx`)).getDownloadUrl();
            source += `<a href="${convert_url}" target="_blank">${c.getSheetName()}<\/a><\/br>`;
        }
    });

    if (source != '') {
        const html = HtmlService.createHtmlOutput(source);
        SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');
    } else {
        SpreadsheetApp.getUi().alert(`주문이 없습니다.`);
    }

    if (channel[0].getLastRow() + channel[1].getLastRow() - 2 > 0) {
        ss.getSheetByName('주문현황').getRange(2, order_form.get('출고지시') + 1, channel[0].getLastRow() + channel[1].getLastRow() - 2, 1).setValue(Utilities.formatDate(new Date(), 'GMT+9', 'HH:mm'));
    }

    channel.forEach((c) => {
        if (c.getLastRow() > 1) {
            c.deleteRows(2, c.getLastRow() - 1);
        }
    }) */
}