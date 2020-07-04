var ref = get_Ref();
var client = get_Client();
var invoice_form = get_Invoice_form();
var order_form = get_Order_form();

async function push_Invoice() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();
    const target = SpreadsheetApp.openById(ref.get('송장저장'));
    target.getSheets().forEach((t, i) => {
        if (i > 0) {
            target.deleteSheet(t);
        }
    })

    const push_table = new Map();
    table.map((t) => {
        if (!t[order_form.get('송장번호')] || t[order_form.get('출고일시')]) {
            return t;
        }

        if (!push_table.get(t[order_form.get('셀러명')])) {
            push_table.set(t[order_form.get('셀러명')], new Array());
        }
        let temp = push_table.get(t[order_form.get('셀러명')]);
        
        let x = new Array();
        invoice_form.get(client.get(t[order_form.get('셀러명')]).get('업로드양식')).forEach((f) => {
            if (t[order_form.get(f)]) {
                x.push(t[order_form.get(f)]);
            } else {
                x.push(f);
            }
        })

        temp.push(x);
        push_table.set(t[order_form.get('셀러명')], temp);

        t[order_form.get('출고일시')] = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm');

        return t;
    });

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().setValues(table);

    target.rename(`${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}_출고완료`);

    push_table.forEach(async (t, k) => {
        let row = 2;
        if (!target.getSheetByName(k)) {
            target.insertSheet().setName(k);
            target.getSheetByName(k).getRange(1, 1, 1, t[0].length).setValues([invoice_form.get(client.get(k).get('업로드양식'))]);
        }

        if (k == '공식몰') {row = 3;}
        if (k == '엔분의일') {
            target.getSheetByName(k).getRange(1, 2).setValue('배송방법');
        }

        if (row < target.getSheetByName(k).getLastRow()) {
            target.getSheetByName(k).deleteRows(row, target.getSheetByName(k).getLastRow() - row + 1);
        }
        
        target.getSheetByName(k).insertRowsAfter(row - 1, t.length);
        target.getSheetByName(k).getRange(row, 1, t.length, t[0].length).setNumberFormat('@').setValues(t);

    });
}

async function send_Invoice() {
    const ss = SpreadsheetApp.openById(ref.get('송장저장'));
    const channel = ss.getSheets();

    let source = new String();

    channel.forEach((c) => {
        const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${ss.getId()}\/export?gid=${c.getSheetId()}`;
        const response = UrlFetchApp.fetch(url, {headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`}});
        let x = DriveApp.getFolderById(ref.get('다운로드/아카이브')).createFile(response.getBlob().setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}_출고완료_${c.getName()}.xlsx`));

        if (client.get(c.getName()).has('출고이메일')) {
            MailApp.sendEmail({
                to: client.get(c.getName()).get('출고이메일'),
                cc: 'service@goqual.com',
                subject: `[헤이홈] ${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}일자 운송장 정보`,
                htmlBody: '<div dir="ltr">안녕하세요.<br>주식회사 고퀄의 커머스팀입니다.<br><br>금일 주문 건에 대한 운송장 정보 전달드립니다.<br><br>감사합니다 :)</div>',
                //attachments: [{fileName: x.getName(), content: response.getContent()}]
                attachments: [x]
            });
        } else {
            source += `<a href="${url}" target="_blank">${c.getName()}<\/a><\/br>`;
            if (c.getName() == '엔분의일') {
                c.setName('발송처리');
            }
        }
    })

    if (source.length > 0) {
        const html = HtmlService.createHtmlOutput(source);
        SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');
    }
}