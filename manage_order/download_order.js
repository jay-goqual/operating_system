var order_form = get_Order_form();

async function download_order() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const channel = [ss.getSheetByName('굿스코아'), ss.getSheetByName('제이에스비즈')];
    
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
    let source = new String();
    channel.forEach((c) => {
        let url = `https:\/\/docs.google.com\/spreadsheets\/d\/${ss.getId()}\/export?gid=${c.getSheetId()}&access_token=${ScriptApp.getOAuthToken()}`;
        source += `<a href="${url}" target="_blank">${c.getSheetName()}<\/a><\/br>`;
    });
    const html = HtmlService.createHtmlOutput(source);

    SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');

    ss.getSheetByName('주문현황').getRange(2, order_form.get('출고지시'), ss.getSheetByName('주문현황').getLastRow(), 1).setValue(true);
}