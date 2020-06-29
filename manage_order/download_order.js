async function download_order() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const channel_1 = ss.getSheetByName('굿스코아');
    const target = SpreadsheetApp.openById(`1xXUAVL3S0NCytOegk52zgLAeiVU89skUyMh_hMElBwg`);
    
    channel_1.copyTo(target, {contentsOnly: true});
    let now = Utilities.formatDate(new Date(), 'GMT+9', 'MMddHHmm');
    target.rename(`${now}_굿스코아_출고지시`);
    target.deleteSheet(target.getSheets()[0]);
    target.getSheets()[0].rename(`출고지시내역`);

    /* const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${ss.getId()}\/export?format=xlsx&contentsonly=true&gid=${channel_1.getSheetId()}&access_token=${ScriptApp.getOAuthToken()}`;
    const html = HtmlService.createHtmlOutput(`<input type="button" value="Download" onClick="location.href='${url}'" >`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'test'); */


    // const res = UrlFetchApp.fetch(url, {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}});
    const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${target.getId()}\/export?access_token=${ScriptApp.getOAuthToken()}`;
    const html = HtmlService.createHtmlOutput(`<a href="${url}" target="_blank">오른쪽 클릭 후 [새 탭에서 열기] 클릭<\/a>`);
    SpreadsheetApp.getUi().showModalDialog(html, '다운로드');
}

function saveAsCSV() {
    const url = SpreadsheetApp.openById(`1xXUAVL3S0NCytOegk52zgLAeiVU89skUyMh_hMElBwg`).getDownloadUrl().replace('?e=download&gd=true', '');
    return url;
}