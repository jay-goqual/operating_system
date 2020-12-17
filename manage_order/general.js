var ref = get_Ref();

//스프레드시트 열릴시
function Init() {
    //ui 생성
    SpreadsheetApp.getUi()
    .createMenu('출고관리')
    .addItem('주문취합', 'fetch_Order_button')
    .addItem('주문제출', 'submit_Order_button')
    .addItem('출고요청', 'download_general_Order')
    .addSeparator()
    .addItem('커튼출고요청', 'download_curtain_Order')
    .addSeparator()
    .addItem('송장입력', 'fetch_Invcoie_button')
    .addItem('송장추출', 'push_Invoice_button')
    .addItem('송장전달', 'send_Invoice_button')
    .addSeparator()
    .addItem('출고종료', 'delete_Archive')
    .addToUi();
}

async function download_general_Order() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    await download_Order([ss.getSheetByName('굿스코아'), ss.getSheetByName('박스풀')]);
}

async function download_curtain_Order() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    await download_Order([ss.getSheetByName('제이에스비즈'), ss.getSheetByName('건인디앤씨')]);
}

async function fetch_Order_button() {
    await check_Upload();
    await fetch_Order_from_sheet();
}

async function submit_Order_button() {
    await fetch_Additional_info();
    await submit_Order();
}

async function fetch_Invcoie_button() {
    await check_Upload()
}

async function push_Invoice_button() {
    await push_Invoice();
}

async function send_Invoice_button() {
    await send_Invoice();
}

async function delete_Archive() {
    let key = ['다운로드/아카이브', '업로드/아카이브'];
    key.forEach((k) => {
        let folder = DriveApp.getFolderById(ref.get(k));
        let files = folder.getFiles();

        while(files.hasNext()) {
            let file = files.next();
            Drive.Files.remove(file.getId());
        }
    })
}

/* async function add_Trigger() {

    let triggers = ScriptApp.getProjectTriggers().filter((x) => x.getHandlerFunction() == 'check_Upload');

    if (triggers.length == 0) {
        ScriptApp.newTrigger('check_Upload')
        .timeBased()
        .everyMinutes(5)
        .create();
    }

    let d_triggers = ScriptApp.getProjectTriggers().filter((x) => x.getHandlerFunction() == 'delete_Archive');

    if (d_triggers.length == 0) {
        ScriptApp.newTrigger('delete_Archive')
        .timeBased()
        .everyDays(1)
        .atHour(18)
        .nearMinute(45)
        .create();
    }

    let r_triggers = ScriptApp.getProjectTriggers().filter((x) => x.getHandlerFunction() == 'remove_Trigger');

    if (r_triggers.length == 0) {
        ScriptApp.newTrigger('remove_Trigger')
        .timeBased()
        .everyDays(1)
        .atHour(19)
        .nearMinute(45)
        .create();
    }
} */

/* async function remove_Trigger() {
    let triggers = ScriptApp.getProjectTriggers().filter((x) => x.getHandlerFunction() == 'check_Upload' || x.getHandlerFunction() == 'remove_Trigger' || x.getHandlerFunction() == 'delete_Archive');

    triggers.forEach((t) => {
        ScriptApp.deleteTrigger(t);
    });
} */