// UI 메뉴를 생성하고 각 함수에 연결시켜주는 파일입니다.

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
    .addItem('커튼송장입력', 'fetch_curtain_button')
    .addSeparator()
    .addItem('반품/회수 다운로드', 'download_Refund')
    .addSeparator()
    .addItem('송장입력', 'fetch_Invcoie_button')
    .addItem('송장추출', 'push_Invoice_button')
    .addItem('송장전달', 'send_Invoice_button')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('CX팀')
        .addItem('커튼확인완료', 'submit_Check')
        .addItem('커튼주문검색', 'find_order')
        .addItem('주문검색', 'find_order2')
        .addItem('CS처리', 'submit_cs'))
    .addToUi();
}

async function download_general_Order() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    await download_Order([ss.getSheetByName('굿스코아'), ss.getSheetByName('박스풀')]);
}

async function download_curtain_Order() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    await download_Order([ss.getSheetByName('제이에스비즈'), ss.getSheetByName('건인디앤씨'), ss.getSheetByName('드림캐쳐')]);
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
    await check_Upload();
}

async function fetch_curtain_button() {
    await fetch_Invoice_curtain();
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