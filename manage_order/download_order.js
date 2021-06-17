// 주문현황의 데이터를 각 출고처별로 분류하여 다운로드 받는 함수들이 있는 파일입니다.

// 전역함수 로드
var order_form = get_Order_form();
var ref = get_Ref();

// 일반 3PL 업체의 주문건을 다운로드하는 함수입니다.
// channels = [출고관리] 시트 내의 각 출고처별 시트
async function download_Order(channels) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // [출고요청] 스프레드시트 오픈
    const target = SpreadsheetApp.openById(ref.get('출고요청'));
    const former = target.getSheets();
    // 기존에 있는 시트의 이름을 사본으로 변경
    former.forEach((f, i) => {
        f.setName(`${f.getSheetName()}_사본`);
    });

    // [주문현황] 시트의 데이터 오픈
    const time = ss.getSheetByName('주문현황').getDataRange().getValues();

    // [출고요청] 스프레드시트에 각 출고처별 시트 생성 및 데이터 복제
    channels.forEach((c) => {
        
        // 출고처 주문이 없을 경우 리턴
        if (c.getLastRow() <= 1) {
            return;
        }

        // 출고처 주문 데이터를 오픈
        let table = c.getDataRange().getValues();

        // [출고요청] 스프레드시트에 신규 시트 생성, 이름 변경 (출고처명), 데이터 복사
        target.insertSheet().setName(c.getName()).getRange(1, 1, table.length, table[0].length).setNumberFormat('@').setValues(table);

        table.forEach((t) => {

            // [주문현황] 시트에 출고요청 일자를 기록하는 과정

            // [주문현황] 시트의 데이터와 [출고요청]에 새로 생성된 시트의 데이터 중 상품주문번호가 같은 index 값을 temp로 반환
            let temp = time.findIndex((x) => {
                // 출고처가 커튼일 경우 상품주문번호 열 조정
                if (c.getName() == '제이에스비즈' || c.getName() == '건인디앤씨' || c.getName() == '드림캐쳐') {
                    return x[order_form.get('상품주문번호')] == t[1];
                } 
                // 출고처가 박스풀일 경우 상품주문번호 열 조정
                else if (c.getName() == '박스풀') {
                    return x[order_form.get('상품주문번호')] == t[4];
                } else {
                    return x[order_form.get('상품주문번호')] == t[order_form.get('상품주문번호')];
                }
            });

            // temp 값이 존재 할 경우
            if (temp && temp != -1) {

                // [주문현황] 시트의 index(temp) 행에 현재 시간값을 작성
                time[temp][order_form.get('출고요청')] = Utilities.formatDate(new Date(), 'GMT+9', 'yy/MM/dd HH:mm');
            }
        });
    });

    // 출고처 주문이 없었을 경우 리턴
    if (former.length == target.getSheets().length) {
        SpreadsheetApp.getUi().alert(`주문이 없습니다.`);
        return;
    }

    // 기존 시트를 삭제
    former.forEach((f) => {
        target.deleteSheet(f);
    });

    // [출고요청] 스프레드시트의 각 시트를 다운로드하는 링크 생성
    const ex = target.getSheets();
    target.rename(`${Utilities.formatDate(new Date(), 'GMT+9', 'MMddHHmm_출고요청')}`);

    // source = html 코드
    let source = '';

    // 각 시트별로 접근하여 링크 코드 생성 후 source에 추가
    ex.forEach((x) => {

        // url = 다운로드 링크, response = urlfetch 값
        const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${target.getId()}\/export?gid=${x.getSheetId()}`;
        const response = UrlFetchApp.fetch(url, {headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`}});
        
        // 출고처가 제이에스비즈일 경우 메일발송
        if (x.getSheetName() == '제이에스비즈') {
            MailApp.sendEmail({
                to: 'jj-smart@naver.com',
                cc: 'service@goqual.com',
                subject: `[헤이홈] ${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}일자 주문 내역`,
                htmlBody: '<div dir="ltr">안녕하세요.<br>주식회사 고퀄의 커머스팀입니다.<br><br>금일 주문 접수 건 공유드립니다.<br><br>감사합니다 :)<br>헤이홈 드림.</div>',
                //attachments: [{fileName: x.getName(), content: response.getContent()}]
                attachments: [response.getBlob().setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}_${x.getName()}.xlsx`)]
            });
        }
        // 아닐 경우에는 링크 생성후 source에 추가
        else {
            source += `<a href="${url}" target="_blank">${x.getSheetName()}<\/a><\/br>`;
        }
    });

    // source 코드를 UI에 출력
    const html = HtmlService.createHtmlOutput(source);
    SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');

    // 정리한 출고요청 시간을 [주문현황] 시트에 기입 후, 해당 열을 기준으로 sorting
    ss.getSheetByName('주문현황').getDataRange().setValues(time);
    ss.getSheetByName('주문현황').getRange(2, 1, ss.getSheetByName('주문현황').getLastRow() - 1, ss.getSheetByName('주문현황').getLastColumn() - 1).sort({column: order_form.get('출고요청') + 1, ascending: false});
}


// CX operating (CX_대시보드) 개발진행 중인 내역
// CX_대시보드에 생성되어 있는 회수, 출고 파일을 다운로드하는 함수
function download_Refund() {
    const out = ['굿스코아_회수', '박스풀_회수', '굿스코아_출고', '박스풀_출고'];

    let data = {'굿스코아_회수': [], '박스풀_회수': [], '굿스코아_출고': [], '박스풀_출고': []};
    let source = '';

    let orig = SpreadsheetApp.openById('1YZyzu3Awg-qiIULVt4VoOC2lBGwcxNJtkbkWQ62LrXc').getSheetByName('회수필요');
    let orig_data = orig.getDataRange().getValues();

    orig_data.forEach((d) => {
        if (d[7] == '굿스코아' || d[7] == '출고채널') {
            data['굿스코아_회수'].push(d);
        }
        if (d[7] == '박스풀' || d[7] == '출고채널') {
            data['박스풀_회수'].push(d);
        }
    });

    let orig2 = SpreadsheetApp.openById('1YZyzu3Awg-qiIULVt4VoOC2lBGwcxNJtkbkWQ62LrXc').getSheetByName('출고필요');
    let orig_data2 = orig2.getDataRange().getValues();

    orig_data2.forEach((d) => {
        if (d[7] == '굿스코아' || d[7] == '출고채널') {
            data['굿스코아_출고'].push(d);
        }
        if (d[7] == '박스풀' || d[7] == '출고채널') {
            data['박스풀_출고'].push(d);
        }
    });

    out.forEach((o) => {
        const sheet = SpreadsheetApp.openById('1YZyzu3Awg-qiIULVt4VoOC2lBGwcxNJtkbkWQ62LrXc').getSheetByName(`${o}`);
        sheet.clear();
        if (data[o].length > 0) {
            sheet.getRange(1, 1, data[o].length, data[o][0].length).setValues(data[o]);
        }
    });

    out.forEach((o) => {
        const sheet = SpreadsheetApp.openById('1YZyzu3Awg-qiIULVt4VoOC2lBGwcxNJtkbkWQ62LrXc').getSheetByName(`${o}`);
        if (sheet.getLastRow() <= 1) {
            return;
        }
        const ssid = SpreadsheetApp.openById('1YZyzu3Awg-qiIULVt4VoOC2lBGwcxNJtkbkWQ62LrXc').getId();
        const sid = sheet.getSheetId();

        const url = `https:\/\/docs.google.com\/spreadsheets\/d\/${ssid}\/export?gid=${sid}`;
        source += `<a href="${url}" target="_blank">${o}<\/a><\/br>`;
    });

    if (source.length > 0) {
        const html = HtmlService.createHtmlOutput(source);
        SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');
        orig.deleteRows(2, orig.getLastRow() - 1);
        orig2.deleteRows(2, orig2.getLastRow() - 1);
    }
}