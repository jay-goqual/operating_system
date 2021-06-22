// 출고가 완료된 이후 송장번호가 작성된 내역을 각 판매처별로 메일링/다운로드 하는 파일입니다.


// 전역함수 호출
var ref = get_Ref();
var client = get_Client();
var invoice_form = get_Invoice_form();
var order_form = get_Order_form();

// 메일링 및 다운로드 전, [출고완료] 스프레드시트에 각 판매처별 시트를 생성하고 데이터를 복사하는 함수입니다.
async function push_Invoice() {
    // [주문현황] 시트 불러오기
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();

    // [출고완료] 시트 불러오기
    const target = SpreadsheetApp.openById(ref.get('송장저장'));
    // 첫번째 시트를 제외하고 모든 시트 삭제하기
    target.getSheets().forEach((t, i) => {
        if (i > 0) {
            target.deleteSheet(t);
        }
    });

    // [주문현황]의 모든 데이터에 접근
    const push_table = new Map();
    table.map((t) => {
        // 송장번호가 없거나 출고일시는 있는 경우 리턴
        if (!t[order_form.get('송장번호')] || t[order_form.get('출고일시')]) {
            return t;
        }

        // [셀러관리] 시트에 등록된 셀러명이 아닌경우, 리턴
        if (!client.has(t[order_form.get('셀러명')])) {
            return t;
        }

        // [출고완료] 스프레드시트에 셀러명 시트가 없을 경우 새로 생성
        // push_table에 셀러명 배열이 없는 경우 새로 생성
        if (!push_table.get(t[order_form.get('셀러명')])) {
            push_table.set(t[order_form.get('셀러명')], new Array());
        }
        let temp = push_table.get(t[order_form.get('셀러명')]);
        
        // 셀러별 송장전달 양식을 불러온 후, [주문현황] 양식에 있는 데이터일 경우에는 값 복사, 아닐 경우에는 최상위열 이름을 일괄 복사
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

        // 출고일시 작성하기
        t[order_form.get('출고일시')] = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm');

        return t;
    });

    // 작성된 출고일시 기입
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().setValues(table);

    target.rename(`${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}_출고완료`);

    // 셀러별로 분류된 테이블을 [출고완료] 스프레드시트로 복사
    push_table.forEach(async (t, k) => {
        let row = 2;
        // 셀러명으로 된 시트가 없다면 새로 생성 후, 데이터 복사
        if (!target.getSheetByName(k)) {
            target.insertSheet().setName(k);
            // 최상위 제목열은 업로드양식 참고하여 작성
            target.getSheetByName(k).getRange(1, 1, 1, t[0].length).setValues([invoice_form.get(client.get(k).get('업로드양식'))]);
        }

        // 셀러가 천삼백케이일 경우 데이터 특별처리
        if (k == '천삼백케이') {
            t.map((x) => {
                x[1] = x[1].split(`${x[0]}-`).join('');
                return x;
            })
        }

        if (k == '공식몰') {row = 3;}

        // 셀러가 스마트스토어 또는 샵플렛일 경우, 데이터 특별처리
        if (k == '스마트스토어' || k == '샵플렛') {
            target.getSheetByName(k).getRange(1, 2).setValue('배송방법');
        }

        // 기본 설정된 row 값보다 데이터가 많다면 row 값 이후의 모든 행 삭제
        if (row < target.getSheetByName(k).getLastRow()) {
            target.getSheetByName(k).deleteRows(row, target.getSheetByName(k).getLastRow() - row + 1);
        }
        
        // 데이터 복사
        target.getSheetByName(k).insertRowsAfter(row - 1, t.length);
        target.getSheetByName(k).getRange(row, 1, t.length, t[0].length).setNumberFormat('@').setValues(t);

        // 셀러가 카카오스토어일 경우, 데이터 특별처리
        if (k == '카카오스토어') {
            target.getSheetByName(k).getRange(1, 3).setValue('택배사코드');
            target.getSheetByName(k).getRange(1, 5).setValue('수령인명');
        }
    });
}


// [출고완료] 스프레드시트의 각 셀러 시트를 메일링하거나 다운로드 링크 생성
async function send_Invoice() {

    // [출고완료] 스프레드시트 불러오기
    const ss = SpreadsheetApp.openById(ref.get('송장저장'));
    const channel = ss.getSheets();

    // [발주체크] 시트 불러오기
    // const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    // const check_data = check_sheet.getDataRange().getValues();

    // source = html 코드
    let source = new String();

    // [출고완료] 스프레드시트의 각 시트 개별로 불러오기
    channel.forEach((c) => {
        // name = 셀러명
        let name = c.getName();

        // 아카이브용 스프레드시트 생성하기
        const new_sheet = SpreadsheetApp.create(`${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}_출고완료_${name}`);

        // 셀러명 시트를 새로운 스프레드시트로 복사하기
        c.copyTo(new_sheet);
        // 첫번째 더미 시트 삭제
        new_sheet.deleteSheet(new_sheet.getSheets()[0]);

        // 스마트스토어거나 샵플렛일 경우 시트명을 '발송처리'로, 아닐 경우에는 셀러명으로
        if (name == '스마트스토어' || name == '샵플렛') {
            new_sheet.getSheets()[0].setName('발송처리');
        } else {
            new_sheet.getSheets()[0].setName(name);
        }

        // 발주체크 횟수 반영
        // check_data[check_data.findIndex((v) => v[0] == name)][2]++;
        
        // 신규생성된 스프레드시트 아카이브 폴더 이동
        DriveApp.getFolderById(ref.get('다운로드/아카이브')).addFile(DriveApp.getFileById(new_sheet.getId()));
        DriveApp.getRootFolder().removeFile(DriveApp.getFileById(new_sheet.getId()));
        
        // 다운로드 url 생성
        url = `https:\/\/docs.google.com\/spreadsheets\/d\/${new_sheet.getId()}\/export?format=xlsx`;
        const response = UrlFetchApp.fetch(url, {headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`}});

        // [셀러관리] 스프레드시트에 출고이메일이 등록되어 있을 경우에는 메일로 발송, 아니라면 다운로드링크 생성
        if (client.get(name).get('출고이메일')) {
            MailApp.sendEmail({
                to: client.get(name).get('출고이메일'),
                cc: 'service@goqual.com',
                subject: `[헤이홈] ${Utilities.formatDate(new Date(), 'GMT+9', 'yyMMdd')}일자 운송장 정보`,
                htmlBody: '<div dir="ltr">안녕하세요.<br>주식회사 고퀄의 커머스팀입니다.<br><br>금일 주문 건에 대한 운송장 정보 전달드립니다.<br><br>감사합니다 :)</div>',
                attachments: [response.getBlob().setName(`${new_sheet.getName()}.xlsx`)]
            });
        } else {
            // 천삼백케이일 경우, csv 파일로 출력
            if (name == '천삼백케이') {
                url = `https:\/\/docs.google.com\/spreadsheets\/d\/${new_sheet.getId()}\/export?format=csv`;
            }
            source += `<a href="${url}" target="_blank">${name}<\/a><\/br>`;
        }
    })

    // 발주체크 반영
    // check_sheet.getDataRange().setValues(check_data);

    // 다운로드해야할 파일이 있을 경우, UI 생성
    if (source.length > 0) {
        const html = HtmlService.createHtmlOutput(source);
        SpreadsheetApp.getUi().showModalDialog(html, '오른쪽 클릭 후 [새 탭에서 열기] 클릭');
    }
}