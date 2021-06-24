// 송장번호 파일에서 송장번호를 긁어오는 파일입니다.

// 전역함수 로드
var ref = get_Ref();
var order_form = get_Order_form();

// 파일에서 송장번호 데이터만 추출하여 [주문현황] 시트로 복사하는 함수
async function fetch_Invoice(file) {

    // 송장파일 데이터 추출
    const invoice_data = SpreadsheetApp.openById(file.getId()).getDataRange().getValues();

    // [주문현황] 시트 데이터 추출
    const order_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();

    // [발주체크] 시트 불러오기
    // const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    // const check_data = check_sheet.getDataRange().getValues();

    // 송장파일의 양식 생성하기 (invoice_form)
    let invoice_form = new Map();
    invoice_data[0].forEach((d, i) => {
        invoice_form.set(d, i);
    });

    //  check = 출고처 확인, 1 = 굿스코아, 2 = 박스풀
    let check = 0;

    // 송장파일 최상위 열 제거
    invoice_data.splice(0, 1);

    // [주문현황] 시트 데이터 전체 검수
    order_data.map((o) => {
        // 송장번호 기존재 시 리턴
        if (o[order_form.get('송장번호')]) {
            return o;
        }

        let find;
        // 송장파일에 [출고상태] 열이 존재할 경우, 굿스코아 송장으로 인식
        if (invoice_form.get('출고상태')) {
            // 굿스코아 송장파일의 주문번호와 [주문현황] 상품주문번호가 동일한 데이터 검색 후 find 저장
            find = invoice_data.filter(d => (d[invoice_form.get('주문번호')] == o[order_form.get('상품주문번호')]));
            check = 1;
        } else {
            // 박스풀 송장파일의 주문번호와 [주문현황] 주문번호가 동일한 데이터 검색 후 find 저장
            find = invoice_data.filter(d => (d[invoice_form.get('Invoice Number')] == o[order_form.get('주문번호')]));
            check = 2;
        }

        // 동일 주문번호(상품주문번호) 존재 할 경우
        if (find.length > 0) {
            // [주문현황] 데이터에 송장파일의 송장 입력
            if (invoice_form.get('송장번호')) {
                o[order_form.get('송장번호')] = find[0][invoice_form.get('송장번호')];
            } else {
                o[order_form.get('송장번호')] = find[0][invoice_form.get('Tracking Code')];
            }
            
            // 입력 택배사와 출력택배사 다를 경우 조정
            if (o[order_form.get('택배사')] == '2 - 롯데택배') {
                o[order_form.get('택배사')] = '롯데택배'
            }
            if (o[order_form.get('택배사')] == '3 - CJ택배') {
                o[order_form.get('택배사')] = 'CJ대한통운'
            }
        }
        return o;
    });

    // [주문현황] 시트 불러오기
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황')
    // 송장번호가 존재하는 데이터로 덮어쓰기
    sheet.getDataRange().setValues(order_data);

    // [발주체크] 시트의 출고처별 출고횟수 반영
    // if (check == 1) {
        // check_data[1][5]++;
    // } else if (check == 2) {
        // check_data[2][5]++;
    // }
    // check_sheet.getDataRange().setValues(check_data);

    // 업로드 파일 아카이브하기
    DriveApp.getFolderById(ref.get('업로드/아카이브')).addFile(file);
    file.getParents().next().removeFile(file);
}

// 커튼 송장을 처리하는 함수입니다.
async function fetch_Invoice_curtain() {
    
    // 공급사 대시보드 - [송장번호 전달] 시트 데이터 불러오기
    const dashboard = SpreadsheetApp.openById(ref.get('제이에스비즈'));
    const invoice_data = dashboard.getSheetByName('송장번호 전달').getDataRange().getValues();
    // [주문현황] 데이터 불러오기
    const order_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황').getDataRange().getValues();

    // 최상위 제목열 제거
    invoice_data.splice(0, 1);

    // 주문번호 동일한 송장번호를 찾아서 [주문현황] 시트에 입력
    order_data.map((o) => {
        if (o[order_form.get('송장번호')]) {
            return o;
        }
        let find = invoice_data.filter(d => d[0] == o[order_form.get('주문번호')]);
        if (find.length > 0 && o[order_form.get('출고채널')] == '제이에스비즈') {
            o[order_form.get('송장번호')] = find[0][5];
        }
        return o;
    });

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황')
    sheet.getDataRange().setValues(order_data);

    // 공급사 대시보드의 송장번호 전달 시트에서 입력완료된 내역 지우기
    dashboard.getSheetByName('송장번호 전달').getRange(2, 6, dashboard.getSheetByName('송장번호 전달').getLastRow() - 1, 1).clear();
    dashboard.getSheetByName('송장번호 전달').getRange(2, 6, dashboard.getSheetByName('송장번호 전달').getLastRow() - 1, 1).clearFormat();

    // 발주체크 출고횟수 반영
    // const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    // const check_data = check_sheet.getDataRange().getValues();

    check_data[3][5]++;
    check_sheet.getDataRange().setValues(check_data);
}