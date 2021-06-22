// 업무가 종료된 후 데이터를 아카이브하는 함수가 저장된 파일입니다.

// 전역함수 불러오기
var ref = get_Ref();
var order_form = get_Order_form();

// 데이터 아카이브 함수이며, 매일 7~8시 사이에 실행되도록 트리거 걸려있습니다.
async function archive_Data() {
    
    // [주문현황] 시트에서 데이터 불러오기
    const data_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const data = data_sheet.getDataRange().getValues();
    
    // 최상단 제목열 삭제
    data.splice(0, 1);

    // 데이터가 없는 경우 리턴
    if (data.length == 0) {
        return;
    }

    // [출고DB] 폴더에서 해당월 파일 찾기
    let month_file = DriveApp.getFolderById(ref.get('출고DB')).getFilesByName(Utilities.formatDate(new Date(), 'GMT+9', 'yy년 MM월'));
    if (!month_file.hasNext()) {
        // 월별 파일이 없는 경우 새로 생성하기
        const source = {
            title: Utilities.formatDate(new Date(), 'GMT+9', 'yy년 MM월'),
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [{id: ref.get('출고DB')}]
        }
        month_file = DriveApp.getFileById(Drive.Files.insert(source).id);
    } else {
        month_file = month_file.next();
    }

    // 월별파일 열기
    const month_target = SpreadsheetApp.openById(month_file.getId());
    
    // [주문현황] 시트를 월별파일로 복사한 후 이름을 d일로 변경
    let new_sheet = data_sheet.copyTo(month_target);
    new_sheet.setName(`${Utilities.formatDate(new Date(), 'GMT+9', 'd')}일`);

    // 파일 신규생성되어 첫 번째 시트가 신규 빈 시트일 경우 삭제처리
    let check_sheet = month_target.getSheets()[0];
    if (check_sheet.getName().indexOf('일') == -1) {
        month_target.deleteSheet(check_sheet);
    }

    // [이번달DB] 스프레드시트로 데이터 복사
    const manage_sale = SpreadsheetApp.openById(ref.get('이번달DB')).getSheetByName('주문');
    // total = [판매현황] 스프레드시트 - [주문DB] 시트
    const total = SpreadsheetApp.openById('1VM1iKCp9RkiktD_4CXfVkENmA1GLyM66de6OAt9-sg0').getSheetByName('주문DB');
    const sale_form = ['접수일', '셀러명', '셀러코드', '주문번호', '상품주문번호', '상품코드', '수량', '결제일', '판매액', '배송비', '수수료', '송장번호', '출고일시', '주문자', '수령인', '수령인연락처'];

    // [주문현황] 시트 중 출고일시 열에 데이터가 있는 행만 추출
    let push_table = new Array();
    let count = 0;
    data.forEach((d, i) => {
        if (d[order_form.get('출고일시')]) {
            push_table[count] = new Array();
            sale_form.forEach((f) => {
                push_table[count].push(d[order_form.get(f)]);
            });
            count++;
        }
    });

    // 추출한 데이터의 수가 1 이상일 경우, [이번달DB]-[주문] 시트로 데이터 복사
    if (push_table.length > 0) {
        manage_sale.insertRowsAfter(1, push_table.length);
        manage_sale.getRange(2, 1, push_table.length, sale_form.length).setValues(push_table);

        // [판매현황] 시트에 데이터 복사
        total.insertRowsAfter(1, push_table.length);
        total.getRange(2, 1, push_table.length, sale_form.length).setValues(push_table);

        // [주문현황] 시트를 출고일시 기준으로 sorting 한 후, 복사한 데이터의 수만큼 행 삭제
        data_sheet.sort(24, false);
        data_sheet.deleteRows(2, push_table.length);
    }

    // [발주체크] 시트 초기화
    // const c_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    // c_sheet.getRange(2, 2, c_sheet.getLastRow() - 1, 2).setValue(0);
    // c_sheet.getRange(2, 6, 4, 1).setValue(0);
}