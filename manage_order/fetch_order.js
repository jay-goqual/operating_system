// 스프레드시트로 변환된 파일의 데이터를 [출고관리] - [에러확인] 시트로 가져오는 파일입니다.

// 전역함수 호출
var fetch_form = get_Fetch_form();
var ref = get_Ref();
var order_form = get_Order_form();
var client = get_Client();

// 인수로 전달받은 파일의 데이터를 주문현황/에러확인으로 데이터 옮김
async function fetch_Order(file) {

    // [에러확인] 시트 불러오기, [발주체크] 시트 불러오기
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    const check_data = check_sheet.getDataRange().getValues();

    // 인수로 전달받은 파일 이름 split
    let separator = file.getName().split('_');

    if (separator[1] == '로켓배송' || separator[0] == 'PO') {
        // 로켓배송 자료일 경우에는 쿠팡 데이터 불러오기 함수 실행
        input_data = await fetch_coupang_Data(file.getId());

        // 불러온 데이터를 [에러확인] 시트에 입력
        target_sheet.insertRowsAfter(1, input_data.length);
        target_sheet.getRange(2, 1, input_data.length, input_data[0].length).setNumberFormat('@').setValues(input_data);

        // 인수 파일 아카이브
        DriveApp.getFolderById(ref.get('업로드/아카이브')).addFile(file);
        file.getParents().next().removeFile(file);

        // 발주체크 횟수 반영
        check_data[check_data.findIndex((v) => v[0] == '쿠팡로켓배송')][1]++;
        check_sheet.getDataRange().setValues(check_data);
        return;
    }

    // _ split 2번째 값(셀러명)으로 셀러 정보 로드
    let client_info = client.get(separator[1]);
    // 업로드(전달) 양식명 불러오기
    let client_form = client_info.get('업로드양식');

    // client_form이 처리가능한 양식일 경우
    if (fetch_form.get(client_form)) {
        
        // 스마트스토어 파일일 경우 최상위 1열 추가 삭제
        if (client_form == '스마트스토어') {
            if (SpreadsheetApp.openById(file.getId()).getSheets()[0].getRange(1, 1).getValue().length > 10) {
                SpreadsheetApp.openById(file.getId()).getSheets()[0].deleteRow(1);
            }
        }
        // 천삼백케이 파일일 경우 최상위 1열 추가 삭제
        if (client_form == '천삼백케이') {
            if (SpreadsheetApp.openById(file.getId()).getSheets()[0].getRange(1, 1).getValue() == '') {
                SpreadsheetApp.openById(file.getId()).getSheets()[0].deleteRow(1);
            }
        }
        // 파일 내의 모든 셀 양식을 [TEXT]로 설정
        await set_Format(file.getId(), '@');

        // 파일에서 데이터 불러오기
        let input_data = await fetch_Data(file.getId(), fetch_form.get(client_form), client_info);

        // [에러확인] 시트에 불러온 데이터 입력
        target_sheet.insertRowsAfter(1, input_data.length);
        target_sheet.getRange(2, 1, input_data.length, input_data[0].length).setNumberFormat('@').setValues(input_data);

        //아카이브로 보내버리기
        DriveApp.getFolderById(ref.get('업로드/아카이브')).addFile(file);
        file.getParents().next().removeFile(file);

        // [발주체크] 시트 횟수 반영
        check_data[check_data.findIndex((v) => v[0] == separator[1])][1]++;
        check_sheet.getDataRange().setValues(check_data);
    }
}


// 쿠팡 데이터 불러오기 함수
async function fetch_coupang_Data(file_id) {
    // 파일 데이터 불러오기
    const data = SpreadsheetApp.openById(file_id).getDataRange().getValues();
    let input_data = new Array();
    
    // 쿠팡 로켓배송 양식에 맞춰 데이터 가져오기
    for (i = 0; i < (data.length - 32) / 2; i++) {
        input_data[i] = new Array();
        input_data[i][order_form.get('셀러명')] = '쿠팡로켓배송';
        input_data[i][order_form.get('주문번호')] = data[9][2];
        input_data[i][order_form.get('상품주문번호')] = `${data[9][2]}-${Utilities.formatString('%02d', i + 1)}`;
        input_data[i][order_form.get('상품코드')] = data[21 + (i * 2)][2];
        input_data[i][order_form.get('수량')] = data[21 + (i * 2)][7];
        input_data[i][order_form.get('주문자')] = `로켓배송_${data[12][2]}`;
        input_data[i][order_form.get('주문자연락처')] = data[13][6];
        input_data[i][order_form.get('수령인')] = `로켓배송_${data[12][2]}`;
        input_data[i][order_form.get('수령인연락처')] = data[12][8];
        input_data[i][order_form.get('주소')] = `${data[12][3]}_${data[9][2]}`;
    }

    // 불러온 데이터 리턴
    return input_data;
}

// 대시보드를 사용하는 업체의 주문 데이터 가져오기
async function fetch_Order_from_sheet() {
    // [에러확인] 시트 불러오기
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');

    // [발주체크] 시트 불러오기
    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('발주체크');
    const check_data = check_sheet.getDataRange().getValues();
    
    // 모든 판매자 데이터 검색
    client.forEach((c, k) => {

        // 첫번째 열일 경우 리턴
        if (k == '셀러명') {
            return;
        }

        // [셀러관리] 스프레드시트의 셀러 행에 [대시보드ID] 값이 존재할 경우
        if (c.get('대시보드ID')) {

            // [대시보드ID] 값을 통해 스프레드시트에 접근, [주문데이터] 시트 불러오기
            const order_sheet = SpreadsheetApp.openById(c.get('대시보드ID')).getSheetByName('주문데이터');
            const order = order_sheet.getDataRange().getValues();

            // 주문이 없는 경우 리턴
            if (order.length <= 1) {
                return;
            }

            // 최상위 제목열 제거
            order.splice(0, 1);

            // [주문데이터] 시트의 데이터를 복사하여 [에리확인] 시트로 입력, [주문데이터] 시트 데이터 삭제
            target_sheet.insertRowsAfter(1, order.length);
            target_sheet.getRange(2, 4, order.length, order[0].length).setNumberFormat('@').setValues(order);
            // 셀러명열에 셀러명 값 일괄입력
            target_sheet.getRange(2, 2, order.length, 1).setNumberFormat('@').setValue(k);

            order_sheet.deleteRows(2, order.length);

            // [발주체크] 횟수 반영
            check_data[check_data.findIndex((v) => v[0] == k)][1]++;
        };

        // 직접발주 셀러일 경우,
        if (k == '직접발주') {

            // [내부발주] 스프레드시트 - [접수내역] 시트 데이터 불러오기
            const order_sheet = SpreadsheetApp.openById('14PVHQMF13UWHnJyvYOXLdIuHptnx0C5iOJbc63LLHzs').getSheetByName('접수내역');
            const order = order_sheet.getDataRange().setNumberFormat('@').getValues();

            // 주문이 없을 경우 리턴
            if (order.length <= 1) {
                return;
            }

            // 최상위 제목열 제거
            order.splice(0, 1);

            // [접수내역] 시트의 데이터를 복사하여 [에리확인] 시트로 입력, [접수내역] 시트 데이터 삭제
            target_sheet.insertRowsAfter(1, order.length);
            target_sheet.getRange(2, 3, order.length, order[0].length).setNumberFormat('@').setValues(order);
            // 셀러명열에 직접발주 값 일괄입력
            target_sheet.getRange(2, 2, order.length, 1).setNumberFormat('@').setValue('직접발주');

            target_sheet.getRange(2, 4, target_sheet.getLastRow() - 1, 2).setNumberFormat('@');

            order_sheet.deleteRows(2, order.length);

            // 발주체크 횟수 반영
            check_data[check_data.findIndex((v) => v[0] == k)][1]++;
        }
        return;
    });

    // 발주체크 데이터 적용
    check_sheet.getDataRange().setValues(check_data);
}

// 시트의 전체 포맷을 변경하는 함수
async function set_Format(id, format) {
    // id 값으로 스프레드시트를 열어 모든 데이터셀의 포맷을 format으로 변경 ('@' = 텍스트)
    SpreadsheetApp.openById(id).getDataRange().setNumberFormat(format);
}

// file_id를 통해 스프레드시트 파일에 접근 후, form에 지정되어있는 열들을 묶어 [에러확인] 시트로 복사하는 함수
async function fetch_Data(file_id, form, client_info) {

    // file_id 스프레드시트의 데이터 가져오기
    const orig_data = SpreadsheetApp.openById(file_id).getDataRange().getValues();
    orig_data.splice(0, 1);

    let input_data = new Array();

    // 데이터 각 행에 접근하여 form에 따라서 데이터 취사 선택 후 저장
    orig_data.forEach((o, index) => {
        // form[0] (주문번호) 에 저장되어있는 알파벳 열을 숫자로 변환
        let i = convert_Column(form[0]);
        
        // 주문번호가 없을 경우에 리턴
        if (!o[i - 1] || o[i - 1] == ' ') {
            return;
        }

        // 15열짜리 배열 생성
        input_data[index] = new Array();
        input_data[index][15] = '';
        // 셀러명 열에 셀러명 입력
        input_data[index][order_form.get('셀러명')] = client_info.get('셀러명');

        // 폼 배열을 하나씩 돌리며, 각 데이터 수집
        form.forEach((f, id) => {

            // 폼 데이터에 저장되어있는 알파벳을 숫자로 변환 후, [주문현황] 양식과 연결
            let column = fetch_form.get('양식')[id];
            let i = convert_Column(f);
            let j = order_form.get(column);
            // form(전달양식)이 none 값일 경우, 데이터 비운채로 리턴
            if (f === 'none') {
                input_data[index][j] = '';
                return;
            }

            // 엔터가 포함된 데이터일 경우 ' ''로 대치
            if (o[i - 1]) {
                o[i - 1] = o[i - 1].split('\n').join(' ');
            }

            // 우편번호 양식 통일
            if (column === '우편번호') {
                input_data[index][j] = Utilities.formatString('%05d', o[i - 1].split('-').join(''));
                return;
            }

            // 전화번호 양식 통일
            if (column === '주문자연락처' || column === '수령인연락처') {
                if (o[i - 1].indexOf('-') == -1) {
                    input_data[index][j] = o[i - 1].split(' ').join('');
                } else {
                    input_data[index][j] = o[i - 1].split('-').join('');
                }
                return;
            }

            // 폼 데이터가 알파벳, 알파벳 (A, B)일 경우 (, 구분자로 구분되어 있을 경우) 첫 번째 알파벳 열 탐색 후 데이터 부재시 두 번째 알파벳 열 탐색
            let comma = f.split(',');
            for (let c of comma) {

                // 폼 데이터가 + 구분자로 구분된 경우, 앞열과 뒷열 데이터를 합쳐서 저장
                let plus = c.split('+');
                let temp = '';
                plus.forEach((p, k) => {
                    let i = convert_Column(p);
                    if (!o[i - 1]) {
                        return;
                    }

                    if (column !== '주소' && column !== '배송메세지' && column !== '옵션정보' && column !== '결제일') {
                        if (o[i - 1]) {
                            o[i - 1] = o[i - 1].split(' ').join('');
                        }
                    }
        
                    if (k == 0) {
                        temp += o[i - 1];
                    } else {
                        temp += '-' + o[i - 1];
                    }
                });
                if (temp != '' && temp != '-') {
                    input_data[index][j] = temp;
                    break;
                }
            }
        });
    });

    return input_data;
}

// 알파벳 >> 숫자 변환 함수
function convert_Column(col) {
    if (typeof col === 'string') {
        let num = 0;
        if (col.length > 1) {
            num += (col.charCodeAt(0) - 64) * 26 + (col.charCodeAt(1) - 64);
        } else {
            num += (col.charCodeAt(0) - 64);
        }
        return num;
    }
    if (typeof col === 'number') {
        let str;
        if (col > 26) {
            str = String.fromCharCode((col / 26) + 64) + String.fromCharCode((col % 26) + 64);
        } else {
            str = String.fromCharCode(col + 64);
        }
        return str;
    }
}