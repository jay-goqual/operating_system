// [에러확인] 시트에 저장된 데이터를 통해 추가적으로 필요한 데이터를 수집(셀러, 상품 데이터)하여 저장하며
// 주문 중복, 데이터 부족 등 처리할 수 없는 행을 걸러낸 후, [주문현황] 시트로 옮기는 파일입니다.

// 전역함수 호출
var ref = get_Ref();
var order_form = get_Order_form();
var delivery = get_Delivery();
var client = get_Client();
var postpone = new Array();

// [에러확인] 시트의 부족한 데이터를 채우고, [주문현황] / [확인요청] 시트로 데이터를 넘기는 함수입니다.
async function submit_Order() {
    // [에러확인], [주문현황], [확인요청] 시트 불러오기

    const error_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문현황');
    const check_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('확인요청');

    const table = error_sheet.getDataRange().getValues();
    const target_table =target_sheet.getDataRange().getValues();

    // [에러확인] 시트의 데이터가 없을 경우 리턴
    if (table.length == 1) {
        return;
    };
    
    // 최상단 제목행 삭제
    table.splice(0, 1);

    // 각 데이터를 저장할 배열 생성
    let total_table = new Array();
    let error_table = new Array();
    let check_table = new Array();

    let count = 0;
    let check = false;

    // [에러확인]의 모든 데이터 탐색
    table.forEach((t, i) => {
        // [주문현황]에 이미 접수된 주문 (상품주문번호와 주문번호로 확인) 일 경우, check = true로 변경하여 에러로 판단
        if (target_table.filter(x =>
            (x[order_form.get('주문번호')] == t[order_form.get('주문번호')] && 
            x[order_form.get('상품주문번호')].split('-')[0] == t[order_form.get('상품주문번호')].split('-')[0])).length > 0) {
                check = true;
                // 행 색상 빨간색으로 변경
                error_sheet.getRange(i + 2, 1, 1, table[0].length).setBackground('#f4cccc');
                // 에러확인 열을 false로 기입
                error_sheet.getRange(i + 2, order_form.get('에러확인') + 1).setValue(false);
                t[order_form.get('에러확인')] = false;
            }

        // [에러확인] 시트에 중복으로 기입된 주문일 경우 에러로 판단
        if (table.filter(x =>
            (x[order_form.get('주문번호')] == t[order_form.get('주문번호')] && 
            x[order_form.get('상품주문번호')] == t[order_form.get('상품주문번호')])).length > 1) {
                check = true;
                error_sheet.getRange(i + 2, 1, 1, table[0].length).setBackground('#f4cccc');
                error_sheet.getRange(i + 2, order_form.get('에러확인') + 1).setValue(false);
                t[order_form.get('에러확인')] = false;
            }

        // 중복되지 않은 경우
        if (t[order_form.get('에러확인')] == true) {

            // 대기가 필요한 상품 (주문제작상품)일 경우, check_table (확인필요) 테이블에 추가
            if (t[order_form.get('출고채널')].indexOf('대기') != -1) {
                check_table.push(t);
            }

            // [주문현황]에 넘어갈 total_table 에도 추가
            total_table.push(t);
        } else {

            // 에러가 확인된 경우에는 error_table에 추가
            count++;
            error_table.push(t);
        }
    });

    // total_table의 데이터를 [주문현황] 시트에 기입
    if (total_table.length > 0) {
        target_sheet.insertRowsAfter(1, total_table.length);
        target_sheet.getRange(2, 1, total_table.length, total_table[0].length).setValues(total_table);

        // [에러확인] 시트의 데이터 중 에러확인이 완료된 데이터를 삭제
        error_sheet.sort(order_form.get('에러확인') + 1);
        error_sheet.deleteRows(count + 2, total_table.length);

        // check_table의 데이터를 [확인필요] 시트에 기입
        if (check_table.length > 0) {
            check_sheet.insertRowsAfter(1, check_table.length);
            check_sheet.getRange(2, 1, check_table.length, check_table[0].length).setValues(check_table);
            check_sheet.getRange(2, check_table[0].length + 1, check_table.length, 1).insertCheckboxes();
        }
    };

    // 에러가 발생했을 경우 UI로 알림
    if (check) {
        SpreadsheetApp.getUi().alert('중복 주문이 감지되었습니다. 에러확인 시트를 확인해 주세요.');
    }
}

// 외부데이터 (상품/셀러) 수집 후, 데이터 이상이 있는 지 체크하는 함수
async function catch_Error(index, order) {
    // 첫 값은 true
    
    order[order_form.get('에러확인')] = true;

    const check = ['셀러명', '주문번호', '상품주문번호', '상품코드', '수량', '주문자', '주문자연락처', '수령인', '수령인연락처', '주소', '상품명', '출고채널', '택배사', '판매액', '배송비', '수수료'];

    // 꼭 값이 존재해야하는 열에 데이터가 없는 경우에 대한 체크
    check.forEach((c) => {
        if (order[order_form.get(c)] === '' || order[order_form.get(c)] === null) {
            // 에러확인 값을 false로 변경
            order[order_form.get('에러확인')] = false;
            // 셀의 색 변경
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인').getRange(index + 2, order_form.get(c) + 1).setBackground('#f4cccc');
        }
    });

    return order;
}

// 추가적인 정보 (상품, 셀러)를 가져오는 함수
async function fetch_Additional_info() {

    // 상품정보 가져오기
    const productInfo = get_Product();

    // [에러확인] 시트의 데이터 가져오기
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('에러확인');
    const order = sheet.getDataRange().getValues();

    // [에러확인] 시트의 데이터가 없을 경우 리턴
    if (order.length == 1) {
        return;
    }

    // 최상위 제목행 삭제
    order.splice(0, 1);

    // 데이터 저장할 배열 생성
    let total = new Map();

    // 모든 데이터 탐색
    order.map((o) => {

        // 접수일 = 현재시각 으로 입력
        o[order_form.get('접수일')] = Utilities.formatDate(new Date(), 'GMT+9', 'yyyy/MM/dd HH:mm');

        // 셀러명이 직접발주가 아닌 경우에는 [셀러관리]에서 셀러코드를 검색하여 입력
        // 직접발주는 셀러코드 기입력 상태
        if (o[order_form.get('셀러명')] != '직접발주') {
            o[order_form.get('셀러코드')] = client.get(o[order_form.get('셀러명')]).get('셀러코드');
        }

        // 상품코드를 통해 상품데이터 fetch
        let code = order_form.get('상품코드');
        let p = productInfo.get(o[code]);

        // [상품관리] 스프레드시트 - [상품DB] 시트에 상품데이터가 존재할 경우
        if (p) {
            // 상품명 기입
            o[order_form.get('상품명')] = p.get('상품명');

            // 출고채널 기입
            if (!o[order_form.get('출고채널')]) {
                o[order_form.get('출고채널')] = p.get('출고채널');

                // 주문제작 제품 특별처리
                // 해피콜이 필요없는 셀러일 경우에는 출고채널에 각 상품의 출고처 입력
                // 주문제작 상품의 기본 출고채널은 '대기_xxx'
                if ((o[order_form.get('셀러코드')][0] == '2' || o[order_form.get('셀러코드')] == '30007' || o[order_form.get('셀러명')] == '직접발주') && o[order_form.get('출고채널')] == '대기_커튼') {
                    o[order_form.get('출고채널')] = '제이에스비즈';
                } else if ((o[order_form.get('셀러코드')][0] == '2' || o[order_form.get('셀러코드')] == '30007' || o[order_form.get('셀러명')] == '직접발주') && o[order_form.get('출고채널')] == '대기_커튼천') {
                    if (o[order_form.get('상품코드')].indexOf('CF201') > 0) {
                        o[order_form.get('출고채널')] = '건인디앤씨';
                    } else {
                        o[order_form.get('출고채널')] = '드림캐쳐';
                    }
                }

                // [택배사] 시트에서 각 출고채널의 택배사 확인하여 입력
                o[order_form.get('택배사')] = delivery.get(p.get('출고채널'));
            }

            // 판매액 계산
            // 직접발주의 경우에는 기입력 상태
            if (o[order_form.get('셀러명')] != '직접발주') {

                // 판매액 = 상품의 판매가 * 주문의 수량
                o[order_form.get('판매액')] = Number(p.get('판매가')) * Number(o[order_form.get('수량')]);

                // 커튼천일 경우에는 옵션정보의 폭 수량을 기반으로 가중처리
                if (o[order_form.get('셀러명')] == '공식몰' && o[order_form.get('상품코드')].indexOf('CF201') > -1) {
                    width = o[order_form.get('옵션정보')].split(' / ')[4].split(' : ')[1].split('폭')[0];
                    o[order_form.get('판매액')] = o[order_form.get('판매액')] * width;
                }
            }

            // 택배비 계산을 위한 주문별 총 주문액 누적
            if (total.has(o[order_form.get('주문번호')])) {
                total.set(o[order_form.get('주문번호')], total.get(o[order_form.get('주문번호')]) + o[order_form.get('판매액')]);
            } else {
                total.set(o[order_form.get('주문번호')], o[order_form.get('판매액')]);
            }

            // 수수료 계산
            if (o[order_form.get('셀러명')] != '직접발주') {
                
                let rate;
                let client_info = client.get(o[order_form.get('셀러명')]);

                // [셀러관리] 스프레드시트의 공급방식에 따란 계산방식 변경
                // 고정수수료일 경우에는 rate 고정
                if (client_info.get('공급방식') == '고정수수료') {
                    rate = Number(client_info.get('고정수수료율'));
                } 
                // 가산수수료일 경우에는, [상품관리]의 상품수수료율 + [셀러관리]의 가산수수료율 적용
                // 단, 상품의 가산수수료율 열이 Y일 경우에만
                else if (client_info.get('공급방식') == '가산수수료') {
                    if (p.get('가산수수료') == 'Y') {
                        rate = Number(p.get('상품수수료율')) + Number(client_info.get('가산수수료율'));
                    } else {
                        rate = Number(p.get('상품수수료율'));
                    }
                } 
                // 모두 아닐 경우에도 rate 고정
                else {
                    rate = Number(client_info.get('고정수수료율'));
                }

                // 위의 조건을 무시하는 조건1 적용
                // [상품관리]의 상품 중 셀러코드와 동일한 열이 존재하고, 그 열에 값이 있는 경우에는 rate를 무시하고 판매액과 수수료 계산
                // 판매액에 반영, 수수료 0원
                if (p.get(String(o[order_form.get('셀러코드')]))) {
                    o[order_form.get('판매액')] = Number(p.get(String(o[order_form.get('셀러코드')]))) * o[order_form.get('수량')];
                    o[order_form.get('수수료')] = 0;
                } 
                // 아닐 경우에는, (rate * 상품의 판매가) (10의 자리에서 올림) * 수량
                else {
                    o[order_form.get('수수료')] = Math.ceil((Number(p.get('판매가')) * rate) / 10) * 10 * o[order_form.get('수량')];
                }
            }
        }
        return o;
    });

    let num = 0;

    // 배송비 책정하기
    order.map(async (o, i) => {
        
        let orderId = o[order_form.get('주문번호')];
        let code = o[order_form.get('상품코드')];

        // t = 주문번호의 총 판매액
        let t = total.get(orderId);
        let p = productInfo.get(code);

        // 직접발주일 경우에는 기입력 상태
        if (o[order_form.get('셀러명')] != '직접발주') {

            // 상품 판매가 가능한 상태일 경우,
            if (p) {
                // [상품관리] 스프레드시트의 무료배송기준보다 총 금액이 높은 경우에는 배송비 0원, 아닐 경우에는 상품별 배송비 기입
                if (Number(t) > Number(p.get('무료배송기준')) || t == -1) {
                    o[order_form.get('배송비')] = 0;
                } else {
                    o[order_form.get('배송비')] = p.get('상품배송비');
                    total.set(orderId, -1);
                }
            }   
        }

        // 결제일 양식 통일 및 변경

        // 결제일 열 위치 = date
        let date = order_form.get('결제일');

        // 결제일이 없는 경우
        if (!o[date] || isNaN(new Date(o[date]).getTime())) {

            // 주문번로를 통한 결제일 역산
            let n;
            if (orderId[0] == 'N') {
                n = 1;
            } else {
                n = 0;
            }
            let assume = new Date(`${orderId.substring(n, n + 4)}-${orderId.substring(n + 4, n + 6)}-${orderId.substring(n + 6, n + 8)} 00:00`);

            // today = 접수일 (현재시각)
            let today = new Date(o[order_form.get('접수일')]);
            o[date] = today;
            
            // 역산된 결제일과 오늘의 차이가 6개월 이내라면 역산된 결제일로 기입
            let oneDay = 24 * 60 * 60 * 1000;
            if (Math.round(Math.abs(assume - today) / oneDay) < 180) {
                o[date] = assume;
            }
        } else {
            o[date] = new Date(o[date]);
        }

        // 포맷 고정
        o[date] = Utilities.formatDate(o[date], 'GMT+9', 'yyyy/MM/dd HH:mm');


        // 상품주문번호가 동일한 경우, 자동적으로 뒤에 -01 -02 붙여줌
        if (i > 0) {
            // 상품주문번호 중복 검색
            // 중복이 있는 경우
            if (order.filter(x => x[order_form.get('상품주문번호')] == o[order_form.get('상품주문번호')]).length > 1) {
                // 이미 계산된 것이 있을 경우에는 num++
                if (o[order_form.get('상품주문번호')].indexOf(order[i - 1][order_form.get('상품주문번호')].split('-')[0]) > -1) {
                    num++;
                } 
                // 없다면 num = 1
                else {
                    num = 1;
                }
            } else {
                // 이미 계산된 것이 있을 경우 num++
                if (o[order_form.get('상품주문번호')].indexOf(order[i - 1][order_form.get('상품주문번호')].split('-')[0]) > -1) {
                    num++;
                } else {
                    // 없다면 num 초기화
                    num = 0;
                }
            }    
        } else {
            if (order.filter(x => x[order_form.get('상품주문번호')] == o[order_form.get('상품주문번호')]).length > 1) {
                num++;
            }
        }

        // 상품주문번호가 중복될 가능성이 있는 업체들일 경우에는 위에 계산된 중복건수(num)를 기존 상품주문번호 뒤에 붙여서 기입
        if (num > 0 && (o[order_form.get('셀러명')] == '직접발주' || o[order_form.get('셀러명')] == '씨씨티비프렌즈' || o[order_form.get('셀러명')] == '나혼자살림' || o[order_form.get('셀러명')] == '도치퀸' || o[order_form.get('셀러명')] == '오늘의집' || o[order_form.get('셀러명')] == '쿠팡마켓플레이스' || o[order_form.get('셀러명')] == '프리스비')) {
            o[order_form.get('상품주문번호')] = `${o[order_form.get('상품주문번호')]}-${Utilities.formatString('%02d', num)}`;
        }

        // 에러 체크
        o = await catch_Error(i, o);
        
        return o;
    });

    sheet.getRange(2, 1, order.length, order[0].length).setValues(order);
}