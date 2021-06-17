// 스프레드시트 = 파일
// 시트 = 스프레드시트 내의 시트

// 전역으로 사용되는 함수들이 작성된 파일입니다.

// 레퍼런스 시트에서 각 연동파일의 ID 값 가져오기
function get_Ref() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().getValues();
    let ref = new Map();

    table.forEach((t) => {
        ref.set(t[0], t[2]);
    });

    return ref;
}

//Ref값 찾기
function find_Ref(key) {
    return ref.get(key);
}

// 셀러관리 스프레드시트 파일에서 각 업체의 정보(이메일, 대시보드 파일 ID, 양식종류 등) 불러오기
function get_Client() {
    const table = SpreadsheetApp.openById(ref.get('셀러관리')).getSheetByName('업체DB').getDataRange().getValues();
    let client = new Map();

    table.forEach((t) => {
        let temp = new Map();
        t.forEach((x, i) => {
            temp.set(table[0][i], x);
        })
        client.set(t[0], temp);
    });

    return client;
}

// 요청양식(파일 업로드시 적용) 시트에서 각 양식 불러오기
function get_Fetch_form() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('요청양식').getDataRange().getValues();

    let fetch_form = new Map();

    table.forEach((t) => {
        fetch_form.set(t[0], t.splice(1, t.length - 1));
    });

    return fetch_form;
}

// 상품관리 스프레드시트에서 상품 정보 불러오기
function get_Product() {
    const table = SpreadsheetApp.openById(ref.get('상품관리')).getSheetByName('상품DB').getDataRange().getValues();
    let product = new Map();

    table.forEach((t) => {
        let temp = new Map();
        t.forEach((x, i) => {
            temp.set(table[0][i], x);
        })
        product.set(t[0], temp);
    });

    return product;
}

// 출고관리 스프레드시트의 주문서 양식을 저장해둔 [주문양식] 시트의 정보 불러오기
function get_Order_form() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문양식').getDataRange().getValues();

    let order_form = new Map();

    table.forEach((t) => {
        order_form.set(t[0], t[1]);
    });

    return order_form;
}

// 각 출고처의 택배사 정보를 [택배사] 시트에서 불러오기
function get_Delivery() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('택배사').getDataRange().getValues();

    let delivery = new Map();

    table.forEach((t) => {
        delivery.set(t[0], t[1]);
    });

    return delivery;
}

// 각 판매처에게 전달될 때 사용되는 양식을 [전달양식] 시트에서 불러오기
function get_Invoice_form() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('전달양식').getDataRange().getValues();
    let invoice_form = new Map();

    table.forEach((t) => {
        invoice_form.set(t[0], t.splice(1, t.length - 1));
    });

    return invoice_form;
}