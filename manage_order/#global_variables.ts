//ref 전역선언
var ref = get_Ref();

//[전역]레퍼런스(id) 가져오기
function get_Ref() {
    const table: Array<Array<string>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().getValues();
    let ref: Map<string, string> = new Map();

    table.forEach((t) => {
        ref.set(t[0], t[2]);
    });

    return ref;
}

//Ref 찾기
function find_Ref(key: string) {
    return ref.get(key);
}

//[전역]이메일 가져오기
//셀러관리 이후 셀러관리에서 셀러정보 가져오는 것으로 변경
function get_Client() {
    //const table: Array<Array<string>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이메일').getDataRange().setNumberFormat('@').getValues();
    const table: Array<Array<string>> = SpreadsheetApp.openById(ref.get('업체관리')).getSheetByName('업체DB').getDataRange().getValues();
    let email_address: Map<string, Array<string>> = new Map();

    table.forEach((t) => {
        email_address.set(t[0], [t[0], t[1], t[2], t[3], t[13]]);
    });

    return email_address;
}

//[전역]변환양식 가져오기
//요청양식에서 변환양식 가져옴
function get_Fetch_form() {
    //const table = SpreadsheetApp.openById(ref['출고요청/요청양식']).getDataRange().getValues();
    const table: Array<Array<string>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('요청양식').getDataRange().getValues();

    let fetch_form: Map<string, Array<string>> = new Map();

    table.forEach((t) => {
        fetch_form.set(t[0], t.splice(1, t.length));
    });

    return fetch_form;
}

//[전역]상품 가져오기
function get_Product() {
    const table: Array<Array<string>> = SpreadsheetApp.openById(ref.get('상품관리')).getSheetByName('상품DB').getDataRange().getValues();
    let product: Map<string, Array<string>> = new Map();

    table.forEach((t) => {
        product.set(t[0], [t[1], t[2], t[4], t[9], t[10]]);
    });

    return product;
}

//[전역]양식정보 가져오기
function get_Order_form() {
    const table: Array<Array<any>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문양식').getDataRange().getValues();

    let order_form: Map<string, number> = new Map();

    table.forEach((t) => {
        order_form.set(t[0], t[1]);
    });

    return order_form;
}

//[전역]택배사 가져오기
function get_Delivery() {
    const table: Array<Array<any>> = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('택배사').getDataRange().getValues();

    let delivery: Map<string, string> = new Map();

    table.forEach((t) => {
        delivery.set(t[0], t[1]);
    });

    return delivery;
}