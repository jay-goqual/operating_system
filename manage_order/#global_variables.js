//ref 전역선언
// var ref = get_Ref();

//[전역]레퍼런스(id) 가져오기
function get_Ref() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().getValues();
    let ref = new Map();

    table.forEach((t) => {
        ref.set(t[0], t[2]);
    });

    return ref;
}

//Ref 찾기
function find_Ref(key) {
    return ref.get(key);
}

//[전역]이메일 가져오기
//셀러관리 이후 셀러관리에서 셀러정보 가져오는 것으로 변경
function get_Client() {
    //const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이메일').getDataRange().setNumberFormat('@').getValues();
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

//[전역]변환양식 가져오기
//요청양식에서 변환양식 가져옴
function get_Fetch_form() {
    //const table = SpreadsheetApp.openById(ref['출고요청/요청양식']).getDataRange().getValues();
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('요청양식').getDataRange().getValues();

    let fetch_form = new Map();

    table.forEach((t) => {
        fetch_form.set(t[0], t.splice(1, t.length - 1));
    });

    return fetch_form;
}

//[전역]상품 가져오기
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

//[전역]양식정보 가져오기
function get_Order_form() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문양식').getDataRange().getValues();

    let order_form = new Map();

    table.forEach((t) => {
        order_form.set(t[0], t[1]);
    });

    return order_form;
}

//[전역]택배사 가져오기
function get_Delivery() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('택배사').getDataRange().getValues();

    let delivery = new Map();

    table.forEach((t) => {
        delivery.set(t[0], t[1]);
    });

    return delivery;
}

function get_Invoice_form() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('전달양식').getDataRange().getValues();
    let invoice_form = new Map();

    table.forEach((t) => {
        invoice_form.set(t[0], t.splice(1, t.length - 1));
    });

    return invoice_form;
}