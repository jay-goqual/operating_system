// 판매데이터를 축적하는 [판매현황] 스프레드시트와 연결되어 있는 apps script로 데이터를 받아오는 함수가 포함되어 있습니다.
// 트리거 없이 수동버튼으로 작업이 이루어지며, manage_order에서 아카이브는 자동적으로 이루어지고 있습니다.

// UI 버튼 생성
function Init() {
    SpreadsheetApp.getUi().createMenu('판매관리')
    .addItem('판매데이터 가져오기', 'button')
    .addToUi()
}

// 레퍼런스 시트 데이터 불러오기
function get_Ref() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().getValues()
    let ref = new Map()

    table.forEach((t) => {
        ref.set(t[0], t[1])
    })

    return ref
}

// 마케팅 데이터 긁어오기 (폐기)
async function fetch_Marketing_data() {
    const ref = get_Ref()
    const today = new Date()
    let year = today.getFullYear()
    let month = today.getMonth()
    let day = 1
    month -= 3;
    if (month < 0) {
        month += 12
        year -= 1
    }
    const time = new Date(`${year}-${month + 1}-${day}`)

    const table = SpreadsheetApp.openById(ref.get('광고관리')).getSheetByName('집행내역').getDataRange().getValues()
    table.splice(0, 1)

    const total = new Array()

    table.forEach((t) => {
        if (t[0] >= time) {
            total.push(t)
        }
    })

    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName('3개월광고DB').getLastRow() > 1) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('3개월광고DB').deleteRows(2, SpreadsheetApp.getActiveSpreadsheet().getSheetByName('3개월광고DB').getLastRow() - 1)
    }
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('3개월광고DB').insertRowsAfter(1, total.length)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('3개월광고DB').getRange(2, 1, total.length, total[0].length).setValues(total)
}

// 주문데이터 가져오기
// 자동 아카이브에 문제가 생겼을 경우, 갖고 있는 모든 주문데이터 파일에서 데이터를 긁어와 복사하는 기능을 합니다.
async function fetch_Order_data() {
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문DB').getLastRow() > 1) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문DB').deleteRows(2, SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문DB').getLastRow() - 1)
    }

    const ref = get_Ref()
    const today = new Date()

    const folder = DriveApp.getFolderById(ref.get('아카이브'))
    const files = folder.getFiles()

    // 이번달, 저번달 시트 id 추가
    let file_id = [ref.get('이번달DB'), ref.get('저번달DB')]

    // 존재하는 모든 시트 id 추가
    while (files.hasNext()) {
        const file = files.next()

        for (i = 2; i < 12; i++) {
            let year = today.getFullYear()
            let month = today.getMonth()

            month -= i;
            if (month < 0) {
                month += 12
                year -= 1
            }

            if (file.getName() == `${year}년 ${month + 1}월`) {
                file_id.push(file.getId())
                break
            }
        }
    }

    const total = new Array();

    // 추가한 모든 시트에 접근하여 데이터 추출
    file_id.forEach((f) => {
        const data = SpreadsheetApp.openById(f).getSheetByName('주문').getDataRange().getValues()

        data.splice(0, 1)
        data.forEach((d, i) => {
            total.push(d);
        })
    })

    // 리턴
    return total;
}

async function button() {
    let total = await fetch_Order_data();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문DB').insertRowsAfter(1, total.length)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('주문DB').getRange(2, 1, total.length, total[0].length).setValues(total)
}