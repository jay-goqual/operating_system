function Init() {
    SpreadsheetApp.getUi().createMenu('판매관리')
    .addItem('이얍', 'fetch_Order_data')
    .addItem('호호', 'fetch_Marketing_data')
    .addToUi()
}

function get_Ref() {
    const table = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('레퍼런스').getDataRange().getValues()
    let ref = new Map()

    table.forEach((t) => {
        ref.set(t[0], t[1])
    })

    return ref
}

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

async function fetch_Order_data() {
    const ref = get_Ref()
    const today = new Date()

    const folder = DriveApp.getFolderById(ref.get('아카이브'))
    let year = [today.getFullYear(), today.getFullYear()]
    let month = [today.getMonth(), today.getMonth()]
    month[0] -= 2;
    if (month[0] < 0) {
        month[0] += 12
        year[0] -= 1
    }
    month[1] -= 3;
    if (month[1] < 0) {
        month[1] += 12
        year[1] -= 1
    }
    const files = folder.getFiles()

    let file_id = [ref.get('이번달DB'), ref.get('저번달DB')]

    while (files.hasNext()) {
        const file = files.next()

        for (i = 2; i < 6; i++) {
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

    file_id.forEach((f) => {
        const data = SpreadsheetApp.openById(f).getSheetByName('주문').getDataRange().getValues()

        data.splice(0, 1)

        data.forEach((d) => {
            total.push(d)
        })
    })

    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName('6개월주문DB').getLastRow() > 1) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('6개월주문DB').deleteRows(2, SpreadsheetApp.getActiveSpreadsheet().getSheetByName('6개월주문DB').getLastRow() - 1)
    }
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('6개월주문DB').insertRowsAfter(1, total.length)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('6개월주문DB').getRange(2, 1, total.length, total[0].length).setValues(total)
}

