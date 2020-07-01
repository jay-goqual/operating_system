var ref = get_Ref();
var client = get_Client();

async function check_Upload() {
    const folder = DriveApp.getFolderById(ref.get('업로드'));
    const files = folder.getFiles();

    while (files.hasNext()) {
        let file = files.next();
        if (file.getMimeType() == MimeType.MICROSOFT_EXCEL_LEGACY || file.getMimeType() == MimeType.MICROSOFT_EXCEL || file.getMimeType() == MimeType.CSV) {
            let blob = file.getBlob();
            let name = file.getName();
            let props = {
                title: name,
                parents: [{
                    id: ref.get('업로드')
                }],
                mimeType: MimeType.GOOGLE_SHEETS,
            }

            file.setTrashed(true);

            let new_id = await insert_file(props, blob);
            let new_file = DriveApp.getFileById(new_id);

            //주문일 경우
            if (client.has(new_file.getName().split('_')[1])) {
                await fetch_Order(new_file);
            }

            //송장일 경우
            if (new_file.getName().indexOf('배송출고현황') != -1 || new_file.getName().indexOf('작업 단위 목록') != -1) {
                await fetch_Invoice(new_file);
            }
        }   
    }
}

async function insert_file(props, blob) {
    return Drive.Files.insert(props, blob).id;
}