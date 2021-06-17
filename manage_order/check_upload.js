// 업로드 폴더에 파일이 업로드되었는지, 업로드가 되었다면 주문파일인지 송장파일인지 확인하는 파일입니다.

// 전역함수 로드
var ref = get_Ref();
var client = get_Client();

// 업로드 폴더 내의 파일을 체크하는 함수
async function check_Upload() {
    const folder = DriveApp.getFolderById(ref.get('업로드'));
    const files = folder.getFiles();

    // 업로드 폴더에 있는 모든 파일을 검색
    while (files.hasNext()) {
        let file = files.next();

        // 파일이 엑셀 혹은 CSV일 경우,
        if (file.getMimeType() == MimeType.MICROSOFT_EXCEL_LEGACY || file.getMimeType() == MimeType.MICROSOFT_EXCEL || file.getMimeType() == MimeType.CSV) {

            // 해당 파일의 데이터를 가진 스프레드시트를 생성 후 업로드폴더로 이동
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

            let new_id = await insert_File(props, blob);
            let new_file = DriveApp.getFileById(new_id);

            // 주문 파일일 경우 (전역 client 정보에 _ split 2번째 인자가 셀러명으로 존재할 경우 혹은 _ split 1번째 인자가 PO일 경우(로켓배송))
            if (client.has(new_file.getName().split('_')[1]) || new_file.getName().split('_')[0] == 'PO') {
                
                // 주문 끌어오기 함수 실행
                await fetch_Order(new_file);
            }

            // 송장일 경우 (파일명에 배송출고현황, 작업 단위 목록, ShipmentReport 가 포함되어 있을 경우)
            if (new_file.getName().indexOf('배송출고현황') != -1 || new_file.getName().indexOf('작업 단위 목록') != -1 || new_file.getName().indexOf('ShipmentReport') != -1) {

                // 송장 끌어오기 함수 실행
                await fetch_Invoice(new_file);
            }
        }

        // 중간 오류 발생 시를 대비하여 구글 스프레드시트 타입일 경우에도 실행
        if (file.getMimeType() == MimeType.GOOGLE_SHEETS) {
            // 주문 파일일 경우 (전역 client 정보에 _ split 2번째 인자가 셀러명으로 존재할 경우 혹은 _ split 1번째 인자가 PO일 경우(로켓배송))
            if (client.has(new_file.getName().split('_')[1]) || new_file.getName().split('_')[0] == 'PO') {
                
                // 주문 끌어오기 함수 실행
                await fetch_Order(new_file);
            }

            // 송장일 경우 (파일명에 배송출고현황, 작업 단위 목록, ShipmentReport 가 포함되어 있을 경우)
            if (new_file.getName().indexOf('배송출고현황') != -1 || new_file.getName().indexOf('작업 단위 목록') != -1 || new_file.getName().indexOf('ShipmentReport') != -1) {

                // 송장 끌어오기 함수 실행
                await fetch_Invoice(new_file);
            }
        }
    }
}


// 엑셀 파일의 데이터를 통해 새로운 스프레드시트파일 생성 후 id 값 리턴
async function insert_File(props, blob) {
    return Drive.Files.insert(props, blob).id;
}