async function check_files() {
    const folder = DriveApp.getFolderById(`1QbOPPcLELS5cnVd84NFhiqx4GcBKnRM9`);
    const files = folder.getFiles();

    while (files.hasNext()) {
        let file = files.next();
        if (file.getMimeType == MimeType.MICROSOFT_EXCEL_LEGACY || file.getMimeType == MimeType.MICROSOFT_EXCEL || file.getMimeType == MimeType.CSV) {
            let blob = file.getBlob();
            let name = file.getName();
            let props = {
                title: name,
                MimeType: MimeType.GOOGLE_SPREAD_SHEETS,
                parents: {
                    id: `1QbOPPcLELS5cnVd84NFhiqx4GcBKnRM9`
                }
            }
            Drive.Files.insert(props, blob);
        }
    }
}