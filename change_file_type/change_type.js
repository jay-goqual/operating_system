async function change_type() {
    const folder = DriveApp.getFolderById(`1QbOPPcLELS5cnVd84NFhiqx4GcBKnRM9`);
    const files = folder.getFiles();

    while (files.hasNext()) {
        let file = files.next();
        console.log(file.getName());
        if (file.getMimeType() == MimeType.MICROSOFT_EXCEL_LEGACY || file.getMimeType() == MimeType.MICROSOFT_EXCEL || file.getMimeType() == MimeType.CSV) {
            let blob = file.getBlob();
            let name = file.getName();
            let props = {
                title: name,
                parents: [{
                    id: `1QbOPPcLELS5cnVd84NFhiqx4GcBKnRM9`
                }],
                mimeType: MimeType.GOOGLE_SHEETS,
            }
            Drive.Files.insert(props, blob);
            file.setTrashed(true);
        }
    }
}