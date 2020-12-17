function myFunction() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('생산관리');
    const target = SpreadsheetApp.openById(`1w-OcPwTV5XgfJ6XKahGcSSkCmuDOxVpDnWiQY9JPbqI`).getSheetByName(`전체 제품`);
    const sheet_data = sheet.getDataRange().getValues();
    const target_data = target.getDataRange().getValues();

    
}
