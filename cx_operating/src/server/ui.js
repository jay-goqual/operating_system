export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('GQ_CX') // edit me!
    // .addItem('Sheet Editor', 'openDialog')
    .addItem('검색', 'openDialogBootstrap')
    // .addItem('접수완료', 'archiveData')
    .addItem('자동매칭', 'matchData')
    .addItem('검수데이터 가져오기', 'getInspection')
    // .addItem('About me', 'openAboutSidebar');

  menu.addToUi();
};

// export const openDialog = () => {
//   const html = HtmlService.createHtmlOutputFromFile('dialog-demo')
//     .setWidth(600)
//     .setHeight(600);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Sheet Editor');
// };

export const openDialogBootstrap = () => {
  const html = HtmlService.createHtmlOutputFromFile('dialog-demo-bootstrap')
    .setWidth(1600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '검색');
};

// export const openAboutSidebar = () => {
//   const html = HtmlService.createHtmlOutputFromFile('sidebar-about-page');
//   SpreadsheetApp.getUi().showSidebar(html);
// };
