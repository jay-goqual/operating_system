export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('GQ_CX') // edit me!
    // .addItem('Sheet Editor', 'openDialog')
    .addItem('검색', 'openDialogBootstrap')
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
