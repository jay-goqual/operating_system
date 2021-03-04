export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('고퀄') // edit me!
    .addItem('사이드바 열기', 'openSidebar')
    .addItem('제출하기', 'pushOrder');

  menu.addToUi();
};

/* export const openDialog = () => {
  const html = HtmlService.createHtmlOutputFromFile('dialog-demo')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sheet Editor');
};

export const openDialogBootstrap = () => {
  const html = HtmlService.createHtmlOutputFromFile('dialog-demo-bootstrap')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sheet Editor (Bootstrap)');
}; */

/* export const openAboutSidebar = () => {
  const html = HtmlService.createHtmlOutputFromFile('sidebar-about-page');
  SpreadsheetApp.getUi().showSidebar(html);
}; */

export const openSidebar = () => {
    const html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('고퀄 내부발주');
    SpreadsheetApp.getUi().showSidebar(html);
}
