// UI와 관련된 함수 저장

export const onOpen = () => {
  const menu = SpreadsheetApp.getUi()
    .createMenu('고퀄') // edit me!
    .addItem('사이드바 열기', 'openSidebar')
    .addItem('제출하기', 'pushOrder');

  menu.addToUi();
};

export const openSidebar = () => {
    const html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('고퀄 내부발주');
    SpreadsheetApp.getUi().showSidebar(html);
}
