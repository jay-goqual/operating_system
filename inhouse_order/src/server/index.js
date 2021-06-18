import * as publicUiFunctions from './ui';
import * as publicSheetFunctions from './sheets';

// sheets.js 파일의 export 함수를 apps script의 함수로 불러오기
global.onOpen = publicUiFunctions.onOpen;
global.openSidebar = publicUiFunctions.openSidebar;
global.pushOrder = publicSheetFunctions.pushOrder;