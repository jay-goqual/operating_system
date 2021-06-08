import * as publicUiFunctions from './ui';
import * as publicSheetFunctions from './sheets';

// Expose public functions by attaching to `global`
global.onOpen = publicUiFunctions.onOpen;
// global.openDialog = publicUiFunctions.openDialog;
global.openDialogBootstrap = publicUiFunctions.openDialogBootstrap;
// global.openAboutSidebar = publicUiFunctions.openAboutSidebar;
// global.getSheetsData = publicSheetFunctions.getSheetsData;
// global.addSheet = publicSheetFunctions.addSheet;
// global.deleteSheet = publicSheetFunctions.deleteSheet;
// global.setActiveSheet = publicSheetFunctions.setActiveSheet;

global.findOrder = publicSheetFunctions.findOrder;
global.getProducts = publicSheetFunctions.getProducts;
global.getData = publicSheetFunctions.getData;
global.getInspection = publicSheetFunctions.getInspection;
global.matchData = publicSheetFunctions.matchData;
global.pushData = publicSheetFunctions.pushData;
global.archiveData = publicSheetFunctions.archiveData;
global.pushArchive = publicSheetFunctions.pushArchive;