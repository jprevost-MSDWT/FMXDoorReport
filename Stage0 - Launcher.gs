// Project Name: Door Report Full
// Project Version: 3.0
// Filename: Stage0 - Launcher.gs
// File Version: 3.01

const CONFIG = {
  sheets: {
    import: "Import",
    helper1: "Output-Helper1",
    helper2: "Output-Helper2",
    report: "AutoReport",
    reportNotes: "AutoReport w/Notes",
    data: "Data"
  },
  reportRanges: {
    standard: 7,
    alt: 14
  }
};

function onOpen() {
  VerifySheets();
  SpreadsheetApp.getUi()
      .createMenu('Report Menu')
      .addItem('Run Full Report', 'FullProcess')
      .addItem('Reprocess Last Import', 'ReProcess')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Manual Steps')
          .addItem('Reprocess Last Import', 'ReProcess')
          .addItem('Run Stage 1', 'Stage1')
          .addItem('Run Stage 2', 'Stage2')
          .addItem('Run Stage 3', 'Stage3'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Testing')
          .addItem('Import Standard (7 days)', 'ImportStandard')
          .addItem('Import Alt (14 days)', 'ImportAlt')
          .addItem('Testing1', 'Testing1')
          .addItem('Testing2', 'Testing2'))
      .addToUi();
}

function VerifySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = ss.getSheets().map(sheet => sheet.getName());
  const requiredSheetNames = Object.values(CONFIG.sheets);

  requiredSheetNames.forEach(sheetName => {
    if (allSheetNames.indexOf(sheetName) === -1) {
      ss.insertSheet(sheetName);
    }
  });
}

function FullProcess(){
  NewImport();
  Stage1();
  Stage2();
  Stage3();
}

function ReProcess(){
  Stage1();
  Stage2();
  Stage3();
}

function NewImport(){
  ImportStandard();
}

function ImportStandard() {
  ImportReport_Auto(CONFIG.reportRanges.standard);
}

function ImportAlt() {
  ImportReport_Auto(CONFIG.reportRanges.alt);
}

function ReImport(){
  runSecondScript();
}

function Stage1(){
  FMX_Doors_AutoImport_V8();
}

function Stage2(){
  stage2_filterProcessedData();
}

function Stage3(){
  copySelectedDataToAutoReport();
}

function Testing1(){
  PrintFormatTesting();
}

function Testing2(){
  
}

