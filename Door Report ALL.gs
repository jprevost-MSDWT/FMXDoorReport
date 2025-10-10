// Project Name: Door Report Full
// Project Version: 5.0
// Filename: Stage0 - Launcher.gs
// File Version: 5.00
// Description: A combined file of all .gs scripts for easy testing.

const CONFIG = {
  sheets: {
    import: "Import",    // Do NOT change. Also used in HTML
    helper1: "Output-Helper1",
    helper2: "Report Prep",
    report: "AutoReport",
    reportNotes: "AutoReport w/Notes",
    data: "Data"
  },
  reportRanges: {
    standard: 7,
    alt: 14
  },
  columnNames: {
    eventTime: ["Event time", "Starts"],
    name: ["Name"],
    buildings: ["Buildings", "Building"],
    status: ["Status"],
    resources: ["Resources"],
    eventDetails: ["Event Details"],
    doorNotes: ["Old ML Door Notes"],
    unlockTime1: ["Unlock Time"],
    lockTime1: ["Lock Time"],
    unlockTime2: ["Unlock Time."],
    lockTime2: ["Lock Time."],
    unlockTime3: ["Unlock Time.."],
    lockTime3: ["Lock Time.."],
    unlockTime4: ["Unlock Time..."],
    lockTime4: ["Lock Time..."],
    unlockTime5: ["Unlock Time...."],
    lockTime5: ["Lock Time...."],
    unlockTime6: ["Unlock Time....."],
    lockTime6: ["Lock Time....."],
    unlockTimeSpecial: ["WCHS Football/Baseball Locker Room Doors Unlock Time"],
    lockTimeSpecial: ["WCHS Football/Baseball Locker Room Doors Lock Time"],
    doorSet1: ['Clinic Doors', 'BV Doors', 'CR Doors', 'ECC Doors', 'EA Doors', 'EdCtr Doors', 'GC Doors', 'HA Doors', 'HP Doors', 'LO Doors', 'LP Doors', 'LA Doors', 'MO Doors', 'SH Doors', 'PO Doors', 'SB Doors', 'RP Doors', 'PR Doors', 'REN Doors', 'WCHS Doors', 'Special Doors'],
    doorSet2: ['EA Doors.', 'BV Doors.', 'CR Doors.', 'ECC Doors.', 'GC Doors.', 'HA Doors.', 'HP Doors.', 'LO Doors.', 'LA Doors.', 'LP Doors.', 'MO Doors.', 'PO Doors.', 'PR Doors.', 'REN Doors.', 'RP Doors.', 'SB Doors.', 'SH Door.', 'WCHS Doors.'],
    doorSet3: ['WCHS Doors..'],
    doorSet4: ['WCHS Doors...'],
    doorSet5: ['WCHS Doors....'],
    doorSet6: ['WCHS Doors.....'],
    doorSetSpecial: ['WCHS Football/Baseball Locker Room Doors']
  }
};

function onOpen() {
  VerifySheets();
  SpreadsheetApp.getUi()
      .createMenu('Report Menu')
      .addItem('Run Full Report', 'FullProcess')
      .addItem('Reprocess', 'ReProcess')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Manual Steps')
          .addItem('Import Standard (7 days)', 'ImportStandard')
          .addItem('Import Alt (14 days)', 'ImportAlt')
          .addItem('Import Box', 'Import')
          .addItem('Import & Proccess', 'ReImport')
          .addItem('Run Stage 1', 'Stage1')
          .addItem('Run Stage 2', 'Stage2')
          .addItem('Run Stage 3', 'Stage3'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Testing')
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

// This function now starts the download, which will trigger the import dialog to run the full process.
function FullProcess(){
  ImportReport_Auto(CONFIG.reportRanges.standard, true);
}

// This menu item now opens the dialog and tells it to run all stages after import.
function ReImport(){
  showImportDialog(true);
}

// This function remains the same, processing already imported data.
function ReProcess(){
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Starting Reprocess...');
    ss.toast('Starting Reprocess...', 'Status', -1);

    Logger.log('Starting Stage 1...');
    ss.toast('Starting Stage 1...', 'Status', -1);
    Stage1();
    
    Logger.log('Starting Stage 2...');
    ss.toast('Starting Stage 2...', 'Status', -1);
    Stage2();
    
    Logger.log('Starting Stage 3...');
    ss.toast('Starting Stage 3...', 'Status', -1);
    Stage3();
    
    Logger.log('Reprocess Complete!');
    ss.toast('Reprocess Complete!', 'Status', 5);

  } catch (e) {
    Logger.log('An error occurred: ' + e.message + '\n' + e.stack);
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('An Error Occurred', 'The script encountered a problem: \n\n' + e.message, ui.ButtonSet.OK);
    } catch (uiError) {
      throw e;
    }
  }
}

// A new central function to run all processing stages. This will be called from the HTML dialog.
function ProcessStages() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Starting Process Stages...');
    ss.toast('Starting Process Stages...', 'Status', -1);

    Logger.log('Starting Stage 1...');
    ss.toast('Starting Stage 1...', 'Status', -1);
    Stage1();
    
    Logger.log('Starting Stage 2...');
    ss.toast('Starting Stage 2...', 'Status', -1);
    Stage2();
    
    Logger.log('Starting Stage 3...');
    ss.toast('Starting Stage 3...', 'Status', -1);
    Stage3();
    
    Logger.log('Process Complete!');
    ss.toast('Process Complete!', 'Status', 5);

  } catch (e) {
    Logger.log('An error occurred during processing: ' + e.message + '\n' + e.stack);
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('An Error Occurred During Processing', 'The script encountered a problem: \n\n' + e.message, ui.ButtonSet.OK);
    } catch (uiError) {
      throw e;
    }
  }
}

function ImportStandard() {
  ImportReport_Auto(CONFIG.reportRanges.standard, false);
}

function ImportAlt() {
  ImportReport_Auto(CONFIG.reportRanges.alt, false);
}

// This menu item opens the dialog for import only.
function Import(){
  showImportDialog(false);
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
  // This function is in a separate testing file.
  // PrintFormatTesting(); 
}

function Testing2(){
  AddBlankDates("AutoReport w/Notes");
}
