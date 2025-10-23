// Project Name: Door Report Full
// Project Version: 5.0
// Filename: Door Report ALL.gs
// File Version: 5.07
// Description: A combined file of all .gs scripts for easy testing.

// =======================================================================================
// --- BEGIN Inserted Code from Stage0 - Launcher.gs ---
// =======================================================================================

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
    alt: 14,
    altSelected: 7 // Used for the "next 7 days" selection feature
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

function onOpenTrigger() {
  // This is triggered by an installed trigger.
  VerifySheets();
  Menu(); 
  SideMenu(); 
}

// ==========================
// Menues
// ==========================

function Menu() {
  SpreadsheetApp.getUi()
    .createMenu('Report Menu')
    .addItem('Run Full Report', 'FullProcess')
    .addItem('ReSort Printsheet', 'ReSortPrint')
    .addItem('Reprocess', 'ReProcess')
    .addItem('Resort Report Prep Sheet', 'ResortStage2')
    .addItem('Reopen Sidebar Menu', 'SideMenu')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Manual Steps')
      .addItem('Import Standard (7 days)', 'ImportStandard')
      .addItem('Import Alt (14 days)', 'ImportAlt')
      .addItem('Import Box', 'showImportDialog')
      .addItem('Import & Proccess', 'ReImport')
      .addItem('Run Stage 1', 'Stage1')
      .addItem('Run Stage 2', 'Stage2')
      .addItem('Run Stage 3', 'Stage3')
      .addItem('Select Next 7 Days', 'altSelectedDays')) // Updated Menu Item Text
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Testing')
      .addItem('Testing1', 'Testing1')
      .addItem('Testing2', 'Testing2'))
    .addToUi();
}


// Sidebar menu features

function SideMenu() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Report Actions');
  SpreadsheetApp.getUi().showSidebar(html);
}

// New navigation functions
function goToAutoReport() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoReport').activate();
}

function goToAutoReportWithNotes() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AutoReport w/Notes').activate();
}

function goToReportPrepSheet() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report Prep').activate();
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
// The original 'Import' function is no longer needed as the menu calls showImportDialog directly.
// function Import(){
//  showImportDialog(false);
// }

function Stage1(){
  FMX_Doors_AutoImport_V8();
}

function Stage2(){
  Stage2_InitialFilter();
  Stage2_ResortAndFormat();
}

function ResortStage2() {
  Stage2_ResortAndFormat();
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

// =======================================================================================
// --- END Inserted Code from Stage0 - Launcher.gs ---
// =======================================================================================


// =======================================================================================
// --- BEGIN Inserted Code from Stage0 Import.gs ---
// =======================================================================================

// This function now receives the shouldProcess flag and passes it to the dialog.
function runSecondScript(shouldProcess) {
  showImportDialog(shouldProcess);
}

function formatDate(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}

// This function now accepts the shouldProcess flag to pass it along the chain.
function ImportReport_Auto(days, shouldProcess) {
  var today = new Date();
  var futureDate = new Date();
  futureDate.setDate(today.getDate() + days);

  var fromDate = formatDate(today);
  var toDate = formatDate(futureDate);

  var url = 'https://warrenk12.gofmx.com/scheduling/occurrences?format=csv&useOnlySelectedColumns=False&from=' + fromDate + '&to=' + toDate;

  // The client-side script now calls runSecondScript with the shouldProcess flag.
  const htmlScript = `
    <script>
      window.open('${url}', '_blank');
      google.script.run
        .withSuccessHandler(google.script.host.close)
        .runSecondScript(${shouldProcess});
    </script>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlScript)
    .setWidth(100)
    .setHeight(100);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Report...');
}

// This is the new central function for showing the dialog.
// It uses an HTML template to pass the 'shouldProcess' variable to the dialog's javascript.
function showImportDialog(shouldProcess) {
  const template = HtmlService.createTemplateFromFile('IMPORTdialog');
  template.shouldProcess = shouldProcess || false; // Pass the flag to the template
  const html = template.evaluate()
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import File from Computer');
}

function importData(fileContent, fileType) {
  const sheetName = 'Import';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  sheet.clear();

  let data;
  let delimiter;

  if (fileType === 'text/csv' || fileContent.includes(',')) {
    delimiter = ',';
  } else if (fileType === 'text/tab-separated-values' || fileContent.includes('\t')) {
    delimiter = '\t';
  } else {
    delimiter = ',';
  }

  try {
    data = Utilities.parseCsv(fileContent, delimiter);

    if (data.length === 0) {
      throw new Error('No data found in the file.');
    }

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    for (let i = 1; i <= data[0].length; i++) {
      sheet.autoResizeColumn(i);
    }

    return `Success! ${data.length} rows imported into the "${sheetName}" sheet.`;
  } catch (e) {
    console.error('Error processing file: ' + e.toString());
    return 'Error: Could not parse the file. Please ensure it is a valid CSV or TXT file.';
  }
}

// =======================================================================================
// --- END Inserted Code from Stage0 Import.gs ---
// =======================================================================================


// =======================================================================================
// --- BEGIN Inserted Code from Stage1.gs ---
// =======================================================================================

// This script processes FMX door data from "Import" to "Output-Helper1", replaces building names, removes duplicates, and sorts.

function FMX_Doors_AutoImport_V8() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName(CONFIG.sheets.import);
  var outputSheet = ss.getSheetByName(CONFIG.sheets.helper1);
  var dataSheet = ss.getSheetByName(CONFIG.sheets.data);

  if (!inputSheet) {
    throw new Error('Source sheet "' + CONFIG.sheets.import + '" not found!');
  }
  if (!outputSheet) {
    throw new Error('Destination sheet "' + CONFIG.sheets.helper1 + '" not found!');
  }
  if (!dataSheet) {
    throw new Error('Lookup sheet "' + CONFIG.sheets.data + '" not found!');
  }

  // --- Create a lookup map from the "Data" sheet ---
  var buildingMap = {};
  var dataValues = dataSheet.getRange("A2:B" + dataSheet.getLastRow()).getValues();
  for (var i = 0; i < dataValues.length; i++) {
    var originalBuilding = dataValues[i][0];
    var newBuilding = dataValues[i][1];
    if (originalBuilding) {
      buildingMap[originalBuilding] = newBuilding;
    }
  }

  var inputData = inputSheet.getDataRange().getValues();
  var inputHeaders = inputData[0];

  var outputData = [
    [
      'Event Date', 'Event Time', 'Building', 'Areas', 'Name', 'ID', 'Status', 'Door Times',
      'Notes',
      'Combined Door Times (Set 1)',
      'Combined Door Times (Set 2)',
      'Combined Door Times (Set 3)',
      'Combined Door Times (Set 4)',
      'Combined Door Times (Set 5)',
      'Combined Door Times (Set 6)',
      'Combined Door Times (Special)'
    ],
  ];

  // --- Find column indexes using the new helper function and CONFIG ---
  var eventTimeCol = getColumnIndex(inputHeaders, 'eventTime');
  var nameCol = getColumnIndex(inputHeaders, 'name');
  var buildingsCol = getColumnIndex(inputHeaders, 'buildings');
  var statusCol = getColumnIndex(inputHeaders, 'status');
  var resourcesCol = getColumnIndex(inputHeaders, 'resources');
  var eventDetailsCol = getColumnIndex(inputHeaders, 'eventDetails');
  var doorNotesCol = getColumnIndex(inputHeaders, 'doorNotes');
  
  var unlockTimeCol = getColumnIndex(inputHeaders, 'unlockTime1');
  var lockTimeCol = getColumnIndex(inputHeaders, 'lockTime1');
  var unlockTimeDotCol = getColumnIndex(inputHeaders, 'unlockTime2');
  var lockTimeDotCol = getColumnIndex(inputHeaders, 'lockTime2');
  var unlockTimeDotDotCol = getColumnIndex(inputHeaders, 'unlockTime3');
  var lockTimeDotDotCol = getColumnIndex(inputHeaders, 'lockTime3');
  var unlockTimeDotDotDotCol = getColumnIndex(inputHeaders, 'unlockTime4');
  var lockTimeDotDotDotCol = getColumnIndex(inputHeaders, 'lockTime4');
  var unlockTimeDotDotDotDotCol = getColumnIndex(inputHeaders, 'unlockTime5');
  var lockTimeDotDotDotDotCol = getColumnIndex(inputHeaders, 'lockTime5');
  var unlockTimeDotDotDotDotDotCol = getColumnIndex(inputHeaders, 'unlockTime6');
  var lockTimeDotDotDotDotDotCol = getColumnIndex(inputHeaders, 'lockTime6');
  var unlockTimeSpecialCol = getColumnIndex(inputHeaders, 'unlockTimeSpecial');
  var lockTimeSpecialCol = getColumnIndex(inputHeaders, 'lockTimeSpecial');

  var doorColIndexes1 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSet1);
  var doorColIndexes2 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSet2);
  var doorColIndexes3 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSet3);
  var doorColIndexes4 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSet4);
  var doorColIndexes5 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSet5);
  var doorColIndexes6 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSet6);
  var doorColIndexes7 = getColumnIndexes(inputHeaders, CONFIG.columnNames.doorSetSpecial);

  var seenRows = {};
  var timeZone = ss.getSpreadsheetTimeZone();

  for (var i = 1; i < inputData.length; i++) {
    var row = inputData[i];

    var eventTimeString = row[eventTimeCol];
    var formattedEventDate = "";
    var extractedEventTime = "";

    if (eventTimeString) {
      var dateObject = new Date(eventTimeString);

      if (!isNaN(dateObject.getTime())) {
        formattedEventDate = Utilities.formatDate(dateObject, timeZone, "MM/dd");
        extractedEventTime = Utilities.formatDate(dateObject, timeZone, "h:mm a").toLowerCase();
      } else if (typeof eventTimeString === 'string') {
        var datePart = eventTimeString.split(',').slice(0, 3).join(',');
        var fallbackDateObject = new Date(datePart);
        if (!isNaN(fallbackDateObject.getTime())) {
          formattedEventDate = Utilities.formatDate(fallbackDateObject, timeZone, "MM/dd");
        } else {
          formattedEventDate = eventTimeString;
        }
        
        var timePartMatch = eventTimeString.match(/\d{1,2}:\d{2}(am|pm)/i);
        if (timePartMatch) {
          extractedEventTime = timePartMatch[0];
        }
      } else {
        formattedEventDate = eventTimeString;
      }
    }

    var name = cleanString(row[nameCol]);
    var id = "";
    if (name && name.indexOf("-") > -1) {
      var splitName = name.split("-");
      id = splitName[0].trim();
      name = splitName.slice(1).join("-").trim();
    } else if (name) {
      id = name.trim();
      name = "";
    }

    var buildings = cleanString(row[buildingsCol]);
    if (buildingMap[buildings]) {
      buildings = buildingMap[buildings];
    }

    var status = cleanString(row[statusCol]);
    var resources = cleanResourcesString(row[resourcesCol]);
    var eventDetails = cleanString(row[eventDetailsCol]);
    var originalDoorNotes = cleanString(row[doorNotesCol]);
    var combinedNotes = [eventDetails, originalDoorNotes].filter(Boolean).join('\n---\n');

    var doorNotes = cleanString(row[doorNotesCol]);
    if (typeof doorNotes === 'string' && doorNotes) {
      var textToRemove1 = / \(Open\/Close times required - be specific\.\): Yes/g;
      var textToRemove2 = / \(Open\/Close times required - be specific\.\): Yes/g;
      doorNotes = doorNotes.replace(textToRemove1, '').replace(textToRemove2, '').replace(/, Door/g, '\nDoor').trim();
    }

    var doors1 = combineDoorValues(row, CONFIG.columnNames.doorSet1, doorColIndexes1);
    var unlockTime1 = formatTimeValue(row[unlockTimeCol], timeZone);
    var lockTime1 = formatTimeValue(row[lockTimeCol], timeZone);
    var combinedTimes1 = formatDoorTimes(doors1, unlockTime1, lockTime1);

    var doors2 = combineDoorValues(row, CONFIG.columnNames.doorSet2, doorColIndexes2);
    var unlockTime2 = formatTimeValue(row[unlockTimeDotCol], timeZone);
    var lockTime2 = formatTimeValue(row[lockTimeDotCol], timeZone);
    var combinedTimes2 = formatDoorTimes(doors2, unlockTime2, lockTime2);

    var doors3 = combineDoorValues(row, CONFIG.columnNames.doorSet3, doorColIndexes3);
    var unlockTime3 = formatTimeValue(row[unlockTimeDotDotCol], timeZone);
    var lockTime3 = formatTimeValue(row[lockTimeDotDotCol], timeZone);
    var combinedTimes3 = formatDoorTimes(doors3, unlockTime3, lockTime3);

    var doors4 = combineDoorValues(row, CONFIG.columnNames.doorSet4, doorColIndexes4);
    var unlockTime4 = formatTimeValue(row[unlockTimeDotDotDotCol], timeZone);
    var lockTime4 = formatTimeValue(row[lockTimeDotDotDotCol], timeZone);
    var combinedTimes4 = formatDoorTimes(doors4, unlockTime4, lockTime4);

    var doors5 = combineDoorValues(row, CONFIG.columnNames.doorSet5, doorColIndexes5);
    var unlockTime5 = formatTimeValue(row[unlockTimeDotDotDotDotCol], timeZone);
    var lockTime5 = formatTimeValue(row[lockTimeDotDotDotDotCol], timeZone);
    var combinedTimes5 = formatDoorTimes(doors5, unlockTime5, lockTime5);

    var doors6 = combineDoorValues(row, CONFIG.columnNames.doorSet6, doorColIndexes6);
    var unlockTime6 = formatTimeValue(row[unlockTimeDotDotDotDotDotCol], timeZone);
    var lockTime6 = formatTimeValue(row[lockTimeDotDotDotDotDotCol], timeZone);
    var combinedTimes6 = formatDoorTimes(doors6, unlockTime6, lockTime6);

    var doors7 = combineDoorValues(row, CONFIG.columnNames.doorSetSpecial, doorColIndexes7);
    var unlockTime7 = formatTimeValue(row[unlockTimeSpecialCol], timeZone);
    var lockTime7 = formatTimeValue(row[lockTimeSpecialCol], timeZone);
    var combinedTimes7 = formatDoorTimes(doors7, unlockTime7, lockTime7);

    var doorTimesArray = [
      doorNotes,
      combinedTimes1,
      combinedTimes2,
      combinedTimes3,
      combinedTimes4,
      combinedTimes5,
      combinedTimes6,
      combinedTimes7
    ];
    var finalDoorTimes = doorTimesArray.filter(function(value) {
      return value && value.toString().trim() !== '';
    }).join('\n');

    var rowKey = [
        formattedEventDate, extractedEventTime, buildings, resources, name, status, finalDoorTimes
    ].join('|');

    if (!seenRows[rowKey]) {
      seenRows[rowKey] = true;

      outputData.push([
        formattedEventDate, extractedEventTime, buildings, resources, name, id, status, finalDoorTimes,
        combinedNotes,
        combinedTimes1, combinedTimes2, combinedTimes3, combinedTimes4, combinedTimes5, combinedTimes6, combinedTimes7
      ]);
    }
  }

  // No try/catch here; let errors be caught by the parent function (ProcessStages/ReProcess)
  if (outputData.length > 1) {
    outputSheet.clear();
    outputSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    var oldFilter = outputSheet.getFilter();
    if (oldFilter) {
      oldFilter.remove();
    }
    var dataRange = outputSheet.getDataRange();
    dataRange.createFilter();
    var rangeToSort = outputSheet.getRange(2, 1, outputSheet.getLastRow() - 1, outputSheet.getLastColumn());
    rangeToSort.sort([
      {column: 1, ascending: true},
      {column: 2, ascending: true}
    ]);
  }
}

function getColumnIndex(headers, configKey) {
  const names = CONFIG.columnNames[configKey];
  if (!names) {
    return -1;
  }
  for (const name of names) {
    const index = headers.indexOf(name);
    if (index !== -1) {
      return index;
    }
  }
  return -1;
}

function formatTimeValue(value, timeZone) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, timeZone, "hh:mm a");
  }
  return cleanString(value);
}

function formatDoorTimes(doors, unlock, lock) {
  var parts = [];
  if (doors) parts.push(doors);
  var timePart = "";
  if (unlock && lock) timePart = unlock + " - " + lock;
  else if (unlock) timePart = unlock;
  else if (lock) timePart = lock;
  if (timePart) parts.push(timePart);
  return parts.join(" / ");
}

function getColumnIndexes(headers, columnNames) {
  var indexes = {};
  for (var i = 0; i < columnNames.length; i++) {
    var colName = columnNames[i];
    indexes[colName] = headers.indexOf(colName);
  }
  return indexes;
}

function cleanString(value) {
  if (typeof value !== 'string' || !value) {
    return "";
  }
  return value.split(/\r\n|\r|\n/).map(function(line) {
    return line.replace(/\s+/g, ' ').trim();
  }).filter(function(line) {
    return line !== '';
  }).join('\n');
}

function cleanResourcesString(value) {
  if (typeof value !== 'string' || !value) {
    return "";
  }
  return value.split(/\r\n|\r|\n/).map(function(line) {
    return line.replace(/\s+/g, ' ').trim();
  }).filter(function(line) {
    return line !== '';
  }).join(', ');
}

function combineDoorValues(row, columnNames, indexes) {
  var currentDoors = [];
  for (var i = 0; i < columnNames.length; i++) {
    var doorName = columnNames[i];
    var colIndex = indexes[doorName];
    if (colIndex !== -1) {
      var cellValue = cleanString(row[colIndex]);
      if (cellValue !== "") {
        currentDoors.push(cellValue);
      }
    }
  }
  return currentDoors.join(", ");
}


// =======================================================================================
// --- END Inserted Code from Stage1.gs ---
// =======================================================================================


// =======================================================================================
// --- BEGIN Inserted Code from Stage2.gs ---
// =======================================================================================

/**
 * New function to select rows in Report Prep based on the next N days.
 */
function altSelectedDays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.helper2);
  
  if (!sheet || sheet.getLastRow() < 2) {
    ss.toast("No data found on the Report Prep sheet to select.", "Status", 5);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  const selectedColIndex = headers.indexOf("Selected");
  const eventDateColIndex = headers.indexOf("Event Date");
  const doorTimesColIndex = headers.indexOf("Door Times");
  
  if (selectedColIndex === -1 || eventDateColIndex === -1 || doorTimesColIndex === -1) {
    throw new Error("Could not find one of the required columns: 'Selected', 'Event Date', or 'Door Times'.");
  }
  
  const daysToLookForward = CONFIG.reportRanges.altSelected;
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Start of today

  const futureDate = new Date();
  futureDate.setDate(today.getDate() + daysToLookForward);
  futureDate.setHours(23, 59, 59, 999); // End of the 7th day from now
  
  const selections = [];
  for (const row of data) {
    const eventDate = parseDate(row[eventDateColIndex]);
    const hasDoorTimes = row[doorTimesColIndex] && row[doorTimesColIndex].toString().trim() !== '';
    let shouldBeChecked = false;

    if (eventDate && hasDoorTimes) {
      if (eventDate >= today && eventDate <= futureDate) {
        shouldBeChecked = true;
      }
    }
    selections.push([shouldBeChecked]);
  }
  
  // Efficiently update all checkboxes at once
  sheet.getRange(2, selectedColIndex + 1, selections.length, 1).setValues(selections);
  
  ss.toast(`Selections have been updated for the next ${daysToLookForward} days.`, "Complete", 5);
}


function Stage2_InitialFilter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.sheets.helper1);
  const destinationSheet = ss.getSheetByName(CONFIG.sheets.helper2);

  if (!sourceSheet) {
    throw new Error('Source sheet "' + CONFIG.sheets.helper1 + '" not found!');
  }
  if (!destinationSheet) {
    throw new Error('Destination sheet "' + CONFIG.sheets.helper2 + '" not found!');
  }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift();

  const statusColIndex = headers.indexOf('Status');
  if (statusColIndex === -1) {
    throw new Error('"Status" column could not be found in the source sheet "' + CONFIG.sheets.helper1 + '".');
  }
  const excludedStatuses = ["Declined", "Canceled", "Deleted", "Bulk Declined", "Bulk Canceled", "Bulk Deleted"];
  const filteredRows = data.filter(row => !excludedStatuses.includes(row[statusColIndex]));

  const columnsToRemove = headers.reduce((acc, header, index) => {
    if (header.includes("Combined Door Times")) {
      acc.push(index);
    }
    return acc;
  }, []);

  const newHeaders = headers.filter((_, index) => !columnsToRemove.includes(index));
  let processedData = filteredRows.map(row => row.filter((_, index) => !columnsToRemove.includes(index)));

  const targetDate = calculateTargetDate();
  targetDate.setHours(0, 0, 0, 0);

  const eventDateColIndex_forSelect = newHeaders.indexOf('Event Date');
  const doorTimesColIndex_forSelect = newHeaders.indexOf('Door Times');
  newHeaders.unshift("Selected");

  processedData = processedData.map(row => {
    let shouldBeChecked = false;
    const eventDate = eventDateColIndex_forSelect !== -1 ? parseDate(row[eventDateColIndex_forSelect]) : null;
    const hasDoorTimes = doorTimesColIndex_forSelect !== -1 && row[doorTimesColIndex_forSelect] && row[doorTimesColIndex_forSelect].toString().trim() !== '';

    if (eventDate && hasDoorTimes) {
      eventDate.setHours(0, 0, 0, 0);
      if (eventDate <= targetDate) {
        shouldBeChecked = true;
      }
    }
    return [shouldBeChecked, ...row];
  });
  
  const finalData = [newHeaders, ...processedData];
  destinationSheet.clear();
  destinationSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
}

function Stage2_ResortAndFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.helper2);

  if (!sheet || sheet.getLastRow() < 1) {
    // If the sheet is empty or doesn't exist, there's nothing to sort.
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const idColIndex = headers.indexOf('ID');
  const areasColIndex = headers.indexOf('Areas');
  
  let uniqueData = [];
  const seenRows = new Set();
  data.forEach(row => {
    const rowKey = row.filter((_, index) => index !== 0 && index !== idColIndex && index !== areasColIndex).join('|');
    if (!seenRows.has(rowKey)) {
      seenRows.add(rowKey);
      uniqueData.push(row);
    }
  });

  const eventDateColIndex = headers.indexOf('Event Date');
  const buildingColIndex = headers.indexOf('Building');
  const eventTimeColIndex = headers.indexOf('Event Time');
  const statusColIndex_forSort = headers.indexOf('Status');
  const doorTimesColIndex_forSort = headers.indexOf('Door Times');
  const selectedColIndex = 0;

  uniqueData.sort((a, b) => {
    const aIsSelected = a[selectedColIndex] === true;
    const bIsSelected = b[selectedColIndex] === true;
    if (aIsSelected !== bIsSelected) return aIsSelected ? -1 : 1;

    if (!aIsSelected) {
      const aHasDoorTimes = a[doorTimesColIndex_forSort] && a[doorTimesColIndex_forSort].toString().trim() !== '';
      const bHasDoorTimes = b[doorTimesColIndex_forSort] && b[doorTimesColIndex_forSort].toString().trim() !== '';
      if (aHasDoorTimes !== bHasDoorTimes) return aHasDoorTimes ? -1 : 1;
    }

    const statusA = (a[statusColIndex_forSort] || '').toUpperCase();
    const statusB = (b[statusColIndex_forSort] || '').toUpperCase();
    const aIsPending = statusA.includes("PENDING") && statusA.includes("APPROVAL");
    const bIsPending = statusB.includes("PENDING") && statusB.includes("APPROVAL");
    if (aIsPending !== bIsPending) return aIsPending ? -1 : 1;
    
    // Updated Date sorting logic
    const dateA = parseDate(a[eventDateColIndex]);
    const dateB = parseDate(b[eventDateColIndex]);
    if (dateA && !dateB) return -1; // a has a date, b doesn't, so a comes first
    if (!dateA && dateB) return 1;  // b has a date, a doesn't, so b comes first
    if (dateA && dateB && (dateA.getTime() !== dateB.getTime())) {
      return dateA.getTime() - dateB.getTime();
    }
    
    const buildingA = a[buildingColIndex] || '';
    const buildingB = b[buildingColIndex] || '';
    if (buildingA.localeCompare(buildingB) !== 0) return buildingA.localeCompare(buildingB);
    
    const timeA = parseTime(a[eventTimeColIndex]);
    const timeB = parseTime(b[eventTimeColIndex]);
    if (timeA && timeB) return timeA.getTime() - timeB.getTime();

    return 0;
  });

  uniqueData.forEach(row => {
    const status = (row[statusColIndex_forSort] || '').toUpperCase();
    if (status.includes("PENDING") && status.includes("APPROVAL")) {
      row[selectedColIndex] = false;
    }
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (uniqueData.length > 0) {
    sheet.getRange(2, 1, uniqueData.length, headers.length).setValues(uniqueData);
  }
  
  Stage2_format(sheet, uniqueData.length, headers);
}


/**
 * =======================================================================================
 * --- HELPER FUNCTIONS ---
 * =======================================================================================
 */

/**
 * Helper function to robustly parse date values which could be Date objects or strings.
 */
function parseDate(dateVal) {
  if (dateVal instanceof Date) {
    return dateVal;
  }
  if (typeof dateVal === 'string' && dateVal.includes('/')) {
    var parts = dateVal.split('/');
    var currentYear = new Date().getFullYear();
    var date = new Date(currentYear, parseInt(parts[0], 10) - 1, parseInt(parts[1], 10));
    if (date < new Date() && new Date().getMonth() - date.getMonth() > 6) { // Heuristic for year rollover
        date.setFullYear(currentYear + 1);
    }
    return date;
  }
  return null;
}


/**
 * Helper function to parse 'h:mm am/pm' into a comparable Date object.
 */
function parseTime(timeStr) {
  if (!timeStr || typeof timeStr !== 'string') {
    return null;
  }
  return new Date('1970/01/01 ' + timeStr.replace(' ', '').toUpperCase());
}

/**
 * Applies all necessary formatting to the destination sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to format.
 * @param {number} numDataRows The number of data rows (excluding the header).
 * @param {string[]} headers The array of header strings.
 */
function Stage2_format(sheet, numDataRows, headers) {
  var numCols = headers.length;
  if (numCols === 0) return;

  sheet.clearConditionalFormatRules();
  var filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  sheet.getRange(1, 1, 1, numCols).setFontWeight("bold");

  // --- BEGIN JSON-Based Formatting ---
  const formattingInfo = {
    widths: [61, 40, 55, 62, 160, 143, 58, 112, 350, 729],
    fontSizes: [10, 10, 8, 10, 8, 10, 10, 8, 10, 8],
    horizontalAlignments: ["center", "center", "center", "center", "left", "left", "left", "left", "left", "left"],
    verticalAlignments: ["middle", "middle", "middle", "middle", "middle", "middle", "middle", "middle", "middle", "bottom"],
    textWraps: [true, true, true, true, true, true, true, true, true, true]
  };

  for (let i = 0; i < formattingInfo.widths.length && i < numCols; i++) {
    sheet.setColumnWidth(i + 1, formattingInfo.widths[i]);
  }

  if (numDataRows > 0) {
    for (let i = 0; i < numCols; i++) {
      const colRange = sheet.getRange(2, i + 1, numDataRows, 1);
      if (formattingInfo.fontSizes[i]) colRange.setFontSize(formattingInfo.fontSizes[i]);
      if (formattingInfo.horizontalAlignments[i]) colRange.setHorizontalAlignment(formattingInfo.horizontalAlignments[i]);
      if (formattingInfo.verticalAlignments[i]) colRange.setVerticalAlignment(formattingInfo.verticalAlignments[i]);
      if (formattingInfo.textWraps[i] !== undefined) colRange.setWrap(formattingInfo.textWraps[i]);
    }

    const eventDateColIndex = headers.indexOf('Event Date');
    if (eventDateColIndex !== -1) {
      sheet.getRange(2, eventDateColIndex + 1, numDataRows, 1).setNumberFormat('mm/dd');
    }

    const eventTimeColIndex = headers.indexOf('Event Time');
    if (eventTimeColIndex !== -1) {
      sheet.getRange(2, eventTimeColIndex + 1, numDataRows, 1).setNumberFormat('h:mm am/pm');
    }

    var dataRange = sheet.getRange(2, 1, numDataRows, numCols);
    var rules = [];
    var statusColIndex = headers.indexOf('Status');
    
    var notesColIndex = headers.indexOf('Notes');
    if (notesColIndex !== -1) {
      sheet.getRange(2, notesColIndex + 1, numDataRows, 1).setFontColor("#b7b7b7");
    }

    if (statusColIndex !== -1) {
      var statusColumnLetter = String.fromCharCode('A'.charCodeAt(0) + statusColIndex);
      var statusColumnRange = sheet.getRange(2, statusColIndex + 1, numDataRows, 1);
      
      var rule_admin_approval_false = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=FALSE, ISNUMBER(SEARCH("Admin", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#d9ead3").setFontColor("#b7b7b7").setRanges([statusColumnRange]).build();
      rules.push(rule_admin_approval_false);

      var rule_pending_approval_false = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=FALSE, ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#fff2cc").setFontColor("#b7b7b7").setRanges([statusColumnRange]).build();
      rules.push(rule_pending_approval_false);
      
      var rule_admin_approval_true = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=TRUE, ISNUMBER(SEARCH("Admin", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#b6d7a8").setFontColor("#000000").setRanges([dataRange]).build();
      rules.push(rule_admin_approval_true);

      var rule_pending_approval_true = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=TRUE, ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#ffd966").setFontColor("#000000").setRanges([dataRange]).build();
      rules.push(rule_pending_approval_true);
    }

    var rule_A_is_false = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$A2=FALSE")
      .setBackground("#efefef").setFontColor("#b7b7b7").setRanges([dataRange]).build();
    rules.push(rule_A_is_false);

    sheet.setConditionalFormatRules(rules);

    sheet.getRange(2, 1, numDataRows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
  
  sheet.getRange(1, 1, numDataRows + 1, numCols).createFilter();

  var totalRows = numDataRows + 1;
  var maxRows = sheet.getMaxRows();
  if (maxRows > totalRows) {
    sheet.deleteRows(totalRows + 1, maxRows - totalRows);
  }
}

/**
 * Calculates a target date based on the current day of the week.
 */
function calculateTargetDate() {
  var today = new Date();
  var dayOfWeek = today.getDay();
  var targetDate = new Date(today);

  if (dayOfWeek === 5) { // Friday
    targetDate.setDate(today.getDate() + 4); // Next Tuesday
  } else { // Any other day
    var daysUntilFriday = (5 - dayOfWeek + 7) % 7;
    if (daysUntilFriday === 0) daysUntilFriday = 7;
    targetDate.setDate(today.getDate() + daysUntilFriday);
  }
  
  Logger.log('Calculated Target Date: ' + targetDate);
  return targetDate;
}

// =======================================================================================
// --- END Inserted Code from Stage2.gs ---
// =======================================================================================


// =======================================================================================
// --- BEGIN Inserted Code from Stage3.gs ---
// =======================================================================================

// This script generates two reports ("AutoReport" and "AutoReport w/Notes") from Output-Helper2, applying formatting and trimming.

function copySelectedDataToAutoReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.sheets.helper2);

  if (!sourceSheet) {
    throw new Error('The sheet "' + CONFIG.sheets.helper2 + '" was not found. Please check the name and try again.');
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData.shift();

  const selectedColumnIndex = sourceHeaders.indexOf("Selected");
  if (selectedColumnIndex === -1) {
    throw new Error('A column named "Selected" was not found in "' + CONFIG.sheets.helper2 + '".');
  }

  const selectedRows = sourceData.filter(row => row[selectedColumnIndex] === true);

  processAndWriteData(ss, sourceHeaders, selectedRows, CONFIG.sheets.report, false);
  processAndWriteData(ss, sourceHeaders, selectedRows, CONFIG.sheets.reportNotes, true);
}

function processAndWriteData(ss, sourceHeaders, selectedRows, destinationSheetName, includeAllColumns) {
  const destinationSheet = ss.getSheetByName(destinationSheetName);
  if (!destinationSheet) {
    throw new Error(`The destination sheet "${destinationSheetName}" was not found.`);
  }

  let columnMapping = [
    { source: "Event Date", destination: "Date" },
    { source: "Event Time", destination: "Time" },
    { source: "Building", destination: "Building" },
    { source: "Areas", destination: "Areas" },
    { source: "Name", destination: "Name" },
    { source: "ID", destination: "ID" },
    { source: "Door Times", destination: "Door Times" },
    { source: "Status", destination: "Status" },
    { source: "Notes", destination: "Notes" }
  ];
  
  if (!includeAllColumns) {
    columnMapping = columnMapping.filter(mapping => mapping.source !== "Notes" && mapping.source !== "Areas" && mapping.source !== "Status");
  }

  const sourceColumnIndices = columnMapping.map(mapping => {
    const index = sourceHeaders.indexOf(mapping.source);
    if (index === -1) {
      throw new Error(`Column "${mapping.source}" not found in "` + CONFIG.sheets.helper2 + `".`);
    }
    return index;
  });

  const outputData = selectedRows.map(row => {
    return sourceColumnIndices.map(index => row[index]);
  });

  const destinationHeaders = columnMapping.map(mapping => mapping.destination);

  destinationSheet.clear();
  destinationSheet.getRange(1, 1, 1, destinationHeaders.length).setValues([destinationHeaders]);

  if (outputData.length > 0) {
    destinationSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  }
  
  PrintPageFullFormatting(destinationSheetName);
}

function PrintPageFullFormatting(sheetName) {
  PrintPageSort(sheetName);
  AddBlankDates(sheetName);
  PrintPageRows(sheetName);
  PrintPageFormattingONLY(sheetName);
  trimSheet(sheetName);
}

function PrintPageRows(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 2) {
    return; // Not enough data to compare rows
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = headers.indexOf("Date");

  if (dateColIndex === -1) {
    console.log("Date column not found. Cannot insert rows.");
    return;
  }

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const dataValues = dataRange.getValues();

  // Iterate backwards to avoid issues with row index changes after insertion
  for (let i = dataValues.length - 1; i > 0; i--) {
    const currentDate = new Date(dataValues[i][dateColIndex]);
    const previousDate = new Date(dataValues[i - 1][dateColIndex]);

    if (!isNaN(currentDate.getTime()) && !isNaN(previousDate.getTime())) {
      const currentDay = currentDate.getDay();   // Day of the week for the current row (e.g., Sat)
      const previousDay = previousDate.getDay(); // Day of the week for the row above it (e.g., Fri)

      // Check for Friday (5) to Saturday (6) transition
      const isWeekendBreak = previousDay === 5 && currentDay === 6;
      // Check for Sunday (0) to Monday (1) transition
      const isWeekStartBreak = previousDay === 0 && currentDay === 1;

      if (isWeekendBreak || isWeekStartBreak) {
        // The row index is i + 1 because data starts at row 2 and i is 0-indexed.
        sheet.insertRowAfter(i + 1);
      }
    }
  }
}

function AddBlankDates(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return; // Not enough data to process
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = headers.indexOf("Date");

  if (dateColIndex === -1) {
    console.log("Date column not found. Cannot add blank dates.");
    return;
  }

  const dataValues = sheet.getRange(2, dateColIndex + 1, sheet.getLastRow() - 1, 1).getValues();

  // Iterate backwards to safely insert rows
  for (let i = dataValues.length - 1; i > 0; i--) {
    const currentDate = new Date(dataValues[i][0]);
    const previousDate = new Date(dataValues[i - 1][0]);

    if (isNaN(currentDate.getTime()) || isNaN(previousDate.getTime())) {
      continue; // Skip if dates are invalid
    }

    const oneDay = 24 * 60 * 60 * 1000;
    // Calculate the difference in days, ignoring time components
    const diffDays = Math.round((currentDate.setHours(0, 0, 0, 0) - previousDate.setHours(0, 0, 0, 0)) / oneDay);

    if (diffDays > 1) {
      // Loop to insert a row for each missing day
      for (let j = diffDays - 1; j >= 1; j--) {
        const missingDate = new Date(previousDate.getTime());
        missingDate.setDate(missingDate.getDate() + j);

        // Row index in the sheet is i + 1 (data starts at row 2, i is 0-indexed)
        const insertRowIndex = i + 1; 
        sheet.insertRowAfter(insertRowIndex);
        sheet.getRange(insertRowIndex + 1, dateColIndex + 1).setValue(missingDate);
      }
    }
  }
}

function PrintPageSort(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return; // No data to sort
  }

  const range = sheet.getDataRange();
  const dataToSort = range.offset(1, 0, range.getNumRows() - 1);
  dataToSort.sort([
    { column: 1, ascending: true },
    { column: 3, ascending: true },
    { column: 2, ascending: true }
  ]);
}

function PrintPageFormattingONLY(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`The sheet "${sheetName}" was not found for formatting.`);
  }

  if (sheet.getLastRow() === 0) {
    return;
  }

  const range = sheet.getDataRange();

  if (range.getNumRows() <= 1) {
    if (range.getNumRows() === 1) {
      range.setFontColor("#000000");
      sheet.getRange(1, 1, 1, range.getNumColumns()).setBackground("#b7b7b7").setFontWeight("bold");
    }
    range.setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
    return;
  }

  range.setFontColor("#000000");
  sheet.getRange(1, 1, 1, range.getNumColumns()).setBackground("#b7b7b7").setFontWeight("bold");

  const dataRange = sheet.getRange(2, 1, range.getNumRows() - 1, range.getNumColumns());
  const backgrounds = [];
  for (let i = 0; i < dataRange.getNumRows(); i++) {
    backgrounds.push(new Array(dataRange.getNumColumns()).fill(i % 2 === 0 ? "#ffffff" : "#d9d9d9"));
  }
  dataRange.setBackgrounds(backgrounds);

  // --- New Dynamic Formatting Section ---
  const formatConfig = {
    "Date":       { width: 50,  dataFontSize: 10, dataAlign: "center", headerAlign: "center", numberFormat: "m/d",          wrap: false },
    "Time":       { width: 50,  dataFontSize: 8,  dataAlign: "center", headerAlign: "center", numberFormat: "h:mm am/pm",   wrap: false },
    "Building":   { width: 60,  dataFontSize: 10, dataAlign: "center", headerAlign: "center", numberFormat: null,           wrap: false },
    "Name":       { width: 100, dataFontSize: 6,  dataAlign: "left",   headerAlign: "center", numberFormat: null,           wrap: true  },
    "ID":         { width: 50,  dataFontSize: 8,  dataAlign: "center", headerAlign: "center", numberFormat: null,           wrap: false },
    "Door Times": { width: 350, dataFontSize: 10, dataAlign: "left",   headerAlign: "left",   numberFormat: null,           wrap: true  },
    "Notes":      { width: 500, dataFontSize: 6,  dataAlign: "left",   headerAlign: "left",   numberFormat: null,           wrap: true  },
    "Status":     { width: 50,  dataFontSize: 6,  dataAlign: "left",   headerAlign: "center", numberFormat: null,           wrap: true  },
    "Areas":      { width: 50,  dataFontSize: 6,  dataAlign: "left",   headerAlign: "center", numberFormat: null,           wrap: true  }
  };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const numDataRows = range.getNumRows() - 1;

  headers.forEach((header, i) => {
    const colIndex = i + 1;
    const config = formatConfig[header];
    if (config) {
      sheet.setColumnWidth(colIndex, config.width);
      
      const headerRange = sheet.getRange(1, colIndex);
      headerRange.setFontSize(10).setHorizontalAlignment(config.headerAlign);

      if (numDataRows > 0) {
        const dataColRange = sheet.getRange(2, colIndex, numDataRows);
        dataColRange.setFontSize(config.dataFontSize)
                    .setHorizontalAlignment(config.dataAlign)
                    .setWrap(config.wrap);
        if (config.numberFormat) {
          dataColRange.setNumberFormat(config.numberFormat);
        }
      }
    }
  });
  
  range.setVerticalAlignment("middle");
  // --- End of New Section ---

  range.setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);

  const dataValues = sheet.getRange(2, 1, numDataRows, sheet.getLastColumn()).getValues();
  const dateColIndex = headers.indexOf("Date");
  const buildingColIndex = headers.indexOf("Building");

  if (dateColIndex !== -1 && buildingColIndex !== -1 && numDataRows > 1) {
    for (let i = 1; i < dataValues.length; i++) {
      const currentRow = dataValues[i];
      const previousRow = dataValues[i - 1];
      const currentDate = new Date(currentRow[dateColIndex]);
      const previousDate = new Date(previousRow[dateColIndex]);
      const currentBuilding = currentRow[buildingColIndex];
      const previousBuilding = previousRow[buildingColIndex];

      if (currentDate.setHours(0, 0, 0, 0) !== previousDate.setHours(0, 0, 0, 0)) {
        sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).setBorder(true, null, null, null, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      } else if (currentBuilding !== previousBuilding) {
        sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).setBorder(true, null, null, null, false, false, "#000000", SpreadsheetApp.BorderStyle.DASHED);
      }
    }
  }
  
  // --- Efficient Row Resizing ---
  const allData = sheet.getDataRange().getValues(); // Read all data in one go

  for (let i = 0; i < allData.length; i++) {
    const rowNumber = i + 1;
    // Check if every cell in the row is empty
    const isRowBlank = allData[i].every(cell => cell === "");

    if (isRowBlank) {
      // If the row is blank, set its height to 5 pixels
      sheet.setRowHeight(rowNumber, 5);
    } else {
      // If the row has content, auto-resize it
      sheet.autoResizeRows(rowNumber, 1);
    }
  }
}

function trimSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`The sheet "${sheetName}" was not found for trimming.`);
  }

  if (sheet.getLastRow() === 0) {
      return;
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxRows > lastRow) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
  if (maxCols > lastCol) {
    sheet.deleteColumns(lastCol + 1, maxCols - lastCol);
  }
}
// =======================================================================================
// --- END Inserted Code from Stage3.gs ---
// =======================================================================================

