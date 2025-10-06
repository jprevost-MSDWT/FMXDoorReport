// Stage1.gs (Stage 1 - Main Processing Script)
// This script processes FMX door data from "Import" to "Output-Helper1", replaces building names, removes duplicates, and sorts.

function FMX_Doors_AutoImport_V8() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Import");
  var outputSheet = ss.getSheetByName("Output-Helper1");
  var dataSheet = ss.getSheetByName("Data");

  if (!inputSheet) {
    console.error('Error: Source sheet "Import" not found!');
    return;
  }
  if (!outputSheet) {
    console.error('Error: Destination sheet "Output-Helper1" not found!');
    return;
  }
  if (!dataSheet) {
    console.error('Error: Lookup sheet "Data" not found!');
    return;
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

  // --- Find column indexes ---
  var eventTimeCol = inputHeaders.indexOf("Event time");
  if (eventTimeCol === -1) {
    eventTimeCol = inputHeaders.indexOf("Starts");
  }
  var nameCol = inputHeaders.indexOf("Name");
  var buildingsCol = inputHeaders.indexOf("Buildings");
  if (buildingsCol === -1) {
    buildingsCol = inputHeaders.indexOf("Building");
  }
  var statusCol = inputHeaders.indexOf("Status");
  var resourcesCol = inputHeaders.indexOf("Resources");
  var eventDetailsCol = inputHeaders.indexOf("Event Details");
  var doorNotesCol = inputHeaders.indexOf("Old ML Door Notes");
  var unlockTimeCol = inputHeaders.indexOf("Unlock Time");
  var lockTimeCol = inputHeaders.indexOf("Lock Time");
  var unlockTimeDotCol = inputHeaders.indexOf("Unlock Time.");
  var lockTimeDotCol = inputHeaders.indexOf("Lock Time.");
  var unlockTimeDotDotCol = inputHeaders.indexOf("Unlock Time..");
  var lockTimeDotDotCol = inputHeaders.indexOf("Lock Time..");
  var unlockTimeDotDotDotCol = inputHeaders.indexOf("Unlock Time...");
  var lockTimeDotDotDotCol = inputHeaders.indexOf("Lock Time...");
  var unlockTimeDotDotDotDotCol = inputHeaders.indexOf("Unlock Time....");
  var lockTimeDotDotDotDotCol = inputHeaders.indexOf("Lock Time....");
  var lockTimeDotDotDotDotDotCol = inputHeaders.indexOf("Lock Time.....");
  var unlockTimeDotDotDotDotDotCol = inputHeaders.indexOf("Unlock Time.....");
  var unlockTimeSpecialCol = inputHeaders.indexOf("WCHS Football/Baseball Locker Room Doors Unlock Time");
  var lockTimeSpecialCol = inputHeaders.indexOf("WCHS Football/Baseball Locker Room Doors Lock Time");

  var doorColumns1 = ['Clinic Doors', 'BV Doors', 'CR Doors', 'ECC Doors', 'EA Doors', 'EdCtr Doors', 'GC Doors', 'HA Doors', 'HP Doors', 'LO Doors', 'LP Doors', 'LA Doors', 'MO Doors', 'SH Doors', 'PO Doors', 'SB Doors', 'RP Doors', 'PR Doors', 'REN Doors', 'WCHS Doors', 'Special Doors'];
  var doorColumns2 = ['EA Doors.', 'BV Doors.', 'CR Doors.', 'ECC Doors.', 'GC Doors.', 'HA Doors.', 'HP Doors.', 'LO Doors.', 'LA Doors.', 'LP Doors.', 'MO Doors.', 'PO Doors.', 'PR Doors.', 'REN Doors.', 'RP Doors.', 'SB Doors.', 'SH Door.', 'WCHS Doors.'];
  var doorColumns3 = ['WCHS Doors..'];
  var doorColumns4 = ['WCHS Doors...'];
  var doorColumns5 = ['WCHS Doors....'];
  var doorColumns6 = ['WCHS Doors.....'];
  var doorColumns7 = ['WCHS Football/Baseball Locker Room Doors'];

  var doorColIndexes1 = getColumnIndexes(inputHeaders, doorColumns1);
  var doorColIndexes2 = getColumnIndexes(inputHeaders, doorColumns2);
  var doorColIndexes3 = getColumnIndexes(inputHeaders, doorColumns3);
  var doorColIndexes4 = getColumnIndexes(inputHeaders, doorColumns4);
  var doorColIndexes5 = getColumnIndexes(inputHeaders, doorColumns5);
  var doorColIndexes6 = getColumnIndexes(inputHeaders, doorColumns6);
  var doorColIndexes7 = getColumnIndexes(inputHeaders, doorColumns7);

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

    var doors1 = combineDoorValues(row, doorColumns1, doorColIndexes1);
    var unlockTime1 = formatTimeValue(row[unlockTimeCol], timeZone);
    var lockTime1 = formatTimeValue(row[lockTimeCol], timeZone);
    var combinedTimes1 = formatDoorTimes(doors1, unlockTime1, lockTime1);

    var doors2 = combineDoorValues(row, doorColumns2, doorColIndexes2);
    var unlockTime2 = formatTimeValue(row[unlockTimeDotCol], timeZone);
    var lockTime2 = formatTimeValue(row[lockTimeDotCol], timeZone);
    var combinedTimes2 = formatDoorTimes(doors2, unlockTime2, lockTime2);

    var doors3 = combineDoorValues(row, doorColumns3, doorColIndexes3);
    var unlockTime3 = formatTimeValue(row[unlockTimeDotDotCol], timeZone);
    var lockTime3 = formatTimeValue(row[lockTimeDotDotCol], timeZone);
    var combinedTimes3 = formatDoorTimes(doors3, unlockTime3, lockTime3);

    var doors4 = combineDoorValues(row, doorColumns4, doorColIndexes4);
    var unlockTime4 = formatTimeValue(row[unlockTimeDotDotDotCol], timeZone);
    var lockTime4 = formatTimeValue(row[lockTimeDotDotDotCol], timeZone);
    var combinedTimes4 = formatDoorTimes(doors4, unlockTime4, lockTime4);

    var doors5 = combineDoorValues(row, doorColumns5, doorColIndexes5);
    var unlockTime5 = formatTimeValue(row[unlockTimeDotDotDotDotCol], timeZone);
    var lockTime5 = formatTimeValue(row[lockTimeDotDotDotDotCol], timeZone);
    var combinedTimes5 = formatDoorTimes(doors5, unlockTime5, lockTime5);

    var doors6 = combineDoorValues(row, doorColumns6, doorColIndexes6);
    var unlockTime6 = formatTimeValue(row[unlockTimeDotDotDotDotDotCol], timeZone);
    var lockTime6 = formatTimeValue(row[lockTimeDotDotDotDotDotCol], timeZone);
    var combinedTimes6 = formatDoorTimes(doors6, unlockTime6, lockTime6);

    var doors7 = combineDoorValues(row, doorColumns7, doorColIndexes7);
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

  try {
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
      console.log('Script Finished! Please check the "Output-Helper1" sheet.');
    } else {
      console.log('Script Finished! No unique data rows found to process.');
    }
  } catch (e) {
    console.error('A critical error occurred while writing to the sheet: ' + e.message);
  }
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