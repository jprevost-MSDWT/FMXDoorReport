// Project Name: Door Report Full
// Project Version: 5.0
// Filename: Stage2.gs
// File Version: 5.00
// Description: A combined file of all .gs scripts for easy testing.


//This script filters data from Stage1, removes duplicates, sorts it, and applies formatting.
 
function stage2_filterProcessedData() {
  // Get the active spreadsheet and the relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.sheets.helper1); // Data from Stage1
  var destinationSheet = ss.getSheetByName(CONFIG.sheets.helper2); // Where the filtered data will go

  // Basic error handling to ensure sheets exist
  if (!sourceSheet) {
    throw new Error('Source sheet "' + CONFIG.sheets.helper1 + '" not found!');
  }
  if (!destinationSheet) {
    throw new Error('Destination sheet "' + CONFIG.sheets.helper2 + '" not found!');
  }

  // Get all the data from the source sheet
  var data = sourceSheet.getDataRange().getValues();
  
  // The first row contains the headers
  var headers = data.shift(); 
  
  // --- Step 1: Filter Rows based on Status ---
  var statusColIndex = headers.indexOf('Status');
  if (statusColIndex === -1) {
    throw new Error('"Status" column could not be found in the source sheet "' + CONFIG.sheets.helper1 + '".');
  }
  var excludedStatuses = ["Declined", "Canceled", "Deleted", "Bulk Declined", "Bulk Canceled", "Bulk Deleted"];
  var filteredRows = data.filter(function(row) {
    var status = row[statusColIndex]; 
    return !excludedStatuses.includes(status);
  });
  
  // --- Step 2: Filter Columns to remove "Combined Door Times" ---
  var columnsToRemove = [];
  headers.forEach(function(header, index) {
    if (header.includes("Combined Door Times")) {
      columnsToRemove.push(index);
    }
  });
  var newHeaders = headers.filter(function(_, index) {
    return !columnsToRemove.includes(index);
  });
  var processedData = filteredRows.map(function(row) {
    return row.filter(function(_, index) {
      return !columnsToRemove.includes(index);
    });
  });

  // --- Calculate the date range for auto-selection ---
  var targetDate = calculateTargetDate();
  targetDate.setHours(0, 0, 0, 0); // Normalize to compare dates only

  // --- Step 3: Add "Selected" column data based on Target Date & Door Times ---
  var eventDateColIndex_forSelect = newHeaders.indexOf('Event Date');
  var doorTimesColIndex_forSelect = newHeaders.indexOf('Door Times'); // Get Door Times index
  newHeaders.unshift("Selected"); // Add header for the new column

  processedData.forEach(function(row) {
    var shouldBeChecked = false;

    // Condition 1: Check if the date is in range
    var isDateInRange = false;
    if (eventDateColIndex_forSelect !== -1) {
      var eventDate = parseDate(row[eventDateColIndex_forSelect]);
      if (eventDate) {
        eventDate.setHours(0, 0, 0, 0);
        if (eventDate <= targetDate) {
          isDateInRange = true;
        }
      }
    }

    // Condition 2: Check if "Door Times" has data
    var hasDoorTimes = false;
    if (doorTimesColIndex_forSelect !== -1) {
      var doorTimesData = row[doorTimesColIndex_forSelect];
      if (doorTimesData && doorTimesData.toString().trim() !== '') {
        hasDoorTimes = true;
      }
    }

    // A row is only selected if BOTH conditions are true
    if (isDateInRange && hasDoorTimes) {
      shouldBeChecked = true;
    }

    row.unshift(shouldBeChecked);
  });


  // --- Step 4: Remove Duplicates ---
  var uniqueData = [];
  var seenRows = {};
  var idColIndex = newHeaders.indexOf('ID');
  var areasColIndex = newHeaders.indexOf('Areas');
  processedData.forEach(function(row) {
    var rowKey = row.filter(function(cell, index) {
      return index !== 0 && index !== idColIndex && index !== areasColIndex;
    }).join('|');
    if (!seenRows[rowKey]) {
      seenRows[rowKey] = true;
      uniqueData.push(row);
    }
  });

  // --- Step 5: Sort Data ---
  var eventDateColIndex = newHeaders.indexOf('Event Date');
  var buildingColIndex = newHeaders.indexOf('Building');
  var eventTimeColIndex = newHeaders.indexOf('Event Time');
  var statusColIndex_forSort = newHeaders.indexOf('Status');
  var doorTimesColIndex_forSort = newHeaders.indexOf('Door Times');
  var selectedColIndex = 0; // "Selected" is always the first column

  uniqueData.sort(function(a, b) {
    // --- Priority 1: "Selected" Checkbox ---
    var aIsSelected = a[selectedColIndex] === true;
    var bIsSelected = b[selectedColIndex] === true;

    if (aIsSelected && !bIsSelected) {
      return -1; // 'a' comes first
    }
    if (!aIsSelected && bIsSelected) {
      return 1; // 'b' comes first
    }

    // --- Priority 2: For UNSELECTED rows, sort by "Door Times" presence ---
    if (!aIsSelected && !bIsSelected) {
      var aHasDoorTimes = a[doorTimesColIndex_forSort] && a[doorTimesColIndex_forSort].toString().trim() !== '';
      var bHasDoorTimes = b[doorTimesColIndex_forSort] && b[doorTimesColIndex_forSort].toString().trim() !== '';
      if (aHasDoorTimes && !bHasDoorTimes) {
        return -1; // 'a' comes first
      }
      if (!aHasDoorTimes && bHasDoorTimes) {
        return 1; // 'b' comes first
      }
    }

    // --- Priority 3: "Pending Approval" Status (Case-Insensitive) ---
    var statusA = (a[statusColIndex_forSort] || '').toUpperCase();
    var statusB = (b[statusColIndex_forSort] || '').toUpperCase();
    var aIsPending = statusA.includes("PENDING") && statusA.includes("APPROVAL");
    var bIsPending = statusB.includes("PENDING") && statusB.includes("APPROVAL");

    if (aIsPending && !bIsPending) {
      return -1; // 'a' comes first
    }
    if (!aIsPending && bIsPending) {
      return 1; // 'b' comes first
    }
    
    // --- Priority 4: Fallback to Original Sorting Logic ---
    var dateA = parseDate(a[eventDateColIndex]);
    var dateB = parseDate(b[eventDateColIndex]);
    if (dateA && !dateB) return -1;
    if (!dateA && dateB) return 1;
    if (dateA && dateB && (dateA.getTime() !== dateB.getTime())) {
      return dateA.getTime() - dateB.getTime();
    }
    
    var buildingA = a[buildingColIndex] || '';
    var buildingB = b[buildingColIndex] || '';
    var buildingCompare = buildingA.localeCompare(buildingB);
    if (buildingCompare !== 0) return buildingCompare;
    
    var timeA = parseTime(a[eventTimeColIndex]);
    var timeB = parseTime(b[eventTimeColIndex]);
    if (timeA && !timeB) return -1;
    if (!timeA && timeB) return 1;
    if (!timeA && !timeB) return 0;
    return timeA.getTime() - timeB.getTime();
  });

  // --- Step 5.5: Uncheck Pending Approval Rows ---
  // This step runs after sorting to ensure "Pending" events are at the top but unchecked.
  uniqueData.forEach(function(row) {
    var status = (row[statusColIndex_forSort] || '').toUpperCase();
    var isPending = status.includes("PENDING") && status.includes("APPROVAL");
    
    if (isPending) {
      row[selectedColIndex] = false; // Uncheck the box
    }
  });

  // --- Step 6: Write data and call formatter ---
  var finalData = [newHeaders].concat(uniqueData);
  destinationSheet.clear();
  
  if (finalData.length > 1) { // Check if there is at least a header and one row of data
    destinationSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
    Stage2_format(destinationSheet, uniqueData.length, newHeaders);
  } else {
    // Also format the sheet even if there's no data, to ensure it's clean
    Stage2_format(destinationSheet, 0, newHeaders);
  }
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

