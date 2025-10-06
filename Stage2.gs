/**
 * =======================================================================================
 * STAGE 2 - FILTER SCRIPT (Version 3.01)
 * =======================================================================================
 * This script filters data from Stage1, removes duplicates, sorts it, and applies formatting.
 */
function stage2_filterProcessedData() {
  // Get the active spreadsheet and the relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.sheets.helper1); // Data from Stage1
  var destinationSheet = ss.getSheetByName(CONFIG.sheets.helper2); // Where the filtered data will go

  // Basic error handling to ensure sheets exist
  if (!sourceSheet) {
    console.error('Error: Source sheet "' + CONFIG.sheets.helper1 + '" not found!');
    return;
  }
  if (!destinationSheet) {
    console.error('Error: Destination sheet "' + CONFIG.sheets.helper2 + '" not found!');
    return;
  }

  // Get all the data from the source sheet
  var data = sourceSheet.getDataRange().getValues();
  
  // The first row contains the headers
  var headers = data.shift(); 
  
  // --- Step 1: Filter Rows based on Status ---
  var statusColIndex = headers.indexOf('Status');
  if (statusColIndex === -1) {
    console.error('Error: "Status" column could not be found in the source sheet "' + CONFIG.sheets.helper1 + '".');
    return; 
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
    console.log('Stage 2 Filtering Complete! Data written and formatted in "' + CONFIG.sheets.helper2 + '".');
  } else {
    // Also format the sheet even if there's no data, to ensure it's clean
    Stage2_format(destinationSheet, 0, newHeaders);
    console.log('Stage 2 Filtering Complete! No data was left to write after filtering.');
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
    // Assuming MM/DD format, and we need to add a year for a valid Date object.
    // The year doesn't matter for sorting within the same year.
    var currentYear = new Date().getFullYear();
    // To handle year-end rollovers, check if the date is in the past (e.g., it's Jan and date is Dec)
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
 * Applies all necessary formatting to the destination sheet. (Version 1.09)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to format.
 * @param {number} numDataRows The number of data rows (excluding the header).
 * @param {string[]} headers The array of header strings.
 */
function Stage2_format(sheet, numDataRows, headers) {
  var numCols = headers.length;
  if (numCols === 0) return; // Exit if there are no columns

  // Clear any existing conditional formatting rules and filters
  sheet.clearConditionalFormatRules();
  var filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  // 1. Make the header row bold
  sheet.getRange(1, 1, 1, numCols).setFontWeight("bold");

  if (numDataRows > 0) {
    var dataRange = sheet.getRange(2, 1, numDataRows, numCols);
    var rules = [];
    var statusColIndex = headers.indexOf('Status');
    
    // --- Set Default font color for "Notes" column ---
    var notesColIndex = headers.indexOf('Notes');
    if (notesColIndex !== -1) {
      var notesColumnRange = sheet.getRange(2, notesColIndex + 1, numDataRows, 1);
      notesColumnRange.setFontColor("#b7b7b7"); // "Dark gray 1"
    }

    if (statusColIndex !== -1) {
      var statusColumnLetter = String.fromCharCode('A'.charCodeAt(0) + statusColIndex);
      var statusColumnRange = sheet.getRange(2, statusColIndex + 1, numDataRows, 1);
      
      // --- Rule (Admin Override): "Admin"+"Pending"+"Approval" status cells when A is FALSE ---
      // Rule - (rule_admin_approval_false)
      var rule_admin_approval_false = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=FALSE, ISNUMBER(SEARCH("Admin", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#d9ead3")   // "Light Green 3"
        .setFontColor("#b7b7b7")   // "Dark gray 1"
        .setRanges([statusColumnRange])
        .build();
      rules.push(rule_admin_approval_false);

      // --- Rule 1: Override for "Pending Approval" status cells when A is FALSE ---
      // Rule - (rule_pending_approval_false)
      var rule_pending_approval_false = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=FALSE, ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#fff2cc")   // "Light Yellow 3"
        .setFontColor("#b7b7b7")   // "Dark gray 1"
        .setRanges([statusColumnRange])
        .build();
      rules.push(rule_pending_approval_false);
      
      // --- Rule (Admin TRUE Override): "Admin"+"Pending"+"Approval" rows when A is TRUE ---
      // Rule - (rule_admin_approval_true)
      var rule_admin_approval_true = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=TRUE, ISNUMBER(SEARCH("Admin", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#b6d7a8") // "Light Green 1"
        .setFontColor("#000000") // Black
        .setRanges([dataRange])
        .build();
      rules.push(rule_admin_approval_true);

      // --- Rule 2: Formatting for "Pending Approval" rows when A is TRUE ---
      // Rule - (rule_pending_approval_true)
      var rule_pending_approval_true = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A2=TRUE, ISNUMBER(SEARCH("Pending", $${statusColumnLetter}2)), ISNUMBER(SEARCH("Approval", $${statusColumnLetter}2)))`)
        .setBackground("#ffd966")   // "Light Yellow 1"
        .setFontColor("#000000")   // Black
        .setRanges([dataRange])
        .build();
      rules.push(rule_pending_approval_true);
    }

    // --- Rule 3: Default formatting for all unchecked rows (A is False) ---
    // Rule - (rule_A_is_false)
    var rule_A_is_false = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$A2=FALSE")
      .setBackground("#efefef")   // "Light gray 2"
      .setFontColor("#b7b7b7")   // "Dark gray 1"
      .setRanges([dataRange])
      .build();
    rules.push(rule_A_is_false);

    // Set all new rules to the sheet
    sheet.setConditionalFormatRules(rules);

    // Add checkboxes to the first column
    var checkboxRange = sheet.getRange(2, 1, numDataRows, 1);
    var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkboxRange.setDataValidation(checkboxRule);
  }
  
  // Add a filter to the data range (including header)
  sheet.getRange(1, 1, numDataRows + 1, numCols).createFilter();

  // Trim extra rows from the bottom of the sheet
  var totalRows = numDataRows + 1; // +1 for the header
  var maxRows = sheet.getMaxRows();
  if (maxRows > totalRows) {
    sheet.deleteRows(totalRows + 1, maxRows - totalRows);
  }
}

/**
 * Calculates a target date based on the current day of the week.
 * If today is Friday, it finds the next Tuesday.
 * Otherwise, it finds the next Friday.
 */
function calculateTargetDate() {
  var today = new Date();
  var dayOfWeek = today.getDay(); // Sunday=0, Monday=1, ..., Friday=5, Saturday=6
  var targetDate = new Date(today); // Create a copy to modify

  if (dayOfWeek === 5) { // If it's Friday
    // Add 4 days to get to the next Tuesday
    targetDate.setDate(today.getDate() + 4);
  } else { // For any other day
    // Calculate days needed to get to the next Friday
    var daysUntilFriday = (5 - dayOfWeek + 7) % 7;
    // If today is Saturday, daysUntilFriday will be 6. If Sunday, 5, etc.
    // If today is before Friday (e.g. Wed), this will be 2.
    if (daysUntilFriday === 0) daysUntilFriday = 7; // If today is Friday, get next Friday
    targetDate.setDate(today.getDate() + daysUntilFriday);
  }
  
  Logger.log('Calculated Target Date: ' + targetDate);
  return targetDate;
}

