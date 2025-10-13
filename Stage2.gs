// Project Name: Door Report Full
// Project Version: 5.0
// Filename: Stage2.gs
// File Version: 5.02
// Description: Filters, formats, and prepares the data for final reporting.

function Stage2() {
  Stage2_InitialFilter();
  Stage2_ResortAndFormat();
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
    
    const dateA = parseDate(a[eventDateColIndex]);
    const dateB = parseDate(b[eventDateColIndex]);
    if (dateA && !dateB) return -1;
    if (!dateA && dateB) return 1;
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

