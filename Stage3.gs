// Project Name: Door Report Full
// Project Version: 5.0
// Filename: Stage3.gs
// File Version: 5.02
// Description: Generates the final, formatted reports.

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

