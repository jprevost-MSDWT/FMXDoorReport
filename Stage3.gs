// Project Name: Door Report Full
// Project Version: 3.0
// Filename: Stage3.gs
// File Version: 3.01

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
  
  PrintPageFormattingandTrim(destinationSheetName);
}

function PrintPageFormattingandTrim(sheetName) {
  PrintPageFormattingONLY(sheetName);
  trimSheet(sheetName);
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

  const dataToSort = range.offset(1, 0, range.getNumRows() - 1);
  dataToSort.sort([
    { column: 1, ascending: true },
    { column: 3, ascending: true },
    { column: 2, ascending: true }
  ]);

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

