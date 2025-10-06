// Stage2.gs (Stage 2 - Filter and Sort)
// This script filters data from Stage1, removes duplicates, sorts it, and writes output to Output-Helper2.

function stage2_filterProcessedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Output-Helper1");
  var destSheet = ss.getSheetByName("Output-Helper2");

  if (!sourceSheet) {
    console.error('Error: Source sheet "Output-Helper1" not found!');
    return;
  }
  if (!destSheet) {
    console.error('Error: Destination sheet "Output-Helper2" not found!');
    return;
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) {
    destSheet.clear();
    destSheet.appendRow(["No data found."]);
    return;
  }

  var headers = sourceData[0];
  var dataRows = sourceData.slice(1);

  // Remove duplicate rows based on all columns except "Notes"
  var notesIndex = headers.indexOf("Notes");
  var rowKeys = {};
  var filteredRows = [];
  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i].slice();
    var rowForKey = row.slice();
    if (notesIndex !== -1) rowForKey[notesIndex] = ""; // Ignore "Notes" in duplicate key
    var key = rowForKey.join("|");
    if (!rowKeys[key]) {
      rowKeys[key] = true;
      filteredRows.push(row);
    }
  }

  // Add a "Selected" column at the start if not present
  var finalHeaders = headers.slice();
  if (finalHeaders[0] !== "Selected") {
    finalHeaders.unshift("Selected");
  }

  // Add "Selected" column to each row, default to FALSE
  var outputRows = filteredRows.map(function(row) {
    return [false].concat(row);
  });

  // Clear and write output
  destSheet.clear();
  destSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  if (outputRows.length > 0) {
    destSheet.getRange(2, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  // Auto-resize columns for readability
  for (var c = 1; c <= finalHeaders.length; c++) {
    destSheet.autoResizeColumn(c);
  }
}