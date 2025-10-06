// Project Name: Door Report Full
// Project Version: 2.0
// Filename: Stage0 Import.gs
// File Version: 2.01

function runSecondScript() {
  showImportDialog();
}

function formatDate(date) {
  var year = date.getFullYear();
  var month = ('0' + (date.getMonth() + 1)).slice(-2);
  var day = ('0' + date.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}

function ImportReport_Auto(days) {
  var today = new Date();
  var futureDate = new Date();
  futureDate.setDate(today.getDate() + days);

  var fromDate = formatDate(today);
  var toDate = formatDate(futureDate);

  var url = 'https://warrenk12.gofmx.com/scheduling/occurrences?format=csv&useOnlySelectedColumns=False&from=' + fromDate + '&to=' + toDate;

  const htmlScript = `
    <script>
      window.open('${url}', '_blank');
      google.script.run
        .withSuccessHandler(google.script.host.close)
        .runSecondScript();
    </script>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlScript)
    .setWidth(100)
    .setHeight(100);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Report...');
}

function showImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('IMPORTdialog')
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

