function findRowIndexByValue(sheet, colKey, rowVal) {
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headerRow = values[0];
  var colIndex = headerRow.indexOf(colKey);

  // Check if the column key exists
  if (colIndex === -1) {
    return -1;
  }

  // Find the row index with the given value
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[colIndex] === rowVal) {
      return i;
    }
  }

  return -1; // Return -1 if the row value is not found
}

// Add a header row if it doesn't already exist
function addHeader(...headers) {
  Logger.log(headers);
  if (headers.length === 0) {
    return 'Error: At least one header must be provided';
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()+1).getValues()[0];

  // Check if header already exists
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (headerRow.indexOf(header) !== -1) {
      return 'Error: Header already exists';
    }
  }

  // Add headers to the first row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  return 'Header added successfully';
}


function addRow(...values) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if the number of values matches the number of columns
  if (values.length !== headerRow.length) {
    return 'Error: Number of values does not match the number of columns';
  }

  sheet.appendRow(values);
  return 'Row added successfully';
}

function updateRow(colKey, oldVal, newVal) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowIndex = findRowIndexByValue(sheet, colKey, oldVal);

  if (rowIndex !== -1) {
    var colIndex = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                          .getValues()[0]
                          .indexOf(colKey);

    sheet.getRange(rowIndex + 1, colIndex + 1).setValue(newVal);
    return 'Row updated successfully';
  }

  return 'Error: Row value not found';
}

function removeRow(colKey, rowVal) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowIndex = findRowIndexByValue(sheet, colKey, rowVal);

  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
    return 'Row removed successfully';
  }

  return 'Error: Row value not found';
}
