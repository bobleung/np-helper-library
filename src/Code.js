



/**
 * A helper function to determine the last non-empty row in a sheet.
 * Adjust if you want a more robust approach to detect truly empty rows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number} The row index (1-based) of the last non-empty row, or 1 if empty/only headers.
 */
function getLastNonEmptyRow(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Start from the bottom and move upwards until a row with data is found
  for (let row = values.length - 1; row >= 0; row--) {
    // Check if the row is entirely empty
    const rowData = values[row];
    // If any cell is not empty, return the 1-based row index
    if (rowData.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
      return row + 1;
    }
  }
  // If all rows are empty, return 1 (i.e. the very first row)
  return 1;
}



/** Before adding the mute param */
// // Read rows from a sheet and parse them into an array of objects, using header as key
// function sheetToObjects(sheet, startRow = 3, lastColumn) {
//   // var ss = SpreadsheetApp.openById(ssId);
//   // var sheet = ss.getSheetByName(sheetName);
//   var sheetName = sheet.getName();

//   if (!sheet) {
//     Logger.log("Sheet not found");
//     return;
//   }

//   // Determine the last column index based on the `lastColumn` parameter
//   var lastColumnIndex = lastColumn ? columnNameToIndex(lastColumn) : sheet.getLastColumn();

//   // Assuming the header is in the first row
//   var headersRange = sheet.getRange(1, 1, 1, lastColumnIndex);
//   var headers = headersRange.getValues()[0];
  
//   var lastRow = sheet.getLastRow();
//   if (lastRow < startRow) {
//     Logger.log("The sheet '" + sheetName + "' is empty.");
//     return { message: "The sheet '" + sheetName + "' is empty.", data: [] };
//   }

//   var dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length);
//   var data = dataRange.getValues();
  
//   data = data.filter(row => row.some(cell => cell !== ""));

//   if (data.length === 0) {
//     Logger.log("The sheet '" + sheetName + "' is empty.");
//     return { message: "The sheet '" + sheetName + "' is empty.", data: [] };
//   }

//   var objectsArray = data.map(function(row) {
//     var obj = {};
//     headers.forEach(function(header, index) {
//       if (row[index] !== "") {
//         obj[header] = row[index];
//       }
//     });
//     return obj;
//   });

//   objectsArray.forEach(function(rowObject, index) {
//     Logger.log("Row " + (index + startRow) + ": " + JSON.stringify(rowObject));
//   });

//   return objectsArray;
// }



/** Move a particular row based on a match **/
function moveRowToAnotherSheet(sourceSheet, targetSheet, headerToMatch, valueToMatch) {
  // Read the first row (headers) to find column indices in the source sheet
  var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  var columnIndex = headers.indexOf(headerToMatch) + 1;
  
  if (columnIndex === 0) {
    throw new Error('Header not found');
  }

  // Find the row to move
  var dataRange = sourceSheet.getRange(2, columnIndex, sourceSheet.getLastRow() - 1, 1);
  var values = dataRange.getValues();
  var rowToMoveIndex = -1;
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == valueToMatch) {
      rowToMoveIndex = i + 2; // Adjust for header row and zero-based index
      break;
    }
  }
  
  if (rowToMoveIndex === -1) {
    Logger.log("No matching row found");
    return; // Exit if no matching row is found
  }

  // Copy the row to the target sheet
  var rowRange = sourceSheet.getRange(rowToMoveIndex, 1, 1, sourceSheet.getLastColumn());
  var rowData = rowRange.getValues();
  var lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow + 1, 1, 1, rowData[0].length).setValues(rowData);

  // Delete the row from the source sheet
  sourceSheet.deleteRow(rowToMoveIndex);
  SpreadsheetApp.flush(); // Ensure changes are applied immediately
}

/** Add a row to a sheet **/
function addRowToSheet(sheet, newRowData) {
  // Read the first row (headers) to identify column indices
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newRow = new Array(headers.length).fill(""); // Initialize an array filled with empty strings for the new row
  
  // Map newRowData keys to headers and populate newRow array
  Object.keys(newRowData).forEach(function(key) {
    var columnIndex = headers.indexOf(key) + 1; // Find the column index for the key
    if (columnIndex > 0) {
      newRow[columnIndex - 1] = newRowData[key]; // Set the appropriate value at the correct column index
    }
  });
  
  // Append the new row to the sheet
  sheet.appendRow(newRow);
  SpreadsheetApp.flush(); // Ensure changes are applied immediately
}

/** Add single or multiple rows to a sheet. newRowsData can be an object or an array **/
function addRowsToSheet(sheet, newRowsData) {
  // Read the first row (headers) to identify column indices
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Initialize an empty array to hold all new rows
  var allNewRows = [];

  // Helper function to create a new row array from data object
  function createRowFromData(rowData) {
    var newRow = new Array(headers.length).fill(""); // Initialize an array filled with empty strings for the new row
    Object.keys(rowData).forEach(function(key) {
      var columnIndex = headers.indexOf(key) + 1; // Find the column index for the key
      if (columnIndex > 0) {
        newRow[columnIndex - 1] = rowData[key]; // Set the appropriate value at the correct column index
      }
    });
    return newRow;
  }

  // Check if newRowsData is an array and handle accordingly
  if (Array.isArray(newRowsData)) {
    // Iterate through each data object in the array
    newRowsData.forEach(function(rowData) {
      allNewRows.push(createRowFromData(rowData));
    });
    // Bulk add all new rows to the sheet in one API call
    sheet.getRange(sheet.getLastRow() + 1, 1, allNewRows.length, headers.length).setValues(allNewRows);
  } else {
    // Handle a single row data object
    var singleRow = createRowFromData(newRowsData);
    sheet.appendRow(singleRow);
  }

  SpreadsheetApp.flush(); // Ensure changes are applied immediately
}


/** Turns column Name like A, B, C etc into index **/
function columnNameToIndex(columnName) {
  let column = 0;
  for (let i = 0; i < columnName.length; i++) {
    const c = columnName.toUpperCase().charCodeAt(i) - 64; // ASCII value of 'A' is 65
    column = column * 26 + c;
  }
  return column; // 1-based index
}

/**
 * Sets the format of column B in the specified sheet to plain text.
 * @param {string} sheetName The name of the sheet where column B will be formatted.
 */
function setColumnBFormatToPlainText(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }

  var columnB = sheet.getRange('B:B'); // Selects the entire column B
  columnB.setNumberFormat('@'); // Sets the number format to plain text
  Logger.log('Column B in "' + sheetName + '" has been set to plain text format.');
}

/**
 * Deletes the contents of a column based on the header name.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to operate on.
 * @param {string} headerName - The name of the column header.
 * @param {number} [startRow=2] - The row to start deleting from (default is 2).
 */
function deleteColumnContents(sheet, headerName, startRow = 2) {
  // Get all values from the sheet
  var values = sheet.getDataRange().getValues();
  
  // Find the column index based on the header name
  var columnIndex = values[0].indexOf(headerName);
  
  // Check if the header was found
  if (columnIndex === -1) {
    throw new Error("Header '" + headerName + "' not found in the sheet.");
  }
  
  // Get the last row with content
  var lastRow = sheet.getLastRow();
  
  // Calculate the number of rows to clear
  var numRows = lastRow - startRow + 1;
  
  // Clear the contents of the column
  if (numRows > 0) {
    sheet.getRange(startRow, columnIndex + 1, numRows, 1).clearContent();
  }
}

/**
 * Searches for a file by name in a specified folder and returns its ID.
 * @param {string} fileName - The name of the file to search for.
 * @param {string} folderId - The ID of the folder to search in. If null, searches in the current folder.
 * @return {string|null} The ID of the first matching file, or null if not found.
 */
function findFileIdByName(fileName, folderId) {
  try {
    var folder;
    
    if (folderId) {
      // Use the provided folder ID
      folder = DriveApp.getFolderById(folderId);
    } else {
      // If no folder ID is provided, use the current folder
      folder = getCurrentFolder();
    }
    
    // Search for files with the given name in the folder
    var files = folder.getFilesByName(fileName);
    
    // Check if any files were found
    if (files.hasNext()) {
      // Return the ID of the first matching file
      return files.next().getId();
    } else {
      // Return null if no file is found
      return null;
    }
  } catch (e) {
    Logger.log("Error finding file: " + e.toString());
    return null;
  }
}

