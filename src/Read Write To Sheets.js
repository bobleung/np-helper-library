// Clear a sheet, but skip header rows
function clearSheetContentFromRow(sheet, startRow = 3) {
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var sheet = ss.getSheetByName(sheetName);
  var sheetName = sheet.getName();


  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return;
  }
  
  // Calculate the number of rows and columns to clear
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  // If there's content to clear
  if (lastRow >= startRow) {
    // Get the range starting from `startRow` to the last row and column
    var rangeToClear = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn);
    rangeToClear.clearContent(); // This clears the content but not the rows themselves
  }
}

/**
 * Takes an array of objects and puts them into rows in a specified sheet, using headers as keys.
 * @param {Object[]} array - The array of objects to be inserted into the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object where data will be inserted.
 * @param {number} headerRowIndex - The index of the header row.
 * @param {number} startRowIndex - The row index to start processing and inserting data.
 * @param {string} mode - The mode of insertion: "overwrite" (default) or "append".
 */
function objectsToSheet(array, sheet, headerRowIndex = 1, startRowIndex = 3, mode = "overwrite") {
  const sheetName = sheet.getName();
  const headers = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dataToInsert = [];

  // Iterate through the array to prepare data for insertion based on headers
  array.forEach(obj => {
    const row = headers.map(header => obj.hasOwnProperty(header) ? obj[header] : "");
    dataToInsert.push(row);
  });

  if (mode === "overwrite") {
    // Clear the sheet except for the header rows
    clearSheetContentFromRow(sheet, startRowIndex);
    
    // Add the data to the sheet in a batch, starting from the specified row
    if (dataToInsert.length > 0) {
      sheet.getRange(startRowIndex, 1, dataToInsert.length, headers.length).setValues(dataToInsert);
    }
  } else if (mode === "append") {
    // Find the last non-empty row
    const lastRow = getLastNonEmptyRow(sheet);
    
    // Append the data to the sheet in a batch, starting from the row after the last non-empty row
    if (dataToInsert.length > 0) {
      sheet.getRange(lastRow + 1, 1, dataToInsert.length, headers.length).setValues(dataToInsert);
    }
  } else {
    throw new Error("Invalid mode. Use 'overwrite' or 'append'.");
  }
}

/**
 * Takes an array of objects and puts them into rows or columns in a specified sheet
 * while preserving formulas.
 * 
 * Optional settings: 
 *   - Set mode="overwrite" (default), "append" or "overlay" to control insertion behavior.
 *   - In overlay mode, if preserveFormulas is false a batch update is used for efficiency,
 *     but if preserveFormulas is true each row/column is processed individually to restore formulas.
 *   - Set pivot=true to transpose data (default is false).
 *   - Set preserveFormulas=true to preserve existing formulas (default is true).
 * 
 * @param {Object[]} array - The array of objects to be inserted into the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object where data will be inserted.
 * @param {number} headerIndex - In normal mode, the row for headers.
 *                               In pivot mode, this is the column from which header keys are read.
 * @param {number} startIndex - In normal mode, the row where data starts.
 *                              In pivot mode, this becomes the starting column index for object data.
 * @param {Object} [options] - Optional settings.
 * @param {("overwrite"|"append"|"overlay")} [options.mode="overwrite"] - The insertion mode.
 * @param {boolean} [options.pivot=false] - If true, pivot the table (flip rows/columns).
 * @param {boolean} [options.preserveFormulas=true] - If true, preserve existing formulas.
 */
function objectsToSheetV2(array, sheet, headerIndex = 1, startIndex = 3, options = {}) {
  const { mode = "overwrite", pivot = false, preserveFormulas = true } = options;
  
  // Internal helper: Get the last non-empty row in the sheet.
  function getLastNonEmptyRow(sheet) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    for (let row = values.length - 1; row >= 0; row--) {
      if (values[row].some(cell => cell !== "" && cell !== null && cell !== undefined)) {
        return row + 1; // 1-indexed
      }
    }
    return 1;
  }
  
  // Internal helper: Get the last non-empty column starting from a given column.
  function getLastNonEmptyColumn(sheet, startCol) {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol === 0 || lastRow === 0) return startCol;
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    for (let col = lastCol - 1; col >= startCol - 1; col--) {
      for (let row = 0; row < lastRow; row++) {
        if (data[row][col] !== "") {
          return col + 1; // 1-indexed
        }
      }
    }
    return startCol;
  }
  
  if (!pivot) {
    // Normal mode: Write objects as rows.
    const headers = sheet.getRange(headerIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dataToInsert = array.map(obj => headers.map(header => (obj.hasOwnProperty(header) ? obj[header] : "")));
    
    if (mode === "overwrite") {
      const lastRow = sheet.getLastRow();
      
      if (preserveFormulas && lastRow >= startIndex) {
        // Get existing formulas
        const numRows = lastRow - startIndex + 1;
        const range = sheet.getRange(startIndex, 1, numRows, headers.length);
        const formulas = range.getFormulas();
        
        // Clear the range (this will remove both values and formulas)
        range.clearContent();
        
        // Set the new values
        if (dataToInsert.length > 0) {
          sheet.getRange(startIndex, 1, dataToInsert.length, headers.length).setValues(dataToInsert);
        }
        
        // Restore formulas
        for (let i = 0; i < formulas.length; i++) {
          for (let j = 0; j < formulas[i].length; j++) {
            if (formulas[i][j] !== "") {
              sheet.getRange(startIndex + i, j + 1).setFormula(formulas[i][j]);
            }
          }
        }
      } else {
        // Clear and set without preserving formulas
        if (sheet.getLastRow() >= startIndex) {
          sheet.getRange(startIndex, 1, sheet.getLastRow() - startIndex + 1, headers.length).clearContent();
        }
        
        if (dataToInsert.length > 0) {
          sheet.getRange(startIndex, 1, dataToInsert.length, headers.length).setValues(dataToInsert);
        }
      }
      
    } else if (mode === "append") {
      const lastRow = getLastNonEmptyRow(sheet);
      if (dataToInsert.length > 0) {
        sheet.getRange(lastRow + 1, 1, dataToInsert.length, headers.length).setValues(dataToInsert);
      }
      
    } else if (mode === "overlay") {
      // Overlay mode: update only cells where the new object provides a value,
      // leaving existing content (and formulas) intact.
      if (!preserveFormulas) {
        // Batch update overlay mode for efficiency when not preserving formulas.
        const range = sheet.getRange(startIndex, 1, dataToInsert.length, headers.length);
        const currentData = range.getValues();
        // Overlay new values onto currentData
        for (let i = 0; i < dataToInsert.length; i++) {
          for (let j = 0; j < headers.length; j++) {
            if (array[i].hasOwnProperty(headers[j])) {
              currentData[i][j] = array[i][headers[j]];
            }
          }
        }
        // Write updated data back in one batch
        range.setValues(currentData);
      } else {
        // Process each row individually if preserving formulas.
        for (let i = 0; i < dataToInsert.length; i++) {
          const targetRow = startIndex + i;
          const currentRange = sheet.getRange(targetRow, 1, 1, headers.length);
          const currentValues = currentRange.getValues()[0];
          const currentFormulas = currentRange.getFormulas()[0];
          const newRow = [];
          for (let j = 0; j < headers.length; j++) {
            if (array[i].hasOwnProperty(headers[j])) {
              newRow.push(array[i][headers[j]]);
            } else {
              newRow.push(currentValues[j]);
            }
          }
          currentRange.setValues([newRow]);
          for (let j = 0; j < headers.length; j++) {
            if (!array[i].hasOwnProperty(headers[j]) && currentFormulas[j] !== "") {
              sheet.getRange(targetRow, j + 1).setFormula(currentFormulas[j]);
            }
          }
        }
      }
      
    } else {
      throw new Error("Invalid mode. Use 'overwrite', 'append' or 'overlay'.");
    }
    
  } else {
    // Pivot mode: Write data as columns (transposed version of normal mode).
    // Cache last row for header reading.
    const lastRow = sheet.getLastRow();
    const propertyNames = sheet.getRange(1, headerIndex, lastRow, 1).getValues();
    const properties = propertyNames.map(row => row[0]); // Retain empty cells as blanks.
    const numProperties = properties.length;
    
    if (mode === "append") {
      // Updated for efficiency: Use a single batch write.
      const baseCol = getLastNonEmptyColumn(sheet, startIndex) + 1;
      const numNewCols = array.length;
      const batchData = [];
      // Build a 2D array with dimensions numProperties x numNewCols.
      for (let i = 0; i < numProperties; i++) {
        const rowData = [];
        const prop = properties[i];
        for (let j = 0; j < numNewCols; j++) {
          rowData.push(array[j].hasOwnProperty(prop) ? array[j][prop] : "");
        }
        batchData.push(rowData);
      }
      // Write the entire block at once.
      sheet.getRange(1, baseCol, numProperties, numNewCols).setValues(batchData);
      
    } else if (mode === "overwrite") {
      array.forEach((obj, index) => {
        const targetCol = startIndex + index;
        const columnData = properties.map(prop => [obj.hasOwnProperty(prop) ? obj[prop] : ""]);
        
        if (preserveFormulas) {
          // Get existing formulas
          const range = sheet.getRange(1, targetCol, numProperties, 1);
          const formulas = range.getFormulas();
          
          // Clear the range
          range.clearContent();
          
          // Set the new values
          sheet.getRange(1, targetCol, numProperties, 1).setValues(columnData);
          
          // Restore formulas
          for (let i = 0; i < formulas.length; i++) {
            if (formulas[i][0] !== "") {
              sheet.getRange(i + 1, targetCol).setFormula(formulas[i][0]);
            }
          }
        } else {
          // Clear and set without preserving formulas
          sheet.getRange(1, targetCol, numProperties, 1).setValues(columnData);
        }
      });
      
    } else if (mode === "overlay") {
      // Overlay mode for pivot: update only cells where the new object provides a value,
      // leaving existing content (and formulas) intact.
      if (!preserveFormulas) {
        // Batch update overlay mode for pivot when not preserving formulas.
        const numNewCols = array.length;
        const range = sheet.getRange(1, startIndex, numProperties, numNewCols);
        const currentData = range.getValues();
        for (let j = 0; j < numNewCols; j++) {
          const obj = array[j];
          for (let i = 0; i < numProperties; i++) {
            const prop = properties[i];
            if (obj.hasOwnProperty(prop)) {
              currentData[i][j] = obj[prop];
            }
          }
        }
        range.setValues(currentData);
      } else {
        // Process each column individually if preserving formulas.
        array.forEach((obj, index) => {
          const targetCol = startIndex + index;
          const currentRange = sheet.getRange(1, targetCol, numProperties, 1);
          const currentValues = currentRange.getValues();
          const currentFormulas = currentRange.getFormulas();
          const newColData = [];
          
          for (let i = 0; i < numProperties; i++) {
            const prop = properties[i];
            if (obj.hasOwnProperty(prop)) {
              newColData.push([obj[prop]]);
            } else {
              newColData.push(currentValues[i]);
            }
          }
          
          currentRange.setValues(newColData);
          
          for (let i = 0; i < numProperties; i++) {
            if (!obj.hasOwnProperty(properties[i]) && currentFormulas[i][0] !== "") {
              sheet.getRange(i + 1, targetCol).setFormula(currentFormulas[i][0]);
            }
          }
        });
      }
      
    } else {
      throw new Error("Invalid mode. Use 'overwrite', 'append' or 'overlay'.");
    }
  }
}

/**
 * Converts a Google Sheet into an array of objects.
 *
 * By default, the function reads the header row as the property keys and each row as the values.
 * With the pivot option set to true, it reads the header column as the property keys and each column as the values.
 *
 * @param {Sheet} sheet - The Google Sheet to process.
 * @param {number} headerIndex - In normal mode, the row number containing headers.
 *                               In pivot mode, the column number containing headers.
 * @param {number} startIndex - In normal mode, the row number where data starts.
 *                              In pivot mode, the column number where data starts.
 * @param {Object} [options={}] - Options object to configure behaviour:
 *                                Set lastColumn to limit processing 
 *                               (e.g., "Z"), mute=true (default) to suppress logging, useDisplayDates=true (default) 
 *                               to replace date objects with displayed text, pivot=false (default) to control 
 *                               table orientation.
 * @param {string|null} [options.lastColumn=null] - Last column to process, e.g., "Z".
 * @param {boolean} [options.mute=true] - Suppress logging.
 * @param {boolean} [options.useDisplayDates=true] - Replace date objects with displayed text.
 * @param {boolean} [options.pivot=false] - If true, pivot the table (flip rows/columns) before processing.
 * @return {Object[]} - Array of objects representing the sheet data.
 */
function sheetToObjectsV2(sheet, headerIndex = 1, startIndex = 3, options = {}) {
  const baseOptions = {
    lastColumn: null,
    mute: true,
    useDisplayDates: true,
    pivot: false
  };

  const { lastColumn, mute, useDisplayDates, pivot } = { ...baseOptions, ...options };

  if (!sheet) {
    Logger.log("Sheet not found");
    return;
  }

  const sheetName = sheet.getName();
  const lastColumnIndex = lastColumn ? columnNameToIndex(lastColumn) : sheet.getLastColumn();

  let headers, data, displayData;

  // Helper: Transpose a 2D array.
  function transpose(matrix) {
    return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
  }

  // Helper: Clear a sheet's content from a specified row (skipping header rows).
  function clearSheetContentFromRow(sheet, startRow = 3) {
    var sheetName = sheet.getName();
    if (!sheet) {
      Logger.log("Sheet not found: " + sheetName);
      return;
    }
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow >= startRow) {
      var rangeToClear = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn);
      rangeToClear.clearContent();
    }
  }

  // Helper: Get the last non-empty row in a sheet.
  function getLastNonEmptyRow(sheet) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    for (let row = values.length - 1; row >= 0; row--) {
      const rowData = values[row];
      if (rowData.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        return row + 1; // 1-indexed
      }
    }
    return 1;
  }

  if (!pivot) {
    // Normal mode: Read headers from a row and data from rows.
    const headersRange = sheet.getRange(headerIndex, 1, 1, sheet.getLastColumn());
    headers = headersRange.getValues()[0];

    const lastRow = sheet.getLastRow();
    if (lastRow < startIndex) {
      Logger.log(`The sheet '${sheetName}' is empty.`);
      return null;
    }

    const dataRange = sheet.getRange(startIndex, 1, lastRow - startIndex + 1, headers.length);
    data = dataRange.getValues();
    if (useDisplayDates) {
      displayData = dataRange.getDisplayValues();
    }
  } else {
    // Pivot mode: Read headers from a column and data from columns.
    // Then, transpose the data so that the rest of the function can process it as if it were row-based.
    headers = sheet.getRange(1, headerIndex, sheet.getLastRow(), 1).getValues().flat();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const dataRange = sheet.getRange(1, startIndex, lastRow, lastCol - startIndex + 1);
    const rawData = dataRange.getValues();
    data = transpose(rawData);
    if (useDisplayDates) {
      const dispRange = sheet.getRange(1, startIndex, lastRow, lastCol - startIndex + 1);
      const rawDispData = dispRange.getDisplayValues();
      displayData = transpose(rawDispData);
    }
  }

  // Filter out any empty rows.
  const filteredData = data.filter(row => row.some(cell => cell !== ""));
  if (filteredData.length === 0) {
    Logger.log(`The sheet '${sheetName}' is empty.`);
    return { message: `The sheet '${sheetName}' is empty.`, data: [] };
  }

  // Convert each row of data into an object using the headers.
  const objectsArray = filteredData.map((row, rowIndex) => {
    const obj = {};
    headers.forEach((header, colIndex) => {
      if (!header) return; // skip empty headers

      const value = row[colIndex];
      if (typeof value === "object") {
        // Check if this is a genuine Date object.
        const isDate = Object.prototype.toString.call(value) === "[object Date]";
        if (isDate && useDisplayDates && displayData) {
          obj[header] = displayData[rowIndex][colIndex];
          return;
        }
      }
      if (value !== null && value !== "" && value !== undefined) {
        obj[header] = value;
      }
    });
    return obj;
  });

  if (!mute) {
    objectsArray.forEach((rowObject, index) => {
      Logger.log(`Row ${index + startIndex}: ${JSON.stringify(rowObject)}`);
    });
  }

  return objectsArray;
}

/**
 * Processes sheet data with columns as properties
 * 
 * @param {Sheet} sheet - The Google Sheet to process
 * @param {number} headerRow - Row number containing headers
 * @param {number} startRow - Row number where data starts
 * @param {number} lastColumnIndex - Last column index to process
 * @param {boolean} useDisplayDates - Whether to use display values for dates
 * @param {boolean} mute - Whether to suppress logging
 * @param {string} sheetName - Name of the sheet
 * @return {Object[]} - Array of objects representing sheet columns
 */
function processColumnsAsProperties(sheet, headerRow, startRow, lastColumnIndex, useDisplayDates, mute, sheetName) {
  // Get the first column which will contain our "headers" (property names)
  const lastRow = sheet.getLastRow();
  if (lastRow < headerRow) {
    Logger.log(`The sheet '${sheetName}' is empty.`);
    return null;
  }
  
  // Get the property names from the first column
  const propertyNamesRange = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, 1);
  const propertyNames = propertyNamesRange.getValues().map(row => row[0]);
  
  // Filter out empty property names
  const filteredPropertyNames = [];
  const propertyIndices = [];
  
  propertyNames.forEach((prop, index) => {
    if (prop && prop !== "") {
      filteredPropertyNames.push(prop);
      propertyIndices.push(index + headerRow);
    }
  });
  
  if (filteredPropertyNames.length === 0) {
    Logger.log(`No valid property names found in the first column of '${sheetName}'.`);
    return [];
  }
  
  // Get all data including the header row (which will be column names)
  const columnsData = sheet.getRange(headerRow, 2, lastRow - headerRow + 1, lastColumnIndex - 1).getValues();
  
  // Get display values if needed
  let displayData;
  if (useDisplayDates) {
    displayData = sheet.getRange(headerRow, 2, lastRow - headerRow + 1, lastColumnIndex - 1).getDisplayValues();
  }
  
  // Extract column names (these will be the object keys)
  const columnNames = columnsData[0];
  
  // Remove columns where all cells are empty (except the header)
  const validColumnIndices = [];
  
  for (let colIndex = 0; colIndex < columnNames.length; colIndex++) {
    if (columnNames[colIndex] && columnNames[colIndex] !== "") {
      // Check if the column has any data
      const hasData = columnsData.slice(1).some(row => {
        const value = row[colIndex];
        return value !== null && value !== "" && value !== undefined;
      });
      
      if (hasData) {
        validColumnIndices.push(colIndex);
      }
    }
  }
  
  if (validColumnIndices.length === 0) {
    Logger.log(`No valid data columns found in '${sheetName}'.`);
    return [];
  }
  
  // Now create the objects with columns as properties
  const objectsArray = [];
  
  validColumnIndices.forEach(colIndex => {
    const columnName = columnNames[colIndex];
    if (!columnName) return; // Skip columns with no name
    
    const obj = {};
    
    filteredPropertyNames.forEach((propName, index) => {
      const rowIndex = propertyIndices[index] - headerRow;
      
      if (rowIndex < columnsData.length) {
        const value = columnsData[rowIndex][colIndex];
        
        if (typeof value === "object") {
          // Check if it's a date
          const isDate = Object.prototype.toString.call(value) === "[object Date]";
          if (isDate && useDisplayDates && displayData) {
            obj[propName] = displayData[rowIndex][colIndex];
            return;
          }
        }
        
        // Use the raw value if it is not null/empty
        if (value !== null && value !== "" && value !== undefined) {
          obj[propName] = value;
        }
      }
    });
    
    // Only add the object if it has properties
    if (Object.keys(obj).length > 0) {
      obj._columnName = columnName; // Add the column name as a reference
      objectsArray.push(obj);
    }
  });
  
  if (!mute) {
    objectsArray.forEach((colObject, index) => {
      Logger.log(`Column ${index + 2}: ${JSON.stringify(colObject)}`);
    });
  }
  
  return objectsArray;
}

/** #### WARNING: #### TO BE DEPRECIATED */
function sheetToObjects(sheet, startRow = 3, lastColumn, mute = true, headerRow = 1) {
  Logger.log("WARNING : FUNCTION DEPRECIATED, USE INSTEAD sheetToObjectsV2")
  var sheetName = sheet.getName();

  if (!sheet) {
    Logger.log("Sheet not found");
    return;
  }

  // Determine the last column index based on the `lastColumn` parameter
  var lastColumnIndex = lastColumn ? columnNameToIndex(lastColumn) : sheet.getLastColumn();

  // Adjust to fetch headers from the specified `headerRow`
  var headersRange = sheet.getRange(headerRow, 1, 1, lastColumnIndex);
  var headers = headersRange.getValues()[0];

  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log("The sheet '" + sheetName + "' is empty.");
    return null;
  }

  // Get data starting from `startRow` and adjust column count based on headers length
  var dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length);
  var data = dataRange.getValues();

  // Remove rows where all cells are empty
  data = data.filter(row => row.some(cell => cell !== ""));

  if (data.length === 0) {
    Logger.log("The sheet '" + sheetName + "' is empty.");
    return { message: "The sheet '" + sheetName + "' is empty.", data: [] };
  }

  // Convert data rows into objects using headers as keys, skipping empty headers
  var objectsArray = data.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      if (header !== "" && row[index] !== "") { // Skip if header is empty
        obj[header] = row[index];
      }
    });
    return obj;
  });

  if (!mute) {
    objectsArray.forEach(function(rowObject, index) {
      Logger.log("Row " + (index + startRow) + ": " + JSON.stringify(rowObject));
    });
  }

  return objectsArray;
}

// // Read rows from a sheet and parse them into an array of objects, using header as key
// function sheetToObjects(sheet, startRow = 3, lastColumn, mute = true) {
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
//     // return { message: "The sheet '" + sheetName + "' is empty.", data: [] };
//     return null;
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

//   if (!mute) {
//     objectsArray.forEach(function(rowObject, index) {
//       Logger.log("Row " + (index + startRow) + ": " + JSON.stringify(rowObject));
//     });
//   }

//   return objectsArray;
// }

/**
 * Matches rows in a Google Sheet with data from an array of objects and updates specific columns.
 * Skips non-existent columns and logs warnings.
 * Uses batch writing per column to minimize API calls while preserving untouched cells.
 *
 * @param {Sheet} sheet - The Google Sheet object to process.
 * @param {Array<Object>} data - Array of objects to match with rows in the sheet.
 * @param {string} columnToMatch - The header name of the column to match.
 * @param {Array<string>} columnsToAdd - Array of header names of the columns to update.
 */
function findAndUpdateRows(sheet, data, columnToMatch, columnsToAdd) {
  // Parameter validation
  if (
    !sheet ||
    !Array.isArray(data) ||
    typeof columnToMatch !== "string" ||
    !Array.isArray(columnsToAdd)
  ) {
    throw new Error("Invalid parameters. Check that the inputs are valid.");
  }

  // Get all data from the sheet
  const sheetData = sheet.getDataRange().getValues();
  if (sheetData.length < 2) return; // No data or only headers in the sheet

  const headers = sheetData[0]; // First row contains the headers
  const columnIndexToMatch = headers.indexOf(columnToMatch);

  if (columnIndexToMatch === -1) {
    throw new Error(`Column '${columnToMatch}' not found in sheet headers.`);
  }

  // Get indices for each column to update, skipping missing columns
  const columnsInfo = columnsToAdd
    .map((column) => {
      const index = headers.indexOf(column);
      if (index === -1) {
        Logger.log(
          `Warning: Column '${column}' does not exist in the sheet. Skipping this column.`
        );
        return null; // Mark missing columns with `null`
      }
      return { columnIndex: index, columnName: column };
    })
    .filter(Boolean); // Remove `null` entries

  if (columnsInfo.length === 0) {
    Logger.log("No valid columns to update. Exiting function.");
    return; // Exit if there are no valid columns to update
  }

  // Create a lookup map from the data array
  const dataLookup = new Map(
    data.map((obj) => [obj[columnToMatch], obj])
  );

  // Track which rows need to be updated
  const rowsToUpdate = []; // Array of row indices (0-based, excluding header)
  for (let i = 1; i < sheetData.length; i++) {
    const rowValue = sheetData[i][columnIndexToMatch];
    if (dataLookup.has(rowValue)) {
      rowsToUpdate.push(i);
    }
  }

  if (rowsToUpdate.length === 0) {
    Logger.log("No matching rows found to update.");
    return;
  }

  // For each column to add, prepare and set the updated values
  columnsInfo.forEach(({ columnIndex, columnName }) => {
    // Get existing values for the column
    const columnRange = sheet.getRange(2, columnIndex + 1, sheetData.length - 1, 1);
    const columnValues = columnRange.getValues(); // 2D array

    // Update the necessary rows
    rowsToUpdate.forEach((rowIndex) => {
      const matchedObject = dataLookup.get(sheetData[rowIndex][columnIndexToMatch]);
      columnValues[rowIndex - 1][0] = matchedObject[columnName] || ""; // Update with value or empty string
    });

    // Set the updated values back to the column
    columnRange.setValues(columnValues);
  });

  Logger.log(`Updated ${rowsToUpdate.length} rows for columns: ${columnsToAdd.join(", ")}`);
}

/** DEPRECIATED (use findAndUpdateRows) : Update a particular row based on a match */

function objectToRow(sheet, headerToMatch, valueToMatch, newRowData) {
  // Read the first row (headers) to find column indices
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnIndex = headers.indexOf(headerToMatch) + 1;
  
  if (columnIndex === 0) {
    throw new Error('Header not found');
  }

  // Find the row to update
  var dataRange = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1);
  var values = dataRange.getValues();
  var rowToEdit = -1;
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == valueToMatch) {
      rowToEdit = i + 2; // Adjust for header row and zero-based index
      break;
    }
  }
  
  if (rowToEdit === -1) {
    Logger.log("No matching row found");
    return; // Exit if no matching row is found
  }

  // Prepare data for the update
  var updateRange = sheet.getRange(rowToEdit, 1, 1, headers.length);
  var rowData = updateRange.getValues()[0];

  // Map newRowData keys to headers and set values using columnNameToIndex if needed
  Object.keys(newRowData).forEach(function(key) {
    var colIndex = headers.indexOf(key) + 1; // This still needs to be found by header matching
    if (colIndex > 0) {
      rowData[colIndex - 1] = newRowData[key]; // Adjust for zero-based index of rowData array
    }
  });

  // Write back the updated data to the sheet in one go
  updateRange.setValues([rowData]);
  SpreadsheetApp.flush(); // Ensure changes are applied immediately
}

/**
 * Upserts rows in a Google Sheet by matching on a particular column and:
 *  - updating data if there is a match, or
 *  - inserting a new row if there is no match.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheet object to process.
 * @param {Object[]} data - An array of objects to match with rows in the sheet.
 * @param {string} columnToMatch - The header name of the column to match on.
 * @param {number} [headerRowIndex=1] - The index of the header row (1-based).
 */
function upsertRows(sheet, data, columnToMatch, headerRowIndex = 1) {
  // --- Step 0: Basic validation ---
  if (!sheet || !Array.isArray(data) || typeof columnToMatch !== 'string') {
    throw new Error('Invalid parameters. Check that the inputs are valid.');
  }
  if (data.length === 0) {
    Logger.log('No data to upsert. Exiting function.');
    return;
  }

  // --- Step 1: Retrieve headers and existing sheet data ---
  const sheetData = sheet.getDataRange().getValues(); 
  // If there's no data in the sheet, sheetData might be [[]]; handle that case
  if (!sheetData || sheetData.length === 0) {
    Logger.log('No data in the sheet; only headers or sheet is empty.');
  }

  const headers = sheetData[0] || [];
  const columnIndexToMatch = headers.indexOf(columnToMatch);
  if (columnIndexToMatch === -1) {
    throw new Error(`Column '${columnToMatch}' not found in sheet headers.`);
  }

  // Create a quick lookup for all rows in the sheet
  // Map "value in match column" => row index
  // row index is 1-based in the sheet, but 0-based in sheetData
  const existingRowsMap = new Map();
  for (let i = 1; i < sheetData.length; i++) {
    const rowVal = sheetData[i][columnIndexToMatch];
    if (rowVal !== '') {
      existingRowsMap.set(rowVal, i);
    }
  }

  // --- Step 2: Build column set from the sheet headers for easier reference ---
  // We only update columns that exist in the sheet's headers
  const headerSet = new Set(headers);

  // --- Step 3: Split data into "existing to update" vs "new to append" ---
  const existingUpdates = []; // Will hold objects of form { rowIndex, rowObject }
  const newRecords = []; // Will hold rowObjects that need to be appended

  data.forEach((rowObject) => {
    // The value we want to match on
    const matchValue = rowObject[columnToMatch];
    // If there's no matchValue, we consider it a new record
    if (!matchValue || !existingRowsMap.has(matchValue)) {
      newRecords.push(rowObject);
    } else {
      const rowIndex = existingRowsMap.get(matchValue);
      existingUpdates.push({ rowIndex, rowObject });
    }
  });

  // --- Step 4: Perform updates for existing rows ---
  // We update all columns that appear in both the rowObject and headers
  if (existingUpdates.length > 0) {
    // We'll do a column-by-column update to minimise overwriting any cells that are not part of the data
    existingUpdates.forEach(({ rowIndex, rowObject }) => {
      // rowIndex is the 0-based index in sheetData. The actual sheet row is rowIndex + 1
      for (const [key, val] of Object.entries(rowObject)) {
        if (headerSet.has(key)) {
          const colIndex = headers.indexOf(key);
          // Write the value to the sheet
          sheet.getRange(rowIndex + 1, colIndex + 1).setValue(val);
        }
      }
    });
    Logger.log(`Updated ${existingUpdates.length} existing rows.`);
  } else {
    Logger.log('No existing rows matched for update.');
  }

  // --- Step 5: Append new records where no match was found ---
  if (newRecords.length > 0) {
    // We find the last non-empty row by scanning from bottom if needed, 
    // but a quick approach is to use the data in memory:
    // (sheetData.length - 1) is the last content row (0-based for header),
    // so the next row is sheetData.length + 1 if the sheet has at least 1 data row.
    const lastNonEmptyRow = getLastNonEmptyRow(sheet);
    const appendStartRow = lastNonEmptyRow + 1;

    // Build a 2D array of values to insert
    // For each record, create an array with values in the order of the sheet's headers
    const rowsToAppend = newRecords.map((rowObject) =>
      headers.map((header) => rowObject.hasOwnProperty(header) ? rowObject[header] : '')
    );

    // Insert them all at once
    sheet
      .getRange(appendStartRow, 1, rowsToAppend.length, headers.length)
      .setValues(rowsToAppend);

    Logger.log(`Appended ${newRecords.length} new records from row ${appendStartRow}.`);
  } else {
    Logger.log('No new rows to append.');
  }
}