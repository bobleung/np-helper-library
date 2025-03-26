/**
 * Formats specified columns in a given sheet as plain text. Changes are applied and saved automatically.
 *
 * @param {string} sheetId The ID of the spreadsheet.
 * @param {string} sheetName The name of the sheet to format.
 * @param {Array} columnArray An array of column letters to format as plain text.
 */
function formatColumnsAsPlainText(sheetId, sheetName, columnArray) {
  // Open the spreadsheet by ID and get the sheet by name
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // Loop through each column letter in the array
  columnArray.forEach(function(columnLetter) {
    // Convert the column letter to a full column range (e.g., "A:A")
    var columnRange = columnLetter + ":" + columnLetter;
    // Get the range for the column
    var range = sheet.getRange(columnRange);
    // Set the number format of the range to plain text
    range.setNumberFormat("@");
  });
  
  // Ensure all pending changes are applied to the spreadsheet
  SpreadsheetApp.flush();
}