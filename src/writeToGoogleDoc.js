/**
 * Replaces placeholders within a Google Docs document with specified values. The function
 * searches through the entire document for placeholders and replaces them with the corresponding
 * values provided in the data object. Placeholders in the document should be uniquely identifiable
 * and wrapped in double curly braces, e.g., "{{placeholderName}}".
 * 
 * @param {string} docId - The ID of the Google Docs document where the placeholders are located.
 * @param {Object} data - An object where each key is the name of a placeholder (excluding curly braces),
 *                        and the corresponding value is the text to replace the placeholder with. 
 *                        The placeholder names should match exactly with those in the document,
 *                        but without the curly braces. For example, if the document contains a 
 *                        placeholder "{{username}}", the data object should have a key "username".
 * 
 * Usage:
 * 1. Prepare your Google Docs document by inserting placeholders you wish to replace. Placeholders
 *    should be uniquely identifiable and wrapped in double curly braces, e.g., "{{username}}".
 * 2. Create a data object with keys as the placeholder names (without curly braces) and values as
 *    the text you want to replace the placeholders with. For example:
 *    { username: 'JohnDoe', age: '30', city: 'New York' }
 * 3. Call the populatePlaceholdersInDocument function with the document ID and the data object.
 *    For instance: populatePlaceholdersInDocument('1a2b3c', { username: 'JohnDoe', age: '30', city: 'New York' });
 * 
 * Note:
 * - The function only replaces placeholders if a corresponding value is found in the data object
 *   and the value is not null/undefined.
 * - This approach allows dynamic content insertion into documents by replacing predefined placeholders
 *   with actual data.
 * - Ensure that your placeholders are unique within your document to avoid unintended replacements.
 * - The document is saved and closed after all replacements are made.
 */

function populatePlaceholdersInDocument(docId, replacements) {
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();

  // Replace each placeholder with its corresponding value, if the value is not null/undefined
  for (var placeholder in replacements) {
    if (replacements[placeholder] !== null && replacements[placeholder] !== undefined) {
      body.replaceText(placeholder, replacements[placeholder]);
    }
  }

  // Save and close the document
  doc.saveAndClose();
}

/**
 * Populates a table within a Google Docs document with specified data. The function
 * targets a table identified by a unique tag placed in the second row, first column of the table.
 * 
 * @param {string} docId - The ID of the Google Docs document where the table is located.
 * @param {Array<Array<string | number>>} data - An array of arrays, where each inner array represents
 *                                                a row of data to be added to the table. Each element
 *                                                in an inner array represents a cell's content in that row.
 *                                                Contents can be strings or numbers.
 * @param {string} tag - The unique tag that identifies the target table. The tag should be placed
 *                       in the second row, first column of the table and enclosed in double curly braces
 *                       in the document itself, e.g., "{{sessions}}". However, when calling this function,
 *                       the tag should be provided without curly braces, e.g., "sessions".
 * 
 * Usage:
 * 1. Prepare your Google Docs document by inserting a table. Place a unique tag, e.g., "{{sessions}}",
 *    in the second row's first cell of the table you wish to populate.
 * 2. Structure your data according to the expected format. For example, to add two rows of data into the
 *    table, where the first row has three columns and the second row has two, your data parameter should look
 *    like this: [['Row1Col1', 'Row1Col2', 'Row1Col3'], ['Row2Col1', 'Row2Col2']].
 * 3. Call the populateTableInDocument function with the document ID, the structured data, and the tag (without braces).
 *    For instance: populateTableInDocument('1a2b3c', [['John Doe', 'Developer', 5000], ['Jane Doe', 'Manager']], 'sessions');
 * 
 * Note:
 * - The tag row will be removed after the table is identified.
 * - If no table with the specified tag is found, a log message will be generated.
 * - In each row, the text in the last column will be aligned to the right.
 */


function populateTableInDocument(docId, data, tag) {
  try {
    // Start timing for performance monitoring (optional)
    const startTime = new Date();
    
    // Open the document and get the body
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    
    // Retrieve all tables in the document once
    const tables = body.getTables();
    
    let targetTable = null;
    let tagRow = null;
    let tagRowIndex = -1;
    const tagRowAttributes = [];

    // Step 1: Locate the target table and tag row
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      // Check if the table has at least two rows (header and tag)
      if (table.getNumRows() > 1) {
        const potentialTagRow = table.getRow(1);
        const tagCell = potentialTagRow.getCell(0);
        const tagCellText = tagCell.getText();
        
        if (tagCellText.includes(tag)) {
          targetTable = table;
          tagRow = potentialTagRow;
          tagRowIndex = 1; // Since it's the second row (0-based indexing)
          break; // Exit loop once the target table is found
        }
      }
    }

    // Exit early if the target table or tag row isn't found
    if (!targetTable || !tagRow) {
      Logger.log('Table with specified tag not found.');
      doc.saveAndClose();
      return;
    }

    // Step 2: Extract formatting attributes from the tag row
    const numCells = tagRow.getNumCells();
    for (let c = 0; c < numCells; c++) {
      const cell = tagRow.getCell(c);
      const cellAttributes = cell.getAttributes();
      const paragraph = cell.getChild(0).asParagraph();
      const paragraphAttributes = paragraph.getAttributes();
      
      // Store both cell and paragraph attributes
      tagRowAttributes.push({
        cell: cellAttributes,
        paragraph: paragraphAttributes
      });
    }

    // Step 3: Determine the insertion point (right after the tag row)
    const insertionIndex = tagRowIndex; // Since the tag row will be removed, this is the correct insertion point

    // Step 4: Remove the tag row in a single call
    targetTable.removeRow(tagRowIndex);

    // Step 5: Prepare all new rows first to minimize DOM interactions
    const newRows = data.map(rowData => {
      // Create a new row object with cell data
      return rowData.map(cellData => cellData.toString());
    });

    // Step 6: Insert all new rows at the insertion index
    newRows.forEach((rowData, rowOffset) => {
      const currentIndex = insertionIndex + rowOffset;
      const newRow = targetTable.insertTableRow(currentIndex);
      
      // Add cells to the new row
      rowData.forEach((cellData, colIndex) => {
        // Create a new cell with data
        const newCell = newRow.appendTableCell(cellData);
        
        // Apply cell-level attributes
        newCell.setAttributes(tagRowAttributes[colIndex].cell);
        
        // Apply paragraph-level attributes
        const newParagraph = newCell.getChild(0).asParagraph();
        newParagraph.setAttributes(tagRowAttributes[colIndex].paragraph);
        
        // **Removed Alignment Logic**
        // Previously, the last column was right-aligned. This has been removed as per the new requirement.
        if (colIndex === rowData.length - 1) {
          newParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        }
      });
    });

    // Optional: Log the number of rows inserted
    Logger.log(`${newRows.length} rows inserted successfully.`);

    // End timing and log execution time (optional)
    const endTime = new Date();
    const timeDiff = (endTime - startTime) / 1000; // in seconds
    Logger.log(`populateTableInDocument executed in ${timeDiff} seconds.`);

    // Save and close the document to apply changes
    doc.saveAndClose();
  } catch (error) {
    Logger.log(`Error in populateTableInDocument: ${error.message}`);
  }
}

// /**
//  * Populates a table within a Google Docs document with specified data.
//  * Inserts new rows between the tag row and the total row, preserving formatting.
//  * 
//  * @param {string} docId - The ID of the Google Docs document where the table is located.
//  * @param {Array<Array<string | number>>} data - An array of arrays, where each inner array represents
//  *                                                a row of data to be added to the table.
//  * @param {string} tag - The unique tag that identifies the target table (without curly braces).
//  */
// function populateTableInDocument(docId, data, tag) {
//   var doc = DocumentApp.openById(docId);
//   var body = doc.getBody();
//   var tables = body.getTables();
//   var targetTable = null;
//   var tagRow = null;
//   var tagRowAttributes = [];
//   var tagRowIndex = -1;

//   // Step 1: Locate the target table by searching for the tag in the second row, first column
//   for (var i = 0; i < tables.length; i++) {
//     var table = tables[i];
//     if (table.getNumRows() > 1) { // Ensure the table has at least two rows
//       var tagCellText = table.getRow(1).getCell(0).getText();
//       if (tagCellText.includes(tag)) {
//         targetTable = table;
//         tagRow = table.getRow(1);
//         tagRowIndex = i; // Note: This is the table index, not the row index
//         break;
//       }
//     }
//   }

//   if (!targetTable || !tagRow) {
//     Logger.log('Table with specified tag not found.');
//     doc.saveAndClose();
//     return;
//   }

//   // Step 2: Extract formatting attributes from each cell in the tag row
//   for (var c = 0; c < tagRow.getNumCells(); c++) {
//     var cell = tagRow.getCell(c);
//     var cellAttributes = cell.getAttributes();
//     var paragraph = cell.getChild(0).asParagraph();
//     var paragraphAttributes = paragraph.getAttributes();
    
//     // Store both cell and paragraph attributes for comprehensive formatting
//     tagRowAttributes.push({
//       cell: cellAttributes,
//       paragraph: paragraphAttributes
//     });
//   }

//   // Step 3: Determine the row index of the tag row within the table
//   // Google Apps Script uses 0-based indexing for rows
//   var tagRowPosition = targetTable.getChildIndex(tagRow);

//   // Step 4: Remove the tag row from the table
//   targetTable.removeRow(tagRowPosition);

//   // Step 5: Insert new rows at the position where the tag row was
//   // This ensures that new rows are inserted before the total row
//   data.forEach(function(rowData, rowIndex) {
//     var newRow = targetTable.insertTableRow(tagRowPosition + rowIndex);

//     rowData.forEach(function(cellData, colIndex) {
//       // Insert a new cell with the data
//       var newCell = newRow.appendTableCell(cellData.toString());

//       // Apply cell-level formatting from the tag row
//       newCell.setAttributes(tagRowAttributes[colIndex].cell);

//       // Apply paragraph-level formatting from the tag row
//       var newParagraph = newCell.getChild(0).asParagraph();
//       newParagraph.setAttributes(tagRowAttributes[colIndex].paragraph);

//       // If this is the last column, align the text to the right
//       if (colIndex === rowData.length - 1) {
//         newParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
//       }
//     });
//   });

//   // Step 6: Save and close the document to apply changes
//   doc.saveAndClose();
// }

// /**
//  * Populates a table within a Google Docs document with specified data. The function
//  * targets a table identified by a unique tag placed in the second row, first column of the table.
//  * 
//  * @param {string} docId - The ID of the Google Docs document where the table is located.
//  * @param {Array<Array<string | number>>} data - An array of arrays, where each inner array represents
//  *                                                a row of data to be added to the table. Each element
//  *                                                in an inner array represents a cell's content in that row.
//  *                                                Contents can be strings or numbers.
//  * @param {string} tag - The unique tag that identifies the target table. The tag should be placed
//  *                       in the second row, first column of the table and enclosed in double curly braces
//  *                       in the document itself, e.g., "{{sessions}}". However, when calling this function,
//  *                       the tag should be provided without curly braces, e.g., "sessions".
//  * 
//  * Usage:
//  * 1. Prepare your Google Docs document by inserting a table. Place a unique tag, e.g., "{{sessions}}",
//  *    in the second row's first cell of the table you wish to populate.
//  * 2. Structure your data according to the expected format. For example, to add two rows of data into the
//  *    table, where the first row has three columns and the second row has two, your data parameter should look
//  *    like this: [['Row1Col1', 'Row1Col2', 'Row1Col3'], ['Row2Col1', 'Row2Col2']].
//  * 3. Call the populateTableInDocument function with the document ID, the structured data, and the tag (without braces).
//  *    For instance: populateTableInDocument('1a2b3c', [['John Doe', 'Developer', 5000], ['Jane Doe', 'Manager']], 'sessions');
//  * 
//  * Note:
//  * - The tag row will be removed after the table is identified.
//  * - If no table with the specified tag is found, a log message will be generated.
//  * - In each row, the text in the last column will be aligned to the right.
//  */

// function populateTableInDocument(docId, data, tag) {
//   var doc = DocumentApp.openById(docId);
//   var body = doc.getBody();
//   var tables = body.getTables();
//   var targetTable = null;

//   // Search for the table with the specified tag in the second row, first column
//   tables.forEach(function(table) {
//     if (table.getNumRows() > 1) {
//       var tagCellText = table.getRow(1).getCell(0).getText();
//       if (tagCellText.includes(tag)) {
//         targetTable = table;
//         // Remove the tag row
//         targetTable.removeRow(1);
//         return;
//       }
//     }
//   });

//   if (targetTable) {
//     // Append data to the table
//     data.forEach(function(rowData) {
//       var row = targetTable.appendTableRow();
//       rowData.forEach(function(cellData, index) {
//         var cell = row.appendTableCell(cellData);
//         // If this is the last column, align the text to the right
//         if (index === rowData.length - 1) {
//           cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
//         }
//       });
//     });
//   } else {
//     Logger.log('Table with specified tag not found.');
//   }

//   doc.saveAndClose();
// }

