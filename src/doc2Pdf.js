function doc2Pdf(fileId, deleteSource = false, deleteDuplicate = false) {
  // Open the Google Doc by Id
  var doc = DriveApp.getFileById(fileId);
  var docName = doc.getName();

  // Define the PDF file name
  var pdfName = docName + ".pdf";

  // Get the folder containing the Google Doc
  var folders = doc.getParents();
  if (folders.hasNext()) {
    var folder = folders.next();
    
    // Check for an existing PDF file with the same name and delete it if deleteDuplicate is true
    if (deleteDuplicate) {
      var files = folder.getFilesByName(pdfName);
      while (files.hasNext()) {
        var file = files.next();
        if (file.getMimeType() === MimeType.PDF) {
          file.setTrashed(true);
        }
      }
    }

    // Convert Google Doc to PDF
    var blob = DocumentApp.openById(fileId).getAs('application/pdf').setName(pdfName);

    // Save the PDF in the same folder and get the new file
    var pdfFile = folder.createFile(blob);

    // Delete the original file if deleteSource is true
    if (deleteSource) {
      doc.setTrashed(true);
    }

    // Return the fileId of the new PDF
    return pdfFile.getId();
  } else {
    throw new Error("Folder not found for the given file.");
  }
}

/** Before adding delete duplicate feature */
// function doc2Pdf(fileId, deleteSource = false) {
//   // Open the Google Doc by Id
//   var doc = DriveApp.getFileById(fileId);
//   var docName = doc.getName();

//   // Define the PDF file name
//   var pdfName = docName + ".pdf";

//   // Convert Google Doc to PDF
//   var blob = DocumentApp.openById(fileId).getAs('application/pdf').setName(pdfName);

//   // Get the folder containing the Google Doc
//   var folders = doc.getParents();
//   if (folders.hasNext()) {
//     var folder = folders.next();

//     // Save the PDF in the same folder and get the new file
//     var pdfFile = folder.createFile(blob);

//     // Delete the original file if deleteSource is true
//     if (deleteSource) {
//       doc.setTrashed(true);
//     }

//     // Return the fileId of the new PDF
//     return pdfFile.getId();
//   } else {
//     throw new Error("Folder not found for the given file.");
//   }
// }