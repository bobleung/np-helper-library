/**
 * Duplicates a file in Google Drive, renames it, moves it to a specified folder,
 * and returns the new file's ID. Optionally overwrites existing file with the same name.
 *
 * @param {string} templateFileId The ID of the template file to duplicate.
 * @param {string} fileName The new name for the duplicated file.
 * @param {string} folderId The ID of the folder where the duplicated file will be placed.
 * @param {boolean} [overwrite=false] If true, overwrite existing file with the same name.
 * @return {string} The ID of the newly created file.
 */

function duplicateFile(templateFileId, fileName, folderId, overwrite = false) {
  
  // Get the folder by ID
  var folder = DriveApp.getFolderById(folderId);
  
  // If overwrite is true, check if a file with the same name exists and delete it
  if (overwrite) {
    var files = folder.getFilesByName(fileName);
    while (files.hasNext()) {
      var existingFile = files.next();
      existingFile.setTrashed(true);  // Delete the file by moving it to trash
    }
  }

  // Duplicate the file
  var file = DriveApp.getFileById(templateFileId);
  var newFile = file.makeCopy();
  
  // Rename the new file
  newFile.setName(fileName);
  
  // Move the new file to the specified folder
  folder.addFile(newFile);
  
  // Remove the new file from its current (root) location
  DriveApp.getRootFolder().removeFile(newFile);

  // Return the ID of the new file
  return newFile.getId();
}

/**
 * Checks for an existing file with the given name in the specified folder.
 * If it exists, returns its ID. If not, duplicates the template file,
 * renames it, moves it to the specified folder, and returns the new file's ID.
 *
 * @param {string} templateFileId The ID of the template file to duplicate.
 * @param {string} fileName The name for the file (existing or new).
 * @param {string} folderId The ID of the folder to check/place the file.
 * @return {string} The ID of the existing or newly created file.
 */
function duplicateFileIfNotExists(templateFileId, fileName, folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByName(fileName);
  
  // Check if a file with the same name already exists
  if (files.hasNext()) {
    // If it exists, return its ID
    return files.next().getId();
  } else {
    // If it doesn't exist, duplicate the template file
    var file = DriveApp.getFileById(templateFileId);
    var newFile = file.makeCopy();

    // Rename the new file
    newFile.setName(fileName);

    // Move the new file to the specified folder
    folder.addFile(newFile);
    
    // Remove the new file from its current (root) location
    DriveApp.getRootFolder().removeFile(newFile);

    // Return the ID of the new file
    return newFile.getId();
  }
}