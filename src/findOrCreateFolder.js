/**
 * Searches for a folder named with the current year within a specified root folder.
 * If the folder exists, returns its ID. If it doesn't, creates it and returns the new folder's ID.
 * 
 * @param {string} rootFolderId - The ID of the root folder where the search will be performed.
 * @returns {string} The ID of the found or created folder.
 */
function findOrCreateCurrentYearFolder(rootFolderId) {
  const currentYear = new Date().getFullYear().toString();
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const folders = rootFolder.getFoldersByName(currentYear);

  if (folders.hasNext()) {
    const folder = folders.next();
    return folder.getId();
  } else {
    const newFolder = rootFolder.createFolder(currentYear);
    return newFolder.getId();
  }
}

/**
 * Searches for a folder named with the year as of last month within a specified root folder.
 * If the folder exists, returns its ID. If it doesn't, creates it and returns the new folder's ID.
 * 
 * @param {string} rootFolderId - The ID of the root folder where the search will be performed.
 * @returns {string} The ID of the found or created folder.
 */
function findOrCreateYearAsOfLastMonthFolder(rootFolderId) {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const yearOfLastMonth = lastMonth.getFullYear().toString();

  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const folders = rootFolder.getFoldersByName(yearOfLastMonth);

  if (folders.hasNext()) {
    const folder = folders.next();
    return folder.getId();
  } else {
    const newFolder = rootFolder.createFolder(yearOfLastMonth);
    return newFolder.getId();
  }
}

/**
 * Finds or creates a folder with the specified name under the root folder identified by rootFolderId.
 * @param {string} folderName - The name of the folder to find or create.
 * @param {string} rootFolderId - The ID of the root folder to search within.
 * @return {string} The ID of the found or created folder.
 */
function findOrCreateFolder(folderName, rootFolderId) {
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const folders = rootFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    // Folder exists, return its ID
    const folder = folders.next();
    return folder.getId();
  } else {
    // Folder doesn't exist, create it and return its ID
    const newFolder = rootFolder.createFolder(folderName);
    return newFolder.getId();
  }
}
