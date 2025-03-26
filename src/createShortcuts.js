/** NOT YET WORKING AS LIBRARY **/

/**
 * Create a shortcut of a folder inside another folder.
 * 
 * @param {string} sourceFolderId - The ID of the folder to create a shortcut of.
 * @param {string} targetFolderId - The ID of the folder where the shortcut will be placed.
 * @returns {string} - The ID of the created shortcut.
 */
function createFolderShortcut(sourceFolderId, targetFolderId) {
  if (typeof sourceFolderId !== "string" || typeof targetFolderId !== "string") {
    throw new Error("Invalid input: Both sourceFolderId and targetFolderId must be strings.");
  }

  Logger.log(`Creating shortcut for sourceFolderId: ${sourceFolderId} inside targetFolderId: ${targetFolderId}`);

  // Define shortcut metadata for the Drive API
  const shortcutMetadata = {
    name: `Shortcut to ${sourceFolderId}`, // Name of the shortcut
    mimeType: 'application/vnd.google-apps.shortcut',
    parents: [{ id: targetFolderId }], // Target folder ID
    shortcutDetails: {
      targetId: sourceFolderId // Source folder ID
    }
  };

  try {
    // Use Drive API to create the shortcut
    const shortcut = Drive.Files.insert(shortcutMetadata);
    Logger.log(`Shortcut created successfully with ID: ${shortcut.id}`);
    return shortcut.id;
  } catch (error) {
    Logger.log(`Failed to create shortcut: ${error.message}`);
    throw new Error(`Failed to create shortcut: ${error.message}`);
  }
}