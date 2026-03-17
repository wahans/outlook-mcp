/**
 * Email folder utilities
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName) {

  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  try {
    // Try to find the folder by name
    const folderId = await getFolderIdByName(accessToken, folderName);
    if (folderId) {
      const path = `me/mailFolders/${folderId}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    // If not found, fall back to inbox
    console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
    return WELL_KNOWN_FOLDERS['inbox'];
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    return WELL_KNOWN_FOLDERS['inbox'];
  }
}

/**
 * Get the ID of a mail folder by its name (searches subfolders too)
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find (can be "subfolder" or "Parent/subfolder")
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  try {
    console.error(`Looking for folder with name "${folderName}"`);

    // Handle path notation like "Inbox/subfolder"
    if (folderName.includes('/')) {
      const parts = folderName.split('/');
      let currentFolderId = null;

      for (const part of parts) {
        const searchEndpoint = currentFolderId
          ? `me/mailFolders/${currentFolderId}/childFolders`
          : 'me/mailFolders';

        const response = await callGraphAPI(
          accessToken,
          'GET',
          searchEndpoint,
          null,
          { $filter: `displayName eq '${part}'` }
        );

        if (response.value && response.value.length > 0) {
          currentFolderId = response.value[0].id;
        } else {
          // Try case-insensitive
          const allResponse = await callGraphAPI(accessToken, 'GET', searchEndpoint, null, { $top: 100 });
          const match = allResponse.value?.find(f => f.displayName.toLowerCase() === part.toLowerCase());
          if (match) {
            currentFolderId = match.id;
          } else {
            console.error(`Folder part "${part}" not found in path "${folderName}"`);
            return null;
          }
        }
      }

      console.error(`Resolved path "${folderName}" to ID: ${currentFolderId}`);
      return currentFolderId;
    }

    // First try with exact match filter on top-level
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${folderName}'` }
    );

    if (response.value && response.value.length > 0) {
      console.error(`Found folder "${folderName}" with ID: ${response.value[0].id}`);
      return response.value[0].id;
    }

    // Search all folders including subfolders
    console.error(`No exact match found for "${folderName}", searching all folders including subfolders`);
    const allFolders = await getAllFolders(accessToken);
    const lowerFolderName = folderName.toLowerCase();

    const matchingFolder = allFolders.find(
      folder => folder.displayName.toLowerCase() === lowerFolderName
    );

    if (matchingFolder) {
      console.error(`Found match for "${folderName}" with ID: ${matchingFolder.id}`);
      return matchingFolder.id;
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  try {
    // Get top-level folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { 
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      }
    );
    
    if (!response.value) {
      return [];
    }
    
    // Get child folders for folders with children
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);
    
    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          { 
            $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
          }
        );
        
        return childResponse.value || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });
    
    const childFolders = await Promise.all(childFolderPromises);
    
    // Combine top-level folders and all child folders
    return [...response.value, ...childFolders.flat()];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders
};
