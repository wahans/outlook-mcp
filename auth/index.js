/**
 * Authentication module for Outlook MCP server
 */
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

// Create singleton instance for token management with auto-refresh
const tokenStorage = new TokenStorage();

/**
 * Ensures the user is authenticated and returns an access token
 * Uses token-storage which automatically refreshes expired tokens
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    // Force re-authentication
    throw new Error('Authentication required');
  }

  // Get valid access token (auto-refreshes if expired)
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenStorage,
  authTools,
  ensureAuthenticated
};
