/************************************************************
 * BOX DOCUMENTS - Search Box.com for vendor-related documents
 *
 * Standalone module for Battle Station
 * Searches Box.com API for files matching vendor names (fuzzy)
 *
 * SETUP:
 * 1. Create a Box Developer App at https://app.box.com/developers/console
 *    - Choose "Custom App" with "User Authentication (OAuth 2.0)"
 *    - Set redirect URI to: https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
 *    - Note your Client ID and Client Secret
 * 2. Run authorizeBox() from the menu to complete OAuth flow
 * 3. Test with testBoxConnection() or searchBoxForVendor("Company Name")
 *
 * After testing, integrate with Battle Station by adding to getBoxDocuments_()
 ************************************************************/

const BOX_CFG = {
  // ========== BOX APP CREDENTIALS ==========
  CLIENT_ID: PropertiesService.getScriptProperties().getProperty('BOX_CLIENT_ID') || '',
  CLIENT_SECRET: PropertiesService.getScriptProperties().getProperty('BOX_CLIENT_SECRET') || '',

  // OAuth endpoints
  AUTH_URL: 'https://account.box.com/api/oauth2/authorize',
  TOKEN_URL: 'https://api.box.com/oauth2/token',

  // API base
  API_BASE: 'https://api.box.com/2.0',

  // Search settings
  DEFAULT_LIMIT: 20,
  FILE_TYPES: ['pdf', 'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt'],  // Optional filter

  // Property keys for storing tokens
  PROP_ACCESS_TOKEN: 'BOX_ACCESS_TOKEN',
  PROP_REFRESH_TOKEN: 'BOX_REFRESH_TOKEN',
  PROP_TOKEN_EXPIRY: 'BOX_TOKEN_EXPIRY'
};


/************************************************************
 * OAUTH 2.0 AUTHENTICATION
 ************************************************************/

/**
 * Get the OAuth2 service for Box
 * Uses the OAuth2 library: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 */
function getBoxService_() {
  return OAuth2.createService('box')
    .setAuthorizationBaseUrl(BOX_CFG.AUTH_URL)
    .setTokenUrl(BOX_CFG.TOKEN_URL)
    .setClientId(BOX_CFG.CLIENT_ID)
    .setClientSecret(BOX_CFG.CLIENT_SECRET)
    .setCallbackFunction('boxAuthCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('root_readonly')  // Read access to files and folders
    .setParam('access_type', 'offline');  // Get refresh token
}

/**
 * Start the Box authorization flow
 * Run this from the Script Editor or menu
 */
function authorizeBox() {
  const service = getBoxService_();

  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('✅ Box is already authorized!\n\nRun testBoxConnection() to verify.');
    return;
  }

  const authUrl = service.getAuthorizationUrl();

  const html = HtmlService.createHtmlOutput(
    '<h2>Box Authorization Required</h2>' +
    '<p>Click the link below to authorize access to your Box account:</p>' +
    '<p><a href="' + authUrl + '" target="_blank" style="font-size: 18px; color: blue;">🔗 Authorize Box Access</a></p>' +
    '<p>After authorizing, close this dialog and run <code>testBoxConnection()</code></p>'
  )
  .setWidth(450)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'Box Authorization');
}

/**
 * OAuth callback handler
 */
function boxAuthCallback(request) {
  const service = getBoxService_();
  const authorized = service.handleCallback(request);

  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<h2>✅ Success!</h2>' +
      '<p>Box has been authorized. You can close this window.</p>' +
      '<p>Run <code>testBoxConnection()</code> to verify the connection.</p>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<h2>❌ Authorization Failed</h2>' +
      '<p>Please try again or check your Box app configuration.</p>'
    );
  }
}

/**
 * Revoke Box authorization (for testing/reset)
 */
function revokeBoxAuth() {
  const service = getBoxService_();
  service.reset();
  SpreadsheetApp.getUi().alert('Box authorization has been revoked.\n\nRun authorizeBox() to re-authorize.');
}


/************************************************************
 * BOX API FUNCTIONS
 ************************************************************/

/**
 * Make an authenticated request to Box API
 */
function boxApiRequest_(endpoint, options = {}) {
  const service = getBoxService_();

  if (!service.hasAccess()) {
    throw new Error('Box is not authorized. Run authorizeBox() first.');
  }

  const url = endpoint.startsWith('http') ? endpoint : BOX_CFG.API_BASE + endpoint;

  const fetchOptions = {
    method: options.method || 'get',
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (options.payload) {
    fetchOptions.payload = JSON.stringify(options.payload);
  }

  const response = UrlFetchApp.fetch(url, fetchOptions);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code >= 200 && code < 300) {
    return JSON.parse(text);
  } else {
    Logger.log(`Box API Error (${code}): ${text}`);
    throw new Error(`Box API Error (${code}): ${text.substring(0, 200)}`);
  }
}

/**
 * Search Box for files matching a query (vendor name)
 * Searches both file NAMES and file CONTENT
 *
 * @param {string} query - The search query (vendor name)
 * @param {object} options - Optional settings
 * @returns {array} Array of matching files with metadata
 */
function searchBox(query, options = {}) {
  const limit = options.limit || BOX_CFG.DEFAULT_LIMIT;
  const type = options.type || 'file';  // 'file', 'folder', or 'web_link'

  // Build search URL with parameters
  // Search both name and content to find files that mention the vendor
  let searchUrl = `/search?query=${encodeURIComponent(query)}&limit=${limit}`;

  // Optional: filter by type
  if (type) {
    searchUrl += `&type=${type}`;
  }

  // Optional: filter by file extensions
  if (options.fileExtensions && options.fileExtensions.length > 0) {
    searchUrl += `&file_extensions=${options.fileExtensions.join(',')}`;
  }

  // Optional: limit to specific folder(s)
  if (options.folderIds && options.folderIds.length > 0) {
    searchUrl += `&ancestor_folder_ids=${options.folderIds.join(',')}`;
  }

  // Request additional fields including parent folder info
  searchUrl += '&fields=id,name,type,description,created_at,modified_at,size,parent,path_collection,shared_link';

  try {
    const result = boxApiRequest_(searchUrl);

    return result.entries.map(item => ({
      id: item.id,
      name: item.name,
      type: item.type,
      description: item.description || '',
      createdAt: item.created_at,
      modifiedAt: item.modified_at,
      size: item.size,
      folderPath: item.path_collection?.entries?.map(e => e.name).join('/') || '',
      parentFolder: item.parent?.name || '',
      parentFolderId: item.parent?.id || '',
      parentFolderUrl: item.parent?.id ? `https://app.box.com/folder/${item.parent.id}` : '',
      sharedLink: item.shared_link?.url || null,
      webUrl: item.type === 'folder' ? `https://app.box.com/folder/${item.id}` : `https://app.box.com/file/${item.id}`
    }));

  } catch (e) {
    Logger.log(`Box search error: ${e.message}`);
    return [];
  }
}

/**
 * Search Box with fuzzy matching for vendor name
 * Tries multiple search strategies to find relevant documents
 *
 * @param {string} vendorName - The vendor name to search for
 * @returns {array} Array of matching documents
 */
function searchBoxForVendor(vendorName) {
  if (!vendorName || vendorName.trim() === '') {
    Logger.log('No vendor name provided');
    return [];
  }

  const cleanName = vendorName.trim();

  // Remove common suffixes for alternate search
  const nameWithoutSuffix = cleanName
    .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.|L\.?L\.?C\.?)$/i, '')
    .trim();

  Logger.log(`Searching Box for exact phrase: "${cleanName}"`);

  let results = [];

  // Strategy 1: Exact phrase search (quoted) - primary name
  results = searchBox(`"${cleanName}"`);

  if (results.length > 0) {
    Logger.log(`Found ${results.length} results with exact search for "${cleanName}"`);
    return results;
  }

  // Strategy 2: Try without suffix (e.g., "Ion Solar" instead of "Ion Solar, LLC")
  if (nameWithoutSuffix !== cleanName) {
    Logger.log(`Trying exact search without suffix: "${nameWithoutSuffix}"`);
    results = searchBox(`"${nameWithoutSuffix}"`);

    if (results.length > 0) {
      Logger.log(`Found ${results.length} results with exact search for "${nameWithoutSuffix}"`);
      return results;
    }
  }

  Logger.log('No Box documents found for vendor');
  return [];
}

/**
 * Get current user info (useful for testing)
 */
function getBoxCurrentUser() {
  return boxApiRequest_('/users/me');
}


/************************************************************
 * TEST FUNCTIONS
 ************************************************************/

/**
 * Test Box connection and display user info
 */
function testBoxConnection() {
  try {
    const user = getBoxCurrentUser();

    const msg = `✅ Box Connected!\n\n` +
      `User: ${user.name}\n` +
      `Email: ${user.login}\n` +
      `Space Used: ${(user.space_used / 1024 / 1024 / 1024).toFixed(2)} GB\n` +
      `Space Total: ${(user.space_amount / 1024 / 1024 / 1024).toFixed(2)} GB`;

    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);

    return user;

  } catch (e) {
    const msg = `❌ Box Connection Failed\n\n${e.message}\n\nRun authorizeBox() to authorize.`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return null;
  }
}

/**
 * Test search for a specific vendor
 * Run this with different vendor names from your list
 */
function testSearchVendor() {
  // Change this to test different vendors
  const testVendor = 'EnergyPal';  // Or any vendor from your List

  Logger.log(`\n========== TESTING BOX SEARCH ==========`);
  Logger.log(`Vendor: ${testVendor}`);
  Logger.log(`=========================================\n`);

  const results = searchBoxForVendor(testVendor);

  if (results.length === 0) {
    Logger.log('No documents found.');
    return;
  }

  Logger.log(`Found ${results.length} document(s):\n`);

  results.forEach((doc, i) => {
    Logger.log(`${i + 1}. ${doc.name}`);
    Logger.log(`   Type: ${doc.type}`);
    Logger.log(`   Path: ${doc.folderPath}`);
    Logger.log(`   Modified: ${doc.modifiedAt}`);
    Logger.log(`   URL: ${doc.webUrl}`);
    Logger.log('');
  });

  return results;
}

/**
 * Test with multiple vendors from your list
 * Pulls vendors from the List sheet and tests each
 */
function testMultipleVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName('List');

  if (!listSheet) {
    Logger.log('List sheet not found');
    return;
  }

  // Get first 10 vendors for testing
  const vendors = listSheet.getRange('A2:A11').getValues()
    .flat()
    .filter(v => v && v.toString().trim() !== '');

  Logger.log(`\n========== TESTING ${vendors.length} VENDORS ==========\n`);

  const summary = [];

  vendors.forEach(vendor => {
    const results = searchBoxForVendor(vendor);
    summary.push({
      vendor: vendor,
      found: results.length,
      firstDoc: results[0]?.name || '-'
    });

    Logger.log(`${vendor}: ${results.length} document(s) found`);

    // Rate limiting - Box has 10 req/sec limit
    Utilities.sleep(200);
  });

  Logger.log('\n========== SUMMARY ==========');
  summary.forEach(s => {
    Logger.log(`${s.vendor}: ${s.found} docs${s.found > 0 ? ` (first: ${s.firstDoc})` : ''}`);
  });

  return summary;
}


/************************************************************
 * BATTLE STATION INTEGRATION
 *
 * These functions are designed to integrate with Battle Station
 * Once tested, add calls to these from your main BattleStation.gs
 ************************************************************/

/**
 * Get Box documents for a vendor - formatted for Battle Station display
 *
 * @param {string} vendorName - The vendor name
 * @returns {object} Object with documents array and formatted HTML
 */
function getBoxDocumentsForBattleStation(vendorName) {
  const docs = searchBoxForVendor(vendorName);

  return {
    documents: docs,
    count: docs.length,
    hasDocuments: docs.length > 0,
    // Pre-formatted for display in Battle Station
    displayRows: docs.map(doc => ({
      name: doc.name,
      folder: doc.parentFolder || doc.folderPath.split('/').pop() || 'Root',
      modified: doc.modifiedAt ? new Date(doc.modifiedAt).toLocaleDateString() : '',
      url: doc.webUrl
    }))
  };
}
