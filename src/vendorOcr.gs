/************************************************************
 * VENDOR OCR UPLOAD - Extract vendor names from screenshots
 *
 * This module allows users to upload screenshots from platforms
 * that cannot be API'd (Microsoft Teams, Telegram, WhatsApp, etc.)
 * and uses OCR to identify vendor names from conversations.
 *
 * FEATURES:
 * - Upload images via file picker
 * - Google Cloud Vision API for OCR
 * - Fuzzy matching against known vendor list
 * - Display matched vendors with confidence scores
 *
 * REQUIREMENTS:
 * - Cloud Vision API enabled in Google Cloud Console
 * - Script has external request permission (already in manifest)
 ************************************************************/

const OCR_CFG = {
  // Cloud Vision API endpoint
  VISION_API_URL: 'https://vision.googleapis.com/v1/images:annotate',

  // Minimum confidence for OCR text (0-1)
  MIN_CONFIDENCE: 0.6,

  // Minimum characters for vendor name match
  MIN_VENDOR_LENGTH: 3,

  // Folder for uploaded screenshots (in Google Drive)
  UPLOADS_FOLDER_NAME: 'Battle Station OCR Uploads',

  // Settings sheet columns for OCR configuration (1-based)
  // Column V: OCR Blacklist - text patterns to ignore in OCR results
  // Column W: OCR Alias - alternate names/spellings for vendors
  // Column X: OCR Maps To - the actual vendor name the alias maps to
  SETTINGS_OCR_BLACKLIST_COL: 22,  // Column V
  SETTINGS_OCR_ALIAS_COL: 23,      // Column W
  SETTINGS_OCR_MAPS_TO_COL: 24     // Column X
};


/************************************************************
 * MAIN ENTRY POINTS (called from menu)
 ************************************************************/

/**
 * Open the OCR upload dialog
 * Users can upload images and find vendor names
 */
function openVendorOcrUpload() {
  const html = HtmlService.createHtmlOutput(getOcrUploadHtml_())
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vendor OCR Upload');
}

/**
 * Process an uploaded image and extract vendor names
 * Called from the HTML dialog via google.script.run
 *
 * @param {string} imageData - Base64 encoded image data
 * @param {string} fileName - Original file name
 * @returns {object} Results object with extracted text and matched vendors
 */
function processOcrImage(imageData, fileName) {
  try {
    Logger.log(`Processing OCR for: ${fileName}`);

    // Step 1: Extract text using Cloud Vision API
    const extractedText = extractTextFromImage_(imageData);

    if (!extractedText || extractedText.trim() === '') {
      return {
        success: false,
        error: 'No text could be extracted from the image. Please try a clearer screenshot.',
        extractedText: '',
        vendors: []
      };
    }

    Logger.log(`Extracted text length: ${extractedText.length} characters`);

    // Step 2: Get vendor list from the spreadsheet
    const vendors = getVendorList_();
    Logger.log(`Loaded ${vendors.length} vendors for matching`);

    // Step 3: Find vendor matches in the extracted text
    const matches = findVendorMatches_(extractedText, vendors);
    Logger.log(`Found ${matches.length} vendor matches`);

    // Step 4: Optionally save the screenshot to Drive
    const savedFile = saveScreenshotToDrive_(imageData, fileName);

    return {
      success: true,
      extractedText: extractedText,
      vendors: matches,
      savedFile: savedFile ? {
        name: savedFile.getName(),
        url: savedFile.getUrl()
      } : null
    };

  } catch (e) {
    Logger.log(`OCR Error: ${e.message}`);
    return {
      success: false,
      error: e.message,
      extractedText: '',
      vendors: []
    };
  }
}


/************************************************************
 * CLOUD VISION API - OCR TEXT EXTRACTION
 ************************************************************/

/**
 * Extract text from an image using Google Cloud Vision API
 *
 * @param {string} imageData - Base64 encoded image data (without data URL prefix)
 * @returns {string} Extracted text from the image
 */
function extractTextFromImage_(imageData) {
  // Remove data URL prefix if present (e.g., "data:image/png;base64,")
  const base64Data = imageData.replace(/^data:image\/[a-z]+;base64,/, '');

  // Build the Vision API request
  const requestBody = {
    requests: [{
      image: {
        content: base64Data
      },
      features: [{
        type: 'TEXT_DETECTION',
        maxResults: 10
      }]
    }]
  };

  // Get OAuth token for authenticated request
  const token = ScriptApp.getOAuthToken();

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(OCR_CFG.VISION_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    Logger.log(`Vision API Error (${responseCode}): ${responseText}`);
    throw new Error(`Vision API error: ${responseText.substring(0, 200)}`);
  }

  const result = JSON.parse(responseText);

  // Extract the full text annotation
  const annotations = result.responses?.[0]?.textAnnotations;

  if (!annotations || annotations.length === 0) {
    return '';
  }

  // First annotation contains the full extracted text
  return annotations[0].description || '';
}


/************************************************************
 * VENDOR MATCHING
 ************************************************************/

/**
 * Get the list of vendors from the List sheet
 *
 * @returns {array} Array of vendor objects {name, type, status}
 */
function getVendorList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName('List');

  if (!listSheet) {
    throw new Error('List sheet not found. Please run "Build List" first.');
  }

  const lastRow = listSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  // Get vendor name (A), source (C), status (D) columns
  const data = listSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  const vendors = [];
  for (const row of data) {
    const name = String(row[0] || '').trim();
    if (name && name.length >= OCR_CFG.MIN_VENDOR_LENGTH) {
      vendors.push({
        name: name,
        source: String(row[2] || '').trim(),
        status: String(row[3] || '').trim()
      });
    }
  }

  return vendors;
}

/**
 * Read OCR blacklist patterns from Settings sheet (Column G)
 * These patterns will be removed from OCR text before matching
 *
 * @returns {array} Array of blacklist patterns (lowercased)
 */
function getOcrBlacklist_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Settings');
  const blacklist = [];

  if (!sh) return blacklist;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return blacklist;

  // Check header
  const header = String(sh.getRange(1, OCR_CFG.SETTINGS_OCR_BLACKLIST_COL).getValue() || '').trim().toLowerCase();
  if (!header.includes('ocr') || !header.includes('blacklist')) {
    Logger.log('OCR Blacklist column not found in Settings (expected header in column V)');
    return blacklist;
  }

  const values = sh.getRange(2, OCR_CFG.SETTINGS_OCR_BLACKLIST_COL, lastRow - 1, 1).getValues().flat();

  for (const v of values) {
    const pattern = String(v || '').trim();
    if (pattern) {
      blacklist.push(pattern.toLowerCase());
    }
  }

  Logger.log(`Loaded ${blacklist.length} OCR blacklist patterns`);
  return blacklist;
}

/**
 * Read OCR alias mappings from Settings sheet (Columns H & I)
 * Aliases are alternate names that map to actual vendor names
 *
 * @returns {Map} Map of alias (lowercased) -> vendor name
 */
function getOcrAliases_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Settings');
  const aliases = new Map();

  if (!sh) return aliases;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return aliases;

  // Check headers
  const aliasHeader = String(sh.getRange(1, OCR_CFG.SETTINGS_OCR_ALIAS_COL).getValue() || '').trim().toLowerCase();
  const mapsToHeader = String(sh.getRange(1, OCR_CFG.SETTINGS_OCR_MAPS_TO_COL).getValue() || '').trim().toLowerCase();

  if (!aliasHeader.includes('alias') || !mapsToHeader.includes('maps')) {
    Logger.log('OCR Alias columns not found in Settings (expected headers in columns W & X)');
    return aliases;
  }

  const values = sh.getRange(2, OCR_CFG.SETTINGS_OCR_ALIAS_COL, lastRow - 1, 2).getValues();

  for (const row of values) {
    const alias = String(row[0] || '').trim();
    const mapsTo = String(row[1] || '').trim();

    if (alias && mapsTo) {
      aliases.set(alias.toLowerCase(), mapsTo);
    }
  }

  Logger.log(`Loaded ${aliases.size} OCR alias mappings`);
  return aliases;
}

/**
 * Apply blacklist to OCR text - remove blacklisted phrases
 *
 * @param {string} text - Original OCR text
 * @param {array} blacklist - Array of patterns to remove
 * @returns {string} Cleaned text
 */
function applyOcrBlacklist_(text, blacklist) {
  if (!blacklist || blacklist.length === 0) return text;

  let cleanedText = text;
  for (const pattern of blacklist) {
    // Create case-insensitive regex to remove pattern
    const regex = new RegExp(escapeRegex_(pattern), 'gi');
    cleanedText = cleanedText.replace(regex, ' ');
  }

  return cleanedText;
}

/**
 * Escape special regex characters in a string
 */
function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Find vendor matches in the extracted text
 * Uses both exact and fuzzy matching
 * Respects blacklist and alias settings from Settings sheet
 *
 * @param {string} text - Extracted OCR text
 * @param {array} vendors - Array of vendor objects
 * @returns {array} Array of matched vendors with context
 */
function findVendorMatches_(text, vendors) {
  // Load settings
  const blacklist = getOcrBlacklist_();
  const aliases = getOcrAliases_();

  // Apply blacklist to clean the text
  const cleanedText = applyOcrBlacklist_(text, blacklist);
  const textLower = cleanedText.toLowerCase();

  const matches = [];
  const matchedNames = new Set();

  // Build a lookup map for vendor info (for alias matching)
  const vendorInfoMap = new Map();
  for (const vendor of vendors) {
    vendorInfoMap.set(vendor.name.toLowerCase(), vendor);
  }

  // Strategy 0: Check aliases first (highest priority - user-defined mappings)
  for (const [alias, mapsTo] of aliases) {
    if (textLower.includes(alias) && !matchedNames.has(mapsTo.toLowerCase())) {
      // Look up the vendor info for the mapped name
      const vendorInfo = vendorInfoMap.get(mapsTo.toLowerCase());
      matchedNames.add(mapsTo.toLowerCase());
      matches.push({
        name: mapsTo,
        source: vendorInfo?.source || 'alias',
        status: vendorInfo?.status || '',
        matchType: `alias ("${alias}")`,
        confidence: 1.0,
        context: extractContext_(cleanedText, alias)
      });
    }
  }

  for (const vendor of vendors) {
    const vendorLower = vendor.name.toLowerCase();

    // Skip if already matched (case-insensitive deduplication)
    if (matchedNames.has(vendorLower)) continue;

    // Strategy 1: Exact match (case-insensitive)
    if (textLower.includes(vendorLower)) {
      matchedNames.add(vendorLower);
      matches.push({
        name: vendor.name,
        source: vendor.source,
        status: vendor.status,
        matchType: 'exact',
        confidence: 1.0,
        context: extractContext_(cleanedText, vendor.name)
      });
      continue;
    }

    // Strategy 2: Match without common suffixes
    const nameWithoutSuffix = vendor.name
      .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.|L\.?L\.?C\.?)$/i, '')
      .trim();

    if (nameWithoutSuffix.length >= OCR_CFG.MIN_VENDOR_LENGTH &&
        nameWithoutSuffix.toLowerCase() !== vendorLower) {
      if (textLower.includes(nameWithoutSuffix.toLowerCase())) {
        matchedNames.add(vendorLower);
        matches.push({
          name: vendor.name,
          source: vendor.source,
          status: vendor.status,
          matchType: 'partial (no suffix)',
          confidence: 0.9,
          context: extractContext_(cleanedText, nameWithoutSuffix)
        });
        continue;
      }
    }

    // Strategy 3: Word-based fuzzy match for multi-word vendor names
    const vendorWords = vendorLower.split(/\s+/).filter(w => w.length >= 3);
    if (vendorWords.length >= 2) {
      // Check if at least 2 significant words match
      const matchingWords = vendorWords.filter(word => {
        // Skip common words
        const skipWords = ['the', 'and', 'inc', 'llc', 'corp', 'company', 'co', 'services', 'group'];
        if (skipWords.includes(word)) return false;
        return textLower.includes(word);
      });

      if (matchingWords.length >= 2) {
        matchedNames.add(vendorLower);
        matches.push({
          name: vendor.name,
          source: vendor.source,
          status: vendor.status,
          matchType: `fuzzy (${matchingWords.length} words)`,
          confidence: 0.7,
          context: extractContext_(cleanedText, matchingWords[0])
        });
      }
    }
  }

  // Sort by confidence (highest first)
  matches.sort((a, b) => b.confidence - a.confidence);

  return matches;
}

/**
 * Extract surrounding context for a match
 *
 * @param {string} text - Full text
 * @param {string} match - Matched string
 * @returns {string} Context snippet (50 chars before and after)
 */
function extractContext_(text, match) {
  const lowerText = text.toLowerCase();
  const lowerMatch = match.toLowerCase();
  const index = lowerText.indexOf(lowerMatch);

  if (index === -1) return '';

  const start = Math.max(0, index - 50);
  const end = Math.min(text.length, index + match.length + 50);

  let context = text.substring(start, end);

  // Add ellipsis if truncated
  if (start > 0) context = '...' + context;
  if (end < text.length) context = context + '...';

  return context.replace(/\n/g, ' ').trim();
}


/************************************************************
 * FILE STORAGE
 ************************************************************/

/**
 * Save the uploaded screenshot to Google Drive
 *
 * @param {string} imageData - Base64 encoded image data
 * @param {string} fileName - Original file name
 * @returns {File} The saved Google Drive file
 */
function saveScreenshotToDrive_(imageData, fileName) {
  try {
    // Get or create the uploads folder
    const folder = getOrCreateUploadsFolder_();

    // Remove data URL prefix if present
    const base64Data = imageData.replace(/^data:image\/[a-z]+;base64,/, '');

    // Decode base64 to blob
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'image/png',
      `${new Date().toISOString().replace(/[:.]/g, '-')}_${fileName}`
    );

    // Save to Drive
    const file = folder.createFile(blob);
    Logger.log(`Saved screenshot: ${file.getName()} (${file.getUrl()})`);

    return file;

  } catch (e) {
    Logger.log(`Failed to save screenshot: ${e.message}`);
    return null;
  }
}

/**
 * Get or create the uploads folder in Google Drive
 *
 * @returns {Folder} The uploads folder
 */
function getOrCreateUploadsFolder_() {
  const folderName = OCR_CFG.UPLOADS_FOLDER_NAME;

  // Use shared folder as parent if configured
  return createSubfolderInShared_(folderName);
}


/************************************************************
 * HTML DIALOG
 ************************************************************/

/**
 * Generate the HTML for the OCR upload dialog
 *
 * @returns {string} HTML content
 */
function getOcrUploadHtml_() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    * {
      box-sizing: border-box;
      font-family: 'Google Sans', Arial, sans-serif;
    }

    body {
      margin: 0;
      padding: 20px;
      background: #f8f9fa;
    }

    .header {
      text-align: center;
      margin-bottom: 20px;
    }

    .header h2 {
      color: #1a73e8;
      margin: 0 0 8px 0;
    }

    .header p {
      color: #5f6368;
      margin: 0;
      font-size: 14px;
    }

    .upload-area {
      border: 2px dashed #dadce0;
      border-radius: 8px;
      padding: 40px 20px;
      text-align: center;
      background: white;
      cursor: pointer;
      transition: all 0.2s;
      margin-bottom: 20px;
    }

    .upload-area:hover {
      border-color: #1a73e8;
      background: #e8f0fe;
    }

    .upload-area.dragover {
      border-color: #1a73e8;
      background: #e8f0fe;
    }

    .upload-area.has-file {
      border-color: #34a853;
      background: #e6f4ea;
    }

    .upload-icon {
      font-size: 48px;
      margin-bottom: 10px;
    }

    .upload-text {
      color: #5f6368;
      margin-bottom: 10px;
    }

    .upload-hint {
      color: #9aa0a6;
      font-size: 12px;
    }

    .file-input {
      display: none;
    }

    .preview-container {
      display: none;
      margin-bottom: 20px;
    }

    .preview-container.show {
      display: block;
    }

    .preview-image {
      max-width: 100%;
      max-height: 200px;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    .preview-name {
      color: #5f6368;
      font-size: 12px;
      margin-top: 8px;
    }

    .btn {
      padding: 12px 24px;
      border: none;
      border-radius: 4px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
    }

    .btn-primary {
      background: #1a73e8;
      color: white;
      width: 100%;
    }

    .btn-primary:hover {
      background: #1557b0;
    }

    .btn-primary:disabled {
      background: #dadce0;
      cursor: not-allowed;
    }

    .results {
      display: none;
      margin-top: 20px;
    }

    .results.show {
      display: block;
    }

    .results-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }

    .results-header h3 {
      margin: 0;
      color: #202124;
    }

    .results-count {
      background: #e8f0fe;
      color: #1a73e8;
      padding: 4px 12px;
      border-radius: 16px;
      font-size: 12px;
      font-weight: 500;
    }

    .vendor-list {
      max-height: 250px;
      overflow-y: auto;
      background: white;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    .vendor-item {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 12px 16px;
      border-bottom: 1px solid #f1f3f4;
    }

    .vendor-item:last-child {
      border-bottom: none;
    }

    .vendor-name {
      font-weight: 500;
      color: #202124;
    }

    .vendor-meta {
      font-size: 12px;
      color: #5f6368;
      margin-top: 2px;
    }

    .vendor-badge {
      font-size: 11px;
      padding: 2px 8px;
      border-radius: 4px;
      background: #e8f0fe;
      color: #1a73e8;
    }

    .vendor-badge.exact {
      background: #e6f4ea;
      color: #137333;
    }

    .vendor-badge.partial {
      background: #fef7e0;
      color: #b06000;
    }

    .vendor-badge.fuzzy {
      background: #fce8e6;
      color: #c5221f;
    }

    .vendor-context {
      font-size: 11px;
      color: #9aa0a6;
      margin-top: 4px;
      font-style: italic;
    }

    .no-results {
      text-align: center;
      padding: 30px;
      color: #5f6368;
    }

    .extracted-text {
      margin-top: 20px;
    }

    .extracted-text summary {
      cursor: pointer;
      color: #5f6368;
      font-size: 13px;
    }

    .extracted-text pre {
      background: white;
      padding: 12px;
      border-radius: 8px;
      font-size: 11px;
      max-height: 150px;
      overflow-y: auto;
      white-space: pre-wrap;
      word-break: break-word;
    }

    .loading {
      display: none;
      text-align: center;
      padding: 20px;
    }

    .loading.show {
      display: block;
    }

    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #1a73e8;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 10px;
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    .error {
      background: #fce8e6;
      color: #c5221f;
      padding: 12px;
      border-radius: 8px;
      margin-top: 10px;
      display: none;
    }

    .error.show {
      display: block;
    }

    .saved-file {
      background: #e6f4ea;
      padding: 8px 12px;
      border-radius: 4px;
      font-size: 12px;
      margin-top: 10px;
    }

    .saved-file a {
      color: #137333;
    }
  </style>
</head>
<body>
  <div class="header">
    <h2>Vendor OCR Upload</h2>
    <p>Upload screenshots from Teams, Telegram, WhatsApp, etc. to find vendor names</p>
  </div>

  <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
    <div class="upload-icon">📷</div>
    <div class="upload-text">Click to upload or drag & drop</div>
    <div class="upload-hint">PNG, JPG, or JPEG (max 10MB)</div>
  </div>

  <input type="file" id="fileInput" class="file-input" accept="image/*" onchange="handleFileSelect(this)">

  <div class="preview-container" id="previewContainer">
    <img id="previewImage" class="preview-image">
    <div id="previewName" class="preview-name"></div>
  </div>

  <button class="btn btn-primary" id="processBtn" onclick="processImage()" disabled>
    Find Vendors in Image
  </button>

  <div class="loading" id="loading">
    <div class="spinner"></div>
    <div>Analyzing image...</div>
  </div>

  <div class="error" id="error"></div>

  <div class="results" id="results">
    <div class="results-header">
      <h3>Matched Vendors</h3>
      <span class="results-count" id="resultsCount">0 found</span>
    </div>
    <div class="vendor-list" id="vendorList"></div>
    <div class="saved-file" id="savedFile" style="display: none;"></div>
    <details class="extracted-text">
      <summary>View extracted text</summary>
      <pre id="extractedText"></pre>
    </details>
  </div>

  <script>
    let selectedFile = null;
    let imageData = null;

    // Drag and drop handlers
    const uploadArea = document.getElementById('uploadArea');

    uploadArea.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
      uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadArea.classList.remove('dragover');

      const files = e.dataTransfer.files;
      if (files.length > 0) {
        handleFile(files[0]);
      }
    });

    function handleFileSelect(input) {
      if (input.files && input.files[0]) {
        handleFile(input.files[0]);
      }
    }

    function handleFile(file) {
      // Validate file type
      if (!file.type.startsWith('image/')) {
        showError('Please select an image file (PNG, JPG, JPEG)');
        return;
      }

      // Validate file size (10MB max)
      if (file.size > 10 * 1024 * 1024) {
        showError('File is too large. Maximum size is 10MB.');
        return;
      }

      selectedFile = file;

      // Read file and create preview
      const reader = new FileReader();
      reader.onload = (e) => {
        imageData = e.target.result;

        // Show preview
        document.getElementById('previewImage').src = imageData;
        document.getElementById('previewName').textContent = file.name;
        document.getElementById('previewContainer').classList.add('show');
        document.getElementById('uploadArea').classList.add('has-file');
        document.getElementById('processBtn').disabled = false;

        // Hide any previous results/errors
        document.getElementById('results').classList.remove('show');
        document.getElementById('error').classList.remove('show');
      };
      reader.readAsDataURL(file);
    }

    function processImage() {
      if (!imageData || !selectedFile) {
        showError('Please select an image first');
        return;
      }

      // Show loading
      document.getElementById('loading').classList.add('show');
      document.getElementById('processBtn').disabled = true;
      document.getElementById('results').classList.remove('show');
      document.getElementById('error').classList.remove('show');

      // Call server-side function
      google.script.run
        .withSuccessHandler(handleResults)
        .withFailureHandler(handleError)
        .processOcrImage(imageData, selectedFile.name);
    }

    function handleResults(result) {
      document.getElementById('loading').classList.remove('show');
      document.getElementById('processBtn').disabled = false;

      if (!result.success) {
        showError(result.error || 'Failed to process image');
        return;
      }

      // Show extracted text
      document.getElementById('extractedText').textContent = result.extractedText || '(No text extracted)';

      // Show vendor matches
      const vendorList = document.getElementById('vendorList');
      vendorList.innerHTML = '';

      if (result.vendors && result.vendors.length > 0) {
        document.getElementById('resultsCount').textContent = result.vendors.length + ' found';

        result.vendors.forEach(vendor => {
          const item = document.createElement('div');
          item.className = 'vendor-item';

          let badgeClass = 'vendor-badge';
          if (vendor.matchType === 'exact') badgeClass += ' exact';
          else if (vendor.matchType.includes('partial')) badgeClass += ' partial';
          else badgeClass += ' fuzzy';

          item.innerHTML =
            '<div>' +
              '<div class="vendor-name">' + escapeHtml(vendor.name) + '</div>' +
              '<div class="vendor-meta">' + escapeHtml(vendor.source || '') + ' | ' + escapeHtml(vendor.status || 'Unknown') + '</div>' +
              (vendor.context ? '<div class="vendor-context">"' + escapeHtml(vendor.context) + '"</div>' : '') +
            '</div>' +
            '<span class="' + badgeClass + '">' + escapeHtml(vendor.matchType) + '</span>';

          vendorList.appendChild(item);
        });
      } else {
        vendorList.innerHTML = '<div class="no-results">No vendors found in the image</div>';
        document.getElementById('resultsCount').textContent = '0 found';
      }

      // Show saved file link if available
      const savedFileDiv = document.getElementById('savedFile');
      if (result.savedFile) {
        savedFileDiv.innerHTML = 'Screenshot saved: <a href="' + result.savedFile.url + '" target="_blank">' + escapeHtml(result.savedFile.name) + '</a>';
        savedFileDiv.style.display = 'block';
      } else {
        savedFileDiv.style.display = 'none';
      }

      document.getElementById('results').classList.add('show');
    }

    function handleError(error) {
      document.getElementById('loading').classList.remove('show');
      document.getElementById('processBtn').disabled = false;
      showError(error.message || 'An unexpected error occurred');
    }

    function showError(message) {
      const errorDiv = document.getElementById('error');
      errorDiv.textContent = message;
      errorDiv.classList.add('show');
    }

    function escapeHtml(text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }
  </script>
</body>
</html>
`;
}


/************************************************************
 * TESTING / DEBUGGING
 ************************************************************/

/**
 * Test OCR with a specific image URL
 * Run this from the script editor to test
 */
function testOcrWithUrl() {
  // Test with a sample image (replace with your own)
  const testImageUrl = 'https://via.placeholder.com/400x300.png?text=Test+Vendor+Name';

  try {
    const response = UrlFetchApp.fetch(testImageUrl);
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    const result = processOcrImage('data:image/png;base64,' + base64, 'test-image.png');
    Logger.log(JSON.stringify(result, null, 2));

    return result;
  } catch (e) {
    Logger.log('Test failed: ' + e.message);
    return null;
  }
}

/**
 * Test vendor matching with sample text
 * Run this from the script editor to test matching logic
 */
function testVendorMatching() {
  const sampleText = `
    Hi John,

    I spoke with EnergyPal yesterday about the solar campaign.
    They mentioned that Ion Solar LLC is also interested.

    Best regards,
    Andy
  `;

  const vendors = getVendorList_();
  const matches = findVendorMatches_(sampleText, vendors);

  Logger.log('=== TEST VENDOR MATCHING ===');
  Logger.log('Sample text: ' + sampleText.substring(0, 100) + '...');
  Logger.log('Vendors loaded: ' + vendors.length);
  Logger.log('Matches found: ' + matches.length);
  matches.forEach(m => {
    Logger.log(`  - ${m.name} (${m.matchType}, confidence: ${m.confidence})`);
  });

  return matches;
}


/************************************************************
 * OCR SETTINGS SETUP
 ************************************************************/

/**
 * Setup OCR columns in the Settings sheet
 * Adds headers for OCR Blacklist, OCR Alias, and OCR Maps To columns
 * Run this once to initialize the settings structure
 */
function setupOcrSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Settings');

  if (!sh) {
    sh = ss.insertSheet('Settings');
    Logger.log('Created Settings sheet');
  }

  // Set up OCR headers in columns G, H, I
  const headers = [
    ['OCR Blacklist', 'OCR Alias', 'OCR Maps To']
  ];

  // Check if headers already exist
  const existingG = String(sh.getRange(1, OCR_CFG.SETTINGS_OCR_BLACKLIST_COL).getValue() || '').trim();

  if (existingG.toLowerCase().includes('ocr')) {
    SpreadsheetApp.getUi().alert('OCR settings columns already exist in Settings sheet.');
    return;
  }

  // Set headers
  sh.getRange(1, OCR_CFG.SETTINGS_OCR_BLACKLIST_COL, 1, 3).setValues(headers);
  sh.getRange(1, OCR_CFG.SETTINGS_OCR_BLACKLIST_COL, 1, 3)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');

  // Add example data
  const exampleData = [
    ['Meeting scheduled', 'energypal', 'EnergyPal'],
    ['You:', 'ion solar', 'Ion Solar LLC'],
    ['Sent from my iPhone', '', ''],
    ['Typing...', '', '']
  ];

  sh.getRange(2, OCR_CFG.SETTINGS_OCR_BLACKLIST_COL, exampleData.length, 3).setValues(exampleData);

  // Auto-resize columns
  sh.autoResizeColumn(OCR_CFG.SETTINGS_OCR_BLACKLIST_COL);
  sh.autoResizeColumn(OCR_CFG.SETTINGS_OCR_ALIAS_COL);
  sh.autoResizeColumn(OCR_CFG.SETTINGS_OCR_MAPS_TO_COL);

  SpreadsheetApp.getUi().alert(
    'OCR Settings Initialized!\n\n' +
    'Column V (OCR Blacklist): Add text patterns to ignore in OCR results\n' +
    'Column W (OCR Alias): Add alternate vendor names/spellings\n' +
    'Column X (OCR Maps To): The actual vendor name the alias should match\n\n' +
    'Example entries have been added - modify as needed.'
  );

  Logger.log('OCR settings columns initialized in Settings sheet');
}

