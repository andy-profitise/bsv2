/************************************************************
 * BATTLE STATION v2 - Label-Agnostic Vendor Review Dashboard (auto-deployed via GitHub Actions)
 *
 * Features:
 * - LABEL-AGNOSTIC: Configure any Gmail label structure via Settings
 * - EMAIL LOGGING: Records all emails to Google Sheet for analysis
 * - Navigate through vendors sequentially via menu
 * - FAST MODE: Email-focused loading, skips Box/GDrive/Airtable/Calendar
 * - View vendor details, notes, status, contacts
 * - See all related emails and monday.com tasks (live search)
 * - View helpful links from monday.com
 * - Update monday.com notes directly
 * - Mark vendors as reviewed/complete
 * - Email contacts directly from Battle Station
 * - Analyze emails with Claude AI (inline links)
 * - Smart Briefing: AI-powered cross-vendor priority advisor
 * - Auto-summarize: Claude summarizes vendor state to notes
 * - Draft Reply: Claude-powered email reply drafting
 * - Email Rules: Automatic actions on inbound emails
 * - Claude API key management via Script Properties
 *
 * LABEL CONFIG: See labelConfig.gs - all label references are configurable
 * EMAIL LOG: See emailLog.gs - all emails recorded to separate Google Sheet
 ************************************************************/

const BS_CFG = {
  // Sheet names
  LIST_SHEET: 'List',
  BATTLE_SHEET: 'Battle Station',
  GMAIL_OUTPUT_SHEET: 'Gmail Review Output',
  TASKS_SHEET: 'monday.com tasks',
  
  // List sheet columns (0-based) - v2: dual Gmail links for both team members
  L_VENDOR: 0,
  L_SOURCE: 1,
  L_STATUS: 2,
  L_NOTES: 3,
  L_M1_GMAIL: 4,      // Team member 1 Gmail link
  L_M1_NO_SNOOZE: 5,  // Team member 1 no-snooze link
  L_M2_GMAIL: 6,      // Team member 2 Gmail link
  L_M2_NO_SNOOZE: 7,  // Team member 2 no-snooze link
  L_PROCESSED: 8,
  
  // Battle Station layout
  HEADER_ROWS: 3,
  DATA_START_ROW: 5,
  
  // Modern Color Palette - sleeker, more professional look
  COLOR_HEADER: '#1a73e8',        // Google Blue - main header
  COLOR_SUBHEADER: '#e8f0fe',     // Light blue - section headers
  COLOR_EMAIL: '#fef7e0',         // Warm cream - email section
  COLOR_TASK: '#e6f4ea',          // Fresh mint - tasks section
  COLOR_LINKS: '#f3e8fd',         // Soft lavender - helpful links
  COLOR_BUTTON: '#e8f0fe',        // Light blue - buttons
  COLOR_WARNING: '#fce8e6',       // Soft coral - warnings
  COLOR_SUCCESS: '#ceead6',       // Success green
  COLOR_SNOOZED: '#e1f5fe',       // Ice blue - snoozed emails
  COLOR_WAITING: '#fff8e1',       // Warm yellow - waiting
  COLOR_MISSING: '#fff3e0',       // Light amber - missing data
  COLOR_PHONEXA: '#ffe0b2',       // Peach - Phonexa waiting
  COLOR_OVERDUE: '#ffcdd2',       // Light red - overdue

  // Row highlight colors for skip/traverse
  COLOR_ROW_CHANGED: '#c8e6c9',   // Green - vendor has changes
  COLOR_ROW_SKIPPED: '#fff9c4',   // Yellow - vendor unchanged

  // Section styling
  COLOR_SECTION_BG: '#fafafa',    // Light gray for section backgrounds
  COLOR_TABLE_HEADER: '#f5f5f5',  // Table header background
  COLOR_TABLE_ALT: '#fafafa',     // Alternating row color
  COLOR_BORDER: '#e0e0e0',        // Border color
  COLOR_TEXT_MUTED: '#757575',    // Muted text
  COLOR_TEXT_LINK: '#1a73e8',     // Link color
  
  // Overdue threshold (default - can be overridden via Settings label config)
  OVERDUE_BUSINESS_HOURS: 16,

  
  // API Keys - stored in Script Properties for security
  // Set via: ⚡ Battle Station → ⚙️ Set Claude API Key (or Script Properties editor)
  MONDAY_API_TOKEN: PropertiesService.getScriptProperties().getProperty('MONDAY_API_TOKEN') || '',
  CLAUDE_API_KEY: PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY') || '',
  
  // Search terms to skip (too generic, cause false positives)
  SKIP_SEARCH_TERMS: ['LLC', 'Inc', 'Inc.', 'Corp', 'Corp.', 'Co', 'Co.', 'Ltd', 'Ltd.', 'LP', 'LLP', 'PC', 'PLLC', 'NA', 'N/A'],
  
  // monday.com Board IDs
  BUYERS_BOARD_ID: '9007735194',
  AFFILIATES_BOARD_ID: '9007716156',
  TASKS_BOARD_ID: '9007661294',
  CONTACTS_BOARD_ID: '9304296922',
  HELPFUL_LINKS_BOARD_ID: '18389463592',
  
  // monday.com Column IDs
  BUYERS_NOTES_COLUMN: 'text_mkqnvsqh',
  AFFILIATES_NOTES_COLUMN: 'text_mkrdahqz',
  BUYERS_CONTACTS_COLUMN: 'board_relation_mky0bt0z',
  AFFILIATES_CONTACTS_COLUMN: 'board_relation_mky0n0rf',
  
  // Helpful Links Column IDs
  HELPFUL_LINKS_LINK_COLUMN: 'link_mky0anm4',
  HELPFUL_LINKS_BUYERS_COLUMN: 'board_relation_mky03dt4',
  HELPFUL_LINKS_AFFILIATES_COLUMN: 'board_relation_mky0gxak',
  HELPFUL_LINKS_NOTES_COLUMN: 'text_mky08ybx',

  // Phonexa Link Column IDs
  BUYERS_PHONEXA_COLUMN: 'link_mksmwprd',
  AFFILIATES_PHONEXA_COLUMN: 'link_mksmgnc0',

  // Live Verticals Column IDs
  BUYERS_LIVE_VERTICALS_COLUMN: 'tag_mkskgt84',
  AFFILIATES_LIVE_VERTICALS_COLUMN: 'tag_mkskrddx',

  // Other Verticals Column IDs
  BUYERS_OTHER_VERTICALS_COLUMN: 'tag_mkskewmq',
  AFFILIATES_OTHER_VERTICALS_COLUMN: 'tag_mkskfs70',

  // Live Modalities Column IDs
  BUYERS_LIVE_MODALITIES_COLUMN: 'tag_mkskfmf3',
  AFFILIATES_LIVE_MODALITIES_COLUMN: 'tag_mksk7whx',

  // States Column IDs (Buyers only)
  BUYERS_STATES_COLUMN: 'dropdown_mkyam4qw',
  BUYERS_DEAD_STATES_COLUMN: 'dropdown_mkyazy2j',

  // Other Name Column ID (for alternate vendor names)
  BUYERS_OTHER_NAME_COLUMN: 'text_mkvkr178',
  AFFILIATES_OTHER_NAME_COLUMN: 'text_mksmcrpw',

  // Add to BS_CFG:
  TASKS_PROJECT_COLUMN: 'board_relation_mkqbg3mb',
  
    // Large-Scale Projects ID to Name mapping
    PROJECT_MAP: {
      '9520665110': 'Home Services',
      '9520665261': 'ACA',
      '9618492546': 'Vertical Activation',
      '9071022704': 'Monthly Returns',
      '9268820620': 'CPL/Zip Optimizations',
      '9520671333': 'Accounting/Invoices',
      '9521113689': 'System Admin',
      '9754457415': 'URL Whitelist',
      '9007621458': 'Outbound Communication',
      '9520669726': 'Pre-Onboarding',
      '9007619323': 'Appointments',
      '9080883844': 'Onboarding - Buyer',
      '9323973905': 'Onboarding - Affiliate',
      '9587318546': 'Onboarding - Vertical',
      '9080886761': 'Templates',
      '9549663466': 'Morning Meeting',
      '9681907462': 'Week of 07/28/25',
    },

  // Last Updated Column IDs
  BUYERS_LAST_UPDATED_INDEX: 15,      // Column P (1-based 16) -> 0-based 15
  AFFILIATES_LAST_UPDATED_INDEX: 16,  // Column Q (1-based 17) -> 0-based 16

  // Checksums
  CHECKSUMS_SHEET: 'BS_Checksums',
  
  // Cache sheet for Airtable/Box data
  CACHE_SHEET: 'BS_Cache',
  CACHE_MAX_AGE_HOURS: 24,  // Refresh cache if older than this

  // Airtable Contracts Configuration
  AIRTABLE_API_TOKEN: PropertiesService.getScriptProperties().getProperty('AIRTABLE_API_TOKEN') || '',
  AIRTABLE_BASE_ID: 'appc6xu9qLlOP5G5m',
  
  // Contracts 2025
  AIRTABLE_CONTRACTS_TABLE_2025: 'Contracts 2025',
  AIRTABLE_CONTRACTS_TABLE_ID_2025: 'tblREBd6zFUUZV5eU',
  AIRTABLE_CONTRACTS_VIEW_ID_2025: 'viw8X7acqwTJEUi1R',
  
  // Contracts 2024
  AIRTABLE_CONTRACTS_TABLE_2024: 'Contracts 2024',
  AIRTABLE_CONTRACTS_TABLE_ID_2024: 'tblYn8yBux9xe6sO0',
  AIRTABLE_CONTRACTS_VIEW_ID_2024: 'viwfGEvHlo8mT5FBX',
  
  AIRTABLE_VENDOR_FIELD: 'Vendor Name',
  AIRTABLE_STATUS_FIELD: 'Status',
  AIRTABLE_CONTRACT_TYPE_FIELD: 'Contract Type',
  AIRTABLE_NOTES_FIELD: 'Notes',
  AIRTABLE_SUBMITTED_BY_FIELD: 'Submitted By',
  AIRTABLE_VERTICAL_FIELD: 'Vertical',
  AIRTABLE_CREATED_DATE_FIELD: 'Created Date',
  
  // Filter values for contracts
  AIRTABLE_ALLOWED_SUBMITTERS: ['Andy Worford', 'Aden Ritz'],
  AIRTABLE_ALLOWED_VERTICALS: ['Home Services', 'Solar'],
  
  AIRTABLE_API_BASE_URL: 'https://api.airtable.com/v0',
  
  // Google Drive Vendors Folder
  GDRIVE_VENDORS_FOLDER_ID: '1fZzQZ_srKJFZab73zE_6hqDLZrn7ud_C',
  
  // Max characters for notes display (truncate with "..." if longer)
  MAX_NOTES_LENGTH: 400
};

/**
 * Add menu to Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Battle Station')
    .addItem('🔧 Setup Battle Station', 'setupBattleStation')
    .addItem('🔧 Build List', 'buildListWithGmailAndNotes')
    .addItem('🔄 Sync monday.com Data', 'syncMondayComBoards')
    .addItem('🔍 Check Duplicate Vendors', 'checkDuplicateVendors')
    .addSeparator()
    .addItem('🧠 Smart Briefing (What to do next)', 'battleStationSmartBriefing')
    .addItem('📝 Summarize & Update Notes (Claude)', 'battleStationSummarizeToNotes')
    .addItem('✉️ Draft Reply (Claude)', 'battleStationDraftReply')
    .addItem('🤖 Analyze Emails (Claude)', 'battleStationAnalyzeEmails')
    .addSeparator()
    .addItem('⏭️ Skip Unchanged', 'skipToNextChanged')
    .addItem('🔄 Skip 5 & Return (Start/Continue)', 'skip5AndReturn')
    .addItem('↩️ Return to Origin (Skip 5)', 'continueSkip5AndReturn')
    .addItem('❌ Cancel Skip 5 Session', 'cancelSkip5Session')
    .addItem('🔁 Auto-Traverse All', 'autoTraverseVendors')
    .addItem('▶ Next Vendor (Fast)', 'battleStationNext')
    .addItem('◀ Previous Vendor (Fast)', 'battleStationPrevious')
    .addSeparator()
    .addItem('⚡ Quick Refresh (Email Only)', 'battleStationQuickRefresh')
    .addItem('🔄 Refresh (Full)', 'battleStationRefresh')
    .addItem('🔄 Hard Refresh (Clear Cache)', 'battleStationHardRefresh')
    .addSeparator()
    .addItem('💾 Update monday.com Notes', 'battleStationUpdateMondayNotes')
    .addItem('✓ Mark as Reviewed', 'battleStationMarkReviewed')
    .addItem('⚑ Flag/Unflag Vendor', 'battleStationToggleFlag')
    .addItem('📧 Open Gmail Search', 'battleStationOpenGmail')
    .addItem('✉️ Email Contacts', 'battleStationEmailContacts')
    .addSeparator()
    .addItem('📧 Manage Email Rules', 'battleStationManageEmailRules')
    .addItem('📧 Process Email Rules', 'battleStationProcessEmailRules')
    .addSeparator()
    .addItem('📷 OCR Vendor Upload', 'openVendorOcrUpload')
    .addItem('⚙️ Setup Label Config', 'setupLabelConfig')
    .addItem('⚙️ Setup OCR Settings', 'setupOcrSettings')
    .addItem('⚙️ Set Claude API Key', 'battleStationSetClaudeApiKey')
    .addItem('📊 Scan Inbox to Log (for 2nd user)', 'scanInboxToLog')
    .addItem('📊 Open Email Log', 'battleStationOpenEmailLog')
    .addItem('📊 Re-log Emails for Vendor', 'battleStationRelogEmails')
    .addItem('📁 Move Spreadsheet to Shared Folder', 'moveMainSpreadsheetToSharedFolder')
    .addItem('🔍 Go to Specific Vendor...', 'battleStationGoTo')
    .addToUi();
}

/************************************************************
 * STYLING HELPER FUNCTIONS
 * Reduce repetitive styling code and ensure visual consistency
 ************************************************************/

/**
 * Apply section header styling (main sections like VENDOR INFO, EMAILS)
 * @param {Range} range - The range to style
 * @param {string} text - Header text
 * @param {string} [bgColor] - Optional background color (defaults to SUBHEADER)
 */
function styleHeader_(range, text, bgColor) {
  range.setValue(text)
    .setBackground(bgColor || BS_CFG.COLOR_SUBHEADER)
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  return range;
}

/**
 * Apply sub-section header styling (smaller headers within sections)
 * @param {Range} range - The range to style
 * @param {string} text - Header text
 */
function styleSubHeader_(range, text) {
  range.setValue(text)
    .setBackground(BS_CFG.COLOR_SECTION_BG)
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor(BS_CFG.COLOR_TEXT_LINK)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  return range;
}

/**
 * Apply table header styling (column headers in tables)
 * @param {Range} range - The range to style
 * @param {string} text - Header text
 */
function styleTableHeader_(range, text) {
  range.setValue(text)
    .setFontWeight('bold')
    .setFontSize(9)
    .setBackground(BS_CFG.COLOR_TABLE_HEADER)
    .setHorizontalAlignment('left');
  return range;
}

/**
 * Apply link cell styling
 * @param {Range} range - The range to style
 * @param {string} url - URL for the hyperlink
 * @param {string} displayText - Text to display
 */
function styleLink_(range, url, displayText) {
  range.setFormula(`=HYPERLINK("${url}", "${displayText.replace(/"/g, '""')}")`)
    .setFontColor(BS_CFG.COLOR_TEXT_LINK);
  return range;
}

/**
 * Apply empty/no data styling
 * @param {Range} range - The range to style
 * @param {string} text - Text to display (e.g., "No data found")
 */
function styleEmpty_(range, text) {
  range.setValue(text)
    .setFontStyle('italic')
    .setFontColor(BS_CFG.COLOR_TEXT_MUTED)
    .setBackground(BS_CFG.COLOR_SECTION_BG)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  return range;
}

/**
 * Apply warning/missing data styling
 * @param {Range} range - The range to style
 * @param {string} text - Warning text
 * @param {string} [linkUrl] - Optional link to fix the issue
 */
function styleWarning_(range, text, linkUrl) {
  if (linkUrl) {
    range.setFormula(`=HYPERLINK("${linkUrl}", "${text}")`)
      .setBackground(BS_CFG.COLOR_MISSING)
      .setFontColor(BS_CFG.COLOR_TEXT_LINK);
  } else {
    range.setValue(text)
      .setBackground(BS_CFG.COLOR_WARNING)
      .setFontColor('#c62828');
  }
  return range;
}

/**
 * Apply label styling (left column labels like "Vendor:", "Status:")
 * @param {Range} range - The range to style
 * @param {string} text - Label text
 */
function styleLabel_(range, text) {
  range.setValue(text)
    .setFontWeight('bold')
    .setFontColor('#424242');
  return range;
}

/**
 * Set column divider styling (thin black separator)
 * @param {Sheet} sheet - The sheet to style
 * @param {number} col - Column number for the divider
 * @param {number} startRow - Starting row
 * @param {number} numRows - Number of rows
 */
function styleColumnDivider_(sheet, col, startRow, numRows) {
  sheet.getRange(startRow, col, numRows, 1)
    .setBackground('#424242');
}

/**
 * Batch set multiple cell values and styles efficiently
 * @param {Sheet} sheet - The sheet
 * @param {Array} cells - Array of {row, col, value, styles} objects
 *   styles can include: bg, fontWeight, fontSize, fontColor, align, wrap
 */
function batchStyleCells_(sheet, cells) {
  for (const cell of cells) {
    const range = sheet.getRange(cell.row, cell.col);

    if (cell.value !== undefined) {
      if (cell.formula) {
        range.setFormula(cell.value);
      } else {
        range.setValue(cell.value);
      }
    }

    const s = cell.styles || {};
    if (s.bg) range.setBackground(s.bg);
    if (s.fontWeight) range.setFontWeight(s.fontWeight);
    if (s.fontSize) range.setFontSize(s.fontSize);
    if (s.fontColor) range.setFontColor(s.fontColor);
    if (s.align) range.setHorizontalAlignment(s.align);
    if (s.vAlign) range.setVerticalAlignment(s.vAlign);
    if (s.wrap) range.setWrap(s.wrap);
    if (s.fontStyle) range.setFontStyle(s.fontStyle);
    if (s.numberFormat) range.setNumberFormat(s.numberFormat);
  }
}

/**
 * Helper function: Get current vendor index from the display row
 */
function getCurrentVendorIndex_() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) return null;

  // Navigation bar is in row 3, format: "◀  X / Y  ▶"
  const cellValue = String(bsSh.getRange(3, 1).getValue() || '');
  const match = cellValue.match(/(\d+)\s*\/\s*\d+/);

  if (!match) {
    Logger.log(`Could not parse index from navigation: "${cellValue}"`);
    return null;
  }

  return parseInt(match[1]);
}

/**
 * Create or reset the Battle Station sheet
 */
function setupBattleStation() {
  const ss = SpreadsheetApp.getActive();
  let bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  
  if (!bsSh) {
    bsSh = ss.insertSheet(BS_CFG.BATTLE_SHEET);
  } else {
    bsSh.clear();
    bsSh.clearConditionalFormatRules();
  }
  
  // Narrower column widths for better screen fit at higher zoom
  // Left side columns (1-4)
  bsSh.setColumnWidth(1, 130);  // Labels (was 200)
  bsSh.setColumnWidth(2, 180);  // Values (was 250)
  bsSh.setColumnWidth(3, 90);   // Secondary labels (was 150)
  bsSh.setColumnWidth(4, 180);  // Secondary values (was 300)
  
  // Divider column (5) - thin black separator
  bsSh.setColumnWidth(5, 3);
  
  // Right side columns (6-9) for Contracts, Links, Documents
  bsSh.setColumnWidth(6, 130);  // Name (was 180)
  bsSh.setColumnWidth(7, 80);   // Type (was 120)
  bsSh.setColumnWidth(8, 90);   // Status (was 150)
  bsSh.setColumnWidth(9, 180);  // Notes/Folder (was 300)
  
  loadVendorData(1);
  
  SpreadsheetApp.getUi().alert('Battle Station initialized!\n\nUse the ⚡ Battle Station menu to navigate:\n- ▶ Next Vendor\n- ◀ Previous Vendor\n- 💾 Update monday.com Notes\n- ✓ Mark as Reviewed\n- ✉️ Email Contacts\n- 🤖 Analyze Emails (Claude)');
}

/**
 * Load and display data for a specific vendor by index
 */
function loadVendorData(vendorIndex, options) {
  // Default options
  options = options || {};
  const useCache = options.useCache !== undefined ? options.useCache : false;
  const forceChanged = options.forceChanged || false;  // If true, skip the ✅ indicator (used when skipToNextChanged detected a change)
  const loadMode = options.loadMode || 'full';
  const isFastMode = loadMode === 'fast';
  
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    throw new Error('Required sheets not found');
  }
  
  const totalVendors = listSh.getLastRow() - 1;
  
  if (vendorIndex < 1) vendorIndex = 1;
  if (vendorIndex > totalVendors) vendorIndex = totalVendors;
  
  const listRow = vendorIndex + 1;
  const vendorData = listSh.getRange(listRow, 1, 1, 9).getValues()[0];

  const vendor = vendorData[BS_CFG.L_VENDOR] || '';
  const source = vendorData[BS_CFG.L_SOURCE] || '';
  const status = vendorData[BS_CFG.L_STATUS] || '';
  const notes = vendorData[BS_CFG.L_NOTES] || '';
  const processed = vendorData[BS_CFG.L_PROCESSED] || false;
  
  const mondayBoardId = source.toLowerCase().includes('buyer') ? `${BS_CFG.BUYERS_BOARD_ID} (Buyers)` : 
                        source.toLowerCase().includes('affiliate') ? `${BS_CFG.AFFILIATES_BOARD_ID} (Affiliates)` : 
                        `${BS_CFG.BUYERS_BOARD_ID} (Buyers - default)`;
  
  const processedDisplay = processed ? '✅ Yes (Reviewed)' : '⚠️ No (Needs Review)';
  
  // Clear entire sheet (9 columns now - includes divider)
  const lastRow = bsSh.getMaxRows();
  if (lastRow > 0) {
    bsSh.getRange(1, 1, lastRow, 9).clearContent().clearFormat().clearDataValidations();
  }
  
  let currentRow = 1;

  // Title - full width, modern blue header with subtle shadow effect
  bsSh.getRange(currentRow, 1, 1, 9).merge()
    .setValue(isFastMode ? `⚡ BATTLE STATION [FAST]` : `⚡ BATTLE STATION`)
    .setFontSize(16).setFontWeight('bold')
    .setBackground(BS_CFG.COLOR_HEADER)
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 40);
  currentRow++;

  // Vendor name banner - prominent display (with flag if flagged)
  const vendorDisplay = isVendorFlagged_(vendor) ? `${vendor} ⚑` : vendor;
  bsSh.getRange(currentRow, 1, 1, 9).merge()
    .setValue(vendorDisplay)
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#e3f2fd')
    .setFontColor('#1565c0')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 32);
  currentRow++;

  // Navigation bar - cleaner, more modern
  const navText = `◀  ${vendorIndex} / ${totalVendors}  ▶`;
  bsSh.getRange(currentRow, 1, 1, 9).merge()
    .setValue(navText)
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#fafafa')
    .setFontColor(BS_CFG.COLOR_TEXT_MUTED);
  bsSh.setRowHeight(currentRow, 22);
  currentRow++;

  bsSh.setFrozenRows(currentRow - 1);

  // Spacer row
  bsSh.setRowHeight(currentRow, 6);
  currentRow++;

  // VENDOR INFO SECTION - using helper
  bsSh.getRange(currentRow, 1, 1, 4).merge();
  styleHeader_(bsSh.getRange(currentRow, 1), `📊 VENDOR INFO`)
    .setFontSize(11)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 26);
  
  // Track right column starting row - starts at same row as VENDOR header
  const rightColumnStartRow = currentRow;
  let rightColumnRow = rightColumnStartRow;
  
  currentRow++;
  
  // Get contacts and notes from monday.com
  ss.toast('Loading vendor details...', '📊 Loading', 2);
  const contactData = getVendorContacts_(vendor, listRow);
  const mondayNotes = contactData.notes || notes;
  const contacts = contactData.contacts;
  
  // Vendor details - build Phonexa link display
  let phonexaDisplay = '';
  let phonexaFormula = '';
  let phonexaMissing = false;
  
  if (contactData.phonexaLink) {
    phonexaDisplay = contactData.phonexaLink;
    phonexaFormula = `=HYPERLINK("${contactData.phonexaLink}", "Open in Phonexa")`;
  } else {
    // Generate monday.com link filtered by vendor name
    const encodedVendor = encodeURIComponent(vendor);
    const mondayFilterLink = `https://profitise-company.monday.com/boards/${contactData.boardId || (source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID)}?term=${encodedVendor}`;
    phonexaDisplay = '(not set)';
    phonexaFormula = `=HYPERLINK("${mondayFilterLink}", "⚠️ Add in monday.com →")`;
    phonexaMissing = true;
  }
  
  // Use live status from monday.com, fallback to List sheet
  const liveStatus = contactData.liveStatus || vendorData[BS_CFG.L_STATUS] || '';
  
  // Build Status link to appropriate board
  const vendorBoardId = source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID;
  const encodedVendorForStatus = encodeURIComponent(vendor);
  const statusLink = `https://profitise-company.monday.com/boards/${vendorBoardId}?term=${encodedVendorForStatus}`;
  const statusFormula = `=HYPERLINK("${statusLink}", "${liveStatus}${status === 'Dead' ? ' ⚠️' : ' ✅'}")`;
  
  // VENDOR INFO - 2 COLUMN LAYOUT
  // Left column: Label (col 1) + Value (col 2)
  // Right column: Label (col 3) + Value (col 4)
  
  // Row 1: Vendor | Status
  bsSh.getRange(currentRow, 1).setValue('Vendor:').setFontWeight('bold');
  bsSh.getRange(currentRow, 2).setValue(vendor);
  bsSh.getRange(currentRow, 3).setValue('Status:').setFontWeight('bold');
  const statusCell = bsSh.getRange(currentRow, 4);
  statusCell.setFormula(statusFormula).setFontColor('#1a73e8');
  if (status === 'Dead') {
    statusCell.setBackground(BS_CFG.COLOR_WARNING);
  }
  currentRow++;
  
  // Row 2: Source
  bsSh.getRange(currentRow, 1).setValue('Source:').setFontWeight('bold');
  bsSh.getRange(currentRow, 2).setValue(source);
  currentRow++;
  
  // Row 3: Live Verticals | Live Modalities
  bsSh.getRange(currentRow, 1).setValue('Live Verticals:').setFontWeight('bold');
  const liveVertCell = bsSh.getRange(currentRow, 2).setValue(contactData.liveVerticals || '(none)');
  if (!contactData.liveVerticals) liveVertCell.setBackground(BS_CFG.COLOR_MISSING);
  bsSh.getRange(currentRow, 3).setValue('Live Modalities:').setFontWeight('bold');
  const liveModCell = bsSh.getRange(currentRow, 4).setValue(contactData.liveModalities || '(none)');
  if (!contactData.liveModalities) liveModCell.setBackground(BS_CFG.COLOR_MISSING);
  currentRow++;
  
  // Row 4: Other Verticals | Phonexa Link
  bsSh.getRange(currentRow, 1).setValue('Other Verticals:').setFontWeight('bold');
  bsSh.getRange(currentRow, 2).setValue(contactData.otherVerticals || '(none)');
  bsSh.getRange(currentRow, 3).setValue('Phonexa Link:').setFontWeight('bold');
  const phonexaCell = bsSh.getRange(currentRow, 4);
  phonexaCell.setFormula(phonexaFormula).setFontColor('#1a73e8');
  if (phonexaMissing) phonexaCell.setBackground(BS_CFG.COLOR_MISSING);
  currentRow++;
  
  // Row 5: State(s) - full width (all 4 columns)
  bsSh.getRange(currentRow, 1).setValue('State(s):').setFontWeight('bold');
  const statesCell = bsSh.getRange(currentRow, 2, 1, 3).merge();
  if (contactData.states) {
    statesCell.setValue(contactData.states);
  } else if (!source.toLowerCase().includes('affiliate')) {
    // Only show as missing for Buyers (Affiliates don't have states)
    const encodedVendor = encodeURIComponent(vendor);
    const mondayStatesLink = `https://profitise-company.monday.com/boards/${contactData.boardId || BS_CFG.BUYERS_BOARD_ID}?term=${encodedVendor}`;
    statesCell.setFormula(`=HYPERLINK("${mondayStatesLink}", "⚠️ Add in monday.com")`);
    statesCell.setBackground(BS_CFG.COLOR_WARNING).setFontColor('#1a73e8');
  } else {
    statesCell.setValue('N/A');
  }
  currentRow++;
  
  // Row 6: Dead State(s) - full width (Buyers only) - strikethrough to indicate dead
  bsSh.getRange(currentRow, 1).setValue('Dead State(s):').setFontWeight('bold').setFontLine('line-through');
  const deadStatesCell = bsSh.getRange(currentRow, 2, 1, 3).merge().setFontLine('line-through');
  if (contactData.deadStates) {
    deadStatesCell.setValue(contactData.deadStates);
  } else if (!source.toLowerCase().includes('affiliate')) {
    deadStatesCell.setValue('(none)').setFontStyle('italic').setFontColor('#999999');
  } else {
    deadStatesCell.setValue('N/A');
  }
  currentRow++;
  
  // Row 7: Last Updated | Processed
  bsSh.getRange(currentRow, 1).setValue('Last Updated:').setFontWeight('bold');
  
  // lastUpdated now comes formatted from API as "Dec 3, 2025 10:25 PM"
  const lastUpdDisplay = contactData.lastUpdated || '(not available)';
  
  const lastUpdCell = bsSh.getRange(currentRow, 2).setValue(lastUpdDisplay).setHorizontalAlignment('left');
  if (!contactData.lastUpdated) lastUpdCell.setBackground(BS_CFG.COLOR_MISSING);
  bsSh.getRange(currentRow, 3).setValue('Processed:').setFontWeight('bold');
  const processedCell = bsSh.getRange(currentRow, 4).setValue(processedDisplay);
  if (processed) {
    processedCell.setBackground(BS_CFG.COLOR_SUCCESS);
  } else {
    processedCell.setBackground('#fff4e5');
  }
  currentRow++;
  
  currentRow++;
  
  // ========== LEFT SIDE (Columns 1-4) ==========
  
  // Track row where Contacts starts (for Helpful Links alignment)
  let helpfulLinksStartRow = currentRow;
  
  // CONTACTS SECTION (moved above calendar)
  if (contacts.length > 0) {
    // Define contact type priority order
    const contactTypePriority = {
      'Primary': 1,
      'Technical': 2,
      'Contracts': 3,
      'Accounting': 4,
      'Management': 5,
      'Other/Unknown': 6
    };

    // Sort contacts by: Status (Active first) -> Type priority -> Name
    contacts.sort((a, b) => {
      // First: Sort by status (Active before Not Active)
      const aActive = (a.status && a.status.toLowerCase() !== 'not active') ? 0 : 1;
      const bActive = (b.status && b.status.toLowerCase() !== 'not active') ? 0 : 1;
      
      if (aActive !== bActive) {
        return aActive - bActive;
      }
      
      // Second: Sort by contact type priority
      const aPriority = contactTypePriority[a.contactType] || 999;
      const bPriority = contactTypePriority[b.contactType] || 999;
      
      if (aPriority !== bPriority) {
        return aPriority - bPriority;
      }
      
      // Third: Sort alphabetically by name
      return a.name.localeCompare(b.name);
    });
    
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue(`👤 CONTACTS (${contacts.length})`)
      .setBackground('#f8f9fa')
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontColor('#1a73e8')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    bsSh.setRowHeight(currentRow, 24);
    currentRow++;
    
    // Header row for contacts
    bsSh.getRange(currentRow, 1).setValue('Name').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    bsSh.getRange(currentRow, 2).setValue('Email / Phone').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    bsSh.getRange(currentRow, 3).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    bsSh.getRange(currentRow, 4).setValue('Type').setFontWeight('bold').setBackground('#f3f3f3').setFontSize(9);
    currentRow++;
    
    for (const contact of contacts) {
      // Name column - ALWAYS clickable and links to Contacts board
      const encodedContact = encodeURIComponent(contact.name);
      const contactsFilterLink = `https://profitise-company.monday.com/boards/${BS_CFG.CONTACTS_BOARD_ID}?term=${encodedContact}`;
      const nameCell = bsSh.getRange(currentRow, 1)
        .setFormula(`=HYPERLINK("${contactsFilterLink}", "${contact.name}")`)
        .setFontSize(10)
        .setBackground('#f0f8ff')
        .setFontColor('#1a73e8');
      
      // Email / Phone column - highlight if both missing
      // Format phone: normalize to (XXX) XXX-XXXX
      let phone = contact.phone || '';
      if (phone) {
        // Strip all non-digits
        let digits = phone.replace(/\D/g, '');
        // Remove leading "1" if 11 digits (country code)
        if (digits.length === 11 && digits.startsWith('1')) {
          digits = digits.substring(1);
        }
        // Format as (XXX) XXX-XXXX if we have 10 digits
        if (digits.length === 10) {
          phone = digits.replace(/(\d{3})(\d{3})(\d{4})/, '($1) $2-$3');
        }
      }
      
      const emailPhone = [contact.email, phone].filter(x => x).join(' / ');
      const emailPhoneCell = bsSh.getRange(currentRow, 2).setValue(emailPhone || '(missing)').setFontSize(10);
      
      if (!emailPhone) {
        emailPhoneCell.setFormula(`=HYPERLINK("${contactsFilterLink}", "⚠️ Add in monday.com")`);
        emailPhoneCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
      } else {
        emailPhoneCell.setBackground('#f0f8ff');
      }
      
      // Status column - highlight if missing
      const statusCell = bsSh.getRange(currentRow, 3).setValue(contact.status || '(missing)').setFontSize(10);
      if (!contact.status) {
        statusCell.setFormula(`=HYPERLINK("${contactsFilterLink}", "⚠️ Add")`);
        statusCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
      } else {
        statusCell.setBackground('#f0f8ff');
      }
      
      // Contact Type column - highlight if missing
      const typeCell = bsSh.getRange(currentRow, 4).setValue(contact.contactType || '(missing)').setFontSize(10);
      if (!contact.contactType) {
        typeCell.setFormula(`=HYPERLINK("${contactsFilterLink}", "⚠️ Add")`);
        typeCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
      } else {
        typeCell.setBackground('#f0f8ff');
      }
      
      // If Not Active, strikethrough the entire row
      if (contact.status && contact.status.toLowerCase() === 'not active') {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setFontLine('line-through')
          .setFontColor('#999999');
      }
      
      currentRow++;
    }
    currentRow++;
  }
  
  // Default values for sections that may be skipped in fast mode
  let meetings = [];
  let totalMeetingCount = 0;
  let helpfulLinks = [];
  let contractsData = { hasContracts: false, contractCount: 0, contracts: [] };
  let contractsMatchedOn = '';
  let boxDocs = [];
  let gDriveFiles = [];
  let gDriveFolderFound = false;
  let gDriveFolderUrl = null;
  let gDriveMatchedOn = '';

  if (!isFastMode) {
  // Track row where Upcoming Meetings starts (for Box Documents alignment)
  const upcomingMeetingsStartRow = currentRow;

  // CALENDAR MEETINGS SECTION
  ss.toast('Checking calendar...', '📅 Loading', 2);
  // Extract contact emails to search for in calendar events
  const contactEmails = (contacts || []).map(c => c.email).filter(e => e && e.includes('@'));
  const meetingsResult = getUpcomingMeetingsForVendor_(vendor, contactEmails);
  meetings = meetingsResult.meetings || [];
  totalMeetingCount = meetingsResult.totalCount || 0;
  
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📅 UPCOMING MEETINGS (${meetings.length})`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  if (meetings.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No upcoming meetings found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    bsSh.setRowHeight(currentRow, 25);
    currentRow++;
  } else {
    // Meeting headers
    bsSh.getRange(currentRow, 1).setValue('Event').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(currentRow, 3).setValue('Time').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(currentRow, 4).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    currentRow++;
    
    for (const meeting of meetings.slice(0, 10)) {
      // Event title - clickable link
      if (meeting.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${meeting.link}", "${meeting.title.replace(/"/g, '""')}")`)
          .setFontColor('#1a73e8');
      } else {
        bsSh.getRange(currentRow, 1).setValue(meeting.title);
      }
      
      bsSh.getRange(currentRow, 2).setValue(meeting.date).setNumberFormat('@').setHorizontalAlignment('left');
      bsSh.getRange(currentRow, 3).setValue(meeting.time).setHorizontalAlignment('left');
      bsSh.getRange(currentRow, 4).setValue(meeting.status).setHorizontalAlignment('left');
      
      // Color code by timing
      if (meeting.isToday) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#fff2cc'); // Yellow for today
      } else if (meeting.isPast) {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#f3f3f3'); // Gray for past
      } else {
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#d9ead3'); // Green for upcoming
      }
      
      currentRow++;
    }
    
    // Show "more meetings" link with Google Calendar search
    if (totalMeetingCount > meetings.length || meetings.length > 10) {
      const moreCount = totalMeetingCount > meetings.length ? totalMeetingCount - meetings.length : meetings.length - 10;
      const calSearchUrl = `https://calendar.google.com/calendar/r/search?q=${encodeURIComponent(vendor)}`;
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setFormula(`=HYPERLINK("${calSearchUrl}", "🔍 ${moreCount}+ more meetings - Search in Google Calendar")`)
        .setFontStyle('italic')
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left');
      currentRow++;
    }
  }
  
  currentRow++;
  
  // ========== RIGHT SIDE (Columns 5-8) ==========
  
  // HELPFUL LINKS SECTION (right side - aligned with VENDOR INFO at top)
  ss.toast('Loading helpful links...', '🔗 Loading', 2);
  helpfulLinks = getHelpfulLinksForVendor_(vendor, listRow);
  
  const helpfulLinksUrl = `https://profitise-company.monday.com/boards/${BS_CFG.HELPFUL_LINKS_BOARD_ID}`;
  bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
    .setFormula(`=HYPERLINK("${helpfulLinksUrl}", "🔗 HELPFUL LINKS (${helpfulLinks.length})")`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(rightColumnRow, 24);
  rightColumnRow++;
  
  if (helpfulLinks.length === 0) {
    bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
      .setValue('No helpful links found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(rightColumnRow, 25);
    rightColumnRow++;
  } else {
    // Single column layout
    bsSh.getRange(rightColumnRow, 6).setValue('Description').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(rightColumnRow, 7, 1, 3).merge().setValue('Link').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    rightColumnRow++;
    
    for (const link of helpfulLinks.slice(0, 8)) {
      bsSh.getRange(rightColumnRow, 6).setValue(link.notes || '(no description)').setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      if (link.url) {
        bsSh.getRange(rightColumnRow, 7, 1, 3).merge()
          .setFormula(`=HYPERLINK("${link.url}", "${link.url.substring(0, 50)}${link.url.length > 50 ? '...' : ''}")`)
          .setFontColor('#1a73e8')
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top');
      } else {
        bsSh.getRange(rightColumnRow, 7, 1, 3).merge().setValue('(no URL)').setHorizontalAlignment('left').setVerticalAlignment('top');
      }
      
      rightColumnRow++;
    }
    
    if (helpfulLinks.length > 8) {
      bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
        .setValue(`... and ${helpfulLinks.length - 8} more links`)
        .setFontStyle('italic')
        .setHorizontalAlignment('left');
      rightColumnRow++;
    }
  }
  
  rightColumnRow++;
  
  // CONTRACTS SECTION (right side - aligned with CONTACTS)
  ss.toast('Checking contracts...', '📋 Loading', 2);
  contractsData = getVendorContracts_(vendor);
  contractsMatchedOn = contractsData.hasContracts ? vendor : '';
  
  // If no contracts found, try Other Name(s)
  if (!contractsData.hasContracts && contactData.otherName) {
    // First try the full Other Name value (in case it's like "Profitise, LLC")
    Logger.log(`No contracts for "${vendor}", trying full Other Name: "${contactData.otherName}"`);
    const fullResult = getVendorContracts_(contactData.otherName);
    
    if (fullResult.hasContracts) {
      contractsData = fullResult;
      contractsMatchedOn = contactData.otherName;
    } else if (contactData.otherName.includes(',')) {
      // If full name found nothing, try splitting by comma for multiple values
      const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
      Logger.log(`Full name found nothing, trying individual values: ${otherNames.join(', ')}`);
      
      for (const altName of otherNames) {
        // Skip generic terms that cause false positives
        if (BS_CFG.SKIP_SEARCH_TERMS.some(term => term.toLowerCase() === altName.toLowerCase())) {
          Logger.log(`Skipping generic term: "${altName}"`);
          continue;
        }
        Logger.log(`Searching Airtable contracts for: "${altName}"`);
        const altResult = getVendorContracts_(altName);
        if (altResult.hasContracts) {
          contractsData = altResult;
          contractsMatchedOn = altName;
          break; // Found contracts, stop searching
        }
      }
    }
  }
  
  // Use helpfulLinksStartRow to align with Contacts section
  let contractsRow = helpfulLinksStartRow;
  
  const airtableContractsUrl = 'https://airtable.com/appc6xu9qLlOP5G5m/tblREBd6zFUUZV5eU/viw8X7acqwTJEUi1R?blocks=hide';
  // Escape quotes in matchedOn for use in formula
  const escapedMatchedOn = contractsMatchedOn ? contractsMatchedOn.replace(/"/g, '""') : '';
  const matchedDisplay = escapedMatchedOn && contractsMatchedOn !== vendor ? ` (matched ""${escapedMatchedOn}"")` : '';
  bsSh.getRange(contractsRow, 6, 1, 4).merge()
    .setFormula(`=HYPERLINK("${airtableContractsUrl}", "📋 AIRTABLE CONTRACTS (${contractsData.contractCount})${matchedDisplay}")`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(contractsRow, 24);
  contractsRow++;
  
  if (!contractsData.hasContracts) {
    bsSh.getRange(contractsRow, 6, 1, 4).merge()
      .setValue('No contracts found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(contractsRow, 25);
    contractsRow++;
  } else {
    // Contract headers
    bsSh.getRange(contractsRow, 6).setValue('Contract').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(contractsRow, 7).setValue('Type').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(contractsRow, 8).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(contractsRow, 9).setValue('Notes').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    contractsRow++;
    
    for (const contract of contractsData.contracts.slice(0, 10)) {
      // Contract name - clickable link to Airtable
      const contractTitle = (contract.vendorName || 'View Contract')
        .replace(/"/g, '""')  // Escape double quotes for formula
        .replace(/\n/g, ' ')  // Replace newlines with space
        .replace(/\r/g, '');  // Remove carriage returns
      
      // Clean the URL - remove any problematic characters
      const cleanUrl = (contract.airtableUrl || '')
        .replace(/"/g, '%22')  // URL encode any quotes in URL
        .replace(/\s/g, '%20'); // URL encode spaces
      
      if (cleanUrl) {
        const formula = `=HYPERLINK("${cleanUrl}", "${contractTitle}")`;
        Logger.log(`Contract formula: ${formula}`);
        bsSh.getRange(contractsRow, 6)
          .setFormula(formula)
          .setFontColor('#1a73e8')
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top');
      } else {
        bsSh.getRange(contractsRow, 6).setValue(contract.vendorName || 'Contract').setHorizontalAlignment('left').setVerticalAlignment('top');
      }
      
      bsSh.getRange(contractsRow, 7).setValue(contract.contractType || '').setHorizontalAlignment('left').setVerticalAlignment('top');

      // Default blank status to "Waiting on Legal"
      const displayStatus = contract.status || 'Waiting on Legal';
      bsSh.getRange(contractsRow, 8).setValue(displayStatus).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
      bsSh.getRange(contractsRow, 9).setValue(contract.notes || '').setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Color code by status
      const status = displayStatus.toLowerCase();
      if (status.includes('active') || status.includes('signed') || status.includes('executed')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#d9ead3'); // Green for active
      } else if (status.includes('pending') || status.includes('draft')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#fff2cc'); // Yellow for pending
      } else if (status.includes('expired') || status.includes('terminated')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#f3f3f3'); // Gray for expired
      } else if (status.includes('waiting')) {
        bsSh.getRange(contractsRow, 6, 1, 4).setBackground('#fce5cd'); // Light orange for waiting
      }
      
      contractsRow++;
    }
    
    if (contractsData.contractCount > 10) {
      bsSh.getRange(contractsRow, 6, 1, 4).merge()
        .setValue(`... and ${contractsData.contractCount - 10} more contracts`)
        .setFontStyle('italic')
        .setHorizontalAlignment('left');
      contractsRow++;
    }
  }
  
  // Update rightColumnRow to track furthest row used on right side
  rightColumnRow = Math.max(rightColumnRow, contractsRow);
  
  contractsRow++; // Add spacing
  
  // BOX DOCUMENTS SECTION (right side - aligned with Upcoming Meetings)
  ss.toast('Searching Box...', '📦 Loading', 2);
  boxDocs = [];
  let boxRow = upcomingMeetingsStartRow;
  
  // Get blacklist from Settings sheet
  const boxBlacklist = getBoxBlacklist_();
  
  // Check cache for Box docs if useCache is true
  let boxDocsFromCache = null;
  if (useCache) {
    boxDocsFromCache = getCachedData_('box', vendor);
    if (boxDocsFromCache) {
      boxDocs = boxDocsFromCache;
      Logger.log(`Box docs loaded from cache: ${boxDocs.length} documents`);
    }
  }
  
  // Only search Box if not loaded from cache
  if (!boxDocsFromCache) {
    try {
      // Check if Box is authorized before searching
      const boxService = getBoxService_();
      if (boxService.hasAccess()) {
        // Search with primary vendor name
        const primaryDocs = searchBoxForVendor(vendor);
        Logger.log(`Box search for "${vendor}" found ${primaryDocs.length} results`);
      
      // Tag each result with the search term that found it
      for (const doc of primaryDocs) {
        doc.matchedOn = vendor;
        boxDocs.push(doc);
      }
      
      // Also try Other Name(s)
      if (contactData.otherName) {
        const existingIds = new Set(boxDocs.map(d => d.id));
        
        // First try the full Other Name value (in case it's like "Profitise, LLC")
        Logger.log(`Searching Box for full Other Name: "${contactData.otherName}"`);
        const fullNameDocs = searchBoxForVendor(contactData.otherName);
        Logger.log(`Box search for "${contactData.otherName}" found ${fullNameDocs.length} results`);
        
        for (const doc of fullNameDocs) {
          if (!existingIds.has(doc.id)) {
            doc.matchedOn = contactData.otherName;
            boxDocs.push(doc);
            existingIds.add(doc.id);
          }
        }
        
        // If full name found nothing, also try splitting by comma for multiple values
        if (fullNameDocs.length === 0 && contactData.otherName.includes(',')) {
          const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
          Logger.log(`Full name found nothing, trying individual values: ${otherNames.join(', ')}`);
          
          for (const altName of otherNames) {
            // Skip generic terms that cause false positives
            if (BS_CFG.SKIP_SEARCH_TERMS.some(term => term.toLowerCase() === altName.toLowerCase())) {
              Logger.log(`Skipping generic term: "${altName}"`);
              continue;
            }
            
            Logger.log(`Searching Box for: "${altName}"`);
            const otherNameDocs = searchBoxForVendor(altName);
            Logger.log(`Box search for "${altName}" found ${otherNameDocs.length} results`);
            
            // Add any new results (not already in boxDocs)
            for (const doc of otherNameDocs) {
              if (!existingIds.has(doc.id)) {
                doc.matchedOn = altName;
                boxDocs.push(doc);
                existingIds.add(doc.id);
              }
            }
          }
        }
        Logger.log(`Combined Box results: ${boxDocs.length} unique documents`);
      }
      
      // Apply blacklist - remove files blacklisted for this vendor
      if (boxBlacklist[vendor]) {
        const blacklistedIds = boxBlacklist[vendor];
        const beforeCount = boxDocs.length;
        boxDocs = boxDocs.filter(doc => !blacklistedIds.includes(doc.id));
        if (beforeCount !== boxDocs.length) {
          Logger.log(`Blacklist removed ${beforeCount - boxDocs.length} files for ${vendor}`);
        }
      }
      
      // Sort Box results: 
      // 1. Vendor name matches first, then Other Name matches (in order they were searched)
      // 2. Then by Modified date DESC
      // 3. Then by Folder name DESC
      if (boxDocs.length > 0) {
        // Build priority map: vendor name = 0, then each other name in order
        const matchPriority = { [vendor]: 0 };
        if (contactData.otherName) {
          // Full other name gets priority 1
          matchPriority[contactData.otherName] = 1;
          // Individual values get subsequent priorities
          if (contactData.otherName.includes(',')) {
            const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
            otherNames.forEach((name, idx) => {
              if (!(name in matchPriority)) {
                matchPriority[name] = idx + 2;
              }
            });
          }
        }
        
        boxDocs.sort((a, b) => {
          // First: sort by modified date DESC (newest first)
          const dateA = a.modifiedAt || '';
          const dateB = b.modifiedAt || '';
          if (dateA !== dateB) return dateB.localeCompare(dateA);
          
          // Second: sort by match priority ASC (vendor name first, then other names in order)
          const priorityA = matchPriority[a.matchedOn] ?? 999;
          const priorityB = matchPriority[b.matchedOn] ?? 999;
          if (priorityA !== priorityB) return priorityA - priorityB;
          
          // Third: sort by document name ASC
          const nameA = a.name || '';
          const nameB = b.name || '';
          return nameA.localeCompare(nameB);
        });
        
        Logger.log(`Sorted Box results by modified DESC, then matched ASC, then document ASC`);
      }
      
      // Cache the Box results
      setCachedData_('box', vendor, boxDocs);
      
    } else {
      Logger.log('Box not authorized - skipping Box search');
    }
  } catch (e) {
    Logger.log(`Box search error: ${e.message}`);
  }
  } // End of !boxDocsFromCache block
  
  bsSh.getRange(boxRow, 6, 1, 4).merge()
    .setValue(`📦 BOX DOCUMENTS (${boxDocs.length})`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(boxRow, 24);
  boxRow++;
  
  if (boxDocs.length === 0) {
    bsSh.getRange(boxRow, 6, 1, 4).merge()
      .setValue('No Box documents found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(boxRow, 25);
    boxRow++;
  } else {
    // Box document headers
    bsSh.getRange(boxRow, 6).setValue('Document').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(boxRow, 7).setValue('Folder').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(boxRow, 8).setValue('Modified').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(boxRow, 9).setValue('Matched').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    boxRow++;
    
    for (const doc of boxDocs.slice(0, 10)) {
      // Document name - clickable link to Box
      const docName = doc.name.length > 40 ? doc.name.substring(0, 37) + '...' : doc.name;
      bsSh.getRange(boxRow, 6)
        .setFormula(`=HYPERLINK("${doc.webUrl}", "${docName.replace(/"/g, '""')}")`)
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('top');
      
      // Folder - show full path with underscores, clickable link to parent folder
      // folderPath is like "All Files/Profitise/Company Name" 
      // Skip "All Files" and join with " > "
      let folderDisplayName = 'Root';
      if (doc.folderPath) {
        const pathParts = doc.folderPath.split('/').filter(p => p && p !== 'All Files');
        folderDisplayName = pathParts.join(' > ') || 'Root';
      } else if (doc.parentFolder) {
        folderDisplayName = doc.parentFolder;
      }
      
      // Truncate if too long but keep the path structure visible
      if (folderDisplayName.length > 45) {
        const parts = folderDisplayName.split(' > ');
        if (parts.length > 2) {
          folderDisplayName = parts[0] + ' > ... > ' + parts[parts.length - 1];
        } else {
          folderDisplayName = folderDisplayName.substring(0, 42) + '...';
        }
      }
      
      const folderUrl = doc.parentFolderUrl || '';
      if (folderUrl) {
        bsSh.getRange(boxRow, 7)
          .setFormula(`=HYPERLINK("${folderUrl}", "${folderDisplayName.replace(/"/g, '""')}")`)
          .setFontColor('#1a73e8')
          .setHorizontalAlignment('left')
          .setVerticalAlignment('top');
      } else {
        bsSh.getRange(boxRow, 7).setValue(folderDisplayName).setHorizontalAlignment('left').setVerticalAlignment('top');
      }
      
      // Modified date
      const modDate = doc.modifiedAt ? Utilities.formatDate(new Date(doc.modifiedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      bsSh.getRange(boxRow, 8).setValue(modDate).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Matched search term (in quotes)
      const matchedTerm = doc.matchedOn ? `"${doc.matchedOn}"` : '';
      bsSh.getRange(boxRow, 9).setValue(matchedTerm).setFontStyle('italic').setFontColor('#666666').setHorizontalAlignment('left').setVerticalAlignment('top');
      
      boxRow++;
    }
    
    if (boxDocs.length > 10) {
      bsSh.getRange(boxRow, 6, 1, 4).merge()
        .setValue(`... and ${boxDocs.length - 10} more documents`)
        .setFontStyle('italic')
        .setHorizontalAlignment('left');
      boxRow++;
    }
  }
  
  // Update rightColumnRow to track furthest row used on right side
  rightColumnRow = Math.max(rightColumnRow, boxRow);
  
  // Fetch Google Drive files now, but display section later (aligned with EMAILS)
  ss.toast('Searching Google Drive...', '📁 Loading', 2);
  gDriveFiles = [];
  gDriveFolderFound = false;
  gDriveFolderUrl = null;
  gDriveMatchedOn = '';
  
  // Check cache for GDrive files if useCache is true
  let gDriveFromCache = null;
  if (useCache) {
    gDriveFromCache = getCachedData_('gdrive', vendor);
    if (gDriveFromCache) {
      gDriveFiles = gDriveFromCache.files || [];
      gDriveFolderFound = gDriveFromCache.folderFound || false;
      gDriveFolderUrl = gDriveFromCache.folderUrl || null;
      gDriveMatchedOn = gDriveFromCache.matchedOn || '';
      Logger.log(`GDrive files loaded from cache: ${gDriveFiles.length} files`);
    }
  }
  
  // Only search GDrive if not loaded from cache
  if (!gDriveFromCache) {
    try {
      const result = getGDriveFilesForVendor_(vendor);
      gDriveFiles = result.files || [];
      gDriveFolderFound = result.folderFound || false;
      gDriveFolderUrl = result.folderUrl || null;
      if (gDriveFolderFound) gDriveMatchedOn = vendor;
      
      // Only try Other Name(s) if NO FOLDER was found (not just empty folder)
      if (!gDriveFolderFound && contactData.otherName) {
        // First try the full Other Name value (in case it's like "Profitise, LLC")
        Logger.log(`No GDrive folder for "${vendor}", trying full Other Name: "${contactData.otherName}"`);
        const fullResult = getGDriveFilesForVendor_(contactData.otherName);
        
        if (fullResult.folderFound) {
          gDriveFiles = fullResult.files || [];
          gDriveFolderFound = true;
          gDriveFolderUrl = fullResult.folderUrl || null;
          gDriveMatchedOn = contactData.otherName;
        } else if (contactData.otherName.includes(',')) {
          // If full name found nothing, try splitting by comma for multiple values
          const otherNames = contactData.otherName.split(',').map(n => n.trim()).filter(n => n.length > 0);
          Logger.log(`Full name found nothing, trying individual values: ${otherNames.join(', ')}`);
          
          for (const altName of otherNames) {
            // Skip generic terms that cause false positives
            if (BS_CFG.SKIP_SEARCH_TERMS.some(term => term.toLowerCase() === altName.toLowerCase())) {
              Logger.log(`Skipping generic term: "${altName}"`);
              continue;
            }
            Logger.log(`Searching GDrive for: "${altName}"`);
            const altResult = getGDriveFilesForVendor_(altName);
            if (altResult.folderFound) {
              gDriveFiles = altResult.files || [];
              gDriveFolderFound = true;
              gDriveFolderUrl = altResult.folderUrl || null;
              gDriveMatchedOn = altName;
              break; // Found a folder, stop searching
            }
          }
        }
      }
      
      // Cache the GDrive results
      setCachedData_('gdrive', vendor, {
        files: gDriveFiles,
        folderFound: gDriveFolderFound,
        folderUrl: gDriveFolderUrl,
        matchedOn: gDriveMatchedOn
      });
      
    } catch (e) {
      Logger.log(`Google Drive search error: ${e.message}`);
    }
  }
  
  } // end if (!isFastMode) - Calendar, Helpful Links, Contracts, Box, GDrive fetch

  // In fast mode, show compact right-side message
  if (isFastMode) {
    bsSh.getRange(rightColumnRow, 6, 1, 4).merge()
      .setValue('⚡ Fast Mode — Use 🔄 Refresh for Box, GDrive, Airtable, Calendar')
      .setBackground('#f5f5f5')
      .setFontStyle('italic')
      .setFontColor('#888888')
      .setFontSize(9)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    rightColumnRow++;
  }

  // ========== CONTINUE LEFT SIDE ==========

  // Notes section
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue('📝 NOTES')
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  const notesCell = bsSh.getRange(currentRow, 1, 2, 4).merge()
    .setValue(mondayNotes || '(no notes)')
    .setWrap(true)
    .setVerticalAlignment('top');
  
  if (!mondayNotes || mondayNotes === '(no notes)') {
    // Link to monday.com board filtered by vendor
    const encodedVendor = encodeURIComponent(vendor);
    const mondayFilterLink = `https://profitise-company.monday.com/boards/${contactData.boardId || (source.toLowerCase().includes('affiliate') ? BS_CFG.AFFILIATES_BOARD_ID : BS_CFG.BUYERS_BOARD_ID)}?term=${encodedVendor}`;
    notesCell.setFormula(`=HYPERLINK("${mondayFilterLink}", "⚠️ Add notes in monday.com")`);
    notesCell.setBackground(BS_CFG.COLOR_MISSING).setFontColor('#1a73e8');
  } else {
    notesCell.setBackground('#fafafa');
  }
  
  currentRow += 2;
  currentRow++;
  
  // Track row where Emails starts (for Google Drive alignment)
  const emailsStartRow = currentRow;
  
  if (!isFastMode) {
  // GOOGLE DRIVE FOLDER SECTION (right side - starts after Box section OR aligned with Emails, whichever is later)
  let gDriveRow = Math.max(emailsStartRow, rightColumnRow);
  
  // Get folder URL - use tracked URL or default to Vendors folder
  const displayFolderUrl = gDriveFolderUrl || `https://drive.google.com/drive/folders/${BS_CFG.GDRIVE_VENDORS_FOLDER_ID}`;
  
  bsSh.getRange(gDriveRow, 6, 1, 4).merge()
    .setFormula(`=HYPERLINK("${displayFolderUrl}", "📁 GOOGLE DRIVE (${gDriveFiles.length})")`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  bsSh.setRowHeight(gDriveRow, 24);
  gDriveRow++;
  
  if (!gDriveFolderFound) {
    // No folder found - show link to create
    const vendorsFolderUrl = `https://drive.google.com/drive/folders/${BS_CFG.GDRIVE_VENDORS_FOLDER_ID}`;
    bsSh.getRange(gDriveRow, 6, 1, 4).merge()
      .setFormula(`=HYPERLINK("${vendorsFolderUrl}", "📂 No folder found - Click to create in Vendors")`)
      .setFontStyle('italic')
      .setFontColor('#1a73e8')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(gDriveRow, 25);
    gDriveRow++;
  } else if (gDriveFiles.length === 0) {
    // Folder found but empty
    bsSh.getRange(gDriveRow, 6, 1, 4).merge()
      .setFormula(`=HYPERLINK("${displayFolderUrl}", "📂 Folder empty - Click to open")`)
      .setFontStyle('italic')
      .setFontColor('#1a73e8')
      .setBackground('#fafafa')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    bsSh.setRowHeight(gDriveRow, 25);
    gDriveRow++;
  } else {
    // Google Drive file headers - show matched term in header if available
    const matchedDisplay = gDriveMatchedOn ? ` (matched "${gDriveMatchedOn}")` : '';
    bsSh.getRange(gDriveRow, 6).setValue('File').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(gDriveRow, 7).setValue('Type').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(gDriveRow, 8).setValue('Modified').setFontWeight('bold').setBackground('#f3f3f3').setHorizontalAlignment('left');
    bsSh.getRange(gDriveRow, 9).setValue(matchedDisplay).setFontStyle('italic').setFontColor('#666666').setBackground('#f3f3f3').setHorizontalAlignment('left');
    gDriveRow++;
    
    for (const file of gDriveFiles.slice(0, 10)) {
      // File name - clickable link
      const fileName = file.name.length > 40 ? file.name.substring(0, 37) + '...' : file.name;
      bsSh.getRange(gDriveRow, 6)
        .setFormula(`=HYPERLINK("${file.url}", "${fileName.replace(/"/g, '""')}")`)
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left')
        .setVerticalAlignment('top');
      
      // File type
      bsSh.getRange(gDriveRow, 7).setValue(file.type).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Modified date
      bsSh.getRange(gDriveRow, 8).setValue(file.modified).setHorizontalAlignment('left').setVerticalAlignment('top');
      
      // Clear column 9 (no per-file matched term for Google Drive - it's at folder level)
      bsSh.getRange(gDriveRow, 9).setValue('').setBackground(null);
      
      gDriveRow++;
    }
    
    if (gDriveFiles.length > 10) {
      // Link to the vendor's folder in Google Drive
      const folderUrl = gDriveFiles[0].folderUrl || `https://drive.google.com/drive/folders/${BS_CFG.GDRIVE_VENDORS_FOLDER_ID}`;
      bsSh.getRange(gDriveRow, 6, 1, 4).merge()
        .setFormula(`=HYPERLINK("${folderUrl}", "... and ${gDriveFiles.length - 10} more files - Open Folder")`)
        .setFontStyle('italic')
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('left');
      gDriveRow++;
    }
  }
  
  // Update rightColumnRow to track furthest row used on right side
  rightColumnRow = Math.max(rightColumnRow, gDriveRow);
  } // end if (!isFastMode) - GDrive render

  // EMAILS SECTION
  ss.toast('Searching Gmail...', '📧 Loading', 2);
  const emails = getEmailsForVendor_(vendor, listRow);
  
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📧 EMAILS (${emails.length})  |  🔵 Snoozed  🔴 Overdue  🟠 Phonexa  🟢 Accounting  🟡 Waiting`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  if (emails.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No emails found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    bsSh.setRowHeight(currentRow, 25);
    currentRow++;
  } else {
    // Email headers
    bsSh.getRange(currentRow, 1).setValue('Subject').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Count').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Labels').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;
    
    for (const email of emails.slice(0, 20)) {
      bsSh.getRange(currentRow, 1).setValue(email.subject);
      const emailDateCell = bsSh.getRange(currentRow, 2);
      emailDateCell.setNumberFormat('@'); // Set format BEFORE value to prevent auto-parsing
      emailDateCell.setValue(email.date);
      bsSh.getRange(currentRow, 3).setValue(email.count).setNumberFormat('0'); // Force number format
      bsSh.getRange(currentRow, 4).setValue(email.labels);
      
      if (email.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${email.link}", "${email.subject.replace(/"/g, '""')}")`);
      }
      
      // LABEL-AGNOSTIC: Use configured labels for color coding
      const emailStatus = getEmailStatusCategory_(email);
      const cfg = getLabelConfig_();
      const hasPriority = cfg.priority_label && email.labels && email.labels.includes(cfg.priority_label);

      // Color priority: Snoozed > OVERDUE > External Wait > Accounting > Customer Wait > Active
      switch (emailStatus) {
        case 'snoozed':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_SNOOZED);
          break;
        case 'overdue':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_OVERDUE);
          bsSh.getRange(currentRow, 1, 1, 4).setFontWeight('bold');
          break;
        case 'waiting_external':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_PHONEXA);
          break;
        case 'accounting':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground('#d9ead3'); // Green
          break;
        case 'waiting_customer':
        case 'priority_waiting':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_WAITING);
          break;
        default:
          bsSh.getRange(currentRow, 1, 1, 4).setBackground('#ffffff');
      }

      // If priority label is configured and missing, make text grey
      if (cfg.priority_label && !hasPriority) {
        bsSh.getRange(currentRow, 1, 1, 4).setFontColor('#999999');
      }

      currentRow++;
    }

    if (emails.length > 20) {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`... and ${emails.length - 20} more emails (showing first 20)`)
        .setFontStyle('italic')
        .setHorizontalAlignment('center');
      currentRow++;
    }
  }

  // Show email log stats if available
  try {
    const logStats = getEmailLogStats_(vendor);
    if (logStats.total > 0) {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`📊 Email Log: ${logStats.total} total (${logStats.singlePerson} single-person) | Latest: ${logStats.latestDate ? Utilities.formatDate(logStats.latestDate, 'America/Los_Angeles', 'MMM d, yyyy') : 'N/A'}`)
        .setFontSize(8)
        .setFontStyle('italic')
        .setFontColor('#666666')
        .setBackground('#f9f9f9');
      currentRow++;
    }
  } catch (e) {
    // Non-fatal - skip log stats
  }

  bsSh.setRowHeight(currentRow, 10);
  currentRow++;
  
  // TASKS SECTION
  let tasks = getTasksForVendor_(vendor, listRow);
  
  // Filter out inappropriate onboarding tasks based on source
  // If source is Affiliates, don't show "Onboarding - Buyer" tasks
  // If source is Buyers, don't show "Onboarding - Affiliate" tasks
  const isAffiliate = source.toLowerCase().includes('affiliate');
  const isBuyer = source.toLowerCase().includes('buyer');
  
  tasks = tasks.filter(task => {
    const project = (task.project || '').toLowerCase();
    if (isAffiliate && project.includes('onboarding - buyer')) {
      return false;
    }
    if (isBuyer && project.includes('onboarding - affiliate')) {
      return false;
    }
    return true;
  });
  
  const nonDoneTasks = tasks.filter(t => !t.isDone);
  
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📋 MONDAY.COM TASKS (${nonDoneTasks.length})`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  bsSh.setRowHeight(currentRow, 24);
  currentRow++;
  
  if (tasks.length === 0) {
    // Check if vendor is in Live, Onboarding, or Paused group - should have tasks
    // Use exact matching to avoid "Preonboarding" matching "onboarding"
    const statusLower = (liveStatus || '').trim().toLowerCase();
    const needsTasksWarning = statusLower === 'live' ||
                               statusLower === 'onboarding' ||
                               statusLower === 'paused';

    if (needsTasksWarning) {
      // Show warning with link to Claude task generator
      const vendorType = source.toLowerCase().includes('affiliate') ? 'Affiliate' : 'Buyer';
      const claudeChatUrl = 'https://claude.ai/chat/33d0e36c-23ad-4e7d-b354-bd6cf3692f3f';
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setFormula(`=HYPERLINK("${claudeChatUrl}", "⚠️ No tasks - Click to generate tasks")`)
        .setFontColor('#d32f2f')
        .setBackground('#ffebee')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      bsSh.setRowHeight(currentRow, 25);
      currentRow++;

      // Add copy/paste line for vendor name and type
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`${vendor} (${vendorType})`)
        .setFontStyle('italic')
        .setFontColor('#666666')
        .setBackground('#fafafa')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      bsSh.setRowHeight(currentRow, 22);
      currentRow++;
    } else {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue('No tasks found')
        .setFontStyle('italic')
        .setBackground('#fafafa')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      bsSh.setRowHeight(currentRow, 25);
      currentRow++;
    }
  } else {
    bsSh.getRange(currentRow, 1).setValue('Task').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Created').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Project').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;
    
    for (const task of tasks) {
      // Task name - clickable link to Tasks board filtered by task name
      const encodedTask = encodeURIComponent(task.subject);
      const taskFilterLink = `https://profitise-company.monday.com/boards/${BS_CFG.TASKS_BOARD_ID}?term=${encodedTask}`;
      bsSh.getRange(currentRow, 1)
        .setFormula(`=HYPERLINK("${taskFilterLink}", "${task.subject.replace(/"/g, '""')}")`)
        .setWrap(true)
        .setFontColor('#1a73e8');
      
      // Status - append date if present (e.g., "Waiting on Profitise - 2025-12-10")
      // Status - append date if present (but not for Done tasks)
      const statusDisplay = (task.taskDate && !task.isDone) ? `${task.status} - ${task.taskDate}` : task.status;
      bsSh.getRange(currentRow, 2).setValue(statusDisplay).setWrap(true);
      const taskDateCell = bsSh.getRange(currentRow, 3);
      taskDateCell.setNumberFormat('@'); // Set format BEFORE value to prevent auto-parsing
      taskDateCell.setValue(task.created).setWrap(true);
      bsSh.getRange(currentRow, 4).setValue(task.project).setWrap(true);
      
      // Color coding for task status
      if (task.isDone) {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setFontLine('line-through')
          .setFontColor('#999999');
      } else if (task.status && task.status.toLowerCase().includes('waiting on client')) {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setBackground('#fff2cc');  // Yellow for waiting on client
      }
      
      currentRow++;
    }
  }
  
  
  // Use the greater of currentRow or rightColumnRow for final row count
  const finalRow = isFastMode ? currentRow : Math.max(currentRow, rightColumnRow);

  // Style the divider column (column 5) - black background from row 3 to end (skip in fast mode)
  if (!isFastMode && finalRow > 2) {
    bsSh.getRange(3, 5, finalRow - 2, 1)
      .setBackground('#000000')
      .setValue('');
  }

  if (finalRow > 5) {
    bsSh.autoResizeRows(5, finalRow - 5);
  }

  // Generate and compare checksum to detect changes (skip full checksum in fast mode)
  if (!isFastMode) {
  try {
    // Generate module sub-checksums
    const newModuleChecksums = generateModuleChecksums_(
      vendor, emails, tasks, contactData.notes || '', contactData.liveStatus || '',
      contactData.states || '', contractsData.contracts || [], helpfulLinks || [],
      meetings || [], boxDocs || [], gDriveFiles || [], contacts || []
    );

    // Generate email sub-checksum (for backward compatibility and early exit)
    const newEmailChecksum = newModuleChecksums.emails;

    // Get previously stored checksums
    const storedData = getStoredChecksum_(vendor);
    Logger.log(`Stored data for ${vendor}: ${storedData ? JSON.stringify({checksum: storedData.checksum, hasModules: !!storedData.moduleChecksums}) : 'null'}`);

    // Generate full checksum
    const newChecksum = generateVendorChecksum_(
      vendor, emails, tasks, contactData.notes || '', contactData.liveStatus || '',
      contactData.states || '', contractsData.contracts || [], helpfulLinks || [],
      meetings || [], boxDocs || [], gDriveFiles || []
    );

    Logger.log(`Checksum comparison for ${vendor}: stored=${storedData?.checksum} (type: ${typeof storedData?.checksum}), new=${newChecksum} (type: ${typeof newChecksum})`);
    const isUnchanged = !forceChanged && storedData && String(storedData.checksum) === String(newChecksum);
    Logger.log(`Is unchanged: ${isUnchanged}${forceChanged ? ' (forceChanged=true)' : ''}`);

    // Determine which modules changed
    const changedModules = [];
    if (storedData && storedData.moduleChecksums) {
      const stored = storedData.moduleChecksums;
      if (stored.emails !== newModuleChecksums.emails) changedModules.push('emails');
      if (stored.tasks !== newModuleChecksums.tasks) changedModules.push('tasks');
      if (stored.notes !== newModuleChecksums.notes) changedModules.push('notes');
      if (stored.status !== newModuleChecksums.status) changedModules.push('status');
      if (stored.states !== newModuleChecksums.states) changedModules.push('states');
      if (stored.contracts !== newModuleChecksums.contracts) changedModules.push('contracts');
      if (stored.helpfulLinks !== newModuleChecksums.helpfulLinks) changedModules.push('helpfulLinks');
      if (stored.meetings !== newModuleChecksums.meetings) changedModules.push('meetings');
      if (stored.boxDocs !== newModuleChecksums.boxDocs) changedModules.push('boxDocs');
      if (stored.gDriveFiles !== newModuleChecksums.gDriveFiles) changedModules.push('gDriveFiles');
      if (stored.contacts !== newModuleChecksums.contacts) changedModules.push('contacts');
    }

    if (isUnchanged) {
      // Add ✅ to title row
      const currentTitle = bsSh.getRange(1, 1).getValue();
      bsSh.getRange(1, 1).setValue(`${currentTitle} ✅`);
      Logger.log(`Added ✅ indicator for ${vendor} - no changes`);
    } else if (storedData && changedModules.length > 0) {
      // Highlight changed section headers with 🔄
      Logger.log(`Changed modules for ${vendor}: ${changedModules.join(', ')}`);

      // Map module names to their header row search patterns
      const moduleHeaderMap = {
        'emails': '📧 EMAILS',
        'tasks': '📋 MONDAY.COM TASKS',
        'meetings': '📅 UPCOMING MEETINGS',
        'contracts': '📋 AIRTABLE CONTRACTS',
        'boxDocs': '📦 BOX DOCUMENTS',
        'gDriveFiles': '📁 GOOGLE DRIVE',
        'contacts': '👤 CONTACTS',
        'notes': '📝 NOTES',
        'helpfulLinks': '🔗 HELPFUL LINKS',
        'status': '📊 VENDOR INFO',
        'states': '📊 VENDOR INFO'
      };

      // Find and update headers for changed modules
      const dataRange = bsSh.getDataRange();
      const values = dataRange.getValues();

      for (const moduleName of changedModules) {
        const searchPattern = moduleHeaderMap[moduleName];
        if (!searchPattern) continue;

        for (let row = 0; row < values.length; row++) {
          const cellValue = String(values[row][0] || '');
          if (cellValue.includes(searchPattern) && !cellValue.includes('🔄')) {
            // Add 🔄 indicator to show this section changed
            bsSh.getRange(row + 1, 1).setValue(cellValue + ' 🔄');
            Logger.log(`Marked ${searchPattern} as changed (row ${row + 1})`);
            break;
          }
        }
      }
    } else if (!storedData) {
      Logger.log(`First view for ${vendor} - no previous checksums`);
    }

    // Store the new checksums including module checksums
    storeChecksum_(vendor, newChecksum, newEmailChecksum, newModuleChecksums);
    Logger.log(`Stored checksums for ${vendor}: full=${newChecksum}`);
  } catch (e) {
    Logger.log(`Error with checksum: ${e.message}`);
  }
  } // end if (!isFastMode) - checksum generation

  const modeLabel = isFastMode ? '⚡ Fast' : '✅ Ready';
  ss.toast(`Loaded vendor ${vendorIndex} of ${totalVendors}`, modeLabel, 2);
}

/**
 * Get helpful links for a specific vendor from monday.com
 */
function getHelpfulLinksForVendor_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) return [];
  
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const source = listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '';
  
  // Determine if buyer or affiliate
  const isBuyer = source.toLowerCase().includes('buyer');
  const isAffiliate = source.toLowerCase().includes('affiliate');
  
  Logger.log(`=== HELPFUL LINKS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Source: ${source} (isBuyer: ${isBuyer}, isAffiliate: ${isAffiliate})`);
  
  // Query all helpful links with linked_items
  const query = `
    query {
      boards (ids: [${BS_CFG.HELPFUL_LINKS_BOARD_ID}]) {
        items_page (limit: 200) {
          items {
            id
            name
            column_values {
              id
              type
              text
              value
              ... on BoardRelationValue {
                linked_items {
                  id
                  name
                }
              }
            }
          }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.data?.boards?.[0]?.items_page?.items) {
      Logger.log('No helpful links found');
      return [];
    }
    
    const allLinks = result.data.boards[0].items_page.items;
    const matchingLinks = [];
    
    for (const item of allLinks) {
      let isMatch = false;
      let linkUrl = '';
      let linkNotes = '';
      
      for (const col of item.column_values) {
        // Check Buyers relation
        if (col.id === BS_CFG.HELPFUL_LINKS_BUYERS_COLUMN && col.linked_items) {
          for (const linkedItem of col.linked_items) {
            if (linkedItem.name.toLowerCase() === vendor.toLowerCase()) {
              isMatch = true;
              break;
            }
          }
        }
        
        // Check Affiliates relation
        if (col.id === BS_CFG.HELPFUL_LINKS_AFFILIATES_COLUMN && col.linked_items) {
          for (const linkedItem of col.linked_items) {
            if (linkedItem.name.toLowerCase() === vendor.toLowerCase()) {
              isMatch = true;
              break;
            }
          }
        }
        
        // Get link URL
        if (col.id === BS_CFG.HELPFUL_LINKS_LINK_COLUMN && col.text) {
          linkUrl = col.text;
        }
        
        // Get notes
        if (col.id === BS_CFG.HELPFUL_LINKS_NOTES_COLUMN && col.text) {
          linkNotes = col.text;
        }
      }
      
      if (isMatch) {
        matchingLinks.push({
          name: item.name,
          url: linkUrl,
          notes: linkNotes
        });
        Logger.log(`Found matching link: ${linkNotes || item.name} -> ${linkUrl}`);
      }
    }
    
    Logger.log(`Returning ${matchingLinks.length} helpful links`);
    return matchingLinks;
    
  } catch (e) {
    Logger.log(`Error fetching helpful links: ${e.message}`);
    return [];
  }
}

/**
 * Get contacts, notes, Phonexa link, LIVE STATUS (from group), and last updated for a vendor from monday.com API
 */
function getVendorContacts_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) return { contacts: [], notes: '', phonexaLink: '', lastUpdated: '', liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: null, boardId: null };
  
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const source = listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '';
  
  let boardId, notesColumnId, contactsColumnId, phonexaColumnId, liveVerticalsColumnId, otherVerticalsColumnId, liveModalitiesColumnId, statesColumnId, deadStatesColumnId, otherNameColumnId, isAffiliates;
  
  if (source.toLowerCase().includes('buyer')) {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
    contactsColumnId = BS_CFG.BUYERS_CONTACTS_COLUMN;
    phonexaColumnId = BS_CFG.BUYERS_PHONEXA_COLUMN;
    liveVerticalsColumnId = BS_CFG.BUYERS_LIVE_VERTICALS_COLUMN;
    otherVerticalsColumnId = BS_CFG.BUYERS_OTHER_VERTICALS_COLUMN;
    liveModalitiesColumnId = BS_CFG.BUYERS_LIVE_MODALITIES_COLUMN;
    statesColumnId = BS_CFG.BUYERS_STATES_COLUMN;
    deadStatesColumnId = BS_CFG.BUYERS_DEAD_STATES_COLUMN;
    otherNameColumnId = BS_CFG.BUYERS_OTHER_NAME_COLUMN;
    isAffiliates = false;
  } else if (source.toLowerCase().includes('affiliate')) {
    boardId = BS_CFG.AFFILIATES_BOARD_ID;
    notesColumnId = BS_CFG.AFFILIATES_NOTES_COLUMN;
    contactsColumnId = BS_CFG.AFFILIATES_CONTACTS_COLUMN;
    phonexaColumnId = BS_CFG.AFFILIATES_PHONEXA_COLUMN;
    liveVerticalsColumnId = BS_CFG.AFFILIATES_LIVE_VERTICALS_COLUMN;
    otherVerticalsColumnId = BS_CFG.AFFILIATES_OTHER_VERTICALS_COLUMN;
    liveModalitiesColumnId = BS_CFG.AFFILIATES_LIVE_MODALITIES_COLUMN;
    statesColumnId = null; // Affiliates don't have states
    deadStatesColumnId = null; // Affiliates don't have dead states
    otherNameColumnId = BS_CFG.AFFILIATES_OTHER_NAME_COLUMN;
    isAffiliates = true;
  } else {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
    contactsColumnId = BS_CFG.BUYERS_CONTACTS_COLUMN;
    phonexaColumnId = BS_CFG.BUYERS_PHONEXA_COLUMN;
    liveVerticalsColumnId = BS_CFG.BUYERS_LIVE_VERTICALS_COLUMN;
    otherVerticalsColumnId = BS_CFG.BUYERS_OTHER_VERTICALS_COLUMN;
    liveModalitiesColumnId = BS_CFG.BUYERS_LIVE_MODALITIES_COLUMN;
    statesColumnId = BS_CFG.BUYERS_STATES_COLUMN;
    deadStatesColumnId = BS_CFG.BUYERS_DEAD_STATES_COLUMN;
    otherNameColumnId = BS_CFG.BUYERS_OTHER_NAME_COLUMN;
    isAffiliates = false;
  }
  
  Logger.log(`=== CONTACTS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Board ID: ${boardId}`);
  Logger.log(`Is Affiliates: ${isAffiliates}`);
  
  const itemId = findMondayItemIdByVendor_(vendor, boardId, apiToken);
  
  if (!itemId) {
    Logger.log('Could not find monday.com item for contacts');
    return { contacts: [], notes: '', phonexaLink: '', lastUpdated: '', liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: null, boardId: boardId };
  }
  
  let lastUpdated = '';
  
  // Fetch item data including GROUP, updated_at, and column values
  const query = `
    query {
      items (ids: [${itemId}]) {
        updated_at
        group {
          id
          title
        }
        column_values {
          id
          type
          text
          value
          ... on BoardRelationValue {
            linked_items {
              id
              name
            }
          }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.data?.items?.[0]) {
      Logger.log('No item data found');
      return { contacts: [], notes: '', phonexaLink: '', lastUpdated: lastUpdated, liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: itemId, boardId: boardId };
    }
    
    // Get updated_at from the item (ISO 8601 format)
    const updatedAtRaw = result.data.items[0].updated_at || '';
    if (updatedAtRaw) {
      const updatedDate = new Date(updatedAtRaw);
      const tz = 'America/Los_Angeles';
      lastUpdated = Utilities.formatDate(updatedDate, tz, 'MMM d, yyyy h:mm a');
      Logger.log(`Last Updated (from API): ${lastUpdated}`);
    }
    
    // Get status from GROUP title (not a column!)
    const liveStatus = result.data.items[0].group?.title || '';
    Logger.log(`Live Status from Group: ${liveStatus}`);
    
    const columnValues = result.data.items[0].column_values;
    const contactIds = [];
    let notes = '';
    let phonexaLink = '';
    let liveVerticals = '';
    let otherVerticals = '';
    let liveModalities = '';
    let states = '';
    let deadStates = '';
    let otherName = '';
    
    for (const col of columnValues) {
      // Get contacts from linked_items (for 2-way board relations)
      if (col.id === contactsColumnId && col.linked_items && col.linked_items.length > 0) {
        Logger.log(`Found ${col.linked_items.length} linked contacts`);
        for (const linkedItem of col.linked_items) {
          contactIds.push(linkedItem.id);
          Logger.log(`  Contact ID: ${linkedItem.id}, Name: ${linkedItem.name}`);
        }
      }
      // Notes column
      else if (col.id === notesColumnId && col.text) {
        notes = col.text;
      }
      // Phonexa Link column
      else if (col.id === phonexaColumnId && col.text) {
        phonexaLink = col.text;
        Logger.log(`Phonexa Link: ${phonexaLink}`);
      }
      // Live Verticals column (tags)
      else if (col.id === liveVerticalsColumnId && col.text) {
        liveVerticals = col.text;
        Logger.log(`Live Verticals: ${liveVerticals}`);
      }
      // Other Verticals column (tags)
      else if (col.id === otherVerticalsColumnId && col.text) {
        otherVerticals = col.text;
        Logger.log(`Other Verticals: ${otherVerticals}`);
      }
      // Live Modalities column (tags)
      else if (col.id === liveModalitiesColumnId && col.text) {
        liveModalities = col.text;
        Logger.log(`Live Modalities: ${liveModalities}`);
      }
      // States column (dropdown - Buyers only)
      else if (statesColumnId && col.id === statesColumnId && col.text) {
        states = col.text;
        Logger.log(`States: ${states}`);
      }
      // Dead States column (dropdown - Buyers only)
      else if (deadStatesColumnId && col.id === deadStatesColumnId && col.text) {
        deadStates = col.text;
        Logger.log(`Dead States: ${deadStates}`);
      }
      // Other Name column (text - Buyers only)
      else if (otherNameColumnId && col.id === otherNameColumnId && col.text) {
        otherName = col.text;
        Logger.log(`Other Name: ${otherName}`);
      }
    }
    
    Logger.log(`Found ${contactIds.length} contact IDs`);
    Logger.log(`Notes: ${notes}`);
    
    const contacts = [];
    
    if (contactIds.length > 0) {
      const contactsQuery = `
        query {
          items (ids: [${contactIds.join(', ')}]) {
            id
            name
            column_values {
              id
              text
              type
            }
          }
        }
      `;
      
      const contactsOptions = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': apiToken },
        payload: JSON.stringify({ query: contactsQuery }),
        muteHttpExceptions: true
      };
      
      const contactsResponse = UrlFetchApp.fetch('https://api.monday.com/v2', contactsOptions);
      const contactsResult = JSON.parse(contactsResponse.getContentText());
      
      if (contactsResult.data?.items) {
        for (const item of contactsResult.data.items) {
          const contact = {
            name: item.name,
            email: '',
            phone: '',
            status: '',
            contactType: ''
          };
          
          for (const col of item.column_values) {
            if (col.id === 'email_mkrk53z4') {
              contact.email = col.text || '';
            } else if (col.id === 'phone_mkrkzxq2') {
              contact.phone = col.text || '';
            } else if (col.id === 'status') {
              contact.status = col.text || '';
            } else if (col.id === 'color_mkrkh4bk') {
              contact.contactType = col.text || '';
            }
          }
          
          contacts.push(contact);
          Logger.log(`Contact found: ${contact.name} <${contact.email}> | ${contact.phone} | ${contact.status} | ${contact.contactType}`);
        }
      }
    }
    
    Logger.log(`Returning ${contacts.length} contacts`);
    return { contacts: contacts, notes: notes, phonexaLink: phonexaLink, lastUpdated: lastUpdated, liveStatus: liveStatus, liveVerticals: liveVerticals, otherVerticals: otherVerticals, liveModalities: liveModalities, states: states, deadStates: deadStates, otherName: otherName, mondayItemId: itemId, boardId: boardId };
    
  } catch (e) {
    Logger.log(`Error fetching contacts: ${e.message}`);
    return { contacts: [], notes: '', phonexaLink: '', lastUpdated: lastUpdated, liveStatus: '', liveVerticals: '', otherVerticals: '', liveModalities: '', states: '', deadStates: '', otherName: '', mondayItemId: itemId, boardId: boardId };
  }
}

/**
 * Get emails for a specific vendor by searching Gmail live.
 * LABEL-AGNOSTIC: Uses configured labels from Settings sheet.
 * Also logs all found emails to the Email Log spreadsheet.
 *
 * Strategy:
 * 1. First try Gmail links from the List sheet (if they exist and are valid)
 * 2. Fall back to building query from label config + vendor name/contacts
 */
function getEmailsForVendor_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) {
    Logger.log('List sheet not found');
    return [];
  }

  try {
    Logger.log(`=== GMAIL SEARCH (Label-Agnostic) ===`);
    Logger.log(`Vendor: ${vendor}`);

    // Pick Gmail link columns for the current user (m1 or m2)
    const memberKey = getCurrentTeamMemberKey_();
    const gmailColIdx = memberKey === 'm2' ? BS_CFG.L_M2_GMAIL : BS_CFG.L_M1_GMAIL;
    const noSnoozeColIdx = memberKey === 'm2' ? BS_CFG.L_M2_NO_SNOOZE : BS_CFG.L_M1_NO_SNOOZE;

    const gmailLinkAll = listSh.getRange(listRow, gmailColIdx + 1).getValue();
    const gmailLinkNoSnooze = listSh.getRange(listRow, noSnoozeColIdx + 1).getValue();

    let allEmails = [];
    let noSnoozeThreadIds = new Set();
    let usedListLinks = false;

    // If List sheet has valid Gmail links, use them
    if (gmailLinkAll && gmailLinkAll.toString().includes('#search')) {
      Logger.log(`Using List sheet Gmail link: ${gmailLinkAll}`);
      const threadsAll = searchGmailFromLink_(gmailLinkAll, 'All');
      Logger.log(`Found ${threadsAll.length} threads from List link`);
      allEmails.push(...threadsAll);
      usedListLinks = true;

      if (gmailLinkNoSnooze && gmailLinkNoSnooze.toString().includes('#search')) {
        const threadsNoSnooze = searchGmailFromLink_(gmailLinkNoSnooze, 'No Snooze');
        for (const email of threadsNoSnooze) {
          noSnoozeThreadIds.add(email.threadId);
        }
      }
    }

    // If no List links or they returned nothing, use label config
    if (!usedListLinks || allEmails.length === 0) {
      Logger.log('Using label-agnostic config for email search');

      // Get contact emails for search
      const contactData = getVendorContacts_(vendor, listRow);
      const contactEmails = (contactData.contacts || [])
        .map(c => c.email)
        .filter(e => e && e.includes('@'));

      // Get vendor slug from Settings sublabel map
      const gmailSublabelMap = readGmailSublabelMap_(ss);
      const vendorSlug = gmailSublabelMap.get(vendor.toLowerCase()) || null;

      // Build queries from config
      const queries = buildVendorEmailQuery_(vendor, vendorSlug, contactEmails);

      Logger.log(`All query: ${queries.allQuery}`);
      Logger.log(`No-snooze query: ${queries.noSnoozeQuery}`);

      // Search with the "all" query
      const threadsAll = searchGmailDirect_(queries.allQuery, 'All');
      Logger.log(`Found ${threadsAll.length} threads from config query`);
      allEmails.push(...threadsAll);

      // Search with the "no snooze" query
      const threadsNoSnooze = searchGmailDirect_(queries.noSnoozeQuery, 'No Snooze');
      for (const email of threadsNoSnooze) {
        noSnoozeThreadIds.add(email.threadId);
      }
    }

    // Deduplicate and mark snoozed
    const uniqueEmails = [];
    const seenThreadIds = new Set();

    for (const email of allEmails) {
      if (!seenThreadIds.has(email.threadId)) {
        seenThreadIds.add(email.threadId);
        email.isSnoozed = noSnoozeThreadIds.size > 0 ? !noSnoozeThreadIds.has(email.threadId) : false;
        uniqueEmails.push(email);
      }
    }

    uniqueEmails.sort((a, b) => new Date(b.date) - new Date(a.date));

    Logger.log(`Returning ${uniqueEmails.length} total emails (${uniqueEmails.filter(e => e.isSnoozed).length} snoozed)`);

    // Log emails to the Email Log spreadsheet (async-safe, non-blocking on failure)
    try {
      logEmailsToSheet_(uniqueEmails, vendor);
    } catch (logErr) {
      Logger.log(`Email logging failed (non-fatal): ${logErr.message}`);
    }

    return uniqueEmails;

  } catch (e) {
    Logger.log(`ERROR in getEmailsForVendor_: ${e.message}`);
    return [];
  }
}

/**
 * Search Gmail directly with a query string (not from a URL).
 * Used by the label-agnostic system.
 */
function searchGmailDirect_(searchQuery, querySetName) {
  try {
    Logger.log(`${querySetName} search query: ${searchQuery}`);

    const threads = GmailApp.search(searchQuery, 0, 50);
    Logger.log(`${querySetName} found ${threads.length} threads`);

    const emails = [];

    for (const thread of threads) {
      const messages = thread.getMessages();
      if (messages.length === 0) continue;

      const lastMessage = messages[messages.length - 1];
      const subject = thread.getFirstMessageSubject();
      const date = lastMessage.getDate();
      const labels = thread.getLabels().map(label => label.getName()).join(', ');
      const threadId = thread.getId();
      const threadLink = buildGmailThreadUrl_(threadId);

      let snippet = '';
      try {
        snippet = lastMessage.getPlainBody().substring(0, 200) + '...';
      } catch (e) {
        snippet = '(unable to load snippet)';
      }

      const tz = 'America/Los_Angeles';
      const dateFormatted = Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');

      emails.push({
        threadId: threadId,
        subject: subject || '(no subject)',
        date: dateFormatted,
        count: messages.length,
        labels: labels,
        link: threadLink,
        querySet: querySetName,
        snippet: snippet,
        isSnoozed: false
      });
    }

    return emails;

  } catch (e) {
    Logger.log(`Error searching Gmail for ${querySetName}: ${e.message}`);
    return [];
  }
}

/**
 * Helper function to search Gmail from a link
 */
function searchGmailFromLink_(gmailLink, querySetName) {
  try {
    const gmailLinkStr = gmailLink.toString();
    const urlParts = gmailLinkStr.split('#search/');
    
    if (urlParts.length < 2) {
      Logger.log(`Could not parse Gmail search URL for ${querySetName}`);
      return [];
    }
    
    let searchQuery = decodeURIComponent(urlParts[1]);
    
    if (searchQuery.includes('?')) {
      searchQuery = searchQuery.split('?')[0];
    }
    
    searchQuery = searchQuery.replace(/\+/g, ' ');
    searchQuery = searchQuery.replace(/-is:snoozed/gi, '-label:snoozed');
    
    Logger.log(`${querySetName} search query: ${searchQuery}`);
    
    const threads = GmailApp.search(searchQuery, 0, 50);
    Logger.log(`${querySetName} found ${threads.length} threads`);
    
    const emails = [];
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      if (messages.length === 0) continue;
      
      const lastMessage = messages[messages.length - 1]; // Most recent message
      const subject = thread.getFirstMessageSubject();
      const date = lastMessage.getDate(); // Use last message date
      const labels = thread.getLabels().map(label => label.getName()).join(', ');
      const threadId = thread.getId();
      const threadLink = buildGmailThreadUrl_(threadId);
      
      let snippet = '';
      try {
        snippet = lastMessage.getPlainBody().substring(0, 200) + '...';
      } catch (e) {
        snippet = '(unable to load snippet)';
      }
      
      // Format date in Pacific timezone with leading zeros
      const tz = 'America/Los_Angeles';
      const dateFormatted = Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');
      
      emails.push({
        threadId: threadId,
        subject: subject || '(no subject)',
        date: dateFormatted,
        count: messages.length,
        labels: labels,
        link: threadLink,
        querySet: querySetName,
        snippet: snippet,
        isSnoozed: false
      });
    }
    
    return emails;
    
  } catch (e) {
    Logger.log(`Error searching Gmail for ${querySetName}: ${e.message}`);
    return [];
  }
}

/**
 * Get tasks for a specific vendor from monday.com Tasks board
 */
function getTasksForVendor_(vendor, listRow) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const tasksBoardId = BS_CFG.TASKS_BOARD_ID;
  
  Logger.log(`=== MONDAY.COM TASKS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  
  const escapedVendor = vendor.replace(/"/g, '\\"');
  
  const query = `
    query {
      boards (ids: [${tasksBoardId}]) {
        items_page (
          limit: 100
          query_params: {
            rules: [
              {
                column_id: "name"
                compare_value: ["${escapedVendor}"]
                operator: contains_text
              }
            ]
          }
        ) {
          items {
            id
            name
            group {
              id
              title
            }
            column_values {
              id
              text
              type
              ... on BoardRelationValue {
                linked_items {
                  id
                  name
                }
              }
            }
            created_at
            updated_at
          }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.errors && result.errors.length > 0) {
      Logger.log(`API Error: ${result.errors[0].message}`);
      return [];
    }
    
    if (!result.data?.boards?.[0]?.items_page?.items) {
      Logger.log('No items found in Tasks board');
      return [];
    }
    
    const items = result.data.boards[0].items_page.items;
    Logger.log(`Found ${items.length} tasks matching vendor`);
    
    const tasks = [];
    
    // Group priority for sorting (DESC = higher number first)
    const groupPriority = {
      'topics': 3,                    // Ongoing Projects
      'group_mkqb5pzw': 2,           // Upcoming/Paused Projects
      'group_title': 1,              // Completed Projects
      'group_mkqf4yzy': 0            // Task Templates
    };
    
    // Explicit project order (from Large-Scale Projects board structure)
    const projectPriority = {
      'Home Services': 1,
      'ACA': 2,
      'Vertical Activation': 3,
      'Monthly Returns': 4,
      'CPL/Zip Optimizations': 5,
      'Accounting/Invoices': 6,
      'System Admin': 7,
      'URL Whitelist': 8,
      'Outbound Communication': 9,
      'Pre-Onboarding': 10,
      'Appointments': 11,
      'Onboarding - Buyer': 12,
      'Onboarding - Affiliate': 13,
      'Onboarding - Vertical': 14,
      'Templates': 15,
      'Morning Meeting': 16,
      'Week of 07/28/25': 17
    };
    
    for (const item of items) {
      const itemName = String(item.name || '');
      const groupId = item.group?.id || '';
      const groupTitle = item.group?.title || '';
      
      let status = '';
      let projectName = '';
      let tempInd = null;
      let taskDate = '';  // Date from timeline column
      
      for (const col of item.column_values) {
        // Status column
        if (col.id === 'status' || col.id === 'status4' || col.id === 'status_1') {
          status = col.text || '';
        }
        // Project column (board relation)
        else if (col.id === BS_CFG.TASKS_PROJECT_COLUMN) {
          if (col.linked_items && col.linked_items.length > 0) {
            projectName = col.linked_items.map(li => li.name).join(', ');
          } else if (col.text) {
            projectName = col.text;
          }
        }
        // temp_ind column
        else if (col.id === 'numeric_mkt9pbj0') {
          if (col.text && col.text.trim()) {
            tempInd = parseFloat(col.text);
          }
        }
        // Date column (timeline type - shows as "YYYY-MM-DD - YYYY-MM-DD" or just end date)
        else if (col.id === 'timerange_mkqfmzyf') {
          if (col.text && col.text.trim()) {
            // Timeline format is "Start - End", we want the end date
            const dateParts = col.text.split(' - ');
            taskDate = dateParts.length > 1 ? dateParts[1].trim() : dateParts[0].trim();
          }
        }
      }
      
      const createdDate = new Date(item.created_at);
      const tz = 'America/Los_Angeles';
      const createdFormatted = Utilities.formatDate(createdDate, tz, 'yyyy-MM-dd HH:mm');
      
      tasks.push({
        subject: itemName,
        status: status || 'No Status',
        taskDate: taskDate,  // Date from timeline column
        created: createdFormatted,
        project: projectName || `Group: ${groupTitle}`,
        isDone: (status.toLowerCase() === 'done'),
        
        // Sorting fields
        groupId: groupId,
        groupPriority: groupPriority[groupId] || -1,
        projectName: projectName,
        projectPriority: projectPriority[projectName] || 999,
        tempInd: tempInd,
        tempIndSort: (tempInd === null) ? 999999 : tempInd  // Blanks at end
      });
      
      Logger.log(`Task: ${itemName} | Group: ${groupTitle} | Project: ${projectName} | temp_ind: ${tempInd}`);
    }
    
    // Sort by: Done status (not done first)
    // For NOT DONE: Group (DESC) -> Project (explicit order) -> temp_ind (ASC, blanks last) -> Created Date (DESC)
    // For DONE: Created Date (DESC) only
    tasks.sort((a, b) => {
      // 0. Done status (not done before done)
      if (a.isDone !== b.isDone) {
        return a.isDone ? 1 : -1;
      }
      
      // If BOTH are done, sort only by Created Date DESC
      if (a.isDone && b.isDone) {
        const dateA = a.created ? new Date(a.created.replace(' ', 'T')) : new Date(0);
        const dateB = b.created ? new Date(b.created.replace(' ', 'T')) : new Date(0);
        return dateB - dateA;
      }
      
      // For not-done tasks, use full sorting logic:
      // 1. Group priority (DESC - higher numbers first)
      if (a.groupPriority !== b.groupPriority) {
        return b.groupPriority - a.groupPriority;
      }
      
      // 2. Project (explicit order from Large-Scale Projects)
      if (a.projectPriority !== b.projectPriority) {
        return a.projectPriority - b.projectPriority;
      }
      
      // 3. temp_ind (ASC - low to high, blanks at end)
      if (a.tempIndSort !== b.tempIndSort) {
        return a.tempIndSort - b.tempIndSort;
      }
      
      // 4. Created Date (DESC - newest first)
      const dateA = a.created ? new Date(a.created.replace(' ', 'T')) : new Date(0);
      const dateB = b.created ? new Date(b.created.replace(' ', 'T')) : new Date(0);
      return dateB - dateA;
    });
    
    return tasks.slice(0, 30);
    
  } catch (e) {
    Logger.log(`Error fetching monday.com tasks: ${e.message}`);
    return [];
  }
}

/**
 * Get upcoming calendar meetings for a vendor
 * Searches Google Calendar for events containing the vendor name
 * Searches in: title, description/notes, location, and attendee emails
 * Returns meetings from 30 days ago to 60 days in the future
 */
function getUpcomingMeetingsForVendor_(vendor, contactEmails) {
  Logger.log(`=== CALENDAR SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  if (contactEmails && contactEmails.length > 0) {
    Logger.log(`Contact emails: ${contactEmails.join(', ')}`);
  }
  
  try {
    const now = new Date();
    // Force Pacific timezone for consistent date comparison
    const tz = 'America/Los_Angeles';
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    Logger.log(`Today's date (${tz}): ${todayStr}`);
    Logger.log(`Current time UTC: ${now.toISOString()}`);
    
    const pastDate = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000)); // 30 days ago
    const futureDate = new Date(now.getTime() + (60 * 24 * 60 * 60 * 1000)); // 60 days ahead
    
    const calendars = CalendarApp.getAllCalendars();
    const meetings = [];
    
    // Only search calendars owned by the user (not shared calendars)
    const ownedCalendars = calendars.filter(cal => {
      try {
        // Primary calendar or calendars where user is owner
        return cal.isOwnedByMe();
      } catch (e) {
        return false;
      }
    });
    
    Logger.log(`Searching ${ownedCalendars.length} owned calendars (filtered from ${calendars.length} total)`);
    
    // Search terms - vendor name and variations
    const searchTerms = [vendor.toLowerCase()];
    
    // Also try without parentheses
    const withoutParens = vendor.replace(/\s*\([^)]*\)/g, '').trim().toLowerCase();
    if (withoutParens !== vendor.toLowerCase()) {
      searchTerms.push(withoutParens);
    }
    
    // Also try first word only (company name) - but only if vendor is a single word and 4+ chars
    const words = vendor.split(/[\s\-\(\)]+/).filter(w => w.length > 0);
    if (words.length === 1) {
      const firstWord = words[0].toLowerCase();
      if (firstWord.length >= 4 && !searchTerms.includes(firstWord)) {
        searchTerms.push(firstWord);
      }
    }
    
    // Also try without common suffixes like LLC, Inc, Corp
    const withoutSuffix = vendor.replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.?)$/i, '').trim().toLowerCase();
    if (withoutSuffix !== vendor.toLowerCase() && !searchTerms.includes(withoutSuffix)) {
      searchTerms.push(withoutSuffix);
    }
    
    // Add contact emails as search terms (for matching attendees)
    const emailSearchTerms = [];
    if (contactEmails && contactEmails.length > 0) {
      for (const email of contactEmails) {
        if (email && email.includes('@')) {
          const emailLower = email.toLowerCase().trim();
          if (!emailSearchTerms.includes(emailLower)) {
            emailSearchTerms.push(emailLower);
          }
        }
      }
    }
    
    // For multi-word vendor names, require at least 2 consecutive words to match
    // This prevents "Solar Energy World" from matching just "Solar"
    const isMultiWord = words.length >= 2;
    
    Logger.log(`Search terms: ${searchTerms.join(', ')} (multi-word: ${isMultiWord})`);
    if (emailSearchTerms.length > 0) {
      Logger.log(`Email search terms: ${emailSearchTerms.join(', ')}`);
    }
    
    for (const calendar of ownedCalendars) {
      try {
        const events = calendar.getEvents(pastDate, futureDate);
        
        for (const event of events) {
          const title = event.getTitle() || '';
          const description = event.getDescription() || '';
          const location = event.getLocation() || '';
          
          // Get attendee emails to search
          let attendeeEmails = '';
          try {
            const guests = event.getGuestList();
            attendeeEmails = guests.map(g => g.getEmail()).join(' ');
          } catch (e) {
            // Some events may not allow guest list access
          }
          
          // Combine all searchable text: title, description/notes, location, attendee emails
          const searchText = (title + ' ' + description + ' ' + location + ' ' + attendeeEmails).toLowerCase();
          
          // Check if any search term matches
          let isMatch = false;
          let matchedTerm = '';
          let matchedIn = '';
          
          // First check vendor name search terms against all searchable text
          for (const term of searchTerms) {
            if (searchText.includes(term)) {
              // For multi-word vendors, require at least 2 words from vendor name to match
              // Skip single-word matches like just "solar" for "Solar Energy World"
              if (isMultiWord) {
                const termWords = term.split(/\s+/);
                if (termWords.length < 2) {
                  // Single word term - skip unless it's a unique identifier (not common words)
                  const commonWords = ['solar', 'energy', 'home', 'power', 'green', 'sun', 'electric', 'services', 'solutions', 'group', 'pro', 'usa', 'national'];
                  if (commonWords.includes(term)) {
                    continue; // Skip common single words for multi-word vendors
                  }
                }
              }
              isMatch = true;
              matchedTerm = term;
              matchedIn = title.toLowerCase().includes(term) ? 'title' : 
                         description.toLowerCase().includes(term) ? 'notes' :
                         location.toLowerCase().includes(term) ? 'location' : 'attendees';
              break;
            }
          }
          
          // If no match yet, check contact emails against attendee emails
          if (!isMatch && emailSearchTerms.length > 0 && attendeeEmails) {
            const attendeeEmailsLower = attendeeEmails.toLowerCase();
            for (const email of emailSearchTerms) {
              if (attendeeEmailsLower.includes(email)) {
                isMatch = true;
                matchedTerm = email;
                matchedIn = 'contact email';
                break;
              }
            }
          }
          
          if (isMatch) {
            const startTime = event.getStartTime();
            const endTime = event.getEndTime();
            const isAllDay = event.isAllDayEvent();
            
            // Determine if past, today, or future using script timezone
            const eventStr = Utilities.formatDate(startTime, tz, 'yyyy-MM-dd');
            
            const isPast = startTime < now;
            const isToday = (eventStr === todayStr);
            
            Logger.log(`Event: ${event.getTitle()} | Date: ${eventStr} | Today: ${todayStr} | isToday: ${isToday}`);
            
            // Determine status
            let status = '';
            if (isPast) {
              status = '✓ Past';
            } else if (isToday) {
              status = '🔴 Today';
            } else {
              // Calculate days until using date strings to avoid timezone issues
              const eventDate = new Date(eventStr + 'T00:00:00');
              const todayDate = new Date(todayStr + 'T00:00:00');
              const daysUntil = Math.round((eventDate - todayDate) / (24 * 60 * 60 * 1000));
              status = `📅 In ${daysUntil} day${daysUntil !== 1 ? 's' : ''}`;
            }
            
            // Note where the match was found (use matchedIn from earlier detection)
            const matchSource = matchedIn || 'unknown';
            
            // Format times in Pacific timezone (use tz already defined above)
            const startHH = Utilities.formatDate(startTime, tz, 'HH');
            const startMM = Utilities.formatDate(startTime, tz, 'mm');
            const endHH = Utilities.formatDate(endTime, tz, 'HH');
            const endMM = Utilities.formatDate(endTime, tz, 'mm');
            const timeFormatted = isAllDay ? 'All Day' : `${startHH}:${startMM} - ${endHH}:${endMM}`;
            
            meetings.push({
              title: title,
              date: Utilities.formatDate(startTime, tz, 'yyyy-MM-dd'),
              time: timeFormatted,
              status: status,
              isPast: isPast,
              isToday: isToday,
              startTime: startTime,
              matchSource: matchSource,
              link: `https://calendar.google.com/calendar/event?eid=${Utilities.base64Encode(event.getId().split('@')[0] + ' ' + calendar.getId())}`
            });
            
            Logger.log(`Found meeting: ${title} on ${Utilities.formatDate(startTime, tz, 'yyyy-MM-dd HH:mm')} (matched in ${matchSource})`);
          }
        }
      } catch (calError) {
        Logger.log(`Error searching calendar ${calendar.getName()}: ${calError.message}`);
      }
    }
    
    // Sort by date (upcoming first, then past)
    meetings.sort((a, b) => {
      // Today first
      if (a.isToday && !b.isToday) return -1;
      if (!a.isToday && b.isToday) return 1;
      
      // Future before past
      if (!a.isPast && b.isPast) return -1;
      if (a.isPast && !b.isPast) return 1;
      
      // Within same category, sort by date
      return a.startTime - b.startTime;
    });
    
    // Remove duplicates - only keep the FIRST instance of each meeting title
    // Also filter out past "checkup" meetings
    // For "Day ##" pattern meetings, dedupe based on title without the day number
    const seenTitles = new Set();
    const uniqueMeetings = meetings.filter(m => {
      // Skip past events with "checkup" in the name
      if (m.isPast && m.title.toLowerCase().includes('checkup')) {
        return false;
      }
      
      // Normalize title for deduplication:
      // Remove "Day ##" or "Day #" patterns (e.g., "Solar Energy World - FL (zips) - $51 Checkup - Day 3" -> "Solar Energy World - FL (zips) - $51 Checkup")
      const normalizedTitle = m.title
        .replace(/\s*-?\s*Day\s*\d+\s*$/i, '')  // Remove " - Day 3" or " Day 14" at end
        .replace(/\s*-?\s*Day\s*\d+\s*-/i, ' - ')  // Remove "Day 3 -" in middle
        .trim();
      
      if (seenTitles.has(normalizedTitle)) return false;
      seenTitles.add(normalizedTitle);
      return true;
    });
    
    // Return both unique meetings and total count for "more" link
    Logger.log(`Found ${uniqueMeetings.length} unique meetings for ${vendor} (${meetings.length} total)`);
    return { meetings: uniqueMeetings, totalCount: meetings.length };
    
  } catch (e) {
    Logger.log(`Error fetching calendar meetings: ${e.message}`);
    return { meetings: [], totalCount: 0 };
  }
}

/************************************************************
 * AIRTABLE CONTRACTS FUNCTIONS
 * Polls Airtable API to get Contract records linked to vendors
 ************************************************************/

/**
 * Fuzzy match vendor names (handles variations)
 */
function isFuzzyMatch_(name1, name2) {
  if (!name1 || !name2) return false;
  
  // Normalize both names
  const n1 = name1.toLowerCase().replace(/[^a-z0-9]/g, '');
  const n2 = name2.toLowerCase().replace(/[^a-z0-9]/g, '');
  
  // Exact match after normalization
  if (n1 === n2) return true;
  
  // One contains the other
  if (n1.includes(n2) || n2.includes(n1)) return true;
  
  // Check if significant portion matches (at least 50% of shorter string in longer)
  const shorter = n1.length < n2.length ? n1 : n2;
  const longer = n1.length < n2.length ? n2 : n1;
  
  if (longer.includes(shorter) && shorter.length >= longer.length * 0.5) {
    return true;
  }
  
  return false;
}

/**
 * Get all contracts from Airtable (with pagination)
 * Fetches from both Contracts 2025 and Contracts 2024 tables
 */
function getAllContracts_(useCache) {
  // Check cache first if requested
  if (useCache !== false) {
    const cached = getCachedData_('airtable', 'all_contracts');
    if (cached) {
      return cached;
    }
  }
  
  const allRecords = [];
  
  // Fetch from both tables
  const tables = [
    { name: 'Contracts 2025', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2025, viewId: BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2025 },
    { name: 'Contracts 2024', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2024, viewId: BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2024 }
  ];
  
  for (const table of tables) {
    let offset = null;
    
    do {
      let url = `${BS_CFG.AIRTABLE_API_BASE_URL}/${BS_CFG.AIRTABLE_BASE_ID}/${table.tableId}?maxRecords=100`;
      
      if (offset) {
        url += `&offset=${offset}`;
      }
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${BS_CFG.AIRTABLE_API_TOKEN}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
          Logger.log(`Airtable API error for ${table.name}: ${response.getContentText()}`);
          break;
        }
        
        const data = JSON.parse(response.getContentText());
        
        if (data.records) {
          // Add table info to each record for URL construction
          data.records.forEach(record => {
            record._tableId = table.tableId;
            record._viewId = table.viewId;
            record._tableName = table.name;
          });
          allRecords.push(...data.records);
        }
        
        offset = data.offset || null;
        
        // Respect rate limits (5 requests/second)
        if (offset) {
          Utilities.sleep(250);
        }
      } catch (e) {
        Logger.log(`Error fetching Airtable contracts from ${table.name}: ${e.message}`);
        break;
      }
      
    } while (offset);
  }
  
  Logger.log(`📄 Retrieved ${allRecords.length} total contracts from Airtable (2024 + 2025)`);
  
  // Cache the results
  setCachedData_('airtable', 'all_contracts', allRecords);
  
  return allRecords;
}

/**
 * Get contracts for a specific vendor using fuzzy matching
 * Searches both Vendor Name field AND Notes field for matches
 */
function getContractsForVendor_(vendorName) {
  if (!vendorName) {
    Logger.log('⚠️ No vendor name provided');
    return [];
  }
  
  // Get all contracts and filter locally for fuzzy matching
  const allContracts = getAllContracts_();
  
  Logger.log(`📋 Filtering ${allContracts.length} total contracts for vendor: ${vendorName}`);
  
  const matches = allContracts.filter(record => {
    const contractVendor = record.fields[BS_CFG.AIRTABLE_VENDOR_FIELD] || '';
    const contractNotes = record.fields[BS_CFG.AIRTABLE_NOTES_FIELD] || '';
    
    // Handle Submitted By - could be a user/collaborator field (object with email/name) or linked record
    let submittedBy = '';
    const submittedByRaw = record.fields[BS_CFG.AIRTABLE_SUBMITTED_BY_FIELD];
    if (submittedByRaw) {
      if (typeof submittedByRaw === 'string') {
        submittedBy = submittedByRaw;
      } else if (submittedByRaw.name) {
        // Collaborator field format: {id, email, name}
        submittedBy = submittedByRaw.name;
      } else if (submittedByRaw.email) {
        submittedBy = submittedByRaw.email;
      } else if (Array.isArray(submittedByRaw)) {
        // Could be array of collaborators
        submittedBy = submittedByRaw.map(s => s.name || s.email || String(s)).join(', ');
      } else {
        submittedBy = JSON.stringify(submittedByRaw);
        Logger.log(`📋 Unexpected Submitted By format: ${submittedBy}`);
      }
    }
    
    const vertical = record.fields[BS_CFG.AIRTABLE_VERTICAL_FIELD] || '';
    
    // First check if this contract matches the vendor name at all
    let vendorMatches = false;
    
    if (isFuzzyMatch_(vendorName, contractVendor)) {
      vendorMatches = true;
    } else if (contractNotes && contractNotes.toLowerCase().includes(vendorName.toLowerCase())) {
      vendorMatches = true;
    } else {
      // Try simplified name matching
      const vendorSimple = vendorName.toLowerCase()
        .replace(/\s*\([^)]*\)/g, '')
        .replace(/,?\s*(LLC|Inc\.?|Corp\.?|Corporation|Company|Co\.?)$/i, '')
        .trim();
      if (vendorSimple.length >= 4 && contractNotes.toLowerCase().includes(vendorSimple)) {
        vendorMatches = true;
      }
    }
    
    // If vendor doesn't match, skip early
    if (!vendorMatches) {
      return false;
    }
    
    // Log potential match for debugging
    Logger.log(`📋 Potential match: "${contractVendor}" | Submitted By: "${submittedBy}" | Vertical: "${vertical}"`);
    
    // Filter by Submitted By (Andy Worford OR Aden Ritz)
    const submitterMatch = BS_CFG.AIRTABLE_ALLOWED_SUBMITTERS.some(
      allowed => submittedBy.toLowerCase().includes(allowed.toLowerCase())
    );
    if (!submitterMatch) {
      Logger.log(`   ❌ Filtered out: Submitted By "${submittedBy}" not in allowed list`);
      return false;
    }
    
    // Filter by Vertical (Home Services OR Solar)
    const verticalMatch = BS_CFG.AIRTABLE_ALLOWED_VERTICALS.some(
      allowed => vertical.toLowerCase().includes(allowed.toLowerCase())
    );
    if (!verticalMatch) {
      Logger.log(`   ❌ Filtered out: Vertical "${vertical}" not in allowed list`);
      return false;
    }
    
    Logger.log(`   ✅ Match included!`);
    return true;
  });
  
  Logger.log(`📋 Found ${matches.length} contract(s) for vendor: ${vendorName}`);
  return matches;
}

/**
 * Format contracts for display in Battle Station
 */
function formatContractsForDisplay_(contracts) {
  if (!contracts || contracts.length === 0) {
    return [];
  }
  
  const formatted = contracts.map(record => {
    const fields = record.fields;
    // Use table/view IDs stored on the record, or fall back to 2025 defaults
    const tableId = record._tableId || BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2025;
    const viewId = record._viewId || BS_CFG.AIRTABLE_CONTRACTS_VIEW_ID_2025;
    const tableName = record._tableName || 'Contracts 2025';
    
    // Trim notes to remove trailing newlines/whitespace, and cap length
    const rawNotes = fields[BS_CFG.AIRTABLE_NOTES_FIELD] || '';
    let cleanNotes = rawNotes.trim();
    if (cleanNotes.length > BS_CFG.MAX_NOTES_LENGTH) {
      cleanNotes = cleanNotes.substring(0, BS_CFG.MAX_NOTES_LENGTH) + '...';
    }
    
    // Get created date and format it
    // Try both "Created Date" and "Created" as field names
    const createdDateRaw = fields[BS_CFG.AIRTABLE_CREATED_DATE_FIELD] || fields['Created'] || fields['Created Time'] || '';
    let createdDateFormatted = '';
    if (createdDateRaw) {
      try {
        const date = new Date(createdDateRaw);
        if (!isNaN(date.getTime())) {
          const tz = 'America/Los_Angeles';
          createdDateFormatted = Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm');
        }
      } catch (e) {
        createdDateFormatted = createdDateRaw;
      }
    }
    
    // Combine status with created date
    const baseStatus = fields[BS_CFG.AIRTABLE_STATUS_FIELD] || '';
    const statusWithDate = createdDateFormatted ? `${baseStatus} - ${createdDateFormatted}` : baseStatus;
    
    // Log first contract's available fields for debugging
    if (contracts.indexOf(record) === 0) {
      Logger.log(`📋 Contract fields available: ${Object.keys(fields).join(', ')}`);
      Logger.log(`📋 Created date raw value: "${createdDateRaw}" (field: ${BS_CFG.AIRTABLE_CREATED_DATE_FIELD})`);
      Logger.log(`📋 Status with date: "${statusWithDate}"`);
    }
    
    return {
      id: record.id,
      vendorName: fields[BS_CFG.AIRTABLE_VENDOR_FIELD] || '',
      status: statusWithDate,
      contractType: fields[BS_CFG.AIRTABLE_CONTRACT_TYPE_FIELD] || '',
      notes: cleanNotes,
      tableName: tableName,
      createdDate: createdDateRaw,  // Keep raw for sorting
      // Direct link to this record in Airtable using table ID and view ID
      airtableUrl: `https://airtable.com/${BS_CFG.AIRTABLE_BASE_ID}/${tableId}/${viewId}/${record.id}`
    };
  });
  
  // Sort: "Other" type goes last, then by created date DESC
  formatted.sort((a, b) => {
    const typeA = (a.contractType || '').toLowerCase();
    const typeB = (b.contractType || '').toLowerCase();
    
    // "Other" type always goes last
    if (typeA === 'other' && typeB !== 'other') return 1;
    if (typeA !== 'other' && typeB === 'other') return -1;
    
    // Same type priority, sort by created date DESC (newest first)
    const dateA = a.createdDate || '';
    const dateB = b.createdDate || '';
    return dateB.localeCompare(dateA);
  });
  
  return formatted;
}

/**
 * Get contracts for vendor and format for Battle Station display
 * This is the main function to call from Battle Station
 */
function getVendorContracts_(vendorName) {
  try {
    const contracts = getContractsForVendor_(vendorName);
    const formatted = formatContractsForDisplay_(contracts);
    
    return {
      vendorName: vendorName,
      contractCount: formatted.length,
      contracts: formatted,
      hasContracts: formatted.length > 0
    };
  } catch (e) {
    Logger.log(`Error fetching vendor contracts: ${e.message}`);
    return {
      vendorName: vendorName,
      contractCount: 0,
      contracts: [],
      hasContracts: false
    };
  }
}

/**
 * Test Airtable connection - run this to verify credentials
 */
function testAirtableConnection() {
  const tables = [
    { name: 'Contracts 2025', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2025 },
    { name: 'Contracts 2024', tableId: BS_CFG.AIRTABLE_CONTRACTS_TABLE_ID_2024 }
  ];
  
  let allSuccess = true;
  
  for (const table of tables) {
    try {
      const url = `${BS_CFG.AIRTABLE_API_BASE_URL}/${BS_CFG.AIRTABLE_BASE_ID}/${table.tableId}?maxRecords=1`;
      
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${BS_CFG.AIRTABLE_API_TOKEN}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        const data = JSON.parse(responseText);
        Logger.log(`✅ ${table.name} connection successful!`);
        Logger.log(`   Found ${data.records.length} record(s) in test query`);
      } else {
        Logger.log(`❌ ${table.name} connection failed with status ${responseCode}`);
        Logger.log(`   Response: ${responseText}`);
        allSuccess = false;
      }
    } catch (error) {
      Logger.log(`❌ Error connecting to ${table.name}: ${error.message}`);
      allSuccess = false;
    }
  }
  
  return allSuccess;
}

/************************************************************
 * GOOGLE DRIVE FOLDER FUNCTIONS
 * Searches for vendor folder in Google Drive Vendors folder
 ************************************************************/

/**
 * Get files from vendor's Google Drive folder
 * Searches for a folder matching the vendor name (ignoring YYMMDD prefix)
 * Only returns files at the first level (not in subfolders)
 * 
 * @param {string} vendorName - The vendor name to search for
 * @returns {array} Array of file objects with name, type, modified, url
 */
function getGDriveFilesForVendor_(vendorName) {
  if (!vendorName || vendorName.trim() === '') {
    Logger.log('No vendor name provided for Google Drive search');
    return [];
  }
  
  Logger.log(`=== GOOGLE DRIVE SEARCH ===`);
  Logger.log(`Vendor: ${vendorName}`);
  
  try {
    const parentFolderId = BS_CFG.GDRIVE_VENDORS_FOLDER_ID;
    const cleanVendorName = vendorName.trim().toLowerCase();
    
    Logger.log(`Searching for: "${cleanVendorName}"`);
    
    // Remove common suffixes for alternate search
    const nameWithoutSuffix = cleanVendorName
      .replace(/,?\s*(llc|inc\.?|corp\.?|corporation|company|co\.|l\.?l\.?c\.?)$/i, '')
      .trim();
    
    let vendorFolder = null;
    
    // Strategy 1: Search for exact vendor name phrase
    let searchQuery = `title contains '${vendorName.replace(/'/g, "\\'")}' and '${parentFolderId}' in parents`;
    Logger.log(`Search query: ${searchQuery}`);
    
    let folderIterator = DriveApp.searchFolders(searchQuery);
    
    while (folderIterator.hasNext()) {
      const folder = folderIterator.next();
      const folderName = folder.getName();
      const folderNameLower = folderName.toLowerCase();
      Logger.log(`  Found: "${folderName}"`);
      
      // Check if folder contains the exact vendor name (case-insensitive)
      if (folderNameLower.includes(cleanVendorName)) {
        vendorFolder = folder;
        Logger.log(`Exact match! Using vendor folder: ${folderName}`);
        break;
      }
      
      // Also accept if it contains the name without suffix
      if (nameWithoutSuffix !== cleanVendorName && folderNameLower.includes(nameWithoutSuffix)) {
        vendorFolder = folder;
        Logger.log(`Match without suffix! Using vendor folder: ${folderName}`);
        break;
      }
    }
    
    // Strategy 2: If not found, try searching without suffix
    if (!vendorFolder && nameWithoutSuffix !== cleanVendorName) {
      searchQuery = `title contains '${nameWithoutSuffix.replace(/'/g, "\\'")}' and '${parentFolderId}' in parents`;
      Logger.log(`Trying without suffix: ${searchQuery}`);
      
      folderIterator = DriveApp.searchFolders(searchQuery);
      
      while (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        const folderName = folder.getName();
        const folderNameLower = folderName.toLowerCase();
        Logger.log(`  Found: "${folderName}"`);
        
        if (folderNameLower.includes(nameWithoutSuffix)) {
          vendorFolder = folder;
          Logger.log(`Match! Using vendor folder: ${folderName}`);
          break;
        }
      }
    }
    
    // Strategy 3: Fallback - iterate through parent folder directly (for newly created folders)
    if (!vendorFolder) {
      Logger.log('Search API found nothing, trying direct folder iteration...');
      const parentFolder = DriveApp.getFolderById(parentFolderId);
      folderIterator = parentFolder.getFolders();
      
      while (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        const folderName = folder.getName();
        const folderNameLower = folderName.toLowerCase();
        
        // Check for exact phrase match
        if (folderNameLower.includes(cleanVendorName) || 
            (nameWithoutSuffix !== cleanVendorName && folderNameLower.includes(nameWithoutSuffix))) {
          vendorFolder = folder;
          Logger.log(`Found via direct iteration: ${folderName}`);
          break;
        }
      }
    }
    
    if (!vendorFolder) {
      Logger.log('No matching vendor folder found in Google Drive');
      return { files: [], folderFound: false, folderUrl: null };
    }
    
    // Get files at the first level only (no subfolders)
    const files = [];
    const folderUrl = vendorFolder.getUrl();
    const fileIterator = vendorFolder.getFiles();
    
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      
      try {
        const mimeType = file.getMimeType() || '';
        
        // Skip shortcuts (they appear as application/vnd.google-apps.shortcut)
        if (mimeType.includes('shortcut')) {
          Logger.log(`Skipping shortcut: ${file.getName()}`);
          continue;
        }
        
        // Determine file type from mime type
        let fileType = 'File';
        if (mimeType.includes('pdf')) fileType = 'PDF';
        else if (mimeType.includes('spreadsheet') || mimeType.includes('excel')) fileType = 'Sheet';
        else if (mimeType.includes('document') || mimeType.includes('word')) fileType = 'Doc';
        else if (mimeType.includes('presentation') || mimeType.includes('powerpoint')) fileType = 'Slides';
        else if (mimeType.includes('image')) fileType = 'Image';
        else if (mimeType.includes('folder')) continue; // Skip folders
        
        const modDate = file.getLastUpdated();
        const tz = Session.getScriptTimeZone();
        
        files.push({
          name: file.getName(),
          type: fileType,
          modified: modDate ? Utilities.formatDate(modDate, tz, 'yyyy-MM-dd') : '',
          url: file.getUrl(),
          folderUrl: folderUrl
        });
      } catch (e) {
        // Log error but continue with other files
        Logger.log(`Error processing file ${file.getName()}: ${e.message}`);
        files.push({
          name: file.getName(),
          type: '',
          modified: '',
          url: file.getUrl() || '',
          folderUrl: folderUrl
        });
      }
    }
    
    // Sort by modified date (newest first)
    files.sort((a, b) => b.modified.localeCompare(a.modified));
    
    Logger.log(`Found ${files.length} files in vendor folder`);
    return { files: files, folderFound: true, folderUrl: folderUrl };
    
  } catch (e) {
    Logger.log(`Error searching Google Drive: ${e.message}`);
    return { files: [], folderFound: false, folderUrl: null };
  }
}

/**
 * Find a monday.com item ID by searching for a vendor name
 */
function findMondayItemIdByVendor_(vendor, boardId, apiToken) {
  Logger.log(`=== SEARCHING FOR VENDOR ===`);
  Logger.log(`Search term: "${vendor}"`);
  Logger.log(`Board ID: ${boardId}`);
  
  let query = `
    query {
      boards (ids: [${boardId}]) {
        items_page (limit: 100, query_params: {rules: [{column_id: "name", compare_value: ["${vendor.replace(/"/g, '\\"')}"]}]}) {
          items { id name }
        }
      }
    }
  `;
  
  let itemId = tryFindItem_(query, apiToken, 'Exact match');
  if (itemId) return itemId;
  
  const withoutParens = vendor.replace(/\s*\([^)]*\)/g, '').trim();
  if (withoutParens !== vendor) {
    Logger.log(`Trying without parentheses: "${withoutParens}"`);
    query = `
      query {
        boards (ids: [${boardId}]) {
          items_page (limit: 100, query_params: {rules: [{column_id: "name", compare_value: ["${withoutParens.replace(/"/g, '\\"')}"]}]}) {
            items { id name }
          }
        }
      }
    `;
    
    itemId = tryFindItem_(query, apiToken, 'Without parentheses');
    if (itemId) return itemId;
  }
  
  Logger.log(`Trying contains search...`);
  query = `
    query {
      boards (ids: [${boardId}]) {
        items_page (limit: 500) {
          items { id name }
        }
      }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.data?.boards?.[0]?.items_page?.items) {
      const items = result.data.boards[0].items_page.items;
      
      for (const item of items) {
        const itemName = String(item.name || '').toLowerCase();
        const searchTerm = vendor.toLowerCase();
        const searchTermNoParens = withoutParens.toLowerCase();
        
        if (itemName.includes(searchTerm) || searchTerm.includes(itemName) ||
            itemName.includes(searchTermNoParens) || searchTermNoParens.includes(itemName)) {
          Logger.log(`✓ FOUND MATCH: "${item.name}" (ID: ${item.id})`);
          return item.id;
        }
      }
    }
    
    return null;
  } catch (e) {
    Logger.log(`Error in contains search: ${e}`);
    return null;
  }
}

/**
 * Helper to try a specific query
 */
function tryFindItem_(query, apiToken, attemptName) {
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.data?.boards?.[0]?.items_page?.items?.length > 0) {
      const item = result.data.boards[0].items_page.items[0];
      Logger.log(`✓ ${attemptName} SUCCESS: Found "${item.name}" (ID: ${item.id})`);
      return item.id;
    } else {
      Logger.log(`✗ ${attemptName} failed: No items found`);
    }
    
    return null;
  } catch (e) {
    Logger.log(`✗ ${attemptName} error: ${e}`);
    return null;
  }
}

/**
 * Navigation: Go to next vendor
 */
function battleStationNext() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    loadVendorData(1, { loadMode: 'fast' });
    return;
  }

  const totalVendors = listSh.getLastRow() - 1;

  if (currentIndex >= totalVendors) {
    ss.toast('Already at the last vendor!', '⚠️ End of List', 3);
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue();
  listSh.getRange(listRow, BS_CFG.L_PROCESSED + 1).setValue(true);
  
  ss.toast(`Marked "${vendor}" as reviewed`, '▶️ Next', 2);
  loadVendorData(currentIndex + 1, { loadMode: 'fast' });
}

/**
 * Navigation: Go to previous vendor
 */
function battleStationPrevious() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  if (!currentIndex) {
    loadVendorData(1, { loadMode: 'fast' });
    return;
  }

  if (currentIndex <= 1) {
    SpreadsheetApp.getActive().toast('Already at the first vendor!', '⚠️ Start of List', 3);
    return;
  }

  loadVendorData(currentIndex - 1, { loadMode: 'fast' });
}

/**
 * Refresh current vendor data
 */
function battleStationRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  
  if (!bsSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  loadVendorData(currentIndex || 1);
}

/**
 * Quick Refresh - Email only, uses cached data for Airtable/Box
 * Much faster than full refresh when you've only changed emails
 */
/**
 * Quick Refresh - Email only, just refreshes the EMAILS section without redrawing everything
 * Much faster than full refresh
 */
function battleStationQuickRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }
  
  ss.toast('Refreshing emails only...', '⚡ Quick Refresh', 2);
  
  // Get current vendor info
  const currentIndex = getCurrentVendorIndex_();
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();
  
  // Find the EMAILS section in the sheet
  const dataRange = bsSh.getDataRange();
  const values = dataRange.getValues();
  
  let emailsHeaderRow = -1;
  let emailsEndRow = -1;
  
  // Find emails section start
  for (let row = 0; row < values.length; row++) {
    const cellValue = String(values[row][0] || '');
    if (cellValue.includes('📧 EMAILS')) {
      emailsHeaderRow = row + 1; // 1-indexed
    } else if (emailsHeaderRow > 0 && (cellValue.includes('📋 MONDAY.COM TASKS') || cellValue.includes('📋 MONDAY'))) {
      // Found the next section (TASKS comes after EMAILS)
      emailsEndRow = row; // 0-indexed, so this is the row before TASKS header (1-indexed)
      break;
    }
  }
  
  if (emailsHeaderRow < 0) {
    ss.toast('Could not find EMAILS section', '❌ Error', 3);
    return;
  }
  
  // If we didn't find the end, assume it goes to current row count
  if (emailsEndRow < 0) {
    emailsEndRow = values.length;
  }
  
  // Fetch fresh emails
  const emails = getEmailsForVendor_(vendor, listRow);
  
  // Calculate how many rows we need vs how many we have
  const emailRowsNeeded = emails.length > 0 ? Math.min(emails.length, 20) + 2 : 2; // +2 for header row and column headers (or "no emails" row)
  if (emails.length > 20) emailRowsNeeded + 1; // +1 for "and X more" row
  
  const availableRows = emailsEndRow - emailsHeaderRow;
  
  // Clear existing email content (keep header row formatting)
  const clearStartRow = emailsHeaderRow + 1; // Row after "📧 EMAILS" header
  const clearEndRow = emailsEndRow;
  
  if (clearEndRow > clearStartRow) {
    bsSh.getRange(clearStartRow, 1, clearEndRow - clearStartRow, 4)
      .clearContent()
      .clearFormat()
      .setBackground('#ffffff')
      .setFontWeight('normal')
      .setFontStyle('normal');
    
    // Unmerge any merged cells in the range
    for (let r = clearStartRow; r < clearEndRow; r++) {
      try {
        bsSh.getRange(r, 1, 1, 4).breakApart();
      } catch (e) {
        // Ignore if not merged
      }
    }
  }
  
  // Update the header with new count
  bsSh.getRange(emailsHeaderRow, 1, 1, 4).breakApart();
  bsSh.getRange(emailsHeaderRow, 1, 1, 4).merge()
    .setValue(`📧 EMAILS (${emails.length})  |  🔵 Snoozed  🔴 Overdue  🟠 Phonexa  🟢 Accounting  🟡 Waiting`)
    .setBackground('#f8f9fa')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#1a73e8')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  
  let currentRow = emailsHeaderRow + 1;
  
  if (emails.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No emails found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    currentRow++;
  } else {
    // Email headers
    bsSh.getRange(currentRow, 1).setValue('Subject').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Count').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Labels').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;
    
    for (const email of emails.slice(0, 20)) {
      bsSh.getRange(currentRow, 1).setValue(email.subject);
      const emailDateCell2 = bsSh.getRange(currentRow, 2);
      emailDateCell2.setNumberFormat('@'); // Set format BEFORE value to prevent auto-parsing
      emailDateCell2.setValue(email.date);
      bsSh.getRange(currentRow, 3).setValue(email.count).setNumberFormat('0');
      bsSh.getRange(currentRow, 4).setValue(email.labels);

      if (email.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${email.link}", "${email.subject.replace(/"/g, '""')}")`);
      }
      
      // LABEL-AGNOSTIC: Use configured labels for color coding
      const emailStatus2 = getEmailStatusCategory_(email);
      const cfg2 = getLabelConfig_();
      const hasPriority2 = cfg2.priority_label && email.labels && email.labels.includes(cfg2.priority_label);

      switch (emailStatus2) {
        case 'snoozed':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_SNOOZED);
          break;
        case 'overdue':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_OVERDUE);
          bsSh.getRange(currentRow, 1, 1, 4).setFontWeight('bold');
          break;
        case 'waiting_external':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_PHONEXA);
          break;
        case 'accounting':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground('#d9ead3');
          break;
        case 'waiting_customer':
        case 'priority_waiting':
          bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_WAITING);
          break;
        default:
          bsSh.getRange(currentRow, 1, 1, 4).setBackground('#ffffff');
      }

      if (cfg2.priority_label && !hasPriority2) {
        bsSh.getRange(currentRow, 1, 1, 4).setFontColor('#999999');
      }

      currentRow++;
    }

    if (emails.length > 20) {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`... and ${emails.length - 20} more emails (showing first 20)`)
        .setFontStyle('italic')
        .setHorizontalAlignment('center');
      currentRow++;
    }
  }

  // Update checksum for emails
  const newEmailChecksum = generateEmailChecksum_(emails);
  updateEmailChecksum_(vendor, newEmailChecksum);
  
  ss.toast('Emails refreshed!', '⚡ Done', 2);
}

/**
 * Update just the email checksum for a vendor
 */
function updateEmailChecksum_(vendor, newEmailChecksum) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === vendor) {
      sh.getRange(i + 1, 3).setValue(newEmailChecksum); // Column C = EmailChecksum
      sh.getRange(i + 1, 5).setValue(new Date()); // Column E = Last Viewed
      Logger.log(`Updated email checksum for ${vendor}: ${newEmailChecksum}`);
      return;
    }
  }
  
  // Vendor not found - this shouldn't happen but handle it
  Logger.log(`Warning: Could not find ${vendor} in checksums sheet to update email checksum`);
}

/**
 * Hard Refresh - Clear cache and reload everything fresh
 */
function battleStationHardRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  
  if (!bsSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }
  
  // Clear the cache
  clearBSCache_();
  ss.toast('Cache cleared, refreshing...', '🔄 Hard Refresh', 2);
  
  const currentIndex = getCurrentVendorIndex_();
  loadVendorData(currentIndex || 1, { useCache: false });
}

/**
 * Get or create the cache sheet
 */
function getBSCacheSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(BS_CFG.CACHE_SHEET);
  
  if (!sh) {
    sh = ss.insertSheet(BS_CFG.CACHE_SHEET);
    // Set up headers: Type, Key, Data (JSON), LastUpdated
    sh.getRange(1, 1, 1, 4).setValues([['Type', 'Key', 'Data', 'LastUpdated']]);
    sh.hideSheet();
  }
  
  return sh;
}

/**
 * Clear the BS cache
 */
function clearBSCache_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BS_CFG.CACHE_SHEET);
  
  if (sh) {
    sh.clear();
    sh.getRange(1, 1, 1, 4).setValues([['Type', 'Key', 'Data', 'LastUpdated']]);
  }
  
  Logger.log('BS Cache cleared');
}

/**
 * Get cached data if fresh enough
 * Returns null if cache is stale or doesn't exist
 */
function getCachedData_(type, key) {
  const sh = getBSCacheSheet_();
  const data = sh.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === type && data[i][1] === key) {
      const lastUpdated = new Date(data[i][3]);
      const ageHours = (new Date() - lastUpdated) / (1000 * 60 * 60);
      
      if (ageHours < BS_CFG.CACHE_MAX_AGE_HOURS) {
        try {
          Logger.log(`Cache HIT: ${type}/${key} (${Math.round(ageHours * 10) / 10}h old)`);
          return JSON.parse(data[i][2]);
        } catch (e) {
          Logger.log(`Cache parse error: ${e.message}`);
          return null;
        }
      } else {
        Logger.log(`Cache STALE: ${type}/${key} (${Math.round(ageHours * 10) / 10}h old)`);
        return null;
      }
    }
  }
  
  Logger.log(`Cache MISS: ${type}/${key}`);
  return null;
}

/**
 * Store data in cache
 */
function setCachedData_(type, key, data) {
  const sh = getBSCacheSheet_();
  const existingData = sh.getDataRange().getValues();
  
  // Look for existing row to update
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0] === type && existingData[i][1] === key) {
      sh.getRange(i + 1, 3, 1, 2).setValues([[JSON.stringify(data), new Date()]]);
      Logger.log(`Cache UPDATE: ${type}/${key}`);
      return;
    }
  }
  
  // Add new row
  sh.appendRow([type, key, JSON.stringify(data), new Date()]);
  Logger.log(`Cache SET: ${type}/${key}`);
}

/**
 * Update monday.com notes for current vendor - NO DIALOGS
 */
function battleStationUpdateMondayNotes() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    ss.toast('Required sheets not found', '❌ Error', 3);
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex || isNaN(currentIndex)) {
    ss.toast('Could not determine current vendor', '❌ Error', 3);
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  
  if (!vendor) {
    ss.toast('Could not determine vendor name', '❌ Error', 3);
    return;
  }
  
  // Find the notes row - look for 📝 NOTES header
  let notesRow = -1;
  for (let i = 5; i < 50; i++) {
    const label = String(bsSh.getRange(i, 1).getValue() || '');
    if (label.indexOf('📝 NOTES') !== -1) {
      notesRow = i + 1;  // Notes content is in the row after the header
      break;
    }
  }
  
  if (notesRow === -1) {
    ss.toast('Could not find notes field', '❌ Error', 3);
    return;
  }
  
  const notes = String(bsSh.getRange(notesRow, 1).getValue() || '').trim();
  
  if (!notes || notes === '(no notes)') {
    ss.toast('No notes to sync - edit the notes field first', '⚠️ Empty', 3);
    return;
  }
  
  ss.toast(`Updating ${vendor}...`, '⚡ Syncing', 3);
  
  try {
    listSh.getRange(listRow, BS_CFG.L_NOTES + 1).setValue(notes);
    
    const result = updateMondayComNotesForVendor_(vendor, notes, listRow);
    
    if (result.success) {
      ss.toast(`Notes updated for ${vendor}`, '✅ Success', 3);
      battleStationRefresh();
    } else {
      ss.toast(`Failed: ${result.error}`, '❌ Error', 5);
    }
  } catch (e) {
    ss.toast(`Error: ${e.message}`, '❌ Error', 5);
  }
}

/**
 * Helper function to update monday.com notes via API
 */
function updateMondayComNotesForVendor_(vendor, notes, listRow) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listRow) {
    const currentIndex = getCurrentVendorIndex_();
    if (!currentIndex) return { success: false, error: 'Could not determine vendor index' };
    listRow = currentIndex + 1;
  }
  
  const source = String(listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '');
  
  let boardId, notesColumnId;
  if (source.toLowerCase().includes('buyer')) {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
  } else if (source.toLowerCase().includes('affiliate')) {
    boardId = BS_CFG.AFFILIATES_BOARD_ID;
    notesColumnId = BS_CFG.AFFILIATES_NOTES_COLUMN;
  } else {
    boardId = BS_CFG.BUYERS_BOARD_ID;
    notesColumnId = BS_CFG.BUYERS_NOTES_COLUMN;
  }
  
  const itemId = findMondayItemIdByVendor_(vendor, boardId, apiToken);
  
  if (!itemId) {
    return { success: false, error: `Could not find monday.com item for vendor: ${vendor}` };
  }
  
  const valueJson = JSON.stringify(notes);
  const escapedValue = valueJson.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  
  const mutation = `
    mutation {
      change_column_value (
        board_id: ${boardId},
        item_id: ${itemId},
        column_id: "${notesColumnId}",
        value: "${escapedValue}"
      ) { id }
    }
  `;
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: mutation }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.errors && result.errors.length > 0) {
      return { success: false, error: result.errors[0].message };
    }
    
    if (result.data?.change_column_value?.id) {
      return { success: true, itemId: result.data.change_column_value.id };
    }
    
    return { success: false, error: 'Unexpected API response' };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Mark current vendor as reviewed
 */
function battleStationMarkReviewed() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue();
  
  listSh.getRange(listRow, BS_CFG.L_PROCESSED + 1).setValue(true);
  
  battleStationRefresh();
  ss.toast('Marked as reviewed!', '✅ ' + vendor, 3);
}

/**
 * Open Gmail search for current vendor
 */
function battleStationOpenGmail() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const memberKey = getCurrentTeamMemberKey_();
  const gmailColIdx = memberKey === 'm2' ? BS_CFG.L_M2_GMAIL : BS_CFG.L_M1_GMAIL;
  const gmailLink = listSh.getRange(listRow, gmailColIdx + 1).getValue();

  if (!gmailLink || gmailLink.toString().indexOf('#search') === -1) {
    SpreadsheetApp.getUi().alert('No valid Gmail search link found.');
    return;
  }
  
  const html = `<html><body><script>window.open('${gmailLink}', '_blank');google.script.host.close();</script></body></html>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(200).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Opening Gmail...');
}

/**
 * Create Gmail draft to vendor contacts
 */
function battleStationEmailContacts() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex || isNaN(currentIndex)) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  
  ss.toast('Finding contacts...', '📧 Creating Draft', 2);
  
  const contactData = getVendorContacts_(vendor, listRow);
  
  if (contactData.contacts.length === 0) {
    SpreadsheetApp.getUi().alert('No contacts found for this vendor.');
    return;
  }
  
  const recipients = contactData.contacts.map(c => c.email).filter(e => e).join(', ');
  
  if (!recipients) {
    SpreadsheetApp.getUi().alert('No email addresses found for contacts.');
    return;
  }
  
  try {
    const subject = `Re: ${vendor}`;
    const body = `Hi,\n\n\n\nBest regards,\nAndy Worford\nProfitise`;
    
    GmailApp.createDraft(recipients, subject, body);
    
    ss.toast('Draft created!', '✅ Success', 3);
    SpreadsheetApp.getUi().alert(`✓ Draft created!\n\nTo: ${recipients}\nSubject: ${subject}\n\nCheck your Gmail drafts.`);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error creating draft: ${e.message}`);
  }
}

/**
 * Analyze unsnoozed emails with Claude AI - Individual breakdown with inline links
 */
function battleStationAnalyzeEmails() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  const claudeApiKey = getClaudeApiKey_();

  if (!claudeApiKey) {
    SpreadsheetApp.getUi().alert('No Claude API key configured.\n\nUse menu: ⚡ Battle Station → ⚙️ Set Claude API Key');
    return;
  }
  
  if (!listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  const currentIndex = getCurrentVendorIndex_();
  
  if (!currentIndex || isNaN(currentIndex)) {
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }
  
  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  
  ss.toast('Fetching emails for analysis...', '🤖 Claude Analysis', 3);

  const emails = getEmailsForVendor_(vendor, listRow);
  const unsnoozedEmails = emails.filter(e => !e.isSnoozed);

  // Also get single-person emails from the log (they may not appear in label searches)
  let singlePersonEmails = [];
  try {
    singlePersonEmails = getSinglePersonEmails_(vendor);
    Logger.log(`Found ${singlePersonEmails.length} single-person emails in log for ${vendor}`);
  } catch (e) {
    Logger.log(`Could not fetch single-person emails: ${e.message}`);
  }

  if (unsnoozedEmails.length === 0 && singlePersonEmails.length === 0) {
    SpreadsheetApp.getUi().alert('No unsnoozed emails found for this vendor.');
    return;
  }

  const totalCount = unsnoozedEmails.length + singlePersonEmails.length;
  ss.toast(`Analyzing ${totalCount} emails with Claude (${singlePersonEmails.length} from log)...`, '🤖 Processing', 10);
  
  // Gather email content with metadata
  const emailData = [];
  
  for (const email of unsnoozedEmails.slice(0, 10)) {
    try {
      const thread = GmailApp.getThreadById(email.threadId);
      if (thread) {
        const messages = thread.getMessages();
        let fullContent = '';
        
        for (const msg of messages.slice(-5)) {
          fullContent += `\n--- From: ${msg.getFrom()} | Date: ${msg.getDate()} ---\n`;
          fullContent += msg.getPlainBody().substring(0, 1500);
        }
        
        emailData.push({
          subject: email.subject,
          date: email.date,
          labels: email.labels,
          link: email.link,
          threadId: email.threadId,
          messageCount: email.count,
          content: fullContent
        });
      }
    } catch (e) {
      emailData.push({
        subject: email.subject,
        date: email.date,
        labels: email.labels,
        link: email.link,
        threadId: email.threadId,
        messageCount: email.count,
        content: email.snippet || '(could not fetch content)'
      });
    }
  }
  
  if (emailData.length === 0) {
    SpreadsheetApp.getUi().alert('Could not fetch email content for analysis.');
    return;
  }
  
  // Build prompt - use numbered emails that we can match to links
  const emailsText = emailData.map((e, i) => 
    `\n\n=== EMAIL_${i + 1} ===
Subject: ${e.subject}
Date: ${e.date}
Labels: ${e.labels}
Messages in thread: ${e.messageCount}

Content:
${e.content}`
  ).join('\n');
  
  const prompt = `You are analyzing email communications for a vendor relationship manager at a lead generation company called Profitise.

The vendor being analyzed is: ${vendor}

Here are the recent unsnoozed emails (${emailData.length} threads):
${emailsText}

Please provide your analysis in this EXACT format (keep EMAIL_1, EMAIL_2 etc as markers - they will be replaced with links):

## OVERALL SUMMARY
[2-3 sentences about the overall relationship status]

## EMAIL BREAKDOWN

EMAIL_1
**Status**: [Active / Waiting on them / Waiting on us / FYI / Urgent]
**Summary**: [1-2 sentence summary]
**Action**: [What to do, or "None"]

EMAIL_2
**Status**: [Status]
**Summary**: [Summary]
**Action**: [Action]

[Continue for each email...]

## PRIORITY ACTIONS
[Numbered list of most important actions]

Be concise. Use the exact EMAIL_1, EMAIL_2 markers so they can be linked.`;

  try {
    const response = callClaudeAPI_(prompt, claudeApiKey);
    
    if (response.error) {
      SpreadsheetApp.getUi().alert(`Claude API Error: ${response.error}`);
      return;
    }
    
    // Format content and replace EMAIL_X markers with clickable links
    let formattedContent = response.content
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    
    // Replace EMAIL_X markers with linked subject lines
    emailData.forEach((e, i) => {
      const marker = `EMAIL_${i + 1}`;
      const shortSubject = e.subject.length > 60 ? e.subject.substring(0, 60) + '...' : e.subject;
      const linkedSubject = `<a href="${e.link}" target="_blank" class="email-link">📧 ${shortSubject}</a> <span class="date">(${e.date})</span>`;
      formattedContent = formattedContent.replace(new RegExp(marker, 'g'), linkedSubject);
    });
    
    // Apply formatting
    formattedContent = formattedContent
      .replace(/\n/g, '<br>')
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/## (.*?)<br>/g, '<h3>$1</h3>')
      .replace(/### (.*?)<br>/g, '<h4>$1</h4>');
    
    const htmlContent = `
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; line-height: 1.6; font-size: 13px; }
        h2 { color: #4a86e8; margin-top: 0; margin-bottom: 10px; }
        h3 { color: #4a86e8; margin-top: 20px; margin-bottom: 10px; border-bottom: 2px solid #4a86e8; padding-bottom: 5px; }
        h4 { color: #666; margin-top: 12px; margin-bottom: 5px; }
        .email-link { 
          color: #1a73e8; 
          text-decoration: none; 
          font-weight: bold;
          font-size: 14px;
          display: inline-block;
          margin-top: 10px;
          padding: 5px 10px;
          background: #e8f0fe;
          border-radius: 4px;
        }
        .email-link:hover { background: #d0e1fd; text-decoration: none; }
        .date { color: #888; font-size: 11px; }
        .content { background: #fafafa; padding: 15px; border-radius: 5px; }
        strong { color: #333; }
      </style>
      <h2>🤖 Claude Analysis: ${vendor}</h2>
      <p><em>Analyzed ${emailData.length} unsnoozed email threads</em></p>
      <div class="content">${formattedContent}</div>
    `;
    
    const html = HtmlService.createHtmlOutput(htmlContent).setWidth(750).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, `Claude Analysis: ${vendor}`);
    
    ss.toast('Analysis complete!', '✅ Done', 3);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

/**
 * Call Claude API
 */
function callClaudeAPI_(prompt, apiKey, options) {
  options = options || {};
  const url = 'https://api.anthropic.com/v1/messages';

  const payload = {
    model: options.model || 'claude-sonnet-4-20250514',
    max_tokens: options.maxTokens || 2000,
    messages: [{ role: 'user', content: prompt }]
  };
  if (options.system) {
    payload.system = options.system;
  }
  
  const fetchOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, fetchOptions);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { error: `API returned ${responseCode}: ${responseText}` };
    }
    
    const result = JSON.parse(responseText);
    
    if (result.content && result.content.length > 0) {
      return { content: result.content[0].text };
    }
    
    return { error: 'No content in response' };
    
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Go to a specific vendor by index or name
 */
function battleStationGoTo() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) return;
  
  const totalVendors = listSh.getLastRow() - 1;
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Go to Vendor',
    `Enter vendor index (1-${totalVendors}) or vendor name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    const index = parseInt(input);
    
    // If it's a number, go directly to that index
    if (!isNaN(index) && index >= 1 && index <= totalVendors) {
      loadVendorData(index, { loadMode: 'fast' });
      return;
    }

    // Otherwise, search by name
    const searchTerm = input.toLowerCase();
    const data = listSh.getDataRange().getValues();
    const matches = [];

    for (let i = 1; i < data.length; i++) {
      const vendorName = String(data[i][0] || '').toLowerCase();
      if (vendorName.includes(searchTerm)) {
        matches.push({ index: i, name: data[i][0] });
      }
    }

    if (matches.length === 0) {
      ui.alert(`No vendors found matching "${input}"`);
    } else if (matches.length === 1) {
      loadVendorData(matches[0].index, { loadMode: 'fast' });
    } else {
      // Multiple matches - show list
      let matchList = matches.slice(0, 15).map(m => `${m.index}: ${m.name}`).join('\n');
      if (matches.length > 15) {
        matchList += `\n... and ${matches.length - 15} more`;
      }
      
      const pickResponse = ui.prompt(
        `Found ${matches.length} matches`,
        `${matchList}\n\nEnter the index number to go to:`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (pickResponse.getSelectedButton() === ui.Button.OK) {
        const pickIndex = parseInt(pickResponse.getResponseText());
        if (!isNaN(pickIndex) && pickIndex >= 1 && pickIndex <= totalVendors) {
          loadVendorData(pickIndex, { loadMode: 'fast' });
        }
      }
    }
  }
}

/**
 * Generate a checksum for vendor data
 * Based on: emails, tasks, notes, status, states, contracts, helpful links, meetings, box documents, gdrive files
 */
function generateVendorChecksum_(vendor, emails, tasks, notes, status, states, contracts, helpfulLinks, meetings, boxDocs, gDriveFiles) {
  const data = {
    vendor: vendor,
    status: status || '',
    notes: notes || '',
    states: states || '',
    emails: (emails || []).map(e => ({
      subject: e.subject,
      date: e.date,
      labels: e.labels
    })),
    // Include overdue count so full checksum changes when emails become overdue
    overdueEmailCount: (emails || []).filter(e => isEmailOverdue_(e)).length,
    tasks: (tasks || []).map(t => ({
      name: t.subject,
      status: t.status,
      created: t.created
    })),
    contracts: (contracts || []).map(c => ({
      vendor: c.vendorName,
      status: c.status,
      type: c.contractType,
      notes: c.notes
    })),
    helpfulLinks: (helpfulLinks || []).map(l => ({
      name: l.name,
      url: l.url,
      notes: l.notes
    })),
    meetings: (meetings || []).map(m => ({
      title: m.title,
      date: m.date,
      time: m.time,
      isPast: m.isPast  // Track if meeting moved from future to past
    })),
    boxDocs: (boxDocs || []).map(d => ({
      name: d.name,
      modified: d.modifiedAt
    })),
    gDriveFiles: (gDriveFiles || []).map(f => ({
      name: f.name,
      modified: f.modified
    }))
  };
  
  const jsonStr = JSON.stringify(data);
  return hashString_(jsonStr);
}

/**
 * Simple hash function used by all checksum generators
 */
function hashString_(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return hash.toString(16);
}

/**
 * Generate sub-checksums for each module
 * Returns an object with checksums for each data section
 */
function generateModuleChecksums_(vendor, emails, tasks, notes, status, states, contracts, helpfulLinks, meetings, boxDocs, gDriveFiles, contacts) {
  return {
    emails: generateEmailChecksum_(emails),  // Use the overdue-aware email checksum
    tasks: generateTasksChecksum_(tasks),
    notes: hashString_(JSON.stringify(notes || '')),
    status: hashString_(JSON.stringify(status || '')),
    states: hashString_(JSON.stringify(states || '')),
    contracts: hashString_(JSON.stringify((contracts || []).map(c => ({ vendor: c.vendorName, status: c.status, type: c.contractType })))),
    helpfulLinks: hashString_(JSON.stringify((helpfulLinks || []).map(l => ({ name: l.name, url: l.url })))),
    meetings: generateMeetingsChecksum_(meetings),
    boxDocs: hashString_(JSON.stringify((boxDocs || []).map(d => ({ name: d.name, modified: d.modifiedAt })))),
    gDriveFiles: hashString_(JSON.stringify((gDriveFiles || []).map(f => ({ name: f.name, modified: f.modified })))),
    contacts: generateContactsChecksum_(contacts)
  };
}

/**
 * Generate a sub-checksum for just emails (most volatile data)
 */
function generateEmailChecksum_(emails) {
  // Base checksum - same as before so existing checksums stay valid
  const data = (emails || []).map(e => ({
    subject: e.subject,
    date: e.date,
    labels: e.labels
  }));
  const baseChecksum = hashString_(JSON.stringify(data));
  
  // Count overdue emails - append to checksum so it changes when overdue status changes
  const overdueCount = (emails || []).filter(e => isEmailOverdue_(e)).length;
  
  // Combine: baseChecksum + overdueCount
  // This way, if no emails are overdue (overdueCount=0), checksum matches old format
  // But if any become overdue, the checksum changes
  return overdueCount > 0 ? `${baseChecksum}_OD${overdueCount}` : baseChecksum;
}

/**
 * Check if an email is overdue.
 * LABEL-AGNOSTIC: Uses configured label names from Settings.
 * Falls back to the configurable version in labelConfig.gs.
 */
function isEmailOverdue_(email) {
  return isEmailOverdueConfigurable_(email);
}

/**
 * Parse email date string to Date object
 * Handles formats like "Dec 5, 2024 3:45 PM" or "Dec 5, 3:45 PM"
 */
function parseEmailDate_(dateStr) {
  if (!dateStr) return null;
  
  try {
    // Try direct parse first
    let date = new Date(dateStr);
    if (!isNaN(date.getTime())) return date;
    
    // If no year, assume current year
    const currentYear = new Date().getFullYear();
    date = new Date(dateStr + ' ' + currentYear);
    if (!isNaN(date.getTime())) return date;
    
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Calculate business hours elapsed since a given date
 * Business hours: Monday-Friday, 9 AM - 5 PM Pacific
 */
function getBusinessHoursElapsed_(startDate) {
  if (!startDate) return 0;
  
  const now = new Date();
  const start = new Date(startDate);
  
  if (start >= now) return 0;
  
  let businessHours = 0;
  let current = new Date(start);
  
  // Iterate hour by hour (simplified approach)
  while (current < now) {
    const day = current.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
    const hour = current.getHours();
    
    // Business hours: Monday-Friday (1-5), 9 AM - 5 PM (9-16)
    if (day >= 1 && day <= 5 && hour >= 9 && hour < 17) {
      businessHours++;
    }
    
    current.setHours(current.getHours() + 1);
    
    // Safety limit - don't count more than 1000 hours
    if (businessHours > 1000) break;
  }
  
  return businessHours;
}

/**
 * Get or create the Checksums sheet
 */
function getChecksumsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(BS_CFG.CHECKSUMS_SHEET);

  if (!sh) {
    sh = ss.insertSheet(BS_CFG.CHECKSUMS_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['Vendor', 'Checksum', 'EmailChecksum', 'ModuleChecksums', 'Last Viewed', 'Flagged']]);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    sh.hideSheet();
  } else {
    // Check if we need to add columns (migration)
    const headers = sh.getRange(1, 1, 1, 6).getValues()[0];
    if (headers[3] !== 'ModuleChecksums') {
      const numCols = sh.getLastColumn();
      if (numCols < 5) {
        sh.insertColumnAfter(3);
        sh.getRange(1, 4).setValue('ModuleChecksums').setFontWeight('bold');
        sh.getRange(1, 5).setValue('Last Viewed').setFontWeight('bold');
      }
    }
    // Add Flagged column if missing
    if (headers[5] !== 'Flagged') {
      const numCols = sh.getLastColumn();
      if (numCols < 6) {
        sh.getRange(1, 6).setValue('Flagged').setFontWeight('bold');
      }
    }
  }

  return sh;
}

/**
 * Check if a vendor is flagged for review
 */
function isVendorFlagged_(vendor) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      return data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 'true';
    }
  }
  return false;
}

/**
 * Set or clear the flag for a vendor
 */
function setVendorFlag_(vendor, flagged) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      sh.getRange(i + 1, 6).setValue(flagged);
      return true;
    }
  }

  // Vendor not in checksums yet - add a row
  if (flagged) {
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1).setValue(vendor);
    sh.getRange(lastRow + 1, 6).setValue(true);
  }
  return true;
}

/**
 * Toggle flag for the currently displayed vendor
 */
function battleStationToggleFlag() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    SpreadsheetApp.getUi().alert('Battle Station sheet not found.');
    return;
  }

  // Get vendor name from row 2 (vendor name banner) - remove any existing flag icon
  const rawValue = String(bsSh.getRange(2, 1).getValue() || '').trim();
  const vendor = rawValue.replace(/\s*⚑\s*$/, '').trim();
  if (!vendor) {
    SpreadsheetApp.getUi().alert('No vendor currently displayed.');
    return;
  }

  const currentlyFlagged = isVendorFlagged_(vendor);
  setVendorFlag_(vendor, !currentlyFlagged);

  // Update the display with or without flag icon
  if (!currentlyFlagged) {
    bsSh.getRange(2, 1).setValue(`${vendor} ⚑`);
    ss.toast(`⚑ Flagged "${vendor}" - will stop here on next skip`, '⚑ Flagged', 3);
  } else {
    bsSh.getRange(2, 1).setValue(vendor);
    ss.toast(`Unflagged "${vendor}"`, '⚑ Unflagged', 3);
  }
}

/**
 * Get Box file blacklist from Settings sheet
 * Returns object like { "ANGI": ["fileId1", "fileId2"], "Vendor2": ["fileId3"] }
 * 
 * Settings sheet format (Box Blacklist section - can start in any column):
 * Row: "Box Blacklist" | (empty) | (empty) | (empty)
 * Row: "Vendor Name" | "File ID" | "File Name (for reference)" | "Reason"
 * Row: "ANGI" | "123456789" | "inbound1_IO.docx" | "False positive - ANGI in content"
 */
function getBoxBlacklist_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');
  
  if (!settingsSh) {
    return {};
  }
  
  const data = settingsSh.getDataRange().getValues();
  const blacklist = {};
  let inBlacklistSection = false;
  let startCol = -1;  // Track which column the blacklist starts in
  
  for (let i = 0; i < data.length; i++) {
    // Search for "Box Blacklist" header in any column
    if (!inBlacklistSection) {
      for (let col = 0; col < data[i].length; col++) {
        if (String(data[i][col] || '').trim().toLowerCase() === 'box blacklist') {
          inBlacklistSection = true;
          startCol = col;
          break;
        }
      }
      continue;
    }
    
    // Now we're in the blacklist section, use startCol for data
    const vendorCell = String(data[i][startCol] || '').trim();
    const fileIdCell = String(data[i][startCol + 1] || '').trim();
    
    // Skip header row
    if (vendorCell.toLowerCase() === 'vendor name') {
      continue;
    }
    
    // Exit if we hit an empty row (both vendor and file ID empty)
    if (vendorCell === '' && fileIdCell === '') {
      break;
    }
    
    // Parse blacklist entries
    if (vendorCell !== '' && fileIdCell !== '') {
      if (!blacklist[vendorCell]) {
        blacklist[vendorCell] = [];
      }
      blacklist[vendorCell].push(fileIdCell);
    }
  }
  
  Logger.log(`Box blacklist loaded: ${JSON.stringify(blacklist)}`);
  return blacklist;
}

/**
 * Get stored checksum for a vendor
 * Returns object with { checksum, emailChecksum, moduleChecksums }
 */
function getStoredChecksum_(vendor) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      let moduleChecksums = null;
      try {
        if (data[i][3]) {
          moduleChecksums = JSON.parse(data[i][3]);
        }
      } catch (e) {
        // Invalid JSON, ignore
      }
      return {
        checksum: data[i][1],
        emailChecksum: data[i][2] || null,
        moduleChecksums: moduleChecksums
      };
    }
  }
  
  return null;
}

/**
 * Store checksum for a vendor (including module sub-checksums)
 */
function storeChecksum_(vendor, checksum, emailChecksum, moduleChecksums) {
  const sh = getChecksumsSheet_();
  const data = sh.getDataRange().getValues();
  const now = new Date();
  const moduleJson = moduleChecksums ? JSON.stringify(moduleChecksums) : '';
  
  // Look for existing row
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === vendor.toLowerCase()) {
      sh.getRange(i + 1, 2, 1, 4).setValues([[checksum, emailChecksum || '', moduleJson, now]]);
      return;
    }
  }
  
  // Add new row
  sh.appendRow([vendor, checksum, emailChecksum || '', moduleJson, now]);
}

/************************************************************
 * HELPER FUNCTIONS - Shared utilities for change detection
 ************************************************************/

/**
 * Make a monday.com API request
 * Centralizes all monday API calls for consistency
 * @param {string} query - GraphQL query
 * @param {string} apiToken - API token (optional, defaults to BS_CFG.MONDAY_API_TOKEN)
 * @returns {object} Parsed JSON response
 */
function mondayApiRequest_(query, apiToken) {
  const token = apiToken || BS_CFG.MONDAY_API_TOKEN;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': token },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
  return JSON.parse(response.getContentText());
}

/**
 * Generate checksum for tasks data
 * @param {array} tasks - Array of task objects
 * @returns {string} Hash string
 */
function generateTasksChecksum_(tasks) {
  return hashString_(JSON.stringify((tasks || []).map(t => ({ 
    name: t.subject, 
    status: t.status, 
    created: t.created 
  }))));
}

/**
 * Generate checksum for contacts data (sorted for consistency)
 * @param {array} contacts - Array of contact objects
 * @returns {string} Hash string
 */
function generateContactsChecksum_(contacts) {
  return hashString_(JSON.stringify(
    (contacts || [])
      .map(c => ({ name: c.name, email: c.email, status: c.status }))
      .sort((a, b) => (a.name || '').localeCompare(b.name || ''))
  ));
}

/**
 * Generate checksum for meetings data
 * @param {array} meetings - Array of meeting objects
 * @returns {string} Hash string
 */
function generateMeetingsChecksum_(meetings) {
  return hashString_(JSON.stringify((meetings || []).map(m => ({ 
    title: m.title, 
    date: m.date, 
    time: m.time, 
    isPast: m.isPast 
  }))));
}

/**
 * Filter tasks based on vendor source (affiliate vs buyer)
 * Removes inappropriate onboarding tasks
 * @param {array} tasks - Array of task objects
 * @param {string} source - Vendor source (e.g., "Buyers L1M", "affiliates monday.com")
 * @returns {array} Filtered tasks
 */
function filterTasksBySource_(tasks, source) {
  const sourceLower = (source || '').toLowerCase();
  const isAffiliate = sourceLower.includes('affiliate');
  const isBuyer = sourceLower.includes('buyer');
  
  return (tasks || []).filter(task => {
    const project = (task.project || '').toLowerCase();
    if (isAffiliate && project.includes('onboarding - buyer')) return false;
    if (isBuyer && project.includes('onboarding - affiliate')) return false;
    return true;
  });
}

/**
 * Check if a vendor has changes compared to stored checksums
 * Returns detailed change info for use in skip functions
 * 
 * @param {string} vendor - Vendor name
 * @param {number} listRow - Row number in List sheet
 * @param {string} source - Vendor source for task filtering
 * @returns {object} { hasChanges, changeType, data }
 *   - hasChanges: boolean
 *   - changeType: string describing what changed (or 'first_view' or 'unchanged')
 *   - data: object with fetched data for reuse { emails, tasks, contactData, meetings }
 */
function checkVendorForChanges_(vendor, listRow, source) {
  // Check if vendor is flagged - always stop on flagged vendors
  if (isVendorFlagged_(vendor)) {
    Logger.log(`${vendor}: flagged for review`);
    // Clear the flag since we're stopping here
    setVendorFlag_(vendor, false);
    return {
      hasChanges: true,
      changeType: 'flagged',
      data: null
    };
  }

  const storedData = getStoredChecksum_(vendor);

  // If no stored data, this is a first view
  if (!storedData) {
    Logger.log(`${vendor}: no stored checksums - first view`);
    return {
      hasChanges: true,
      changeType: 'first view',
      data: null 
    };
  }
  
  // Check emails (most volatile)
  const emails = getEmailsForVendor_(vendor, listRow) || [];
  const newEmailChecksum = generateEmailChecksum_(emails);
  
  if (storedData.emailChecksum !== newEmailChecksum) {
    Logger.log(`${vendor}: emails changed (stored=${storedData.emailChecksum}, new=${newEmailChecksum})`);
    return { 
      hasChanges: true, 
      changeType: 'emails changed',
      data: { emails } 
    };
  }
  
  // Check tasks (second most volatile)
  let tasks = getTasksForVendor_(vendor, listRow) || [];
  tasks = filterTasksBySource_(tasks, source);
  const newTasksChecksum = generateTasksChecksum_(tasks);
  
  if (storedData.moduleChecksums && storedData.moduleChecksums.tasks !== newTasksChecksum) {
    Logger.log(`${vendor}: tasks changed`);
    return { 
      hasChanges: true, 
      changeType: 'tasks changed',
      data: { emails, tasks } 
    };
  }
  
  // Check contacts/notes/status
  const contactData = getVendorContacts_(vendor, listRow);
  const newNotesChecksum = hashString_(JSON.stringify(contactData.notes || ''));
  const newStatusChecksum = hashString_(JSON.stringify(contactData.liveStatus || ''));
  const newContactsChecksum = generateContactsChecksum_(contactData.contacts);
  
  if (storedData.moduleChecksums) {
    if (storedData.moduleChecksums.notes !== newNotesChecksum) {
      Logger.log(`${vendor}: notes changed`);
      return { 
        hasChanges: true, 
        changeType: 'notes changed',
        data: { emails, tasks, contactData } 
      };
    }
    if (storedData.moduleChecksums.status !== newStatusChecksum) {
      Logger.log(`${vendor}: status changed`);
      return { 
        hasChanges: true, 
        changeType: 'status changed',
        data: { emails, tasks, contactData } 
      };
    }
    if (storedData.moduleChecksums.contacts !== newContactsChecksum) {
      Logger.log(`${vendor}: contacts changed`);
      return { 
        hasChanges: true, 
        changeType: 'contacts changed',
        data: { emails, tasks, contactData } 
      };
    }
  }
  
  // Check meetings
  const contactEmails = (contactData.contacts || []).map(c => c.email).filter(e => e && e.includes('@'));
  const meetingsResult = getUpcomingMeetingsForVendor_(vendor, contactEmails);
  const meetings = meetingsResult.meetings || [];
  const newMeetingsChecksum = generateMeetingsChecksum_(meetings);
  
  if (storedData.moduleChecksums && storedData.moduleChecksums.meetings !== newMeetingsChecksum) {
    Logger.log(`${vendor}: meetings changed`);
    return { 
      hasChanges: true, 
      changeType: 'meetings changed',
      data: { emails, tasks, contactData, meetings } 
    };
  }
  
  // No changes detected
  Logger.log(`${vendor}: unchanged (emails, tasks, notes, status, contacts, meetings all same)`);
  return { 
    hasChanges: false, 
    changeType: 'unchanged',
    data: { emails, tasks, contactData, meetings } 
  };
}

/**
 * Set row background color in List sheet
 * @param {Sheet} listSh - List sheet
 * @param {number} listRow - Row number (1-based)
 * @param {string} color - Background color
 */
function setListRowColor_(listSh, listRow, color) {
  const numCols = listSh.getLastColumn();
  listSh.getRange(listRow, 1, 1, numCols).setBackground(color);
}

/**
 * Skip to next vendor with changes (different checksum)
 */
function skipToNextChanged(trackComeback) {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  // If trackComeback wasn't explicitly passed, check if there's an existing comeback pending
  if (trackComeback === undefined) {
    const comebackStr = props.getProperty('BS_COMEBACK');
    if (comebackStr) {
      trackComeback = true;  // Auto-enable if comeback is pending
    }
  }
  
  const listData = listSh.getDataRange().getValues();
  const totalVendors = listData.length - 1;
  
  // Get current index using the same function as other navigation
  let currentIdx = getCurrentVendorIndex_() || 1;
  
  let skippedCount = 0;
  
  ss.toast('Searching for changed vendors...', '⏭️ Skipping', -1);
  
  // Loop through vendors looking for one with changes - start from NEXT vendor
  while (true) {
    currentIdx++;  // Move to next vendor FIRST
    
    // Stop at end
    if (currentIdx > totalVendors) {
      ss.toast('');
      SpreadsheetApp.getUi().alert(`Checked all remaining vendors.\nSkipped ${skippedCount} unchanged vendor(s).\nNo more vendors with changes found.`);
      return;
    }
    
    const vendor = listData[currentIdx][BS_CFG.L_VENDOR];
    const source = listData[currentIdx][BS_CFG.L_SOURCE] || '';
    const listRow = currentIdx + 1;
    
    ss.toast(`Checking ${vendor}... (${skippedCount} skipped so far)`, '⏭️ Skipping', -1);
    
    // Use the centralized change detection helper
    const changeResult = checkVendorForChanges_(vendor, listRow, source);
    
    if (changeResult.hasChanges) {
      ss.toast('');
      loadVendorData(currentIdx, { forceChanged: true, loadMode: 'fast' });
      setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_CHANGED);
      if (skippedCount > 0) {
        SpreadsheetApp.getUi().alert(`Skipped ${skippedCount} unchanged vendor(s).\nNow viewing: ${vendor} (${changeResult.changeType})`);
      }
      if (trackComeback) checkComeback_();
      return;
    }
    
    // No changes - mark as skipped (yellow)
    setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_SKIPPED);
    skippedCount++;
  }
}

/**
 * Skip with Comeback - Skip to next changed vendor but mark current vendor to revisit
 * After N vendors are viewed, automatically returns to the marked vendor
 */
function skipWithComeback() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const checkSh = ss.getSheetByName(BS_CFG.CHECKSUMS_SHEET);
  
  if (!listSh || !checkSh) {
    ui.alert('Error', 'Required sheets not found.', ui.ButtonSet.OK);
    return;
  }
  
  const currentIdx = getCurrentVendorIndex_();
  if (!currentIdx || currentIdx < 1) {
    ui.alert('Error', 'No vendor currently loaded.', ui.ButtonSet.OK);
    return;
  }
  
  const listRow = currentIdx + 1;
  const vendor = listSh.getRange(listRow, 1).getValue();
  
  // Ask how many vendors to see before coming back
  const response = ui.prompt(
    '⏰ Skip with Comeback',
    `Current vendor: ${vendor}\n\nHow many vendors do you want to review before coming back to this one?`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const countStr = response.getResponseText().trim();
  const comebackAfter = parseInt(countStr, 10);
  
  if (isNaN(comebackAfter) || comebackAfter < 1 || comebackAfter > 50) {
    ui.alert('Error', 'Please enter a number between 1 and 50.', ui.ButtonSet.OK);
    return;
  }
  
  // Store comeback info in script properties
  const props = PropertiesService.getScriptProperties();
  const comebackData = {
    vendorIndex: currentIdx,
    vendorName: vendor,
    comebackAfter: comebackAfter,
    vendorsSeen: 0,
    setAt: new Date().toISOString()
  };
  props.setProperty('BS_COMEBACK', JSON.stringify(comebackData));
  
  // Color the comeback vendor row with a distinct color (light purple)
  const numCols = listSh.getLastColumn();
  listSh.getRange(listRow, 1, 1, numCols).setBackground('#e1d5e7');
  
  Logger.log(`Comeback set for ${vendor} after ${comebackAfter} vendors`);
  ss.toast(`Will return to ${vendor} after ${comebackAfter} vendors`, '⏰ Comeback Set', 3);
  
  // Now run skipToNextChanged with comeback tracking enabled
  skipToNextChanged(true);
}

/**
 * Check and handle comeback when using regular Skip Unchanged
 * Call this at the end of skipToNextChanged or wrap it
 */
function checkComeback_() {
  const props = PropertiesService.getScriptProperties();
  const comebackStr = props.getProperty('BS_COMEBACK');
  
  if (!comebackStr) return false;
  
  try {
    const comebackData = JSON.parse(comebackStr);
    comebackData.vendorsSeen++;
    
    const ss = SpreadsheetApp.getActive();
    const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
    
    Logger.log(`Comeback tracker: ${comebackData.vendorsSeen}/${comebackData.comebackAfter} vendors seen`);
    
    if (comebackData.vendorsSeen >= comebackData.comebackAfter) {
      // Time to go back!
      props.deleteProperty('BS_COMEBACK');
      
      // Clear the purple highlight
      const numCols = listSh.getLastColumn();
      listSh.getRange(comebackData.vendorIndex + 1, 1, 1, numCols).setBackground(null);
      
      // Load the comeback vendor
      ss.toast(`Returning to ${comebackData.vendorName}...`, '⏰ Comeback!', 2);
      loadVendorData(comebackData.vendorIndex);
      
      SpreadsheetApp.getUi().alert('⏰ Comeback!', `Reviewed ${comebackData.vendorsSeen} vendors.\nNow returning to: ${comebackData.vendorName}`, SpreadsheetApp.getUi().ButtonSet.OK);
      return true; // Indicates comeback was triggered
    } else {
      // Update counter
      props.setProperty('BS_COMEBACK', JSON.stringify(comebackData));
      const remaining = comebackData.comebackAfter - comebackData.vendorsSeen;
      ss.toast(`${remaining} more vendor(s) until returning to ${comebackData.vendorName}`, '⏰ Comeback', 2);
      return false;
    }
  } catch (e) {
    props.deleteProperty('BS_COMEBACK');
    return false;
  }
}

/**
 * Cancel any pending comeback
 */
function cancelComeback() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();
  
  const comebackStr = props.getProperty('BS_COMEBACK');
  
  if (!comebackStr) {
    ui.alert('No Comeback', 'There is no pending comeback to cancel.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const comebackData = JSON.parse(comebackStr);
    
    // Clear the purple highlight
    const numCols = listSh.getLastColumn();
    listSh.getRange(comebackData.vendorIndex + 1, 1, 1, numCols).setBackground(null);
    
    props.deleteProperty('BS_COMEBACK');
    
    ui.alert('Comeback Cancelled', `Cancelled comeback to: ${comebackData.vendorName}`, ui.ButtonSet.OK);
  } catch (e) {
    props.deleteProperty('BS_COMEBACK');
    ui.alert('Comeback Cancelled', 'Pending comeback has been cleared.', ui.ButtonSet.OK);
  }
}

/**
 * View current comeback status
 */
function viewComebackStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const comebackStr = props.getProperty('BS_COMEBACK');
  
  if (!comebackStr) {
    ui.alert('Comeback Status', 'No comeback currently scheduled.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const comebackData = JSON.parse(comebackStr);
    const remaining = comebackData.comebackAfter - comebackData.vendorsSeen;
    
    ui.alert('⏰ Comeback Status', 
      `Vendor: ${comebackData.vendorName}\n` +
      `Vendors seen: ${comebackData.vendorsSeen}/${comebackData.comebackAfter}\n` +
      `Remaining: ${remaining} vendor(s)\n` +
      `Set at: ${comebackData.setAt}`,
      ui.ButtonSet.OK);
  } catch (e) {
    props.deleteProperty('BS_COMEBACK');
    ui.alert('Comeback Status', 'No valid comeback scheduled.', ui.ButtonSet.OK);
  }
}

/**
 * Skip to next changed vendor 5 times, then return to the original
 * Useful for putting things in motion and letting them process while reviewing others
 */
function skip5AndReturn() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const props = PropertiesService.getScriptProperties();
  
  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }
  
  // Check if we're already in a "Skip 5 & Return" session
  const sessionStr = props.getProperty('BS_SKIP5_SESSION');
  let session = null;
  
  if (sessionStr) {
    try {
      session = JSON.parse(sessionStr);
    } catch (e) {
      props.deleteProperty('BS_SKIP5_SESSION');
    }
  }
  
  const totalVendors = listSh.getLastRow() - 1;
  const listData = listSh.getDataRange().getValues();
  
  // If no active session, start a new one
  if (!session) {
    const originIdx = getCurrentVendorIndex_();
    
    if (!originIdx) {
      SpreadsheetApp.getUi().alert('Could not determine current vendor index.');
      return;
    }
    
    const originVendor = listSh.getRange(originIdx + 1, 1).getValue();
    
    session = {
      originIdx: originIdx,
      originVendor: originVendor,
      currentIdx: originIdx,
      changedFound: 0,
      changedTarget: 5,
      skippedCount: 0,
      startedAt: new Date().toISOString()
    };
    
    ss.toast(`Starting Skip 5 & Return from: ${originVendor}`, '🔄 Skip 5 & Return', 3);
  } else {
    ss.toast(`Continuing Skip 5 & Return (${session.changedFound}/${session.changedTarget} found)`, '🔄 Skip 5 & Return', 2);
  }
  
  // Search for next changed vendor
  let currentIdx = session.currentIdx;
  
  while (currentIdx < totalVendors) {
    currentIdx++;
    
    if (currentIdx > totalVendors) {
      break;
    }
    
    const vendor = listData[currentIdx][BS_CFG.L_VENDOR];
    const source = listData[currentIdx][BS_CFG.L_SOURCE] || '';
    const listRow = currentIdx + 1;
    
    ss.toast(`Checking ${vendor}... (${session.changedFound}/${session.changedTarget} changed, ${session.skippedCount} skipped)`, '🔄 Skip 5 & Return', -1);
    
    // Use the centralized change detection helper
    const changeResult = checkVendorForChanges_(vendor, listRow, source);
    
    if (changeResult.hasChanges) {
      session.changedFound++;
      session.currentIdx = currentIdx;
      
      // Load this vendor
      loadVendorData(currentIdx, { forceChanged: true });
      setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_CHANGED);
      
      // Check if we've found all 5
      if (session.changedFound >= session.changedTarget) {
        // Done! Next call will return to origin
        ss.toast(`Found ${session.changedFound} changed vendors. Run again to return to ${session.originVendor}`, '🔄 5/5 Found', 5);
        
        // Mark session as complete (ready to return)
        session.complete = true;
        props.setProperty('BS_SKIP5_SESSION', JSON.stringify(session));
      } else {
        // More to find - save session and stop here
        ss.toast(`Changed vendor ${session.changedFound}/${session.changedTarget}: ${vendor} (${changeResult.changeType})`, '🔄 Skip 5 & Return', 3);
        props.setProperty('BS_SKIP5_SESSION', JSON.stringify(session));
      }
      return;
    }
    
    // No changes - mark as skipped (yellow)
    setListRowColor_(listSh, listRow, BS_CFG.COLOR_ROW_SKIPPED);
    session.skippedCount++;
  }
  
  // If we get here, we ran out of vendors before finding enough changed ones
  // Return to origin
  ss.toast(`Returning to: ${session.originVendor}`, '🔄 Returning', 2);
  Utilities.sleep(500);
  loadVendorData(session.originIdx);
  
  // Clear the session
  props.deleteProperty('BS_SKIP5_SESSION');
  
  SpreadsheetApp.getUi().alert(
    'Skip 5 & Return Complete',
    `Found only ${session.changedFound} changed vendor(s) before end of list.\n` +
    `Skipped ${session.skippedCount} unchanged vendor(s).\n\n` +
    `Returned to: ${session.originVendor}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Continue or complete a Skip 5 & Return session
 * Called automatically or can be called manually
 */
function continueSkip5AndReturn() {
  const props = PropertiesService.getScriptProperties();
  const sessionStr = props.getProperty('BS_SKIP5_SESSION');
  
  if (!sessionStr) {
    SpreadsheetApp.getUi().alert('No Skip 5 & Return session active.\n\nUse "Skip 5 & Return" to start a new session.');
    return;
  }
  
  let session;
  try {
    session = JSON.parse(sessionStr);
  } catch (e) {
    props.deleteProperty('BS_SKIP5_SESSION');
    SpreadsheetApp.getUi().alert('Session data corrupted. Please start a new session.');
    return;
  }
  
  // If session is complete, return to origin
  if (session.complete) {
    const ss = SpreadsheetApp.getActive();
    ss.toast(`Returning to: ${session.originVendor}`, '🔄 Returning', 2);
    Utilities.sleep(500);
    loadVendorData(session.originIdx);
    
    // Clear the session
    props.deleteProperty('BS_SKIP5_SESSION');
    
    ss.toast(`Back to ${session.originVendor} after ${session.changedFound} changed, ${session.skippedCount} skipped`, '✅ Done', 3);
    return;
  }
  
  // Otherwise, continue searching
  skip5AndReturn();
}

/**
 * Cancel an active Skip 5 & Return session
 */
function cancelSkip5Session() {
  const props = PropertiesService.getScriptProperties();
  const sessionStr = props.getProperty('BS_SKIP5_SESSION');
  
  if (!sessionStr) {
    SpreadsheetApp.getUi().alert('No Skip 5 & Return session active.');
    return;
  }
  
  props.deleteProperty('BS_SKIP5_SESSION');
  SpreadsheetApp.getUi().alert('Skip 5 & Return session cancelled.');
}

/**
 * Auto-traverse through all vendors, loading each one to record checksums
 * Processes in batches and asks to continue after each batch
 */
function autoTraverseVendors() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  
  if (!listSh) {
    ui.alert('Error', 'List sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  const totalVendors = listSh.getLastRow() - 1;
  let currentIdx = (getCurrentVendorIndex_() || 0) + 1; // Start on NEXT vendor
  
  // Make sure we don't start past the end
  if (currentIdx > totalVendors) {
    ui.alert('End of List', 'Already at the last vendor.', ui.ButtonSet.OK);
    return;
  }
  
  // Ask for batch size
  const response = ui.prompt(
    '🔁 Auto-Traverse Vendors',
    `Starting from vendor ${currentIdx} of ${totalVendors}.\n\nHow many vendors to process before pausing?\n(Enter a number, or "all" for no pausing)\n\nYou can also cancel the script from the Apps Script editor if needed.`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const inputText = response.getResponseText().trim().toLowerCase();
  let batchSize = 25; // Default batch size
  let noPause = false;
  
  if (inputText === 'all') {
    noPause = true;
    batchSize = totalVendors;
  } else {
    const parsed = parseInt(inputText);
    if (!isNaN(parsed) && parsed > 0) {
      batchSize = parsed;
    }
  }
  
  let processedCount = 0;
  let totalProcessed = 0;
  const startTime = new Date();
  
  while (currentIdx <= totalVendors) {
    const vendor = listSh.getRange(currentIdx + 1, 1).getValue();
    
    ss.toast(`Processing ${currentIdx} of ${totalVendors}: ${vendor}`, '🔁 Auto-Traverse', 2);
    
    try {
      // Load vendor data (this will record checksums)
      loadVendorData(currentIdx, { forceChanged: true });
      processedCount++;
      totalProcessed++;
      
      Logger.log(`Auto-traverse: Processed ${vendor} (${currentIdx}/${totalVendors})`);
    } catch (e) {
      Logger.log(`Auto-traverse error on ${vendor}: ${e.message}`);
    }
    
    currentIdx++;
    
    // Check if we've completed a batch
    if (!noPause && processedCount >= batchSize && currentIdx <= totalVendors) {
      const elapsed = Math.round((new Date() - startTime) / 1000);
      const remaining = totalVendors - currentIdx + 1;
      
      const continueResponse = ui.alert(
        '🔁 Batch Complete',
        `Processed ${totalProcessed} vendors (${processedCount} in this batch).\n` +
        `Time elapsed: ${elapsed} seconds\n` +
        `Current position: ${currentIdx} of ${totalVendors}\n` +
        `Remaining: ${remaining} vendors\n\n` +
        `Continue with next batch?`,
        ui.ButtonSet.YES_NO
      );
      
      if (continueResponse !== ui.Button.YES) {
        ss.toast(`Stopped at vendor ${currentIdx}. Processed ${totalProcessed} total.`, '⏹️ Stopped', 5);
        return;
      }
      
      processedCount = 0; // Reset batch counter
    }
    
    // Safety check - Apps Script has a 6 minute timeout
    const elapsedMs = new Date() - startTime;
    if (elapsedMs > 5 * 60 * 1000) { // 5 minutes
      ui.alert(
        '⏱️ Time Limit',
        `Approaching Apps Script time limit.\n` +
        `Processed ${totalProcessed} vendors.\n` +
        `Stopped at vendor ${currentIdx}.\n\n` +
        `Run Auto-Traverse again to continue from here.`,
        ui.ButtonSet.OK
      );
      return;
    }
  }
  
  const totalTime = Math.round((new Date() - startTime) / 1000);
  ui.alert(
    '✅ Complete!',
    `Auto-traverse finished!\n\nProcessed ${totalProcessed} vendors in ${totalTime} seconds.`,
    ui.ButtonSet.OK
  );
}

/**
 * TEST FUNCTION: Debug Google Drive folder search
 * Run this from the Apps Script editor to see why folders aren't being found
 */
function testGDriveSearch() {
  const vendorName = 'Ion Solar';  // Changed to test Ion Solar
  const parentFolderId = BS_CFG.GDRIVE_VENDORS_FOLDER_ID;
  
  Logger.log('='.repeat(60));
  Logger.log('GOOGLE DRIVE DEBUG TEST');
  Logger.log('='.repeat(60));
  Logger.log(`Vendor: ${vendorName}`);
  Logger.log(`Parent Folder ID: ${parentFolderId}`);
  Logger.log('');
  
  // Test 1: Can we access the parent folder?
  Logger.log('--- TEST 1: Access Parent Folder ---');
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    Logger.log(`✅ Parent folder found: ${parentFolder.getName()}`);
    Logger.log(`   URL: ${parentFolder.getUrl()}`);
  } catch (e) {
    Logger.log(`❌ Cannot access parent folder: ${e.message}`);
    return;
  }
  
  // Test 2: List ALL subfolders in parent (first 20)
  Logger.log('');
  Logger.log('--- TEST 2: List Subfolders (first 20) ---');
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folders = parentFolder.getFolders();
    let count = 0;
    while (folders.hasNext() && count < 20) {
      const folder = folders.next();
      Logger.log(`  ${count + 1}. "${folder.getName()}" (ID: ${folder.getId()})`);
      count++;
    }
    Logger.log(`   Total shown: ${count}`);
  } catch (e) {
    Logger.log(`❌ Error listing folders: ${e.message}`);
  }
  
  // Test 3: Search for folder using searchFolders
  Logger.log('');
  Logger.log('--- TEST 3: searchFolders() with vendor name ---');
  try {
    const query1 = `title contains '${vendorName}' and '${parentFolderId}' in parents`;
    Logger.log(`Query: ${query1}`);
    const results1 = DriveApp.searchFolders(query1);
    let found1 = 0;
    while (results1.hasNext()) {
      const folder = results1.next();
      Logger.log(`  ✅ Found: "${folder.getName()}" (ID: ${folder.getId()})`);
      found1++;
    }
    if (found1 === 0) Logger.log('  ❌ No results');
  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
  
  // Test 4: Search without parent restriction
  Logger.log('');
  Logger.log('--- TEST 4: searchFolders() without parent restriction ---');
  try {
    const query2 = `title contains '${vendorName}' and mimeType = 'application/vnd.google-apps.folder'`;
    Logger.log(`Query: ${query2}`);
    const results2 = DriveApp.searchFolders(query2);
    let found2 = 0;
    while (results2.hasNext() && found2 < 10) {
      const folder = results2.next();
      Logger.log(`  Found: "${folder.getName()}" (ID: ${folder.getId()})`);
      found2++;
    }
    if (found2 === 0) Logger.log('  ❌ No results');
  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
  
  // Test 5: Try to access the known folder directly
  Logger.log('');
  Logger.log('--- TEST 5: Direct access to known SunPower folder ---');
  const knownFolderId = '1UbRPIvG6ZsO3nNJ12h2YgG0W5fJzR5ae';
  try {
    const folder = DriveApp.getFolderById(knownFolderId);
    Logger.log(`✅ Direct access works: "${folder.getName()}"`);
    Logger.log(`   URL: ${folder.getUrl()}`);
    
    // Check its parent
    const parents = folder.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      Logger.log(`   Parent: "${parent.getName()}" (ID: ${parent.getId()})`);
    }
  } catch (e) {
    Logger.log(`❌ Cannot access folder directly: ${e.message}`);
  }
  
  // Test 6: List files in the known folder
  Logger.log('');
  Logger.log('--- TEST 6: Files in known SunPower folder ---');
  try {
    const folder = DriveApp.getFolderById(knownFolderId);
    const files = folder.getFiles();
    let fileCount = 0;
    while (files.hasNext() && fileCount < 10) {
      const file = files.next();
      Logger.log(`  ${fileCount + 1}. "${file.getName()}"`);
      fileCount++;
    }
    Logger.log(`   Total shown: ${fileCount}`);
  } catch (e) {
    Logger.log(`❌ Error: ${e.message}`);
  }
  
  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('TEST COMPLETE - Check logs above');
  Logger.log('='.repeat(60));
}

/************************************************************
 * MONDAY.COM BOARD SYNC
 * Pull data directly from monday.com boards to update
 * the affiliates monday.com and buyers monday.com sheets
 ************************************************************/

/**
 * Main function to sync monday.com boards to sheets
 */
function syncMondayComBoards() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Starting sync...', '🔄 Syncing', 30);
  
  try {
    // Sync Buyers board
    ss.toast('Syncing Buyers board...', '🔄 Syncing', 30);
    const buyersResult = syncBuyersBoard_();
    Logger.log(`Buyers sync complete: ${buyersResult.count} items`);
    
    // Sync Affiliates board
    ss.toast('Syncing Affiliates board...', '🔄 Syncing', 30);
    const affiliatesResult = syncAffiliatesBoard_();
    Logger.log(`Affiliates sync complete: ${affiliatesResult.count} items`);
    
    ss.toast(`Sync complete! Buyers: ${buyersResult.count}, Affiliates: ${affiliatesResult.count}`, '✅ Done', 5);
    
  } catch (e) {
    Logger.log(`Sync error: ${e.message}`);
    ss.toast(`Error: ${e.message}`, '❌ Error', 5);
  }
}

/**
 * Sync Buyers board to buyers monday.com sheet
 */
function syncBuyersBoard_() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'buyers monday.com';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Column IDs from testGetBoardColumns() output
  const columns = [
    { id: 'name', header: 'Name', type: 'name' },
    { id: 'subtasks_mktcnxe', header: 'Subitems', type: 'subtasks' },
    { id: 'color_mktqyter', header: 'Buyer Type', type: 'status' },
    { id: 'text_mkqnvsqh', header: 'BUYER NOTES MAIN', type: 'text' },
    { id: 'rating_mkv5ztjb', header: 'Opportunity Size', type: 'rating' },
    { id: 'tag_mkskgt84', header: 'Live Verticals', type: 'tags' },
    { id: 'tag_mkskewmq', header: 'Other Verticals', type: 'tags' },
    { id: 'board_relation_mky0bt0z', header: 'Contacts', type: 'board_relation' },
    { id: 'tag_mkskfmf3', header: 'Live Modalities', type: 'tags' },
    { id: 'tag_mkskassa', header: 'Other Modalities', type: 'tags' },
    { id: 'link_mksmwprd', header: 'Phonexa Link', type: 'link' },
    { id: 'link_mksmgg2h', header: 'Company URL', type: 'link' },
    { id: 'pulse_log_mkthvn03', header: 'Creation log', type: 'creation_log' },
    { id: 'board_relation_mkvd98v7', header: 'Sourcing', type: 'board_relation' },
    { id: 'text_mkvkr178', header: 'Other Name', type: 'text' },
    { id: 'pulse_updated_mkvqtmew', header: 'Last updated', type: 'last_updated' },
    { id: 'numeric_mkwp5np4', header: 'Rank', type: 'numbers' },
    { id: 'board_relation_mky0j7qj', header: 'link to Helpful Links', type: 'board_relation' },
    { id: 'dropdown_mkyam4qw', header: 'State(s)', type: 'dropdown' },
    { id: 'dropdown_mkyazy2j', header: 'Dead States(s)', type: 'dropdown' }
  ];
  
  return syncBoardToSheet_(BS_CFG.BUYERS_BOARD_ID, sheet, columns, 'Buyers');
}

/**
 * Sync Affiliates board to affiliates monday.com sheet
 */
function syncAffiliatesBoard_() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'affiliates monday.com';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Column IDs from testGetBoardColumns() output
  const columns = [
    { id: 'name', header: 'Name', type: 'name' },
    { id: 'multiple_person_mkt9k1n1', header: 'AM', type: 'people' },
    { id: 'text_mkrdahqz', header: 'AFFILIATE NOTES MAIN', type: 'text' },
    { id: 'rating_mkv5k453', header: 'Opportunity Size', type: 'rating' },
    { id: 'board_relation_mky0n0rf', header: 'Contacts', type: 'board_relation' },
    { id: 'tag_mksm70xw', header: 'Traffic Sources', type: 'tags' },
    { id: 'tag_mkskrddx', header: 'Live Verticals', type: 'tags' },
    { id: 'tag_mkskfs70', header: 'Other Verticals', type: 'tags' },
    { id: 'tag_mksk7whx', header: 'Live Modalities', type: 'tags' },
    { id: 'tag_mkskkszw', header: 'Other Modalities', type: 'tags' },
    { id: 'link_mksmgnc0', header: 'Phonexa Link', type: 'link' },
    { id: 'text_mksmcrpw', header: 'Other Name', type: 'text' },
    { id: 'board_relation_mksnsmsg', header: 'Sourcing', type: 'board_relation' },
    { id: 'color_mksy6tak', header: 'Priority', type: 'status' },
    { id: 'board_relation_mkthwkt6', header: 'URLs - Affiliates', type: 'board_relation' },
    { id: 'subtasks_mkvgk8ab', header: 'Subitems', type: 'subtasks' },
    { id: 'pulse_updated_mkvq53b1', header: 'Last updated', type: 'last_updated' },
    { id: 'boolean_mkxb61bz', header: 'imp', type: 'checkbox' },
    { id: 'board_relation_mky04azc', header: 'link to Helpful Links', type: 'board_relation' }
  ];
  
  return syncBoardToSheet_(BS_CFG.AFFILIATES_BOARD_ID, sheet, columns, 'Affiliates');
}

/**
 * Generic function to sync a monday.com board to a sheet
 */
function syncBoardToSheet_(boardId, sheet, columns, boardName) {
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  
  // Build column IDs for query (excluding 'name' which is automatic)
  const columnIds = columns
    .filter(c => c.id !== 'name')
    .map(c => `"${c.id}"`)
    .join(', ');
  
  // Query all items with pagination
  const allItems = [];
  let cursor = null;
  let pageCount = 0;
  const maxPages = 20; // Safety limit
  
  do {
    pageCount++;
    Logger.log(`Fetching ${boardName} page ${pageCount}...`);
    
    const cursorPart = cursor ? `, cursor: "${cursor}"` : '';
    const query = `
      query {
        boards(ids: [${boardId}]) {
          groups {
            id
            title
          }
          items_page(limit: 500${cursorPart}) {
            cursor
            items {
              id
              name
              group {
                id
                title
              }
              column_values(ids: [${columnIds}]) {
                id
                text
                value
                ... on MirrorValue {
                  display_value
                }
                ... on BoardRelationValue {
                  display_value
                }
                ... on StatusValue {
                  label
                }
                ... on DropdownValue {
                  text
                }
                ... on TagsValue {
                  text
                }
                ... on PersonValue {
                  text
                }
              }
            }
          }
        }
      }
    `;
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': apiToken },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.errors) {
      Logger.log(`API Error: ${JSON.stringify(result.errors)}`);
      throw new Error(result.errors[0].message);
    }
    
    const itemsPage = result.data?.boards?.[0]?.items_page;
    if (!itemsPage) {
      Logger.log('No items_page in response');
      break;
    }
    
    allItems.push(...itemsPage.items);
    cursor = itemsPage.cursor;
    
    Logger.log(`  Fetched ${itemsPage.items.length} items (total: ${allItems.length})`);
    
  } while (cursor && pageCount < maxPages);
  
  Logger.log(`Total ${boardName} items fetched: ${allItems.length}`);
  
  // Build sheet data
  const headers = columns.map(c => c.header);
  // Add Group column at the end
  headers.push('Group');
  
  const rows = allItems.map(item => {
    const row = columns.map(col => {
      if (col.id === 'name') {
        return item.name;
      }
      
      const colValue = item.column_values.find(cv => cv.id === col.id);
      if (!colValue) return '';
      
      // Extract value based on type
      switch (col.type) {
        case 'status':
          return colValue.label || colValue.text || '';
        case 'board_relation':
          return colValue.display_value || colValue.text || '';
        case 'tags':
        case 'dropdown':
        case 'people':
          return colValue.text || '';
        case 'checkbox':
          // Parse checkbox JSON to get checked state
          try {
            const checkData = JSON.parse(colValue.value || '{}');
            return checkData.checked === 'true' || checkData.checked === true ? 'Yes' : '';
          } catch {
            return colValue.text || '';
          }
        case 'rating':
          // Parse rating JSON to get rating value
          try {
            const ratingData = JSON.parse(colValue.value || '{}');
            return ratingData.rating || colValue.text || '';
          } catch {
            return colValue.text || '';
          }
        case 'link':
          // Parse link JSON to get URL
          try {
            const linkData = JSON.parse(colValue.value || '{}');
            return linkData.url || colValue.text || '';
          } catch {
            return colValue.text || '';
          }
        case 'creation_log':
        case 'last_updated':
          // Parse to get date
          try {
            const logData = JSON.parse(colValue.value || '{}');
            if (logData.created_at) {
              return formatMondayDate_(logData.created_at);
            }
            if (logData.updated_at) {
              return formatMondayDate_(logData.updated_at);
            }
            return colValue.text || '';
          } catch {
            return colValue.text || '';
          }
        default:
          return colValue.text || '';
      }
    });
    
    // Add group name
    row.push(item.group?.title || '');
    
    return row;
  });
  
  // Clear sheet and write data
  sheet.clear();
  
  // Write headers
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white');
  }
  
  // Write data rows
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  return { count: rows.length };
}

/**
 * Format monday.com date to readable format
 */
function formatMondayDate_(dateStr) {
  try {
    const date = new Date(dateStr);
    // Format as "MMM D, YYYY H:MM AM/PM"
    return Utilities.formatDate(date, 'America/Los_Angeles', 'MMM d, yyyy h:mm a');
  } catch {
    return dateStr;
  }
}

/**
 * Test function to get column IDs from a board
 * Run this to discover column IDs for new boards
 */
function testGetBoardColumns() {
  const boardId = BS_CFG.AFFILIATES_BOARD_ID; // Change to BUYERS_BOARD_ID or AFFILIATES_BOARD_ID
  const apiToken = BS_CFG.MONDAY_API_TOKEN;
  
  const query = `
    query {
      boards(ids: [${boardId}]) {
        name
        columns {
          id
          title
          type
        }
      }
    }
  `;
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': apiToken },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
  const result = JSON.parse(response.getContentText());
  
  Logger.log('Board: ' + result.data?.boards?.[0]?.name);
  Logger.log('Columns:');
  
  const columns = result.data?.boards?.[0]?.columns || [];
  columns.forEach(col => {
    Logger.log(`  { id: '${col.id}', header: '${col.title}', type: '${col.type}' },`);
  });
}

/************************************************************
 * DUPLICATE VENDOR CHECK
 * Identify vendors appearing in both Buyers and Affiliates boards
 ************************************************************/

/**
 * Check for vendors appearing in both Buyers and Affiliates
 * Also checks for duplicates within the same board
 * Shows results and allows adding to exclusion list
 */
function checkDuplicateVendors() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  ss.toast('Checking for duplicate vendors...', '🔍 Scanning', 5);
  
  // Get vendors from both sheets
  const buyersSheet = ss.getSheetByName('buyers monday.com');
  const affiliatesSheet = ss.getSheetByName('affiliates monday.com');
  
  if (!buyersSheet || !affiliatesSheet) {
    ui.alert('Error', 'Please run "Sync monday.com Data" first to populate the buyers and affiliates sheets.', ui.ButtonSet.OK);
    return;
  }
  
  // Get vendor names from column A (skip header)
  const buyersData = buyersSheet.getRange(2, 1, buyersSheet.getLastRow() - 1, 1).getValues();
  const affiliatesData = affiliatesSheet.getRange(2, 1, affiliatesSheet.getLastRow() - 1, 1).getValues();
  
  // Find duplicates within each board
  const buyerDuplicatesWithinBoard = findDuplicatesInList_(buyersData);
  const affiliateDuplicatesWithinBoard = findDuplicatesInList_(affiliatesData);
  
  const buyerNames = new Set(buyersData.map(row => String(row[0]).trim().toLowerCase()).filter(n => n));
  const affiliateNames = new Set(affiliatesData.map(row => String(row[0]).trim().toLowerCase()).filter(n => n));
  
  // Get exclusion list from Settings
  const exclusions = getDuplicateExclusions_();
  const exclusionSet = new Set(exclusions.map(e => e.toLowerCase()));
  
  // Find cross-board duplicates (in both, not in exclusion list)
  const crossBoardDuplicates = [];
  for (const name of buyerNames) {
    if (affiliateNames.has(name) && !exclusionSet.has(name)) {
      // Get the proper case version
      const properCase = buyersData.find(row => String(row[0]).trim().toLowerCase() === name)?.[0] || name;
      crossBoardDuplicates.push(properCase);
    }
  }
  
  // Also find excluded ones for reference
  const excludedDuplicates = [];
  for (const name of buyerNames) {
    if (affiliateNames.has(name) && exclusionSet.has(name)) {
      const properCase = buyersData.find(row => String(row[0]).trim().toLowerCase() === name)?.[0] || name;
      excludedDuplicates.push(properCase);
    }
  }
  
  Logger.log(`Found ${crossBoardDuplicates.length} cross-board duplicates (${excludedDuplicates.length} excluded)`);
  Logger.log(`Found ${buyerDuplicatesWithinBoard.length} duplicates within Buyers board`);
  Logger.log(`Found ${affiliateDuplicatesWithinBoard.length} duplicates within Affiliates board`);
  
  const totalIssues = crossBoardDuplicates.length + buyerDuplicatesWithinBoard.length + affiliateDuplicatesWithinBoard.length;
  
  if (totalIssues === 0) {
    let message = '✅ No duplicate vendors found!';
    if (excludedDuplicates.length > 0) {
      message += `\n\n(${excludedDuplicates.length} known duplicates are excluded via Settings)`;
    }
    ui.alert('Duplicate Check Complete', message, ui.ButtonSet.OK);
    return;
  }
  
  // Show results
  let message = '';
  
  if (crossBoardDuplicates.length > 0) {
    message += `⚠️ ${crossBoardDuplicates.length} vendor(s) in BOTH Buyers and Affiliates:\n`;
    message += crossBoardDuplicates.slice(0, 10).join(', ');
    if (crossBoardDuplicates.length > 10) {
      message += `, ... +${crossBoardDuplicates.length - 10} more`;
    }
    message += '\n\n';
  }
  
  if (buyerDuplicatesWithinBoard.length > 0) {
    message += `🔵 ${buyerDuplicatesWithinBoard.length} duplicate(s) within Buyers board:\n`;
    message += buyerDuplicatesWithinBoard.slice(0, 10).join(', ');
    if (buyerDuplicatesWithinBoard.length > 10) {
      message += `, ... +${buyerDuplicatesWithinBoard.length - 10} more`;
    }
    message += '\n\n';
  }
  
  if (affiliateDuplicatesWithinBoard.length > 0) {
    message += `🔵 ${affiliateDuplicatesWithinBoard.length} duplicate(s) within Affiliates board:\n`;
    message += affiliateDuplicatesWithinBoard.slice(0, 10).join(', ');
    if (affiliateDuplicatesWithinBoard.length > 10) {
      message += `, ... +${affiliateDuplicatesWithinBoard.length - 10} more`;
    }
    message += '\n\n';
  }
  
  if (excludedDuplicates.length > 0) {
    message += `(${excludedDuplicates.length} known cross-board duplicates already excluded)\n\n`;
  }
  message += 'See "Duplicate Vendors" sheet for full list.';
  
  ui.alert('Duplicate Vendors Found', message, ui.ButtonSet.OK);
  
  // Write to sheet for easier review
  writeDuplicatesToSheet_(crossBoardDuplicates, excludedDuplicates, buyerDuplicatesWithinBoard, affiliateDuplicatesWithinBoard);
}

/**
 * Find duplicate names within a single list
 * Returns array of names that appear more than once
 */
function findDuplicatesInList_(data) {
  const nameCounts = {};
  const duplicates = [];
  
  for (const row of data) {
    const name = String(row[0]).trim();
    const nameLower = name.toLowerCase();
    if (!nameLower) continue;
    
    if (nameCounts[nameLower]) {
      nameCounts[nameLower].count++;
      // Only add to duplicates once (when we see it the second time)
      if (nameCounts[nameLower].count === 2) {
        duplicates.push(nameCounts[nameLower].properCase);
      }
    } else {
      nameCounts[nameLower] = { count: 1, properCase: name };
    }
  }
  
  return duplicates;
}

/**
 * Get duplicate exclusions from Settings sheet
 * Format in Settings:
 * Row: "Duplicate Exclusions" | (empty)
 * Row: "Vendor Name" | "Reason"
 * Row: "SomeVendor" | "Legitimately both buyer and affiliate"
 */
function getDuplicateExclusions_() {
  const ss = SpreadsheetApp.getActive();
  const settingsSh = ss.getSheetByName('Settings');
  
  if (!settingsSh) {
    return [];
  }
  
  const data = settingsSh.getDataRange().getValues();
  const exclusions = [];
  let inExclusionsSection = false;
  let startCol = -1;
  
  for (let i = 0; i < data.length; i++) {
    // Search for "Duplicate Exclusions" header in any column
    if (!inExclusionsSection) {
      for (let col = 0; col < data[i].length; col++) {
        if (String(data[i][col] || '').trim().toLowerCase() === 'duplicate exclusions') {
          inExclusionsSection = true;
          startCol = col;
          break;
        }
      }
      continue;
    }
    
    const vendorCell = String(data[i][startCol] || '').trim();
    
    // Skip header row
    if (vendorCell.toLowerCase() === 'vendor name') {
      continue;
    }
    
    // Exit if we hit an empty row
    if (vendorCell === '') {
      break;
    }
    
    exclusions.push(vendorCell);
  }
  
  Logger.log(`Loaded ${exclusions.length} duplicate exclusions from Settings`);
  return exclusions;
}

/**
 * Write duplicates to a sheet for easy review
 */
function writeDuplicatesToSheet_(crossBoardDuplicates, excludedDuplicates, buyerDuplicatesWithinBoard, affiliateDuplicatesWithinBoard) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'Duplicate Vendors';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  // Headers
  sheet.getRange(1, 1, 1, 4).setValues([['Vendor Name', 'Status', 'Type', 'Notes']]);
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('white');
  
  let row = 2;
  
  // Write cross-board duplicates (yellow)
  for (const vendor of crossBoardDuplicates) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('⚠️ DUPLICATE');
    sheet.getRange(row, 3).setValue('Cross-Board');
    sheet.getRange(row, 4).setValue('In both Buyers AND Affiliates');
    sheet.getRange(row, 1, 1, 4).setBackground('#fff2cc'); // Yellow
    row++;
  }
  
  // Write same-board duplicates - Buyers (blue)
  for (const vendor of buyerDuplicatesWithinBoard) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('🔵 DUPLICATE');
    sheet.getRange(row, 3).setValue('Within Buyers');
    sheet.getRange(row, 4).setValue('Appears multiple times in Buyers board');
    sheet.getRange(row, 1, 1, 4).setBackground('#cfe2f3'); // Light blue
    row++;
  }
  
  // Write same-board duplicates - Affiliates (blue)
  for (const vendor of affiliateDuplicatesWithinBoard) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('🔵 DUPLICATE');
    sheet.getRange(row, 3).setValue('Within Affiliates');
    sheet.getRange(row, 4).setValue('Appears multiple times in Affiliates board');
    sheet.getRange(row, 1, 1, 4).setBackground('#cfe2f3'); // Light blue
    row++;
  }
  
  // Write excluded duplicates (green)
  for (const vendor of excludedDuplicates) {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue('✓ Excluded');
    sheet.getRange(row, 3).setValue('Cross-Board');
    sheet.getRange(row, 4).setValue('In Settings exclusion list');
    sheet.getRange(row, 1, 1, 4).setBackground('#d9ead3'); // Green
    row++;
  }
  
  // Auto-resize
  sheet.autoResizeColumns(1, 4);
  sheet.setFrozenRows(1);
  
  // Add legend and instructions
  const totalRows = crossBoardDuplicates.length + buyerDuplicatesWithinBoard.length + 
                    affiliateDuplicatesWithinBoard.length + excludedDuplicates.length;
  if (totalRows > 0) {
    row += 2;
    sheet.getRange(row, 1).setValue('LEGEND:').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 2).setValues([['🟡 Yellow', 'Cross-board duplicate (in both Buyers and Affiliates)']]);
    sheet.getRange(row, 1, 1, 2).setBackground('#fff2cc');
    row++;
    sheet.getRange(row, 1, 1, 2).setValues([['🔵 Blue', 'Same-board duplicate (appears multiple times on one board)']]);
    sheet.getRange(row, 1, 1, 2).setBackground('#cfe2f3');
    row++;
    sheet.getRange(row, 1, 1, 2).setValues([['🟢 Green', 'Excluded (in Settings exclusion list)']]);
    sheet.getRange(row, 1, 1, 2).setBackground('#d9ead3');
    row += 2;
    
    sheet.getRange(row, 1).setValue('To exclude cross-board duplicates, add to Settings sheet:').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue('Duplicate Exclusions');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue('Vendor Name');
    sheet.getRange(row, 2).setValue('Reason');
    sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#f3f3f3');
    row++;
    sheet.getRange(row, 1).setValue('Example Vendor');
    sheet.getRange(row, 2).setValue('Legitimately both buyer and affiliate');
    sheet.getRange(row, 1, 1, 2).setFontStyle('italic');
  }
  
  ss.toast(`Results written to "${sheetName}" sheet`, '✅ Done', 3);
}

/************************************************************
 * CLAUDE API KEY MANAGEMENT
 * Store and retrieve API key from Script Properties
 ************************************************************/

/**
 * Get Claude API key - checks Script Properties first, falls back to BS_CFG
 */
function getClaudeApiKey_() {
  const props = PropertiesService.getScriptProperties();
  const storedKey = props.getProperty('CLAUDE_API_KEY');
  if (storedKey && storedKey.length > 10) return storedKey;
  if (BS_CFG.CLAUDE_API_KEY && BS_CFG.CLAUDE_API_KEY !== 'YOUR_ANTHROPIC_API_KEY_HERE') {
    return BS_CFG.CLAUDE_API_KEY;
  }
  return null;
}

/**
 * Set Claude API key via dialog prompt - stores in Script Properties
 */
function battleStationSetClaudeApiKey() {
  const ui = SpreadsheetApp.getUi();
  const currentKey = getClaudeApiKey_();
  const masked = currentKey ? '***' + currentKey.slice(-8) : '(not set)';

  const response = ui.prompt(
    '⚙️ Set Claude API Key',
    `Current key: ${masked}\n\nEnter your Anthropic API key (starts with sk-ant-):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const newKey = response.getResponseText().trim();
  if (!newKey || newKey.length < 20) {
    ui.alert('Invalid API key. Key must be at least 20 characters.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY', newKey);
  SpreadsheetApp.getActive().toast('Claude API key saved!', '✅ Success', 3);
}

/************************************************************
 * SMART BRIEFING - AI-powered vendor priority advisor
 * Scans across all vendors and recommends actions
 ************************************************************/

/**
 * Smart Briefing - Claude analyzes all active vendors and tells you what to do
 */
function battleStationSmartBriefing() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const ui = SpreadsheetApp.getUi();

  const apiKey = getClaudeApiKey_();
  if (!apiKey) {
    ui.alert('No Claude API key configured.\n\nUse menu: ⚡ Battle Station → ⚙️ Set Claude API Key');
    return;
  }

  if (!listSh) {
    ui.alert('List sheet not found.');
    return;
  }

  ss.toast('Scanning vendor emails for Smart Briefing...', '🧠 Thinking', 10);

  const allVendors = listSh.getRange(2, 1, listSh.getLastRow() - 1, 8).getValues();
  const vendorSummaries = [];
  const maxVendors = 25;

  for (let i = 0; i < Math.min(allVendors.length, maxVendors); i++) {
    const vendor = allVendors[i][BS_CFG.L_VENDOR];
    const status = allVendors[i][BS_CFG.L_STATUS];
    const notes = allVendors[i][BS_CFG.L_NOTES];
    const listRow = i + 2;

    if (!vendor) continue;

    try {
      const emails = getEmailsForVendor_(vendor, listRow);
      const unsnoozed = emails.filter(e => !e.isSnoozed);
      const overdue = emails.filter(e => isEmailOverdue_(e));

      if (unsnoozed.length === 0 && !overdue.length) continue;

      // Get latest email content for context
      let latestContent = '';
      if (unsnoozed.length > 0 && unsnoozed[0].threadId) {
        try {
          const thread = GmailApp.getThreadById(unsnoozed[0].threadId);
          if (thread) {
            const msgs = thread.getMessages();
            latestContent = msgs[msgs.length - 1].getPlainBody().substring(0, 600);
          }
        } catch (e) { /* skip if thread fetch fails */ }
      }

      vendorSummaries.push({
        vendor: vendor,
        status: status || '(unknown)',
        notes: (notes || '').substring(0, 200),
        emailCount: unsnoozed.length,
        overdueCount: overdue.length,
        latestSubject: unsnoozed[0] ? unsnoozed[0].subject : '',
        latestLabels: unsnoozed[0] ? unsnoozed[0].labels : '',
        latestContent: latestContent
      });
    } catch (e) {
      Logger.log(`Smart Briefing: Error scanning ${vendor}: ${e.message}`);
    }

    if (i % 5 === 0) {
      ss.toast(`Scanned ${i + 1} of ${Math.min(allVendors.length, maxVendors)} vendors...`, '🧠 Scanning', 5);
    }
  }

  if (vendorSummaries.length === 0) {
    ui.alert('No active vendor communications found to analyze.');
    return;
  }

  ss.toast(`Analyzing ${vendorSummaries.length} vendors with Claude...`, '🧠 Processing', 15);

  const vendorText = vendorSummaries.map((v, i) =>
    `\n--- VENDOR ${i + 1}: ${v.vendor} ---
Status: ${v.status}
Notes: ${v.notes || '(none)'}
Active emails: ${v.emailCount}, Overdue: ${v.overdueCount}
Latest email subject: ${v.latestSubject}
Latest labels: ${v.latestLabels}
Content preview: ${v.latestContent}`
  ).join('\n');

  // Build a name lookup for matching Claude's output back to exact vendor names
  const vendorNameList = vendorSummaries.map(v => v.vendor);

  const prompt = `You are an AI assistant for Andy, a vendor relationship manager at Profitise (a lead generation company in Home Services and Solar verticals).

Here's a snapshot of Andy's ${vendorSummaries.length} most active vendor communications:
${vendorText}

Provide a SMART BRIEFING with specific, actionable recommendations:

## 🔥 URGENT — Do Right Now
[List 3-5 specific actions for overdue or time-sensitive items. Name the vendor and what to do.]

## 📋 HIGH PRIORITY — Today
[List 3-5 important follow-ups with specific vendor names and actions.]

## 💡 OPPORTUNITIES
[2-3 things Andy could proactively do to improve business outcomes or relationships.]

## ⚠️ RISKS
[2-3 vendors or situations that need attention before they become problems.]

Be specific: name the vendor, state the action, explain why. Keep it concise and actionable.

IMPORTANT: At the very end of your response, include a section exactly like this:

## PRIORITY_ORDER
[List EVERY vendor name from above, one per line, in the exact order Andy should work through them. Most urgent first. Use the EXACT vendor names as provided. No numbers, no dashes, just the vendor name on each line.]`;

  try {
    const response = callClaudeAPI_(prompt, apiKey, { maxTokens: 3000 });

    if (response.error) {
      ui.alert(`Claude API Error: ${response.error}`);
      return;
    }

    const rawContent = response.content;

    // Parse the PRIORITY_ORDER section from the response
    const priorityNames = parsePriorityOrder_(rawContent, vendorNameList);

    // Reorder the List sheet to match the priority order
    if (priorityNames.length > 0) {
      reorderListByPriority_(listSh, priorityNames);
      ss.toast(`List reordered: ${priorityNames.length} priority vendors at top`, '🧠 Reordered', 3);
    }

    // Strip PRIORITY_ORDER section from display content
    let displayContent = rawContent.replace(/## PRIORITY_ORDER[\s\S]*$/, '').trim();

    let content = displayContent
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\n/g, '<br>')
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/## (.*?)<br>/g, '<h3>$1</h3>');

    const htmlContent = `
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; line-height: 1.6; font-size: 13px; }
        h2 { color: #4a86e8; margin-top: 0; }
        h3 { color: #4a86e8; margin-top: 16px; margin-bottom: 8px; border-bottom: 2px solid #4a86e8; padding-bottom: 4px; }
        strong { color: #333; }
        .meta { color: #888; font-size: 11px; margin-bottom: 10px; }
        .reorder-note { background: #e8f0fe; padding: 8px 12px; border-radius: 4px; margin-bottom: 12px; font-size: 12px; }
      </style>
      <h2>🧠 Smart Briefing</h2>
      <p class="meta">Analyzed ${vendorSummaries.length} vendors | ${new Date().toLocaleString()}</p>
      ${priorityNames.length > 0 ? '<div class="reorder-note">✅ List has been reordered — hit <strong>Next Vendor</strong> to work through them in priority order.</div>' : ''}
      <div>${content}</div>
    `;

    const html = HtmlService.createHtmlOutput(htmlContent).setWidth(800).setHeight(650);
    ui.showModalDialog(html, '🧠 Smart Briefing — What to do next');

    // Load vendor #1 so user can start traversing
    if (priorityNames.length > 0) {
      loadVendorData(1, { loadMode: 'fast' });
    }

  } catch (e) {
    ui.alert(`Error: ${e.message}`);
  }
}

/**
 * Parse PRIORITY_ORDER section from Claude's response
 * Matches vendor names against the known list (fuzzy)
 */
function parsePriorityOrder_(responseText, knownVendors) {
  const prioritySection = responseText.match(/## PRIORITY_ORDER\s*\n([\s\S]*?)$/);
  if (!prioritySection) return [];

  const lines = prioritySection[1].split('\n')
    .map(l => l.replace(/^[\d.\-*•]+\s*/, '').trim()) // Strip numbering/bullets
    .filter(l => l.length > 0);

  // Build lowercase lookup for known vendors
  const vendorLookup = {};
  for (const v of knownVendors) {
    vendorLookup[v.toLowerCase()] = v;
  }

  const matched = [];
  const seen = new Set();

  for (const line of lines) {
    const lower = line.toLowerCase();
    // Exact match first
    if (vendorLookup[lower] && !seen.has(lower)) {
      matched.push(vendorLookup[lower]);
      seen.add(lower);
      continue;
    }
    // Fuzzy: check if any known vendor is contained in the line
    for (const known of knownVendors) {
      const knownLower = known.toLowerCase();
      if (!seen.has(knownLower) && (lower.includes(knownLower) || knownLower.includes(lower))) {
        matched.push(known);
        seen.add(knownLower);
        break;
      }
    }
  }

  Logger.log(`Smart Briefing: Parsed ${matched.length} priority vendors from ${lines.length} lines`);
  return matched;
}

/**
 * Reorder the List sheet so priority vendors appear at the top
 * Non-priority vendors keep their existing relative order below
 */
function reorderListByPriority_(listSh, priorityNames) {
  const lastRow = listSh.getLastRow();
  if (lastRow <= 1) return;

  const numCols = listSh.getLastColumn();
  const dataRange = listSh.getRange(2, 1, lastRow - 1, numCols);
  const allData = dataRange.getValues();
  const allBgs = dataRange.getBackgrounds();

  // Build priority lookup: vendor name -> priority index
  const priorityMap = {};
  for (let i = 0; i < priorityNames.length; i++) {
    priorityMap[priorityNames[i].toLowerCase()] = i;
  }

  // Split into priority rows and non-priority rows
  const priorityRows = new Array(priorityNames.length).fill(null);
  const priorityBgs = new Array(priorityNames.length).fill(null);
  const otherRows = [];
  const otherBgs = [];

  for (let i = 0; i < allData.length; i++) {
    const vendorName = String(allData[i][0] || '').toLowerCase();
    const pIdx = priorityMap[vendorName];
    if (pIdx !== undefined && priorityRows[pIdx] === null) {
      priorityRows[pIdx] = allData[i];
      priorityBgs[pIdx] = allBgs[i];
    } else {
      otherRows.push(allData[i]);
      otherBgs.push(allBgs[i]);
    }
  }

  // Remove any nulls (vendors in priority list but not found in sheet)
  const finalRows = [];
  const finalBgs = [];
  for (let i = 0; i < priorityRows.length; i++) {
    if (priorityRows[i]) {
      finalRows.push(priorityRows[i]);
      finalBgs.push(priorityBgs[i]);
    }
  }

  // Combine: priority vendors first, then the rest
  const combined = finalRows.concat(otherRows);
  const combinedBgs = finalBgs.concat(otherBgs);

  // Write back
  dataRange.setValues(combined);
  dataRange.setBackgrounds(combinedBgs);

  // Highlight priority vendors with a subtle indicator
  if (finalRows.length > 0) {
    listSh.getRange(2, 1, finalRows.length, numCols).setBackground('#e8f0fe');
  }

  Logger.log(`List reordered: ${finalRows.length} priority vendors moved to top, ${otherRows.length} others below`);
}

/************************************************************
 * AUTO-SUMMARIZE TO NOTES
 * Claude summarizes current vendor state and appends to monday.com notes
 ************************************************************/

/**
 * Summarize current vendor activity with Claude and update monday.com notes
 */
function battleStationSummarizeToNotes() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const ui = SpreadsheetApp.getUi();

  const apiKey = getClaudeApiKey_();
  if (!apiKey) {
    ui.alert('No Claude API key configured.\n\nUse menu: ⚡ Battle Station → ⚙️ Set Claude API Key');
    return;
  }

  if (!listSh || !bsSh) {
    ui.alert('Required sheets not found.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();
  if (!currentIndex) {
    ui.alert('No vendor currently loaded.');
    return;
  }

  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  const currentNotes = String(listSh.getRange(listRow, BS_CFG.L_NOTES + 1).getValue() || '');
  const status = String(listSh.getRange(listRow, BS_CFG.L_STATUS + 1).getValue() || '');

  ss.toast(`Analyzing ${vendor}...`, '📝 Summarizing', 5);

  // Get emails for context (from live Gmail search)
  const emails = getEmailsForVendor_(vendor, listRow);
  const unsnoozed = emails.filter(e => !e.isSnoozed);

  // Also get single-person emails from log (may not appear in label searches)
  let singlePersonContext = '';
  try {
    const singleEmails = getSinglePersonEmails_(vendor);
    if (singleEmails.length > 0) {
      singlePersonContext = '\n\nSingle-person emails from log:\n';
      for (const se of singleEmails.slice(0, 5)) {
        singlePersonContext += `Subject: ${se.subject} | From: ${se.from} | Date: ${se.date}\nSnippet: ${(se.snippet || '').substring(0, 200)}\n---\n`;
      }
    }
  } catch (e) {
    Logger.log(`Could not get single-person emails for summary: ${e.message}`);
  }

  let emailContext = '';
  for (const email of unsnoozed.slice(0, 5)) {
    try {
      const thread = GmailApp.getThreadById(email.threadId);
      if (thread) {
        const msgs = thread.getMessages();
        const latestMsg = msgs[msgs.length - 1];
        emailContext += `\nSubject: ${email.subject}\nDate: ${email.date}\nLabels: ${email.labels}\nLatest: ${latestMsg.getPlainBody().substring(0, 500)}\n---`;
      }
    } catch (e) {
      emailContext += `\nSubject: ${email.subject} (${email.date}) [${email.labels}]\n---`;
    }
  }

  const prompt = `You are summarizing vendor activity for a relationship manager's notes at Profitise (lead generation).

Vendor: ${vendor}
Status: ${status}
Current notes: ${currentNotes || '(none)'}

Recent unsnoozed emails (${unsnoozed.length} threads):
${emailContext || '(no emails)'}
${singlePersonContext}

Write a 2-3 sentence summary of what's currently happening with this vendor. Focus on:
- What actions are pending or needed
- Any outstanding issues or requests
- The general state of the relationship

Keep it concise and factual. Do NOT repeat information already in the existing notes.`;

  try {
    const response = callClaudeAPI_(prompt, apiKey, { maxTokens: 500 });

    if (response.error) {
      ui.alert(`Claude API Error: ${response.error}`);
      return;
    }

    const summary = response.content.trim();

    // Append summary to existing notes
    const updatedNotes = currentNotes
      ? `${currentNotes}\n\n${summary}`
      : summary;

    // Update the List sheet
    listSh.getRange(listRow, BS_CFG.L_NOTES + 1).setValue(updatedNotes);

    // Update monday.com
    const result = updateMondayComNotesForVendor_(vendor, updatedNotes, listRow);

    if (result.success) {
      ss.toast(`Notes updated for ${vendor}`, '✅ Summary Added', 3);

      // Update the notes cell on the Battle Station sheet
      for (let i = 5; i < 50; i++) {
        const label = String(bsSh.getRange(i, 1).getValue() || '');
        if (label.indexOf('📝 NOTES') !== -1) {
          bsSh.getRange(i + 1, 1).setValue(updatedNotes);
          break;
        }
      }
    } else {
      ss.toast(`List updated but monday.com sync failed: ${result.error}`, '⚠️ Partial', 5);
    }
  } catch (e) {
    ui.alert(`Error: ${e.message}`);
  }
}

/************************************************************
 * EMAIL RULES ENGINE
 * Automatic actions based on inbound email patterns
 ************************************************************/

/**
 * Get or create the Email Rules sheet for configuring automation
 */
function getEmailRulesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('BS_Email_Rules');

  if (!sh) {
    sh = ss.insertSheet('BS_Email_Rules');
    // Headers
    sh.getRange(1, 1, 1, 8).setValues([[
      'Rule Name', 'Sender Pattern', 'Subject Pattern', 'Label Pattern',
      'Action', 'Action Value', 'Enabled', 'Last Triggered'
    ]]);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#e8f0fe');

    // Example rules (update label names to match your Settings configuration)
    const ruleCfg = getLabelConfig_();
    const acctLabel = ruleCfg.accounting_label || 'accounting';
    const prioLabel = ruleCfg.priority_label || 'priority';
    sh.getRange(2, 1, 4, 8).setValues([
      ['Auto-label invoices', '*@billing*', '*invoice*', '', 'add_label', acctLabel, 'TRUE', ''],
      ['Flag contract requests', '', '*contract*,*agreement*,*MSA*', '', 'add_label', prioLabel, 'TRUE', ''],
      ['Auto-snooze newsletters', '*@newsletter*,*@marketing*', '', '', 'snooze', '3', 'TRUE', ''],
      ['Notify on urgent', '', '*urgent*,*ASAP*,*immediately*', '', 'notify', '', 'TRUE', '']
    ]);

    // Formatting
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 200);
    sh.setColumnWidth(4, 150);
    sh.setColumnWidth(5, 120);
    sh.setColumnWidth(6, 200);
    sh.setColumnWidth(7, 80);
    sh.setColumnWidth(8, 150);

    // Add instructions row at top
    sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, 8).merge()
      .setValue('📧 EMAIL RULES — Patterns use * as wildcard, comma-separate multiple patterns. Actions: add_label, snooze (days), notify (email), flag, auto_reply_draft')
      .setBackground('#fff2cc')
      .setFontStyle('italic')
      .setWrap(true);
    sh.setRowHeight(1, 40);
  }

  return sh;
}

/**
 * Open/create the Email Rules configuration sheet
 */
function battleStationManageEmailRules() {
  const sh = getEmailRulesSheet_();
  SpreadsheetApp.getActive().setActiveSheet(sh);
  SpreadsheetApp.getActive().toast('Edit rules in this sheet, then run "Process Email Rules" to apply them.', '📧 Email Rules', 5);
}

/**
 * Process email rules against recent inbound emails
 */
function battleStationProcessEmailRules() {
  const ss = SpreadsheetApp.getActive();
  const rulesSh = getEmailRulesSheet_();

  const rulesData = rulesSh.getDataRange().getValues();
  const rules = [];

  // Parse rules (skip header rows)
  for (let i = 2; i < rulesData.length; i++) {
    const enabled = String(rulesData[i][6]).toUpperCase();
    if (enabled !== 'TRUE') continue;

    rules.push({
      row: i + 1,
      name: rulesData[i][0],
      senderPatterns: parsePatterns_(rulesData[i][1]),
      subjectPatterns: parsePatterns_(rulesData[i][2]),
      labelPatterns: parsePatterns_(rulesData[i][3]),
      action: String(rulesData[i][4]).toLowerCase().trim(),
      actionValue: String(rulesData[i][5]).trim(),
    });
  }

  if (rules.length === 0) {
    ss.toast('No enabled rules found.', '⚠️ No Rules', 3);
    return;
  }

  ss.toast(`Processing ${rules.length} rules against recent emails...`, '📧 Processing', 10);

  // Get threads from the last 24 hours using configured live email query
  const rulesCfg = getLabelConfig_();
  const rulesLiveQuery = rulesCfg.live_email_query || 'label:inbox';
  const threads = GmailApp.search(`${rulesLiveQuery} newer_than:1d`, 0, 50);
  let matchCount = 0;
  const actions = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    const latest = messages[messages.length - 1];
    const sender = latest.getFrom();
    const subject = thread.getFirstMessageSubject();
    const labels = thread.getLabels().map(l => l.getName()).join(',');

    for (const rule of rules) {
      if (matchesPatterns_(sender, rule.senderPatterns) ||
          matchesPatterns_(subject, rule.subjectPatterns) ||
          matchesPatterns_(labels, rule.labelPatterns)) {

        matchCount++;
        actions.push({ rule: rule, thread: thread, subject: subject });

        try {
          switch (rule.action) {
            case 'add_label':
              const label = GmailApp.getUserLabelByName(rule.actionValue);
              if (label) {
                thread.addLabel(label);
                Logger.log(`Rule "${rule.name}": Added label "${rule.actionValue}" to "${subject}"`);
              }
              break;

            case 'snooze':
              // Move to snooze by archiving (and optionally adding exclude label)
              const days = parseInt(rule.actionValue) || 3;
              const excludeQ = rulesCfg.exclude_query || '';
              // If exclude_query is like "-label:X", extract X and add it
              const excludeMatch = excludeQ.match(/-label:(\S+)/);
              if (excludeMatch) {
                const excludeLabel = GmailApp.getUserLabelByName(excludeMatch[1]);
                if (excludeLabel) thread.addLabel(excludeLabel);
              }
              GmailApp.moveThreadToArchive(thread);
              Logger.log(`Rule "${rule.name}": Snoozed "${subject}" for ${days} days`);
              break;

            case 'notify':
              GmailApp.sendEmail(
                rule.actionValue,
                `[Battle Station Alert] ${rule.name}: ${subject}`,
                `Email rule "${rule.name}" triggered.\n\nSubject: ${subject}\nFrom: ${sender}\n\nView in Gmail to take action.`
              );
              Logger.log(`Rule "${rule.name}": Sent notification for "${subject}"`);
              break;

            case 'flag':
              // Find vendor and flag it
              const vendorName = findVendorForEmail_(sender, subject);
              if (vendorName) {
                setVendorFlag_(vendorName, true);
                Logger.log(`Rule "${rule.name}": Flagged vendor "${vendorName}" for "${subject}"`);
              }
              break;

            case 'auto_reply_draft':
              // Create a draft reply using Claude
              const apiKey = getClaudeApiKey_();
              if (apiKey) {
                const body = latest.getPlainBody().substring(0, 1000);
                const replyPrompt = `Draft a brief, professional reply to this email. Context: You are Andy at Profitise, a lead gen company. Rule: "${rule.name}". Value: "${rule.actionValue}"\n\nFrom: ${sender}\nSubject: ${subject}\nBody: ${body}\n\nWrite ONLY the reply body (no subject, no greeting headers). Keep it under 3 sentences.`;
                const aiResponse = callClaudeAPI_(replyPrompt, apiKey, { maxTokens: 300 });
                if (!aiResponse.error) {
                  thread.createDraftReply(aiResponse.content);
                  Logger.log(`Rule "${rule.name}": Created draft reply for "${subject}"`);
                }
              }
              break;
          }

          // Update last triggered timestamp
          rulesSh.getRange(rule.row, 8).setValue(new Date());

        } catch (e) {
          Logger.log(`Rule "${rule.name}" error: ${e.message}`);
        }

        break; // Only apply first matching rule per thread
      }
    }
  }

  ss.toast(`Processed ${threads.length} emails, ${matchCount} rule matches`, '✅ Rules Applied', 5);

  if (actions.length > 0) {
    const summary = actions.map(a => `• ${a.rule.name}: "${a.subject}"`).join('\n');
    Logger.log(`Email Rules Summary:\n${summary}`);
  }
}

/**
 * Parse comma-separated wildcard patterns
 */
function parsePatterns_(patternStr) {
  if (!patternStr) return [];
  return String(patternStr).split(',').map(p => p.trim().toLowerCase()).filter(p => p.length > 0);
}

/**
 * Check if a string matches any of the wildcard patterns
 */
function matchesPatterns_(text, patterns) {
  if (!patterns || patterns.length === 0) return false;
  const lowerText = String(text || '').toLowerCase();

  for (const pattern of patterns) {
    if (!pattern) continue;

    // Convert wildcard pattern to regex
    const regexStr = pattern
      .replace(/[.+?^${}()|[\]\\]/g, '\\$&') // Escape special regex chars
      .replace(/\*/g, '.*'); // Convert * to .*

    try {
      if (new RegExp(regexStr).test(lowerText)) return true;
    } catch (e) {
      // Fallback: simple includes check
      const cleanPattern = pattern.replace(/\*/g, '');
      if (cleanPattern && lowerText.includes(cleanPattern)) return true;
    }
  }
  return false;
}

/**
 * Try to find which vendor an email belongs to based on sender/subject
 */
function findVendorForEmail_(sender, subject) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  if (!listSh) return null;

  const data = listSh.getDataRange().getValues();
  const searchText = (sender + ' ' + subject).toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const vendorName = String(data[i][BS_CFG.L_VENDOR] || '').toLowerCase();
    if (vendorName && searchText.includes(vendorName)) {
      return data[i][BS_CFG.L_VENDOR];
    }
  }
  return null;
}

/************************************************************
 * CLAUDE-POWERED EMAIL REPLY DRAFTING
 * Smart reply drafting for the current vendor's emails
 ************************************************************/

/**
 * Draft a reply to the most recent unsnoozed email using Claude
 */
function battleStationDraftReply() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);
  const ui = SpreadsheetApp.getUi();

  const apiKey = getClaudeApiKey_();
  if (!apiKey) {
    ui.alert('No Claude API key configured.\n\nUse menu: ⚡ Battle Station → ⚙️ Set Claude API Key');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();
  if (!currentIndex) {
    ui.alert('No vendor currently loaded.');
    return;
  }

  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();
  const notes = String(listSh.getRange(listRow, BS_CFG.L_NOTES + 1).getValue() || '');

  ss.toast(`Finding emails for ${vendor}...`, '✉️ Drafting', 3);

  const emails = getEmailsForVendor_(vendor, listRow);
  const unsnoozed = emails.filter(e => !e.isSnoozed);

  if (unsnoozed.length === 0) {
    ui.alert('No unsnoozed emails found for this vendor.');
    return;
  }

  // Get the most recent unsnoozed email thread
  const targetEmail = unsnoozed[0];
  let threadContent = '';
  let thread = null;

  try {
    thread = GmailApp.getThreadById(targetEmail.threadId);
    if (thread) {
      const messages = thread.getMessages();
      for (const msg of messages.slice(-3)) {
        threadContent += `\nFrom: ${msg.getFrom()}\nDate: ${msg.getDate()}\n${msg.getPlainBody().substring(0, 1000)}\n---`;
      }
    }
  } catch (e) {
    threadContent = `Subject: ${targetEmail.subject}\n(Could not fetch thread content)`;
  }

  ss.toast('Generating reply with Claude...', '🤖 Drafting', 5);

  const prompt = `You are drafting an email reply for Andy Worford at Profitise (a lead generation company in Home Services and Solar).

Vendor: ${vendor}
Notes about this vendor: ${notes || '(none)'}

Email thread:
${threadContent}

Write a professional, concise reply. Be helpful and move the conversation forward. Sign off as "Andy Worford, Profitise".

Write ONLY the email body. No meta-commentary.`;

  try {
    const response = callClaudeAPI_(prompt, apiKey, { maxTokens: 800 });

    if (response.error) {
      ui.alert(`Claude API Error: ${response.error}`);
      return;
    }

    // Create a Gmail draft reply
    if (thread) {
      thread.createDraftReply(response.content);
      ss.toast(`Draft reply created for "${targetEmail.subject}"`, '✅ Draft Ready', 3);
      ui.alert(`✓ Draft reply created!\n\nSubject: Re: ${targetEmail.subject}\n\nCheck your Gmail drafts to review and send.`);
    } else {
      ui.alert(`Reply generated but couldn't create draft.\n\nCopy this reply:\n\n${response.content}`);
    }
  } catch (e) {
    ui.alert(`Error: ${e.message}`);
  }
}

