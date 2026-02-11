/************************************************************
 * EMAIL LOG - Record all emails to a Google Sheet
 *
 * Records every email the system encounters to a designated
 * Google Sheet. This allows:
 * - Single-person emails to be included in analysis/summarizing
 * - Historical tracking of all vendor communications
 * - Full audit trail of email activity
 *
 * The log spreadsheet ID is configured in Settings (email_log_spreadsheet_id).
 * If blank, a new spreadsheet is created and the ID saved.
 *
 * Log Sheet columns:
 *   A: Timestamp (when logged)
 *   B: Thread ID
 *   C: Message ID
 *   D: Subject
 *   E: From
 *   F: To
 *   G: CC
 *   H: Date (email date)
 *   I: Labels
 *   J: Snippet (first 500 chars)
 *   K: Vendor (matched vendor name)
 *   L: Participant Count (unique email addresses)
 *   M: Is Single Person (TRUE if only one external person)
 *   N: Thread Link
 ************************************************************/

const EMAIL_LOG_SHEET_NAME = 'Email Log';
const EMAIL_LOG_HEADERS = [
  'Timestamp', 'Thread ID', 'Message ID', 'Subject', 'From', 'To', 'CC',
  'Date', 'Labels', 'Snippet', 'Vendor', 'Participant Count',
  'Is Single Person', 'Thread Link'
];

/**
 * Get or create the email log spreadsheet.
 * If no ID is configured in Settings, creates a new one.
 *
 * @returns {Spreadsheet|null} The log spreadsheet, or null if disabled
 */
function getEmailLogSpreadsheet_() {
  const cfg = getLabelConfig_();
  let sheetId = cfg.email_log_spreadsheet_id;

  // If blank, create a new spreadsheet
  if (!sheetId) {
    return createEmailLogSpreadsheet_();
  }

  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (e) {
    Logger.log(`Could not open email log spreadsheet (${sheetId}): ${e.message}`);
    // Try to create a new one
    return createEmailLogSpreadsheet_();
  }
}

/**
 * Create a new email log spreadsheet and save its ID to Settings.
 *
 * @returns {Spreadsheet} The newly created spreadsheet
 */
function createEmailLogSpreadsheet_() {
  const ss = SpreadsheetApp.getActive();
  const parentName = ss.getName();

  // Create new spreadsheet
  const logSS = SpreadsheetApp.create(`${parentName} - Email Log`);
  const logSheet = logSS.getActiveSheet();
  logSheet.setName(EMAIL_LOG_SHEET_NAME);

  // Write headers
  logSheet.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length)
    .setValues([EMAIL_LOG_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8f0fe');

  // Set column widths
  logSheet.setColumnWidth(1, 150);  // Timestamp
  logSheet.setColumnWidth(2, 120);  // Thread ID
  logSheet.setColumnWidth(3, 120);  // Message ID
  logSheet.setColumnWidth(4, 300);  // Subject
  logSheet.setColumnWidth(5, 200);  // From
  logSheet.setColumnWidth(6, 200);  // To
  logSheet.setColumnWidth(7, 150);  // CC
  logSheet.setColumnWidth(8, 150);  // Date
  logSheet.setColumnWidth(9, 200);  // Labels
  logSheet.setColumnWidth(10, 300); // Snippet
  logSheet.setColumnWidth(11, 150); // Vendor
  logSheet.setColumnWidth(12, 100); // Participant Count
  logSheet.setColumnWidth(13, 100); // Is Single Person
  logSheet.setColumnWidth(14, 250); // Thread Link

  // Freeze header row
  logSheet.setFrozenRows(1);

  // Save the ID to Settings
  const logId = logSS.getId();
  saveEmailLogSpreadsheetId_(logId);

  Logger.log(`Created email log spreadsheet: ${logSS.getUrl()}`);
  return logSS;
}

/**
 * Save the email log spreadsheet ID to the Settings sheet.
 *
 * @param {string} spreadsheetId - The ID to save
 */
function saveEmailLogSpreadsheetId_(spreadsheetId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Settings');
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const keys = sh.getRange(2, LABEL_CFG_KEY_COL, lastRow - 1, 1).getValues().flat();

  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i]).trim().toLowerCase() === 'email_log_spreadsheet_id') {
      sh.getRange(i + 2, LABEL_CFG_VALUE_COL).setValue(spreadsheetId);
      // Clear cache so next read picks up the new value
      clearLabelConfigCache_();
      return;
    }
  }
}

/**
 * Log a batch of emails for a vendor to the email log sheet.
 * Deduplicates by Thread ID + Message ID to avoid duplicate entries.
 *
 * @param {Array} emails - Array of email objects from getEmailsForVendor_
 * @param {string} vendorName - The vendor name for these emails
 * @param {Array} [threads] - Optional pre-fetched GmailThread objects
 */
function logEmailsToSheet_(emails, vendorName, threads) {
  if (!emails || emails.length === 0) return;

  let logSS;
  try {
    logSS = getEmailLogSpreadsheet_();
    if (!logSS) return;
  } catch (e) {
    Logger.log(`Email logging disabled or failed: ${e.message}`);
    return;
  }

  let logSheet = logSS.getSheetByName(EMAIL_LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = logSS.insertSheet(EMAIL_LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length)
      .setValues([EMAIL_LOG_HEADERS])
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
  }

  // Get existing thread IDs to avoid duplicates
  const lastRow = logSheet.getLastRow();
  const existingThreadIds = new Set();

  if (lastRow > 1) {
    const existing = logSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    for (const id of existing) {
      if (id) existingThreadIds.add(String(id));
    }
  }

  // Get the user's own email for single-person detection
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();

  const newRows = [];
  const now = new Date();

  for (const email of emails) {
    // Skip if already logged
    if (existingThreadIds.has(email.threadId)) continue;

    // Try to get full message details
    let from = '', to = '', cc = '', messageId = '', snippet = '';
    let participantCount = 0;
    let isSinglePerson = false;

    try {
      const thread = threads ?
        threads.find(t => t.getId() === email.threadId) :
        GmailApp.getThreadById(email.threadId);

      if (thread) {
        const messages = thread.getMessages();
        const latestMsg = messages[messages.length - 1];

        from = latestMsg.getFrom();
        to = latestMsg.getTo();
        cc = latestMsg.getCc() || '';
        messageId = latestMsg.getId();

        try {
          snippet = latestMsg.getPlainBody().substring(0, 500);
        } catch (e) {
          snippet = email.snippet || '';
        }

        // Count unique participants (excluding the user)
        const allAddresses = extractEmailAddresses_(
          [from, to, cc].join(', ')
        );
        const externalAddresses = allAddresses.filter(
          addr => addr.toLowerCase() !== myEmail
        );
        participantCount = externalAddresses.length;
        isSinglePerson = participantCount <= 1;
      }
    } catch (e) {
      Logger.log(`Could not fetch thread details for ${email.threadId}: ${e.message}`);
      snippet = email.snippet || '';
    }

    const threadLink = buildGmailThreadUrl_(email.threadId);

    newRows.push([
      now,                    // Timestamp
      email.threadId,         // Thread ID
      messageId,              // Message ID
      email.subject,          // Subject
      from,                   // From
      to,                     // To
      cc,                     // CC
      email.date,             // Date
      email.labels || '',     // Labels
      snippet,                // Snippet
      vendorName,             // Vendor
      participantCount,       // Participant Count
      isSinglePerson,         // Is Single Person
      threadLink              // Thread Link
    ]);

    existingThreadIds.add(email.threadId);
  }

  if (newRows.length > 0) {
    const startRow = logSheet.getLastRow() + 1;
    logSheet.getRange(startRow, 1, newRows.length, EMAIL_LOG_HEADERS.length)
      .setValues(newRows);
    Logger.log(`Logged ${newRows.length} new emails for vendor "${vendorName}"`);
  }
}

/**
 * Extract email addresses from a string (handles "Name <email>" format).
 *
 * @param {string} str - String containing email addresses
 * @returns {Array<string>} Array of extracted email addresses
 */
function extractEmailAddresses_(str) {
  if (!str) return [];

  const emails = [];
  // Match email addresses in "Name <email@domain.com>" or plain "email@domain.com" format
  const regex = /[\w.+-]+@[\w.-]+\.\w+/g;
  let match;

  while ((match = regex.exec(str)) !== null) {
    const email = match[0].toLowerCase();
    if (!emails.includes(email)) {
      emails.push(email);
    }
  }

  return emails;
}

/**
 * Get all single-person emails from the log for a vendor.
 * These are emails where only one external person is involved,
 * meaning they may not show up in label-based filtering but
 * are still relevant for summarization.
 *
 * @param {string} vendorName - The vendor name
 * @returns {Array} Array of email log rows for single-person emails
 */
function getSinglePersonEmails_(vendorName) {
  let logSS;
  try {
    logSS = getEmailLogSpreadsheet_();
    if (!logSS) return [];
  } catch (e) {
    return [];
  }

  const logSheet = logSS.getSheetByName(EMAIL_LOG_SHEET_NAME);
  if (!logSheet) return [];

  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = logSheet.getRange(2, 1, lastRow - 1, EMAIL_LOG_HEADERS.length).getValues();
  const vendorLower = vendorName.toLowerCase();

  return data
    .filter(row => {
      const rowVendor = String(row[10] || '').toLowerCase(); // Column K: Vendor
      const isSingle = row[12] === true || row[12] === 'TRUE'; // Column M: Is Single Person
      return rowVendor === vendorLower && isSingle;
    })
    .map(row => ({
      timestamp: row[0],
      threadId: row[1],
      messageId: row[2],
      subject: row[3],
      from: row[4],
      to: row[5],
      cc: row[6],
      date: row[7],
      labels: row[8],
      snippet: row[9],
      vendor: row[10],
      participantCount: row[11],
      isSinglePerson: row[12],
      threadLink: row[13]
    }));
}

/**
 * Get all logged emails for a vendor (for analysis/summarization).
 * Includes both multi-person and single-person emails.
 *
 * @param {string} vendorName - The vendor name
 * @returns {Array} Array of all logged email objects
 */
function getAllLoggedEmails_(vendorName) {
  let logSS;
  try {
    logSS = getEmailLogSpreadsheet_();
    if (!logSS) return [];
  } catch (e) {
    return [];
  }

  const logSheet = logSS.getSheetByName(EMAIL_LOG_SHEET_NAME);
  if (!logSheet) return [];

  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = logSheet.getRange(2, 1, lastRow - 1, EMAIL_LOG_HEADERS.length).getValues();
  const vendorLower = vendorName.toLowerCase();

  return data
    .filter(row => String(row[10] || '').toLowerCase() === vendorLower)
    .map(row => ({
      timestamp: row[0],
      threadId: row[1],
      messageId: row[2],
      subject: row[3],
      from: row[4],
      to: row[5],
      cc: row[6],
      date: row[7],
      labels: row[8],
      snippet: row[9],
      vendor: row[10],
      participantCount: row[11],
      isSinglePerson: row[12],
      threadLink: row[13]
    }));
}

/**
 * Menu entry: Open/view the email log spreadsheet
 */
function battleStationOpenEmailLog() {
  let logSS;
  try {
    logSS = getEmailLogSpreadsheet_();
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Could not open email log: ${e.message}`);
    return;
  }

  if (!logSS) {
    SpreadsheetApp.getUi().alert('Email logging is not configured.\n\nRun "Setup Label Config" first.');
    return;
  }

  const url = logSS.getUrl();
  const html = HtmlService.createHtmlOutput(
    `<script>window.open("${url}");google.script.host.close();</script>`
  ).setWidth(200).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Email Log...');
}

/**
 * Menu entry: Force re-log all emails for current vendor
 */
function battleStationRelogEmails() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName('List');

  const currentIndex = getCurrentVendorIndex_();
  if (!currentIndex) {
    SpreadsheetApp.getUi().alert('No vendor currently loaded.');
    return;
  }

  const listRow = currentIndex + 1;
  const vendor = String(listSh.getRange(listRow, 1).getValue() || '').trim();

  if (!vendor) {
    SpreadsheetApp.getUi().alert('No vendor found.');
    return;
  }

  ss.toast(`Re-logging emails for ${vendor}...`, 'Email Log', 5);

  const emails = getEmailsForVendor_(vendor, listRow);
  if (emails.length > 0) {
    logEmailsToSheet_(emails, vendor);
    ss.toast(`Logged ${emails.length} emails for ${vendor}`, 'Email Log Complete', 3);
  } else {
    ss.toast(`No emails found for ${vendor}`, 'Email Log', 3);
  }
}

/**
 * Get email log summary stats for a vendor.
 *
 * @param {string} vendorName - The vendor name
 * @returns {Object} { total, singlePerson, multiPerson, latestDate }
 */
function getEmailLogStats_(vendorName) {
  const allEmails = getAllLoggedEmails_(vendorName);

  if (allEmails.length === 0) {
    return { total: 0, singlePerson: 0, multiPerson: 0, latestDate: null };
  }

  const singlePerson = allEmails.filter(e => e.isSinglePerson === true || e.isSinglePerson === 'TRUE');
  const multiPerson = allEmails.filter(e => e.isSinglePerson !== true && e.isSinglePerson !== 'TRUE');

  // Find latest date
  let latestDate = null;
  for (const e of allEmails) {
    const d = new Date(e.date);
    if (!isNaN(d.getTime()) && (!latestDate || d > latestDate)) {
      latestDate = d;
    }
  }

  return {
    total: allEmails.length,
    singlePerson: singlePerson.length,
    multiPerson: multiPerson.length,
    latestDate: latestDate
  };
}
